<#
.SYNOPSIS
    GUI til opslag af brugere i Exchange Online: finder initialer, alias og e-mailadresse ud fra navne.

.DESCRIPTION
    Bruges til at udfylde de tomme felter i "Sorterede_mailbox_rettigheder.xlsx" (fanen "Send on behalf").

    1. Forbind til Exchange Online (moderne login, MFA understøttes).
    2. Indlæs navnene: "Manglende_SendOnBehalf_opslag.csv" (kun rækker uden e-mailadresse), en CSV-eksport
       af fanen, eller indsæt dem direkte fra Excel. Ligger "Manglende_SendOnBehalf_opslag.csv" i samme
       mappe som scriptet, indlæses den automatisk.
    3. Start søgningen. Alle modtagere hentes én gang, og hvert navn matches lokalt i flere trin:
         1) Eksakt Exchange-navn (Name), fx "Birgitte Pedersen_e901a4b14e"
         2) Eksakt visningsnavn (DisplayName)
         3) Navnet efter normalisering (store/små bogstaver, accenter, mellemrum, punktummer og suffikser)
         4) Fornavn + efternavn, hvor mellemnavne/initialer må afvige        -> "Fundet – kontrollér"
         5) Lignende navne (stavevarianter, mellemnavn som fornavn osv.)     -> skal vælges manuelt
       Navne, der ikke findes, slås op blandt slettede postkasser og brugerkonti uden postkasse,
       så scriptet kan forklare, hvorfor de ikke findes.
    4. Gem resultatet som CSV (semikolonsepareret, UTF-8 – åbner direkte i dansk Excel):
         <navn>.csv             fundne brugere med initialer, alias og e-mailadresse
         <navn>_ikke_fundet.csv brugere der ikke blev fundet, med årsag og eventuelle kandidater

.NOTES
    Kræver Windows PowerShell 5.1 eller PowerShell 7 på Windows og modulet ExchangeOnlineManagement v3+
    (scriptet tilbyder at installere det). Kontoen skal kunne læse modtagere i Exchange Online,
    fx med rollen "View-Only Recipients" eller "View-Only Organization Management".
#>
#Requires -Version 5.1
[CmdletBinding()]
param()

if ($PSVersionTable.PSEdition -eq 'Core' -and -not $IsWindows) {
    Write-Error 'Scriptet bruger WPF og kræver Windows.'
    return
}

# WPF kræver en STA-tråd – start scriptet igen med -STA, hvis det ikke allerede kører sådan
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

#region Hjælpeklasser (C#) – navnenormalisering, stavevarianter og rækker til tabellen
if (-not ('SobLookup.NameTools' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace SobLookup
{
    public class Candidate
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string Name { get; set; }
        public string Alias { get; set; }
        public string Email { get; set; }
        public string Type { get; set; }
        public string Info { get; set; }
        public string Norm { get; set; }
        public string[] Tokens { get; set; }

        public string Initialer { get { return NameTools.LocalPart(Email); } }

        public static Candidate Create(string displayName, string name, string alias, string email, string type, string id, string info)
        {
            Candidate c = new Candidate();
            c.DisplayName = (displayName ?? string.Empty).Trim();
            c.Name = (name ?? string.Empty).Trim();
            c.Alias = alias ?? string.Empty;
            c.Email = email ?? string.Empty;
            c.Type = type ?? string.Empty;
            c.Info = info ?? string.Empty;
            c.Id = string.IsNullOrEmpty(id) ? c.Name + "|" + c.Email : id;
            c.Norm = NameTools.Normalize(c.DisplayName.Length > 0 ? c.DisplayName : NameTools.RemoveSuffix(c.Name));
            c.Tokens = NameTools.Tokens(c.Norm);
            return c;
        }

        public override string ToString()
        {
            string s = DisplayName;
            if (!string.IsNullOrEmpty(Email)) { s += " <" + Email + ">"; }
            if (!string.IsNullOrEmpty(Type)) { s += " [" + Type + "]"; }
            if (!string.IsNullOrEmpty(Info)) { s += " (" + Info + ")"; }
            return s;
        }
    }

    public class ResultRow
    {
        public ResultRow() { Candidates = new List<Candidate>(); }

        public string Row { get; set; }
        public string InputName { get; set; }
        public string Status { get; set; }
        public string OriginalStatus { get; set; }
        public string Category { get; set; }
        public string Initialer { get; set; }
        public string Alias { get; set; }
        public string Email { get; set; }
        public string FoundName { get; set; }
        public string Type { get; set; }
        public string Method { get; set; }
        public string Reason { get; set; }
        public List<Candidate> Candidates { get; set; }

        public int RowNumber
        {
            get { int n; return int.TryParse(Row, out n) ? n : int.MaxValue; }
        }

        public string Note
        {
            get { return string.IsNullOrEmpty(Reason) ? Method : Reason; }
        }

        public string CandidatesText
        {
            get
            {
                List<string> parts = new List<string>();
                foreach (Candidate c in Candidates) { parts.Add(c.ToString()); }
                return string.Join("; ", parts.ToArray());
            }
        }
    }

    // Opslagsindeks over modtagere: eksakt Name/DisplayName, normaliseret navn, fornavn og efternavn
    public class NameIndex
    {
        private readonly Dictionary<string, List<Candidate>> byName = new Dictionary<string, List<Candidate>>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, List<Candidate>> byDisplay = new Dictionary<string, List<Candidate>>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, List<Candidate>> byNorm = new Dictionary<string, List<Candidate>>(StringComparer.Ordinal);
        private readonly Dictionary<string, List<Candidate>> byFirst = new Dictionary<string, List<Candidate>>(StringComparer.Ordinal);
        private readonly Dictionary<string, List<Candidate>> byLast = new Dictionary<string, List<Candidate>>(StringComparer.Ordinal);
        private readonly List<Candidate> all;
        private readonly string[] normKeys;

        public NameIndex(IEnumerable<Candidate> candidates)
        {
            all = new List<Candidate>(candidates);
            normKeys = new string[all.Count];
            for (int i = 0; i < all.Count; i++)
            {
                Candidate c = all[i];
                normKeys[i] = c.Norm;
                Add(byName, c.Name, c);
                Add(byDisplay, c.DisplayName, c);
                Add(byNorm, c.Norm, c);
                Add(byNorm, NameTools.Normalize(NameTools.RemoveSuffix(c.Name)), c);
                if (c.Tokens.Length >= 1) { Add(byFirst, c.Tokens[0], c); }
                if (c.Tokens.Length >= 2) { Add(byLast, c.Tokens[c.Tokens.Length - 1], c); }
            }
        }

        public int Count { get { return all.Count; } }

        public List<Candidate> FindByName(string name) { return Get(byName, name); }
        public List<Candidate> FindByDisplayName(string displayName) { return Get(byDisplay, displayName); }
        public List<Candidate> FindByNormalized(string normalized) { return Get(byNorm, normalized); }
        public List<Candidate> WithFirstName(string token) { return Get(byFirst, token); }
        public List<Candidate> WithLastName(string token) { return Get(byLast, token); }

        // Normaliserede navne højst maxDistance tegn fra det søgte – nærmeste først
        public List<Candidate> FindSimilar(string normalized, int maxDistance)
        {
            List<Candidate> result = new List<Candidate>();
            if (string.IsNullOrEmpty(normalized)) { return result; }
            foreach (int i in NameTools.FindClose(normalized, normKeys, maxDistance))
            {
                if (!result.Contains(all[i])) { result.Add(all[i]); }
            }
            return result;
        }

        private static void Add(Dictionary<string, List<Candidate>> map, string key, Candidate c)
        {
            if (string.IsNullOrWhiteSpace(key)) { return; }
            key = key.Trim();
            List<Candidate> list;
            if (!map.TryGetValue(key, out list))
            {
                list = new List<Candidate>();
                map[key] = list;
            }
            if (!list.Contains(c)) { list.Add(c); }
        }

        private static List<Candidate> Get(Dictionary<string, List<Candidate>> map, string key)
        {
            List<Candidate> list;
            if (string.IsNullOrWhiteSpace(key) || !map.TryGetValue(key.Trim(), out list)) { return new List<Candidate>(); }
            return new List<Candidate>(list);
        }
    }

    public static class NameTools
    {
        private static readonly Regex Suffix = new Regex("_[0-9a-fA-F]{10}$");

        // Exchange tilføjer "_" + 10 hex-tegn til Name, når navnet allerede er i brug (fx "Birgitte Pedersen_e901a4b14e")
        public static string RemoveSuffix(string name)
        {
            return string.IsNullOrEmpty(name) ? string.Empty : Suffix.Replace(name, string.Empty);
        }

        // Små bogstaver uden accenter (ä -> a, é -> e). Alt andet end bogstaver og tal bliver til ét mellemrum.
        public static string Normalize(string s)
        {
            if (string.IsNullOrEmpty(s)) { return string.Empty; }
            string d = s.Normalize(NormalizationForm.FormD);
            StringBuilder sb = new StringBuilder(d.Length);
            bool pendingSpace = false;
            foreach (char c in d)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) == UnicodeCategory.NonSpacingMark) { continue; }
                if (char.IsLetterOrDigit(c))
                {
                    if (pendingSpace && sb.Length > 0) { sb.Append(' '); }
                    pendingSpace = false;
                    sb.Append(char.ToLowerInvariant(c));
                }
                else
                {
                    pendingSpace = true;
                }
            }
            return sb.ToString();
        }

        public static string[] Tokens(string normalized)
        {
            if (string.IsNullOrEmpty(normalized)) { return new string[0]; }
            return normalized.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        }

        public static string LocalPart(string address)
        {
            if (string.IsNullOrEmpty(address)) { return string.Empty; }
            int at = address.IndexOf('@');
            return at > 0 ? address.Substring(0, at) : address;
        }

        // Levenshtein-afstand: antal tegn der skal indsættes, slettes eller udskiftes
        public static int Distance(string a, string b)
        {
            if (a == null) { a = string.Empty; }
            if (b == null) { b = string.Empty; }
            int[] prev = new int[b.Length + 1];
            int[] cur = new int[b.Length + 1];
            for (int j = 0; j <= b.Length; j++) { prev[j] = j; }
            for (int i = 1; i <= a.Length; i++)
            {
                cur[0] = i;
                for (int j = 1; j <= b.Length; j++)
                {
                    int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                    cur[j] = Math.Min(Math.Min(cur[j - 1] + 1, prev[j] + 1), prev[j - 1] + cost);
                }
                int[] t = prev; prev = cur; cur = t;
            }
            return prev[b.Length];
        }

        // Indeks på de nøgler, der ligger højst maxDistance fra query – nærmeste først
        public static int[] FindClose(string query, string[] keys, int maxDistance)
        {
            List<int> hits = new List<int>();
            List<int> distances = new List<int>();
            for (int i = 0; i < keys.Length; i++)
            {
                string k = keys[i];
                if (string.IsNullOrEmpty(k) || Math.Abs(k.Length - query.Length) > maxDistance) { continue; }
                int d = Distance(query, k);
                if (d <= maxDistance) { hits.Add(i); distances.Add(d); }
            }
            int[] result = hits.ToArray();
            Array.Sort(distances.ToArray(), result);
            return result;
        }
    }
}
'@
}

# Mørk titellinje via Windows' DWM-API – ignoreres på Windows-versioner uden mørk titellinje
if (-not ('SobLookup.Native' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace SobLookup
{
    public static class Native
    {
        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

        public static void UseDarkTitleBar(IntPtr hwnd)
        {
            int on = 1;
            // 20 = DWMWA_USE_IMMERSIVE_DARK_MODE (Windows 10 20H1+ og 11), 19 = samme på ældre Windows 10
            if (DwmSetWindowAttribute(hwnd, 20, ref on, sizeof(int)) != 0)
            {
                DwmSetWindowAttribute(hwnd, 19, ref on, sizeof(int));
            }
        }
    }
}
'@
}
#endregion

#region Brugerflade (XAML)
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Send on behalf – brugeropslag i Exchange Online"
        Width="1320" Height="830" MinWidth="1040" MinHeight="640"
        WindowStartupLocation="CenterScreen" UseLayoutRounding="True"
        Background="#0B1220" FontFamily="Segoe UI" FontSize="13" Foreground="#E6EDF7">
  <Window.Resources>
    <!-- Mørkt farvetema -->
    <SolidColorBrush x:Key="Accent" Color="#3B8EEA"/>
    <SolidColorBrush x:Key="Text" Color="#E6EDF7"/>
    <SolidColorBrush x:Key="Muted" Color="#94A3B8"/>
    <SolidColorBrush x:Key="Line" Color="#22314A"/>
    <SolidColorBrush x:Key="BarBg" Color="#0F1A2E"/>
    <SolidColorBrush x:Key="CardBg" Color="#111A2B"/>
    <SolidColorBrush x:Key="InputBg" Color="#0B1424"/>
    <SolidColorBrush x:Key="InputBorder" Color="#2A3A55"/>
    <SolidColorBrush x:Key="ButtonBg" Color="#1A2740"/>
    <SolidColorBrush x:Key="HeaderBg" Color="#162238"/>
    <SolidColorBrush x:Key="HeaderHover" Color="#1D2C47"/>
    <SolidColorBrush x:Key="Selection" Color="#1E4A7A"/>
    <SolidColorBrush x:Key="ScrollThumbBg" Color="#33435E"/>
    <!-- Systemfarver bruges af fx hjørnet mellem rullepaneler og markering af tekst -->
    <SolidColorBrush x:Key="{x:Static SystemColors.ControlBrushKey}" Color="#111A2B"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.HighlightBrushKey}" Color="#1E4A7A"/>

    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource CardBg}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Line}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="10"/>
      <Setter Property="Padding" Value="16"/>
    </Style>
    <Style x:Key="StepBadge" TargetType="Border">
      <Setter Property="Width" Value="24"/>
      <Setter Property="Height" Value="24"/>
      <Setter Property="CornerRadius" Value="12"/>
      <Setter Property="Background" Value="{StaticResource Accent}"/>
      <Setter Property="Margin" Value="0,0,10,0"/>
    </Style>
    <Style x:Key="StepNumber" TargetType="TextBlock">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="HorizontalAlignment" Value="Center"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="CardTitle" TargetType="TextBlock">
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="Hint" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Muted}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="Chip" TargetType="Border">
      <Setter Property="CornerRadius" Value="12"/>
      <Setter Property="Padding" Value="10,3"/>
      <Setter Property="Margin" Value="0,2,8,2"/>
    </Style>

    <Style TargetType="Button">
      <Setter Property="Background" Value="{StaticResource ButtonBg}"/>
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="BorderBrush" Value="{StaticResource InputBorder}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="14,7"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Grid x:Name="root">
              <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6"/>
              <Border x:Name="shade" Background="#FFFFFF" CornerRadius="6" Opacity="0"/>
              <ContentPresenter Margin="{TemplateBinding Padding}" HorizontalAlignment="Center"
                                VerticalAlignment="Center" RecognizesAccessKey="True"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="shade" Property="Opacity" Value="0.08"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="shade" Property="Opacity" Value="0.16"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="root" Property="Opacity" Value="0.4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Primary" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="{StaticResource Accent}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Accent}"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Padding" Value="6,5"/>
      <Setter Property="Background" Value="{StaticResource InputBg}"/>
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="BorderBrush" Value="{StaticResource InputBorder}"/>
      <Setter Property="CaretBrush" Value="{StaticResource Text}"/>
      <Setter Property="SelectionBrush" Value="{StaticResource Accent}"/>
    </Style>

    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#CBD5E1"/>
      <Setter Property="Margin" Value="0,8,0,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <Grid Background="Transparent">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Border x:Name="box" Width="16" Height="16" CornerRadius="4" Margin="0,1,8,0" VerticalAlignment="Top"
                      Background="{StaticResource InputBg}" BorderBrush="{StaticResource InputBorder}" BorderThickness="1">
                <Path x:Name="mark" Data="M3,7 L6,10 L11,4" Stroke="White" StrokeThickness="2" Visibility="Collapsed"/>
              </Border>
              <ContentPresenter Grid.Column="1" VerticalAlignment="Center" RecognizesAccessKey="True"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="box" Property="Background" Value="{StaticResource Accent}"/>
                <Setter TargetName="box" Property="BorderBrush" Value="{StaticResource Accent}"/>
                <Setter TargetName="mark" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="box" Property="BorderBrush" Value="{StaticResource Accent}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="ToolTip">
      <Setter Property="Background" Value="#1A2740"/>
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="BorderBrush" Value="{StaticResource InputBorder}"/>
      <Setter Property="Padding" Value="8,5"/>
    </Style>

    <!-- Rullepaneler -->
    <Style x:Key="ScrollThumb" TargetType="Thumb">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Thumb">
            <Border x:Name="thumb" Background="{StaticResource ScrollThumbBg}" CornerRadius="4"/>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="thumb" Property="Background" Value="#46587A"/>
              </Trigger>
              <Trigger Property="IsDragging" Value="True">
                <Setter TargetName="thumb" Property="Background" Value="#5A6E93"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="ScrollPage" TargetType="RepeatButton">
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="IsTabStop" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RepeatButton">
            <Border Background="Transparent"/>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ScrollBar">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Width" Value="10"/>
      <Setter Property="MinWidth" Value="10"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Border Background="{TemplateBinding Background}">
              <Track x:Name="PART_Track" IsDirectionReversed="True">
                <Track.DecreaseRepeatButton>
                  <RepeatButton Style="{StaticResource ScrollPage}" Command="ScrollBar.PageUpCommand"/>
                </Track.DecreaseRepeatButton>
                <Track.IncreaseRepeatButton>
                  <RepeatButton Style="{StaticResource ScrollPage}" Command="ScrollBar.PageDownCommand"/>
                </Track.IncreaseRepeatButton>
                <Track.Thumb>
                  <Thumb Style="{StaticResource ScrollThumb}" Margin="2"/>
                </Track.Thumb>
              </Track>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="Orientation" Value="Horizontal">
          <Setter Property="Width" Value="Auto"/>
          <Setter Property="MinWidth" Value="0"/>
          <Setter Property="Height" Value="10"/>
          <Setter Property="MinHeight" Value="10"/>
          <Setter Property="Template">
            <Setter.Value>
              <ControlTemplate TargetType="ScrollBar">
                <Border Background="{TemplateBinding Background}">
                  <Track x:Name="PART_Track" IsDirectionReversed="False">
                    <Track.DecreaseRepeatButton>
                      <RepeatButton Style="{StaticResource ScrollPage}" Command="ScrollBar.PageLeftCommand"/>
                    </Track.DecreaseRepeatButton>
                    <Track.IncreaseRepeatButton>
                      <RepeatButton Style="{StaticResource ScrollPage}" Command="ScrollBar.PageRightCommand"/>
                    </Track.IncreaseRepeatButton>
                    <Track.Thumb>
                      <Thumb Style="{StaticResource ScrollThumb}" Margin="2"/>
                    </Track.Thumb>
                  </Track>
                </Border>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- Kolonneoverskrifter i tabellen og kandidatlisten -->
    <Style x:Key="HeaderGripper" TargetType="Thumb">
      <Setter Property="Width" Value="8"/>
      <Setter Property="Cursor" Value="SizeWE"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Thumb">
            <Border Background="Transparent"/>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Foreground" Value="#CBD5E1"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridColumnHeader">
            <Grid>
              <Border x:Name="headerBg" Background="{StaticResource HeaderBg}" BorderBrush="{StaticResource Line}"
                      BorderThickness="0,0,1,1" Padding="8,0">
                <DockPanel>
                  <Path x:Name="sortArrow" DockPanel.Dock="Right" Width="8" Height="5" Margin="6,0,0,0" Stretch="Fill"
                        Fill="{StaticResource Muted}" VerticalAlignment="Center" Visibility="Collapsed" Data="M0,0 L1,0 L0.5,1 Z"/>
                  <ContentPresenter VerticalAlignment="Center"/>
                </DockPanel>
              </Border>
              <Thumb x:Name="PART_LeftHeaderGripper" HorizontalAlignment="Left" Style="{StaticResource HeaderGripper}"/>
              <Thumb x:Name="PART_RightHeaderGripper" HorizontalAlignment="Right" Style="{StaticResource HeaderGripper}"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="headerBg" Property="Background" Value="{StaticResource HeaderHover}"/>
              </Trigger>
              <Trigger Property="SortDirection" Value="Ascending">
                <Setter TargetName="sortArrow" Property="Visibility" Value="Visible"/>
                <Setter TargetName="sortArrow" Property="Data" Value="M0,1 L1,1 L0.5,0 Z"/>
              </Trigger>
              <Trigger Property="SortDirection" Value="Descending">
                <Setter TargetName="sortArrow" Property="Visibility" Value="Visible"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="GridViewColumnHeader">
      <Setter Property="Foreground" Value="#CBD5E1"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="GridViewColumnHeader">
            <Grid>
              <Border Background="{StaticResource HeaderBg}" BorderBrush="{StaticResource Line}" BorderThickness="0,0,1,1" Padding="8,6">
                <ContentPresenter HorizontalAlignment="Left" VerticalAlignment="Center"/>
              </Border>
              <Thumb x:Name="PART_HeaderGripper" HorizontalAlignment="Right" Margin="0,0,-4,0" Style="{StaticResource HeaderGripper}"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ListViewItem">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListViewItem">
            <Border x:Name="itemBg" Background="Transparent" Padding="0,4">
              <GridViewRowPresenter VerticalAlignment="Center" Columns="{TemplateBinding GridView.ColumnCollection}"
                                    Content="{TemplateBinding Content}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="itemBg" Property="Background" Value="#16233A"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="itemBg" Property="Background" Value="{StaticResource Selection}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Topbjælke -->
    <Border Grid.Row="0" Background="{StaticResource BarBg}" BorderBrush="{StaticResource Line}" BorderThickness="0,0,0,1" Padding="22,14">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel>
          <TextBlock Text="Send on behalf – brugeropslag" FontSize="20" FontWeight="SemiBold" Foreground="White"/>
          <TextBlock Text="Finder initialer, alias og e-mailadresse i Exchange Online for navne uden match"
                     Foreground="{StaticResource Muted}" Margin="0,2,0,0"/>
        </StackPanel>
        <Border Grid.Column="1" VerticalAlignment="Center" Background="{StaticResource ButtonBg}" BorderBrush="{StaticResource InputBorder}"
                BorderThickness="1" CornerRadius="14" Padding="12,6">
          <StackPanel Orientation="Horizontal">
            <Ellipse x:Name="dotConn" Width="10" Height="10" Fill="#94A3B8" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBlock x:Name="lblConn" Text="Ikke forbundet" Foreground="{StaticResource Text}" VerticalAlignment="Center"/>
          </StackPanel>
        </Border>
      </Grid>
    </Border>

    <Grid Grid.Row="1" Margin="16">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="340"/>
        <ColumnDefinition Width="16"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <Grid Grid.Column="0">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="12"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- 1. Forbindelse -->
        <Border Grid.Row="0" Style="{StaticResource Card}">
          <StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
              <Border Style="{StaticResource StepBadge}"><TextBlock Text="1" Style="{StaticResource StepNumber}"/></Border>
              <TextBlock Text="Forbind til Exchange Online" Style="{StaticResource CardTitle}"/>
            </StackPanel>
            <TextBlock Text="Administratorkonto (valgfri – udfyldes i login-vinduet)" Style="{StaticResource Hint}" Margin="0,0,0,4"/>
            <TextBox x:Name="txtUpn"/>
            <CheckBox x:Name="chkBrowser">
              <TextBlock Text="Log ind via browseren (hvis login-vinduet ikke dukker op)" TextWrapping="Wrap"/>
            </CheckBox>
            <StackPanel Orientation="Horizontal" Margin="0,14,0,0">
              <Button x:Name="btnConnect" Content="Forbind" Style="{StaticResource Primary}"/>
              <Button x:Name="btnDisconnect" Content="Afbryd forbindelse" IsEnabled="False"/>
            </StackPanel>
            <TextBlock x:Name="lblRecipients" Text="Modtagerlisten hentes ved første søgning." Style="{StaticResource Hint}" Margin="0,12,0,0"/>
            <Button x:Name="btnReload" Content="Hent modtagerlisten igen" HorizontalAlignment="Left"
                    Margin="0,6,0,0" Padding="10,4" FontSize="12" IsEnabled="False"/>
          </StackPanel>
        </Border>

        <!-- 2. Navne -->
        <Border Grid.Row="2" Style="{StaticResource Card}">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
              <Border Style="{StaticResource StepBadge}"><TextBlock Text="2" Style="{StaticResource StepNumber}"/></Border>
              <TextBlock Text="Navne der skal slås op" Style="{StaticResource CardTitle}"/>
            </StackPanel>
            <TextBlock Grid.Row="1" Style="{StaticResource Hint}"
                       Text="Ét navn pr. linje – 'Række;Navn' eller bare 'Navn'. Indsæt direkte fra Excel, eller indlæs Manglende_SendOnBehalf_opslag.csv eller en CSV-eksport af fanen."/>
            <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,10,0,0">
              <Button x:Name="btnLoadCsv" Content="Indlæs CSV …"/>
              <Button x:Name="btnClear" Content="Ryd"/>
            </StackPanel>
            <TextBox x:Name="txtNames" Grid.Row="3" Margin="0,10,0,0" AcceptsReturn="True" AcceptsTab="True"
                     TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                     FontFamily="Consolas" FontSize="12"/>
            <TextBlock x:Name="lblNameCount" Grid.Row="4" Text="0 navne" Style="{StaticResource Hint}" Margin="0,4,0,0"/>
            <StackPanel Grid.Row="5">
              <CheckBox x:Name="chkOnlyMissing" IsChecked="True">
                <TextBlock Text="Indlæs kun rækker uden e-mailadresse fra CSV" TextWrapping="Wrap"/>
              </CheckBox>
              <CheckBox x:Name="chkPreferMailbox" IsChecked="True">
                <TextBlock Text="Ved flere match: vælg automatisk brugerpostkassen, hvis kun én af dem er en brugerpostkasse (markeres 'kontrollér')" TextWrapping="Wrap"/>
              </CheckBox>
            </StackPanel>
            <Grid Grid.Row="6" Margin="0,14,0,0">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Button x:Name="btnSearch" Content="Start søgning" Style="{StaticResource Primary}" Padding="14,9"/>
              <Button x:Name="btnStop" Grid.Column="1" Content="Stop" Margin="0" Padding="18,9" IsEnabled="False"/>
            </Grid>
          </Grid>
        </Border>
      </Grid>

      <!-- 3. Resultat -->
      <Border Grid.Column="2" Style="{StaticResource Card}">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="12"/>
            <RowDefinition Height="210"/>
          </Grid.RowDefinitions>

          <Grid Grid.Row="0" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
              <Border Style="{StaticResource StepBadge}"><TextBlock Text="3" Style="{StaticResource StepNumber}"/></Border>
              <TextBlock Text="Resultat" Style="{StaticResource CardTitle}"/>
            </StackPanel>
            <WrapPanel Grid.Column="1" Margin="18,0,8,0" VerticalAlignment="Center">
              <Border Style="{StaticResource Chip}" Background="#12351F"><TextBlock x:Name="chipFound" Text="Fundet: 0" Foreground="#86EFAC"/></Border>
              <Border Style="{StaticResource Chip}" Background="#3A2E0A"><TextBlock x:Name="chipCheck" Text="Kontrollér: 0" Foreground="#FCD34D"/></Border>
              <Border Style="{StaticResource Chip}" Background="#3D2410"><TextBlock x:Name="chipChoose" Text="Kræver valg: 0" Foreground="#FDBA74"/></Border>
              <Border Style="{StaticResource Chip}" Background="#3F1519"><TextBlock x:Name="chipMissing" Text="Ikke fundet: 0" Foreground="#FCA5A5"/></Border>
            </WrapPanel>
            <Button x:Name="btnExport" Grid.Column="2" Content="Gem CSV …" Style="{StaticResource Primary}" Margin="0" IsEnabled="False"/>
          </Grid>

          <DataGrid x:Name="gridResults" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True"
                    CanUserAddRows="False" CanUserDeleteRows="False" SelectionMode="Extended" SelectionUnit="FullRow"
                    HeadersVisibility="Column" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="{StaticResource Line}"
                    BorderBrush="{StaticResource Line}" BorderThickness="1" Background="{StaticResource InputBg}"
                    Foreground="{StaticResource Text}" RowHeight="28" ColumnHeaderHeight="32">
            <DataGrid.RowStyle>
              <Style TargetType="DataGridRow">
                <Setter Property="Background" Value="#0E1626"/>
                <Setter Property="Foreground" Value="{StaticResource Text}"/>
                <Style.Triggers>
                  <DataTrigger Binding="{Binding Category}" Value="Found"><Setter Property="Background" Value="#10291B"/></DataTrigger>
                  <DataTrigger Binding="{Binding Category}" Value="Manual"><Setter Property="Background" Value="#0F2740"/></DataTrigger>
                  <DataTrigger Binding="{Binding Category}" Value="Check"><Setter Property="Background" Value="#33290A"/></DataTrigger>
                  <DataTrigger Binding="{Binding Category}" Value="Choose"><Setter Property="Background" Value="#3A2410"/></DataTrigger>
                  <DataTrigger Binding="{Binding Category}" Value="NotFound"><Setter Property="Background" Value="#3A1518"/></DataTrigger>
                  <DataTrigger Binding="{Binding Category}" Value="Error"><Setter Property="Background" Value="#2E1A3D"/></DataTrigger>
                </Style.Triggers>
              </Style>
            </DataGrid.RowStyle>
            <DataGrid.CellStyle>
              <Style TargetType="DataGridCell">
                <Setter Property="BorderThickness" Value="0"/>
                <Setter Property="Template">
                  <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                      <Border Background="{TemplateBinding Background}" Padding="6,0">
                        <ContentPresenter VerticalAlignment="Center"/>
                      </Border>
                    </ControlTemplate>
                  </Setter.Value>
                </Setter>
                <Style.Triggers>
                  <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{StaticResource Selection}"/>
                    <Setter Property="Foreground" Value="White"/>
                  </Trigger>
                </Style.Triggers>
              </Style>
            </DataGrid.CellStyle>
            <DataGrid.Columns>
              <DataGridTextColumn Header="Række" Binding="{Binding Row}" SortMemberPath="RowNumber" Width="58"/>
              <DataGridTextColumn Header="Navn fra Excel" Binding="{Binding InputName}" Width="1.6*" MinWidth="140"/>
              <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="1.3*" MinWidth="120">
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock"><Setter Property="FontWeight" Value="SemiBold"/></Style>
                </DataGridTextColumn.ElementStyle>
              </DataGridTextColumn>
              <DataGridTextColumn Header="Initialer" Binding="{Binding Initialer}" Width="74"/>
              <DataGridTextColumn Header="Alias" Binding="{Binding Alias}" Width="74"/>
              <DataGridTextColumn Header="E-mailadresse" Binding="{Binding Email}" Width="1.8*" MinWidth="160"/>
              <DataGridTextColumn Header="Fundet som" Binding="{Binding FoundName}" Width="1.5*" MinWidth="130"/>
              <DataGridTextColumn Header="Bemærkning" Binding="{Binding Note}" Width="2.4*" MinWidth="180">
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock">
                    <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
                    <Setter Property="ToolTip" Value="{Binding Note}"/>
                  </Style>
                </DataGridTextColumn.ElementStyle>
              </DataGridTextColumn>
            </DataGrid.Columns>
          </DataGrid>

          <Grid Grid.Row="3">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="16"/>
              <ColumnDefinition Width="1.35*"/>
            </Grid.ColumnDefinitions>
            <DockPanel Grid.Column="0">
              <TextBlock DockPanel.Dock="Top" Text="Detaljer" FontWeight="SemiBold" Margin="0,0,0,4"/>
              <TextBox x:Name="txtDetail" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                       BorderBrush="{StaticResource Line}"/>
            </DockPanel>
            <DockPanel Grid.Column="2">
              <TextBlock DockPanel.Dock="Top" Text="Kandidater – marker den rigtige og klik 'Brug valgte kandidat' (eller dobbeltklik)"
                         FontWeight="SemiBold" Margin="0,0,0,4" TextTrimming="CharacterEllipsis"/>
              <Button x:Name="btnUseCandidate" DockPanel.Dock="Bottom" Content="Brug valgte kandidat"
                      HorizontalAlignment="Right" Margin="0,8,0,0" IsEnabled="False"/>
              <ListView x:Name="lstCandidates" SelectionMode="Single" Background="{StaticResource InputBg}"
                        Foreground="{StaticResource Text}" BorderBrush="{StaticResource Line}">
                <ListView.View>
                  <GridView>
                    <GridViewColumn Header="Visningsnavn" DisplayMemberBinding="{Binding DisplayName}" Width="170"/>
                    <GridViewColumn Header="Initialer" DisplayMemberBinding="{Binding Initialer}" Width="70"/>
                    <GridViewColumn Header="E-mail / UPN" DisplayMemberBinding="{Binding Email}" Width="200"/>
                    <GridViewColumn Header="Type" DisplayMemberBinding="{Binding Type}" Width="110"/>
                    <GridViewColumn Header="Info" DisplayMemberBinding="{Binding Info}" Width="200"/>
                  </GridView>
                </ListView.View>
              </ListView>
            </DockPanel>
          </Grid>
        </Grid>
      </Border>
    </Grid>

    <!-- Statuslinje -->
    <Border Grid.Row="2" Background="{StaticResource BarBg}" BorderBrush="{StaticResource Line}" BorderThickness="0,1,0,0" Padding="16,8">
      <DockPanel>
        <ProgressBar x:Name="progress" DockPanel.Dock="Right" Width="240" Height="6" Minimum="0" Maximum="100"
                     Foreground="{StaticResource Accent}" Background="#1F2A3D" BorderThickness="0" VerticalAlignment="Center"/>
        <TextBlock x:Name="lblStatus" Text="Klar." Foreground="{StaticResource Muted}" VerticalAlignment="Center"
                   TextTrimming="CharacterEllipsis" Margin="0,0,16,0"/>
      </DockPanel>
    </Border>
  </Grid>
</Window>
'@
#endregion

#region Generelle hjælpefunktioner
function New-Brush([string]$Color) {
    (New-Object System.Windows.Media.BrushConverter).ConvertFromString($Color)
}

# Lader vinduet tegne sig selv og reagere på klik (fx "Stop") midt i en længere handling
function Update-UI {
    $window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background)
}

function Set-Status {
    param([string]$Text, [double]$Percent = -1, [switch]$Indeterminate)
    $ui.lblStatus.Text = $Text
    $ui.progress.IsIndeterminate = $Indeterminate.IsPresent
    if ($Percent -ge 0) { $ui.progress.Value = $Percent }
    Update-UI
}

function Show-Message {
    param([string]$Text, [string]$Title = 'Send on behalf-opslag', [string]$Icon = 'Information', [string]$Buttons = 'OK')
    [System.Windows.MessageBox]::Show($window, $Text, $Title, [System.Windows.MessageBoxButton]$Buttons, [System.Windows.MessageBoxImage]$Icon)
}

# Exchange-modulet bruger asynkrone kald internt. WPF's SynchronizationContext fjernes under kaldene,
# så de ikke kan forsøge at vende tilbage til den (blokerede) GUI-tråd og låse vinduet.
function Invoke-WithoutSyncContext([scriptblock]$Action) {
    $previousContext = [System.Threading.SynchronizationContext]::Current
    [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)
    try { & $Action }
    finally { [System.Threading.SynchronizationContext]::SetSynchronizationContext($previousContext) }
}

function Read-TextFile([string]$Path) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    try {
        return (New-Object System.Text.UTF8Encoding($false, $true)).GetString($bytes)
    } catch {
        # Excels "CSV (semikolonsepareret)" gemmes i Windows-1252
        try { return [System.Text.Encoding]::GetEncoding(1252).GetString($bytes) }
        catch { return [System.Text.Encoding]::GetEncoding(28591).GetString($bytes) }
    }
}

function Write-CsvFile([object[]]$Rows, [string[]]$Columns, [string]$Path) {
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    if ($Rows.Count -gt 0) {
        $Rows | Select-Object $Columns | Export-Csv -LiteralPath $Path -Delimiter ';' -NoTypeInformation -Encoding $encoding
    } else {
        ('"{0}"' -f ($Columns -join '";"')) | Set-Content -LiteralPath $Path -Encoding $encoding
    }
}
#endregion

#region Navne ind
# Tekstfeltet: én linje pr. navn, enten "Række;Navn", "Række<tab>Navn" (indsat fra Excel) eller bare "Navn"
function Read-NameEntries([string]$Text) {
    $entries = New-Object 'System.Collections.Generic.List[object]'
    foreach ($line in ($Text -split '\r?\n')) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $row = ''
        $name = $line
        if ($line -match '^\s*"?(\d+)"?\s*[;\t]\s*(.+)$') {
            $row = $Matches[1]
            $name = $Matches[2]
        } elseif ($line.Contains("`t")) {
            $name = [string]($line -split "`t" | Where-Object { $_.Trim() } | Select-Object -First 1)
        }
        $name = $name.Trim().Trim('"').Trim()
        if ($name) { $entries.Add([pscustomobject]@{ Row = $row; Name = $name }) }
    }
    , $entries
}

# Læser navne fra en CSV- eller tekstfil og returnerer linjer på formen "Række;Navn"
function Import-NameFile([string]$Path, [bool]$OnlyMissing) {
    $lines = @((Read-TextFile $Path) -split '\r?\n' | Where-Object { $_.Trim() })
    if ($lines.Count -eq 0) { return , ([string[]]@()) }

    $header = $lines[0]
    $delimiter = $null
    foreach ($d in ';', "`t", ',') { if ($header.Contains($d)) { $delimiter = $d; break } }
    $columns = if ($delimiter) { @($header.Split($delimiter) | ForEach-Object { $_.Trim().Trim('"').Trim() }) } else { @($header.Trim().Trim('"')) }
    $nameColumn = @('NavnFraExcel', 'Navn', 'Name', 'DisplayName', 'Visningsnavn') | Where-Object { $columns -contains $_ } | Select-Object -First 1

    $result = New-Object 'System.Collections.Generic.List[string]'
    if (-not $nameColumn) {
        # Ingen kendt overskrift – hver linje er et navn
        foreach ($l in $lines) { $result.Add($l.Trim().Trim('"')) }
        return , $result.ToArray()
    }

    $rowColumn = @('ExcelRaekke', 'ExcelRække', 'Raekke', 'Række', 'Row') | Where-Object { $columns -contains $_ } | Select-Object -First 1
    $mailColumns = @($columns | Where-Object { $_ -match '^(E-?mail(adresse)?|PrimarySmtpAddress|Mail)$' })
    if (-not $delimiter) { $delimiter = ',' }

    $excelRow = 1   # overskriften er række 1 i Excel
    foreach ($r in ($lines | ConvertFrom-Csv -Delimiter $delimiter)) {
        $excelRow++
        $name = ([string]$r.$nameColumn).Trim()
        if (-not $name) { continue }
        if ($OnlyMissing -and $mailColumns.Count -gt 0) {
            $hasMail = $false
            foreach ($mc in $mailColumns) { if (([string]$r.$mc).Trim()) { $hasMail = $true } }
            if ($hasMail) { continue }
        }
        $rowNo = if ($rowColumn) { ([string]$r.$rowColumn).Trim() } else { [string]$excelRow }
        $result.Add("$rowNo;$name")
    }
    , $result.ToArray()
}
#endregion

#region Opslagsindeks og matchning
# "Henrik Lorentzen - MGS" -> "Henrik Lorentzen", "Lene Petersen." -> "Lene Petersen", "Lone  Brun" -> "Lone Brun"
function Get-CleanName([string]$Name) {
    $n = [SobLookup.NameTools]::RemoveSuffix($Name.Trim())
    $n = $n -replace '\s+-\s+.*$', ''
    $n = $n -replace '[.\s]+$', ''
    $n = $n -replace '\s{2,}', ' '
    $n.Trim()
}

function New-CandidateList {
    , (New-Object 'System.Collections.Generic.List[SobLookup.Candidate]')
}

function ConvertFrom-ExoObject($Object, [string]$Info) {
    $id = [string]$Object.ExternalDirectoryObjectId
    if (-not $id) { $id = [string]$Object.Guid }
    if (-not $id) { $id = [string]$Object.Identity }
    [SobLookup.Candidate]::Create([string]$Object.DisplayName, [string]$Object.Name, [string]$Object.Alias,
        [string]$Object.PrimarySmtpAddress, [string]$Object.RecipientTypeDetails, $id, $Info)
}

function Get-MatchList($Result) {
    $list = New-CandidateList
    if ($Result.Kind -eq 'Unique') { $list.Add($Result.Match) }
    elseif ($Result.Kind -eq 'Multiple') { foreach ($m in $Result.Matches) { $list.Add($m) } }
    , $list
}

function Select-BestMatch {
    param($Hits, [string]$Method, [bool]$PreferUserMailbox = $false, [bool]$Weak = $false, [string]$ForceChoiceNote)
    if ($ForceChoiceNote) {
        return @{ Kind = 'Multiple'; Matches = $Hits; Method = $Method; Note = $ForceChoiceNote }
    }
    if ($Hits.Count -eq 1) {
        return @{ Kind = 'Unique'; Match = $Hits[0]; Method = $Method; Weak = $Weak }
    }
    if ($PreferUserMailbox) {
        $mailboxes = @($Hits | Where-Object { $_.Type -eq 'UserMailbox' })
        if ($mailboxes.Count -eq 1) {
            $others = ($Hits | Where-Object { $_ -ne $mailboxes[0] } | ForEach-Object { $_.Type }) -join ', '
            return @{
                Kind         = 'Unique'
                Match        = $mailboxes[0]
                Method       = "$Method – $($Hits.Count) match, valgte den eneste brugerpostkasse (øvrige: $others)"
                Weak         = $true
                Alternatives = $Hits
            }
        }
    }
    @{ Kind = 'Multiple'; Matches = $Hits; Method = $Method }
}

# Returnerer @{ Kind = 'Unique' | 'Multiple' | 'Suggest' | 'None'; ... }
function Find-NameMatch {
    param(
        [Parameter(Mandatory)][string]$InputName,
        [Parameter(Mandatory)][SobLookup.NameIndex]$Index,
        [bool]$PreferUserMailbox = $false,
        [switch]$NoFuzzy
    )
    $raw = $InputName.Trim()
    $clean = Get-CleanName $raw
    $nRaw = [SobLookup.NameTools]::Normalize($raw)
    $nClean = [SobLookup.NameTools]::Normalize($clean)
    $suffixNote = $null
    if ($raw -match '_([0-9a-fA-F]{10})$') {
        $suffixNote = "Navnet i Excel har Exchange-suffikset '_$($Matches[1])', men ingen modtager har præcis dette Exchange-navn. " +
            'Kandidaten har samme navn uden suffiks og kan derfor være en anden person med samme navn – kontrollér, før du vælger den.'
    }

    # Trin 1: eksakt Exchange-navn (Name)
    $hits = $Index.FindByName($raw)
    if ($hits.Count -gt 0) {
        $result = Select-BestMatch -Hits $hits -Method 'Eksakt match på Exchange-navn (Name)' -PreferUserMailbox $PreferUserMailbox
        # Et navn med suffiks peger entydigt på ét objekt. Uden suffiks advares der, hvis flere har samme visningsnavn.
        if ($result.Kind -eq 'Unique' -and -not $suffixNote) {
            $same = $Index.FindByDisplayName($result.Match.DisplayName)
            if ($same.Count -gt 1) {
                $result.Weak = $true
                $result.Method += " – bemærk: $($same.Count) modtagere har visningsnavnet '$($result.Match.DisplayName)'"
                $result.Alternatives = $same
            }
        }
        return $result
    }

    # Trin 2: eksakt visningsnavn
    $hits = $Index.FindByDisplayName($raw)
    if ($hits.Count -gt 0) {
        return (Select-BestMatch -Hits $hits -Method 'Eksakt match på visningsnavn' -PreferUserMailbox $PreferUserMailbox)
    }

    # Trin 3a: samme navn, når der ses bort fra store/små bogstaver, accenter, mellemrum og tegnsætning
    $hits = $Index.FindByNormalized($nRaw)
    if ($hits.Count -gt 0) {
        return (Select-BestMatch -Hits $hits -Method 'Match efter normalisering (store/små bogstaver, accenter, mellemrum og tegnsætning)' -PreferUserMailbox $PreferUserMailbox)
    }

    # Trin 3b: navnet uden suffikser som "_e901a4b14e" og " - Distrikt Vest"
    if ($nClean -and $nClean -ne $nRaw) {
        $hits = $Index.FindByNormalized($nClean)
        if ($hits.Count -gt 0) {
            return (Select-BestMatch -Hits $hits -Method "Match på navnet uden suffiks ('$clean')" -PreferUserMailbox $PreferUserMailbox -Weak $true -ForceChoiceNote $suffixNote)
        }
    }

    # Trin 4: fornavn + efternavn – mellemnavne og initialer må afvige eller mangle
    $tokens = [SobLookup.NameTools]::Tokens($nClean)
    $first = $null
    $last = $null
    $significant = @()
    if ($tokens.Count -ge 2) {
        $first = $tokens[0]
        $last = $tokens[-1]
        $middle = if ($tokens.Count -gt 2) { $tokens[1..($tokens.Count - 2)] } else { @() }
        $significant = @($tokens | Where-Object { $_.Length -gt 1 })
        $best = 0
        $bestHits = New-CandidateList
        foreach ($c in $Index.WithFirstName($first)) {
            $ct = $c.Tokens
            if ($ct.Count -lt 2) { continue }
            $lastOk = $ct[-1] -eq $last
            $allIn = $true
            foreach ($t in $significant) { if ($ct -notcontains $t) { $allIn = $false; break } }
            if (-not ($lastOk -or $allIn)) { continue }
            # Mellemnavne skal passe – en initial ("D") passer med et helt navn ("Dahl") og omvendt
            $cMiddle = if ($ct.Count -gt 2) { $ct[1..($ct.Count - 2)] } else { @() }
            $middleOk = $true
            foreach ($t in $middle) {
                $ok = $false
                foreach ($cm in $cMiddle) {
                    if (($cm -eq $t) -or ($t.Length -eq 1 -and $cm.StartsWith($t)) -or ($cm.Length -eq 1 -and $t.StartsWith($cm))) { $ok = $true; break }
                }
                if (-not $ok) { $middleOk = $false; break }
            }
            $score = [int]$lastOk + [int]$allIn + [int]$middleOk
            if ($score -gt $best) { $best = $score; $bestHits.Clear() }
            if ($score -eq $best) { $bestHits.Add($c) }
        }
        if ($bestHits.Count -gt 0) {
            return (Select-BestMatch -Hits $bestHits -Method 'Fornavn og efternavn matcher – mellemnavn eller initial afviger/mangler' -PreferUserMailbox $PreferUserMailbox -Weak $true -ForceChoiceNote $suffixNote)
        }
    }

    if ($NoFuzzy) { return @{ Kind = 'None' } }

    # Trin 5: forslag – stavevarianter og navne, der delvist går igen
    $suggestions = New-CandidateList
    $maxDistance = if ($nClean.Length -ge 12) { 2 } else { 1 }
    foreach ($c in $Index.FindSimilar($nClean, $maxDistance)) {
        # Fornavnet må højst afvige med ét tegn – ellers foreslås fx "Jette Petersen" for "Anette Petersen"
        if ($tokens.Count -ge 2 -and $c.Tokens.Count -gt 0 -and [SobLookup.NameTools]::Distance($first, $c.Tokens[0]) -gt 1) { continue }
        $suggestions.Add($c)
    }
    if ($tokens.Count -ge 2) {
        # Samme efternavn, og kandidatens fornavn indgår i navnet (fx "Anne Mette Hansen" -> "Mette Hansen")
        foreach ($c in $Index.WithLastName($last)) {
            if (($tokens -contains $c.Tokens[0]) -and -not $suggestions.Contains($c)) { $suggestions.Add($c) }
        }
        # Samme fornavn og mindst ét navn mere til fælles (fx nyt efternavn)
        $others = @($significant | Where-Object { $_ -ne $first })
        if ($others.Count -gt 0) {
            foreach ($c in $Index.WithFirstName($first)) {
                if ($suggestions.Contains($c)) { continue }
                foreach ($t in $others) { if ($c.Tokens -contains $t) { $suggestions.Add($c); break } }
            }
        }
    }
    if ($suggestions.Count -gt 0) {
        $top = New-CandidateList
        foreach ($c in $suggestions) { if ($top.Count -ge 8) { break }; $top.Add($c) }
        return @{ Kind = 'Suggest'; Matches = $top; Method = 'Lignende navne' }
    }
    @{ Kind = 'None' }
}

# Finder ud af, hvorfor et navn ikke findes blandt de aktive modtagere
function Get-MissingDiagnosis([string]$Name) {
    $diag = @{ Found = $false; Reason = ''; Candidates = (New-CandidateList) }
    $notes = New-Object 'System.Collections.Generic.List[string]'

    # 1) Slettede og inaktive postkasser
    if ($script:State.Deleted) {
        $hits = Get-MatchList (Find-NameMatch -InputName $Name -Index $script:State.Deleted -NoFuzzy)
        if ($hits.Count -gt 0) {
            foreach ($h in $hits) { $diag.Candidates.Add($h) }
            $diag.Found = $true
            $diag.Reason = "Brugeren findes kun som $($hits[0].Info.ToLower()). Brugeren er sandsynligvis fratrådt, og Send on behalf-rettigheden peger på en bruger, der ikke findes mere."
            return $diag
        }
    } elseif ($script:State.DeletedError) {
        $notes.Add("Slettede postkasser kunne ikke kontrolleres ($($script:State.DeletedError)).")
    }

    # 2) Brugerkonti uden postkasse (fx fratrådte eller uden licens)
    try {
        $search = Get-CleanName $Name
        $users = @(Invoke-WithoutSyncContext { Get-User -Anr $search -ResultSize 50 -ErrorAction Stop })
        if ($users.Count -gt 0) {
            $userCandidates = New-CandidateList
            foreach ($u in $users) {
                $type = [string]$u.RecipientTypeDetails
                $info = switch ($type) {
                    'User' { 'Brugerkonto uden postkasse' }
                    'DisabledUser' { 'Deaktiveret brugerkonto uden postkasse' }
                    default { "Brugerkonto ($type)" }
                }
                $id = [string]$u.ExternalDirectoryObjectId
                if (-not $id) { $id = [string]$u.UserPrincipalName }
                $userCandidates.Add([SobLookup.Candidate]::Create([string]$u.DisplayName, [string]$u.Name, '', [string]$u.UserPrincipalName, $type, $id, $info))
            }
            $hits = Get-MatchList (Find-NameMatch -InputName $Name -Index ([SobLookup.NameIndex]::new($userCandidates)) -NoFuzzy)
            if ($hits.Count -gt 0) {
                foreach ($h in $hits) { $diag.Candidates.Add($h) }
                $diag.Found = $true
                $h = $hits[0]
                if ($h.Type -in 'User', 'DisabledUser') {
                    $diag.Reason = "Brugerkontoen findes ($($h.Email) – $($h.Info.ToLower())), men har ingen postkasse eller e-mailadresse i Exchange Online. Brugeren er sandsynligvis fratrådt eller har fået fjernet sin licens."
                } else {
                    $diag.Reason = "Der findes en brugerkonto med navnet ($($h.Email), type $($h.Type)), men den er ikke blandt modtagerne i Exchange Online. Kontrollér brugeren manuelt."
                }
                return $diag
            }
        }
    } catch {
        $notes.Add("Brugerkonti kunne ikke kontrolleres ($($_.Exception.Message)).")
    }

    # 3) Intet fundet
    $reason = 'Ingen modtager, slettet postkasse eller brugerkonto i Exchange Online har dette navn.'
    if ($Name -match '_([0-9a-fA-F]{10})$') {
        $reason += " Navnet har Exchange-suffikset '_$($Matches[1])', som Exchange tilføjer, når to objekter har samme navn – objektet med dette navn findes ikke længere."
    }
    $reason += ' Brugeren er sandsynligvis slettet, eller navnet er stavet anderledes i Exchange.'
    if ($notes.Count -gt 0) { $reason += ' ' + ($notes -join ' ') }
    $diag.Reason = $reason
    $diag
}

function Set-RowMatch($Row, $Candidate, [string]$Category, [string]$Status, [string]$Method) {
    $Row.Category = $Category
    $Row.Status = $Status
    $Row.Method = $Method
    $Row.Reason = ''
    $Row.Initialer = $Candidate.Initialer
    $Row.Alias = $Candidate.Alias
    $Row.Email = $Candidate.Email
    $Row.FoundName = $Candidate.DisplayName
    $Row.Type = $Candidate.Type
}

function Resolve-Entry($Entry, [bool]$PreferUserMailbox) {
    $row = New-Object SobLookup.ResultRow
    $row.Row = [string]$Entry.Row
    $row.InputName = $Entry.Name
    try {
        $m = Find-NameMatch -InputName $Entry.Name -Index $script:State.Index -PreferUserMailbox $PreferUserMailbox

        if ($m.Kind -eq 'Unique') {
            if ($m.Weak) { Set-RowMatch $row $m.Match 'Check' 'Fundet – kontrollér' $m.Method }
            else { Set-RowMatch $row $m.Match 'Found' 'Fundet' $m.Method }
            if ($null -ne $m.Alternatives) { foreach ($a in $m.Alternatives) { $row.Candidates.Add($a) } }
        } elseif ($m.Kind -eq 'Multiple') {
            foreach ($c in $m.Matches) { $row.Candidates.Add($c) }
            $n = $row.Candidates.Count
            $row.Category = 'Choose'
            $row.Status = if ($n -eq 1) { 'Muligt match – vælg' } else { "Flere match ($n) – vælg" }
            $row.Method = $m.Method
            $row.Reason = if ($m.Note) { $m.Note } else {
                "$n modtagere matcher navnet ($($m.Method.ToLower())), så det kan ikke afgøres automatisk, hvem der menes. Vælg den rigtige under Kandidater."
            }
        } else {
            # Ikke fundet (evt. med lignende navne) – find ud af hvorfor
            $diag = Get-MissingDiagnosis $Entry.Name
            foreach ($c in $diag.Candidates) { $row.Candidates.Add($c) }
            $suggestions = @(if ($m.Kind -eq 'Suggest') { $m.Matches })
            foreach ($c in $suggestions) { if (-not $row.Candidates.Contains($c)) { $row.Candidates.Add($c) } }

            if (-not $diag.Found -and $suggestions.Count -gt 0) {
                $row.Category = 'Choose'
                $row.Status = 'Muligt match – vælg'
                $row.Reason = "Navnet findes ikke præcist i Exchange Online, men der er $($suggestions.Count) lignende navn(e) (stavevariant, mellemnavn brugt som fornavn eller nyt efternavn). Vælg under Kandidater, hvis et af dem er den rigtige person."
            } else {
                $row.Category = 'NotFound'
                $row.Status = 'Ikke fundet'
                $row.Reason = $diag.Reason
                if ($diag.Found -and $suggestions.Count -gt 0) {
                    $row.Reason += " Der findes desuden $($suggestions.Count) aktive modtager(e) med lignende navn – se Kandidater."
                }
            }
        }
    } catch {
        $row.Category = 'Error'
        $row.Status = 'Fejl'
        $row.Reason = "Opslaget fejlede: $($_.Exception.Message)"
    }
    $row.OriginalStatus = $row.Status
    $row
}
#endregion

#region Exchange Online
function Initialize-ExoModule {
    $loaded = Get-Module ExchangeOnlineManagement
    if ($loaded -and $loaded.Version -ge [version]'3.0.0') { return $true }
    if ($loaded) { Remove-Module ExchangeOnlineManagement -Force }
    $module = Get-Module -ListAvailable ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $module -or $module.Version -lt [version]'3.0.0') {
        $what = if ($module) { "Den installerede version af ExchangeOnlineManagement ($($module.Version)) er for gammel – der kræves version 3 eller nyere." }
        else { 'PowerShell-modulet ExchangeOnlineManagement er ikke installeret.' }
        $answer = Show-Message "$what`n`nVil du installere den nyeste version fra PowerShell Gallery nu (kun for din bruger)?" 'Exchange Online-modul mangler' 'Question' 'YesNo'
        if ($answer -ne 'Yes') { return $false }
        Set-Status 'Installerer ExchangeOnlineManagement fra PowerShell Gallery …' -Indeterminate
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        if ($PSVersionTable.PSEdition -eq 'Desktop' -and
            -not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [version]'2.8.5.201' })) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
        }
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Set-Status 'Indlæser ExchangeOnlineManagement …' -Indeterminate
    # -MinimumVersion: uden den indlæses den første version i PSModulePath, som kan være en gammel v2
    Import-Module ExchangeOnlineManagement -MinimumVersion 3.0.0 -ErrorAction Stop
    $true
}

function Get-ExoConnection {
    if (-not (Get-Module ExchangeOnlineManagement)) { return $null }
    try {
        Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
    } catch {
        $null
    }
}

function Update-ConnectionUi {
    $connection = Get-ExoConnection
    $script:State.Connected = $null -ne $connection
    if ($connection) {
        $ui.dotConn.Fill = New-Brush '#22C55E'
        $ui.lblConn.Text = "Forbundet: $($connection.UserPrincipalName)"
    } else {
        $ui.dotConn.Fill = New-Brush '#94A3B8'
        $ui.lblConn.Text = 'Ikke forbundet'
    }
    Update-Buttons
}

function Connect-Exo {
    if (-not (Initialize-ExoModule)) { return }
    $params = @{ ShowBanner = $false; ErrorAction = 'Stop' }
    $upn = $ui.txtUpn.Text.Trim()
    if ($upn) { $params.UserPrincipalName = $upn }
    if ($ui.chkBrowser.IsChecked) {
        if ((Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')) { $params.DisableWAM = $true }
        else { Set-Status 'Din version af ExchangeOnlineManagement understøtter ikke browser-login – bruger standardlogin.' }
    }
    Set-Status 'Venter på login til Exchange Online – log ind i det vindue, der åbner …' -Indeterminate
    # Vinduet minimeres under login, så login-dialogen ikke skjules bag det
    $previousState = $window.WindowState
    $window.WindowState = [System.Windows.WindowState]::Minimized
    Update-UI
    try {
        Invoke-WithoutSyncContext { Connect-ExchangeOnline @params | Out-Null }
    } finally {
        $window.WindowState = $previousState
        [void]$window.Activate()
    }
    Clear-RecipientCache
    Update-ConnectionUi
    if (-not $script:State.Connected) { throw 'Forbindelsen til Exchange Online kunne ikke bekræftes.' }
    Set-Status ("Forbundet til Exchange Online som {0}." -f ($ui.lblConn.Text -replace '^Forbundet: ', '')) 0
}

function Disconnect-Exo {
    Set-Status 'Afbryder forbindelsen …' -Indeterminate
    Invoke-WithoutSyncContext { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop }
    Clear-RecipientCache
    Update-ConnectionUi
    Set-Status 'Forbindelsen til Exchange Online er afbrudt.' 0
}

function Clear-RecipientCache {
    $script:State.Index = $null
    $script:State.Deleted = $null
    $script:State.DeletedError = $null
    $ui.lblRecipients.Text = 'Modtagerlisten hentes ved første søgning.'
}

# Henter alle modtagere én gang og bygger et opslagsindeks – hurtigere og mere præcist end ét opslag pr. navn
function Import-RecipientIndex {
    $recipients = New-CandidateList
    Set-Status 'Henter modtagere fra Exchange Online – det kan tage et par minutter …' -Indeterminate
    Invoke-WithoutSyncContext {
        Get-EXORecipient -ResultSize Unlimited -Properties DisplayName, Alias, PrimarySmtpAddress -ErrorAction Stop | ForEach-Object {
            $recipients.Add((ConvertFrom-ExoObject $_ ''))
            if ($recipients.Count % 250 -eq 0) {
                Set-Status ('Henter modtagere fra Exchange Online … {0:N0}' -f $recipients.Count) -Indeterminate
                if ($script:State.Cancel) { throw (New-Object System.OperationCanceledException 'Hentningen af modtagere blev stoppet.') }
            }
        }
    }

    Set-Status ('Bygger opslagsindeks over {0:N0} modtagere …' -f $recipients.Count) -Indeterminate
    $script:State.Index = [SobLookup.NameIndex]::new($recipients)

    # Slettede og inaktive postkasser bruges kun til at forklare, hvorfor et navn ikke findes
    Set-Status 'Henter slettede postkasser …' -Indeterminate
    $deleted = New-CandidateList
    $script:State.Deleted = $null
    $script:State.DeletedError = $null
    $addDeleted = {
        $ws = $_.WhenSoftDeleted
        $when = 'ukendt dato'
        if ($ws) { try { $when = ([datetime]$ws).ToString('dd-MM-yyyy') } catch { $when = [string]$ws } }
        $kind = if ($_.IsInactiveMailbox) { 'Inaktiv postkasse' } else { 'Slettet postkasse' }
        $deleted.Add((ConvertFrom-ExoObject $_ "$kind (slettet $when)"))
    }
    try {
        try {
            Invoke-WithoutSyncContext {
                Get-EXOMailbox -SoftDeletedMailbox -ResultSize Unlimited -Properties DisplayName, Alias, PrimarySmtpAddress, WhenSoftDeleted, IsInactiveMailbox -ErrorAction Stop | ForEach-Object $addDeleted
            }
        } catch {
            $deleted.Clear()
            Invoke-WithoutSyncContext {
                Get-EXOMailbox -SoftDeletedMailbox -ResultSize Unlimited -ErrorAction Stop | ForEach-Object $addDeleted
            }
        }
        $script:State.Deleted = [SobLookup.NameIndex]::new($deleted)
    } catch {
        $script:State.DeletedError = $_.Exception.Message
    }

    $text = '{0:N0} modtagere hentet kl. {1:HH:mm}' -f $recipients.Count, (Get-Date)
    if ($script:State.Deleted) { $text += ' · {0:N0} slettede postkasser' -f $deleted.Count }
    else { $text += ' · slettede postkasser kunne ikke hentes' }
    $ui.lblRecipients.Text = $text
}
#endregion

#region Handlinger i brugerfladen
function Update-Buttons {
    $idle = -not $script:State.Busy
    $ui.btnConnect.IsEnabled = $idle
    $ui.btnDisconnect.IsEnabled = $idle -and $script:State.Connected
    $ui.btnReload.IsEnabled = $idle -and $script:State.Connected
    $ui.btnLoadCsv.IsEnabled = $idle
    $ui.btnClear.IsEnabled = $idle
    $ui.btnSearch.IsEnabled = $idle
    $ui.btnStop.IsEnabled = -not $idle
    $ui.btnExport.IsEnabled = $idle -and $script:Results.Count -gt 0
    $ui.btnUseCandidate.IsEnabled = $idle -and $null -ne $ui.gridResults.SelectedItem -and $null -ne $ui.lstCandidates.SelectedItem
    $ui.txtNames.IsReadOnly = -not $idle
}

function Set-Busy([bool]$Busy) {
    $script:State.Busy = $Busy
    if ($Busy) { $script:State.Cancel = $false }
    $window.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::AppStarting } else { $null }
    Update-Buttons
}

# Kører en handling med låste knapper og viser eventuelle fejl i en dialogboks
function Invoke-UiAction([scriptblock]$Action) {
    if ($script:State.Busy) { return }
    Set-Busy $true
    try {
        & $Action
    } catch {
        if ($script:State.Cancel -or $_.Exception -is [System.OperationCanceledException]) {
            Set-Status 'Handlingen blev stoppet.' 0
        } else {
            Set-Status ('Fejl: {0}' -f $_.Exception.Message) 0
            Show-Message $_.Exception.Message 'Der opstod en fejl' 'Error' | Out-Null
        }
    } finally {
        $ui.progress.IsIndeterminate = $false
        Set-Busy $false
    }
}

function Update-NameCount {
    $n = (Read-NameEntries $ui.txtNames.Text).Count
    $ui.lblNameCount.Text = if ($n -eq 1) { '1 navn' } else { "$n navne" }
}

function Get-ResultCounts {
    $counts = @{ Found = 0; Check = 0; Choose = 0; Missing = 0 }
    foreach ($r in $script:Results) {
        switch ($r.Category) {
            'Found' { $counts.Found++ }
            'Manual' { $counts.Found++ }
            'Check' { $counts.Check++ }
            'Choose' { $counts.Choose++ }
            default { $counts.Missing++ }
        }
    }
    $counts
}

function Update-Summary {
    $counts = Get-ResultCounts
    $ui.chipFound.Text = "Fundet: $($counts.Found)"
    $ui.chipCheck.Text = "Kontrollér: $($counts.Check)"
    $ui.chipChoose.Text = "Kræver valg: $($counts.Choose)"
    $ui.chipMissing.Text = "Ikke fundet: $($counts.Missing)"
    Update-Buttons
}

function Show-RowDetails {
    $row = $ui.gridResults.SelectedItem
    if ($null -eq $row) {
        $ui.txtDetail.Text = 'Vælg en række i tabellen for at se detaljer og kandidater.'
        $ui.lstCandidates.ItemsSource = $null
    } else {
        $lines = New-Object 'System.Collections.Generic.List[string]'
        $title = $row.InputName
        if ($row.Row) { $title += "   (Excel-række $($row.Row))" }
        $lines.Add($title)
        $lines.Add("Status: $($row.Status)")
        if ($row.Email) { $lines.Add("Fundet: $($row.FoundName) – initialer $($row.Initialer), alias $($row.Alias), $($row.Email) [$($row.Type)]") }
        if ($row.Method) { $lines.Add("Metode: $($row.Method)") }
        if ($row.Reason) { $lines.Add("Årsag: $($row.Reason)") }
        $ui.txtDetail.Text = $lines -join [Environment]::NewLine
        $ui.lstCandidates.ItemsSource = $row.Candidates
    }
    Update-Buttons
}

function Use-SelectedCandidate {
    if ($script:State.Busy) { return }
    $row = $ui.gridResults.SelectedItem
    $candidate = $ui.lstCandidates.SelectedItem
    if ($null -eq $row -or $null -eq $candidate) { return }
    $method = "Valgt manuelt blandt $($row.Candidates.Count) kandidat(er) – oprindelig status: $($row.OriginalStatus)"
    if ($candidate.Info) { $method += " – bemærk: $($candidate.Info)" }
    Set-RowMatch -Row $row -Candidate $candidate -Category 'Manual' -Status 'Valgt manuelt' -Method $method
    $ui.gridResults.Items.Refresh()
    Update-Summary
    Show-RowDetails
    Set-Status ('{0} er sat til {1} ({2}).' -f $row.InputName, $candidate.DisplayName, $candidate.Email)
}

function Import-NamesFromFile([string]$Path) {
    $onlyMissing = [bool]$ui.chkOnlyMissing.IsChecked
    $lines = Import-NameFile -Path $Path -OnlyMissing $onlyMissing
    $ui.txtNames.Text = $lines -join [Environment]::NewLine
    $what = if ($onlyMissing) { ' (kun rækker uden e-mailadresse)' } else { '' }
    Set-Status ('Indlæst {0} navne fra {1}{2}.' -f $lines.Count, (Split-Path -Leaf $Path), $what) 0
}

function Import-NamesFromDialog {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Title = 'Vælg fil med navne'
    $dlg.Filter = 'CSV- og tekstfiler (*.csv;*.txt)|*.csv;*.txt|Alle filer (*.*)|*.*'
    if ($script:State.LastFolder -and (Test-Path -LiteralPath $script:State.LastFolder)) { $dlg.InitialDirectory = $script:State.LastFolder }
    if (-not $dlg.ShowDialog($window)) { return }
    $script:State.LastFolder = Split-Path -Parent $dlg.FileName
    Import-NamesFromFile $dlg.FileName
}

function Start-Lookup {
    $entries = Read-NameEntries $ui.txtNames.Text
    if ($entries.Count -eq 0) {
        Show-Message 'Der er ingen navne at søge efter. Indsæt navne, eller indlæs en CSV-fil.' | Out-Null
        return
    }
    Update-ConnectionUi
    if (-not $script:State.Connected) {
        Show-Message 'Forbind til Exchange Online først (trin 1).' | Out-Null
        return
    }
    if (-not $script:State.Index) { Import-RecipientIndex }

    $script:Results.Clear()
    Show-RowDetails
    Update-Summary
    $prefer = [bool]$ui.chkPreferMailbox.IsChecked
    $done = 0
    foreach ($entry in $entries) {
        if ($script:State.Cancel) { break }
        Set-Status ('Søger {0} af {1}: {2}' -f ($done + 1), $entries.Count, $entry.Name) (100 * $done / $entries.Count)
        $script:Results.Add((Resolve-Entry $entry $prefer))
        $done++
        Update-Summary
    }

    $c = Get-ResultCounts
    $summary = '{0} fundet, {1} skal kontrolleres, {2} kræver valg, {3} ikke fundet.' -f $c.Found, $c.Check, $c.Choose, $c.Missing
    if ($done -lt $entries.Count) {
        Set-Status ('Søgningen blev stoppet efter {0} af {1} navne: {2}' -f $done, $entries.Count, $summary) (100 * $done / $entries.Count)
    } else {
        Set-Status ("Færdig – {0} navne gennemsøgt: {1} Gem resultatet med 'Gem CSV …'." -f $entries.Count, $summary) 100
    }
}

function Export-Results {
    if ($script:Results.Count -eq 0) { return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Title = 'Gem resultatet som CSV'
    $dlg.Filter = 'CSV-fil, semikolonsepareret (*.csv)|*.csv'
    $dlg.FileName = 'SendOnBehalf_opslag_{0:yyyyMMdd_HHmm}.csv' -f (Get-Date)
    if ($script:State.LastFolder -and (Test-Path -LiteralPath $script:State.LastFolder)) { $dlg.InitialDirectory = $script:State.LastFolder }
    if (-not $dlg.ShowDialog($window)) { return }

    $foundPath = $dlg.FileName
    $folder = Split-Path -Parent $foundPath
    $script:State.LastFolder = $folder
    $missingPath = Join-Path $folder ('{0}_ikke_fundet.csv' -f [System.IO.Path]::GetFileNameWithoutExtension($foundPath))

    $found = New-Object 'System.Collections.Generic.List[object]'
    $missing = New-Object 'System.Collections.Generic.List[object]'
    foreach ($r in $script:Results) {
        if ($r.Category -in 'Found', 'Check', 'Manual') {
            $found.Add([pscustomobject][ordered]@{
                    ExcelRaekke       = $r.Row
                    NavnFraExcel      = $r.InputName
                    Initialer         = $r.Initialer
                    Alias             = $r.Alias
                    Emailadresse      = $r.Email
                    FundetDisplayName = $r.FoundName
                    Modtagertype      = $r.Type
                    Status            = $r.Status
                    Metode            = $r.Method
                })
        } else {
            $missing.Add([pscustomobject][ordered]@{
                    ExcelRaekke  = $r.Row
                    NavnFraExcel = $r.InputName
                    Status       = $r.Status
                    Aarsag       = $r.Reason
                    Kandidater   = $r.CandidatesText
                })
        }
    }
    Write-CsvFile -Rows $found.ToArray() -Path $foundPath `
        -Columns ExcelRaekke, NavnFraExcel, Initialer, Alias, Emailadresse, FundetDisplayName, Modtagertype, Status, Metode
    Write-CsvFile -Rows $missing.ToArray() -Path $missingPath `
        -Columns ExcelRaekke, NavnFraExcel, Status, Aarsag, Kandidater

    Set-Status ('Gemt: {0} fundne i {1} og {2} ikke fundne i {3}.' -f $found.Count, (Split-Path -Leaf $foundPath), $missing.Count, (Split-Path -Leaf $missingPath)) 0
    $text = "Resultatet er gemt.`n`n{0} fundne brugere (initialer, alias og e-mailadresse):`n{1}`n`n{2} ikke fundne (med årsag og eventuelle kandidater):`n{3}`n`nVil du åbne mappen?" -f
        $found.Count, $foundPath, $missing.Count, $missingPath
    if ((Show-Message $text 'Resultatet er gemt' 'Information' 'YesNo') -eq 'Yes') {
        Start-Process explorer.exe -ArgumentList "/select,`"$foundPath`""
    }
}
#endregion

#region Start
$window = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
$ui = @{}
$xamlNs = New-Object System.Xml.XmlNamespaceManager($xaml.NameTable)
$xamlNs.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
foreach ($node in $xaml.SelectNodes('//*[@x:Name]', $xamlNs)) {
    $name = $node.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
    $element = $window.FindName($name)
    if ($element) { $ui[$name] = $element }
}

$script:State = @{
    Busy         = $false
    Cancel       = $false
    Connected    = $false
    Index        = $null
    Deleted      = $null
    DeletedError = $null
    LastFolder   = $PSScriptRoot
}
$script:Results = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ui.gridResults.ItemsSource = $script:Results

$window.Add_SourceInitialized({
        try { [SobLookup.Native]::UseDarkTitleBar((New-Object System.Windows.Interop.WindowInteropHelper($window)).Handle) } catch { }
    })
$ui.btnConnect.Add_Click({ Invoke-UiAction { Connect-Exo } })
$ui.btnDisconnect.Add_Click({ Invoke-UiAction { Disconnect-Exo } })
$ui.btnReload.Add_Click({ Invoke-UiAction { Clear-RecipientCache; Import-RecipientIndex; Set-Status 'Modtagerlisten er hentet igen.' 0 } })
$ui.btnLoadCsv.Add_Click({ Invoke-UiAction { Import-NamesFromDialog } })
$ui.btnClear.Add_Click({ $ui.txtNames.Clear() })
$ui.btnSearch.Add_Click({ Invoke-UiAction { Start-Lookup } })
$ui.btnStop.Add_Click({ $script:State.Cancel = $true; $ui.lblStatus.Text = 'Stopper …' })
$ui.btnExport.Add_Click({ Invoke-UiAction { Export-Results } })
$ui.btnUseCandidate.Add_Click({ Use-SelectedCandidate })
$ui.lstCandidates.Add_MouseDoubleClick({ Use-SelectedCandidate })
$ui.lstCandidates.Add_SelectionChanged({ Update-Buttons })
$ui.gridResults.Add_SelectionChanged({ Show-RowDetails })
$ui.txtNames.Add_TextChanged({ Update-NameCount })
$window.Add_Closing({
        param($s, $e)
        if ($script:State.Busy) {
            # Vent på at den igangværende handling stopper, før vinduet lukkes
            $script:State.Cancel = $true
            $e.Cancel = $true
            $ui.lblStatus.Text = 'Stopper den igangværende handling … luk vinduet igen om et øjeblik.'
        }
    })

Update-ConnectionUi
Show-RowDetails
Update-Summary
$defaultFile = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'Manglende_SendOnBehalf_opslag.csv' }
if ($defaultFile -and (Test-Path -LiteralPath $defaultFile)) {
    try { Import-NamesFromFile $defaultFile } catch { Set-Status "Kunne ikke indlæse $defaultFile`: $($_.Exception.Message)" }
}
Update-NameCount

Write-Host 'Send on behalf-opslag kører i et separat vindue. Luk ikke dette konsolvindue, før du er færdig.' -ForegroundColor Cyan
[void]$window.ShowDialog()
#endregion
