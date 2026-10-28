# Mailbox-rettigheder

## Find-SendOnBehalfBrugere.ps1

GUI-værktøj der slår navne op i Exchange Online og finder **initialer (det før @), alias og e-mailadresse**.
Det bruges til at udfylde de tomme felter i fanen *Send on behalf* i `Sorterede_mailbox_rettigheder.xlsx`.

### Krav

- Windows med Windows PowerShell 5.1 eller PowerShell 7
- PowerShell-modulet `ExchangeOnlineManagement` v3 eller nyere. Mangler det, tilbyder scriptet at installere det.
- En konto der kan læse modtagere i Exchange Online, fx rollen *View-Only Recipients*

### Sådan bruges det

1. Læg `Manglende_SendOnBehalf_opslag.csv` i samme mappe som scriptet. Den indlæses så automatisk.
2. Dobbeltklik på `Start-SendOnBehalfOpslag.cmd`.
   Du kan også køre `powershell -STA -File .\Find-SendOnBehalfBrugere.ps1`.
   Er filen hentet fra internettet, skal den måske først låses op med `Unblock-File`.
3. **Forbind** til Exchange Online og log ind (MFA understøttes).
4. Navnene kan indlæses på tre måder: automatisk (punkt 1), med **Indlæs CSV …** eller ved at indsætte dem direkte fra Excel.
   Skriv én pr. linje, enten som `Række;Navn` eller bare `Navn`.
   **Indlæs CSV** forstår både `Manglende_SendOnBehalf_opslag.csv` og en CSV-eksport af fanen (`Navn;Alias;E-mailadresse`).
   Som standard indlæses kun rækker uden e-mailadresse.
5. Klik **Start søgning**. Hvert navn får en farvet status:

   | Status | Betydning |
   |---|---|
   | Fundet (grøn) | Entydigt match på Exchange-navn, visningsnavn eller normaliseret navn |
   | Fundet – kontrollér (gul) | Sandsynligt match, fx hvor mellemnavnet afviger, eller hvor flere har samme visningsnavn |
   | Flere match / Muligt match – vælg (orange) | Flere kandidater eller kun lignende navne. Vælg den rigtige under *Kandidater* og klik **Brug valgte kandidat** |
   | Valgt manuelt (blå) | Du har selv valgt en kandidat |
   | Ikke fundet (rød) | Årsagen står i *Bemærkning*, fx slettet postkasse, brugerkonto uden postkasse eller at navnet ikke findes |

6. Klik **Gem CSV …**. Der gemmes to semikolonseparerede UTF-8-filer, som åbner direkte i dansk Excel:
   - `<navn>.csv` med de fundne brugere: `ExcelRaekke;NavnFraExcel;Initialer;Alias;Emailadresse;FundetDisplayName;Modtagertype;Status;Metode`
   - `<navn>_ikke_fundet.csv` med dem, der ikke blev fundet: `ExcelRaekke;NavnFraExcel;Status;Aarsag;Kandidater`

### Sådan matches navnene

Alle modtagere hentes én gang (`Get-EXORecipient`). Derefter prøves hvert navn i denne rækkefølge:

1. Eksakt **Exchange-navn** (`Name`), fx `Birgitte Pedersen_e901a4b14e`. Exchange tilføjer suffikset `_xxxxxxxxxx`, når et navn allerede er i brug.
2. Eksakt **visningsnavn** (`DisplayName`)
3. **Normaliseret navn**: der ses bort fra store og små bogstaver, accenter (ä/a), dobbelte mellemrum, punktummer og suffikser som ` - MGS`
4. **Fornavn og efternavn**: mellemnavne og initialer må afvige eller mangle. Giver status *kontrollér*.
5. **Lignende navne**: stavevarianter, et mellemnavn brugt som fornavn eller et nyt efternavn. Disse skal altid vælges manuelt.

Findes et navn ikke, slås det op blandt slettede og inaktive postkasser (`Get-EXOMailbox -SoftDeletedMailbox`) og blandt brugerkonti uden postkasse (`Get-User`). Så kan scriptet forklare, hvorfor brugeren ikke findes.

CSV- og Excel-filer er udelukket i `.gitignore`, så personoplysninger ikke ender i repoet.
