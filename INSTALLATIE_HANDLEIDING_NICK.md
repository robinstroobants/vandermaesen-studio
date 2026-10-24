# 📘 Handleiding & Installatie: VDM Prijslijst Studio (HP EliteBook / Windows)

Deze handleiding legt stap voor stap (zonder technisch jargon) uit hoe je de **VDM Prijslijst Studio** op je **HP EliteBook** zet en hoe je jouw (gevoelige) Excel-bestanden toevoegt.

> **🔒 100% Veilig & Lokaal:**
> De tool draait volledig op je eigen laptop. Omwille van strikte vertrouwelijkheid bevat de programmacode **geen enkele prijslijst of klantenlijst**. Al jouw gevoelige bestanden (zowel je Excel-bronbestanden als de gegenereerde `output/`) worden bewaard in één afgeschermde map **`confidential/`** en verlaten **nooit** je computer.
> Het **Vandermaesen-logo** (`assets/vdm_logo.png`) en de **standaard e-mailsjablonen** (`templates/mails/`) zitten wél al netjes voorgeprogrammeerd in de app!

---

## Stap 1: Eenmalig Python aanzetten op je HP EliteBook (1 minuut)

Om de Studio op je Windows-laptop te laten draaien, heb je één standaard hulpprogramma van Microsoft nodig (**Python**). Je installeert dit veilig via de ingebouwde winkel van Windows:

1. Klik linksonder op je scherm op het **Windows Start-icoontje** (of druk op de Windows-toets op je toetsenbord).
2. Typ **`Microsoft Store`** en druk op *Enter*.
3. Typ bovenaan in de zoekbalk van de Microsoft Store: **`Python 3.12`** en druk op *Enter*.
4. Klik bij **Python 3.12** op de blauwe knop **Downloaden** (of **Installeren**).
5. Wacht een halve minuut tot er *"Geïnstalleerd"* staat. Je mag de Microsoft Store daarna sluiten — hier hoef je nooit meer naar om te kijken!

---

## Stap 2: De 2 Gevoelige Excel-bestanden toevoegen in `confidential/`

Zet de map **`vandermaesen`** op een vaste plek op je laptop (bijvoorbeeld op je Bureaublad of in `Documenten`) en dubbelklik op **`Start_VDM_Studio_Windows.bat`**.

Omdat jouw prijslijsten en klantenkortingen strikt vertrouwelijk zijn, worden de **2 gevoelige Excel-bestanden** niet meegeleverd in de installatie. Zolang een van deze bestanden ontbreekt in de map **`confidential/`**, meldt de Studio dat meteen bovenaan op je scherm in een **gele waarschuwingsbalk**.

Je kunt deze 2 bestanden op **twee manieren** toevoegen (kies wat jij het makkelijkst vindt):

### Optie A: Rechtstreeks uploaden via het scherm (Makkelijkst!)
Zodra je de Studio opent en een bestand ontbreekt, zie je bovenaan een gele balk met **3 status-kaartjes**:
1. Klik bij **1. Prijsberekening Bronfile (NL)** op **`📤 Selecteer Bronfile (NL)`** en kies jouw `Art. Prijzen_BRON_...xlsx`.
2. Klik bij **2. Master Klantenconfig (.xlsx)** op **`📤 Selecteer Master Klantenconfig`** en kies jouw `00_MASTER_KLANTENCONFIG.xlsx`.
3. *(Optioneel)* Klik bij **3. Bronfile Frans** op **`📤 Selecteer Bronfile (FR)`** als je een Franstalige bronfile gebruikt.

De Studio zet beide bestanden automatisch in de map **`confidential/`** en laadt direct al je klanten in! *(Later een nieuwe bronfile inladen? Klik rechtsboven op **`📁 Bronbestanden Beheren`**).*

### Optie B: Zelf in de map `confidential/` slepen via Windows Verkenner
Wil je de bestanden liever zelf met de hand in de map zetten? Zowel je bron-prijslijst, je klantenconfig als alle gegenereerde output zitten overzichtelijk samen in één map **`confidential/`**:

```text
📁 vandermaesen/
 ├── 🚀 Start_VDM_Studio_Windows.bat          <-- HIER DUBBELKLIK JE OP OM TE STARTEN
 ├── 📄 INSTALLATIE_HANDLEIDING_NICK.pdf      <-- Deze handleiding
 │
 ├── 🔒 confidential/                         <-- ÉÉN MAP VOOR AL JOUW VERTROUWELIJKE DATA:
 │    ├── 📊 00_MASTER_KLANTENCONFIG.xlsx         (Gevoelig 1: Jouw klantenlijst, kortingen & aansprekingen)
 │    ├── 📊 Art. Prijzen_BRON_20260601.xlsx      (Gevoelig 2: Jouw actuele NL Bronfile met formules)
 │    ├── 📊 Art. Prijzen_BRON_FRANS_20260601.xlsx(Optioneel: Jouw actuele FR Bronfile)
 │    │
 │    └── 📁 output/                              (Hier rollen alle kant-en-klare bestanden uit!)
 │         ├── 📁 01_Prijslijsten_Klanten/        (Alle 74 opgeschoonde Excel-prijslijsten per klant)
 │         ├── 📁 02_Outlook_Conceptmails/        (Alle 74 kant-en-klare Outlook .eml mails met bijlage)
 │         └── 📁 03_Navision_Export/             (Kant-en-klare lijst met RGK-kortingen & prijzen)
 │
 └── 📁 assets/ & 📁 templates/mails/         <-- Zitten al ingebouwd! (Logo & standaard e-mailsjablonen)
```

### Goed om te weten over jouw 2 Excel-bestanden in `confidential/`:
1. **Jouw Bron-Prijslijst (`confidential/Art. Prijzen_BRON_...xlsx`):**
   - **Belangrijk voor je gemoedsrust:** In je bronbestand mogen al je formules, koppelingen naar `V:\Documenten\Prijsberekening\...`, basis-meterprijzen (`Euro/m`) en verborgen kolommen gewoon blijven staan. De Studio stript die er automatisch **100% uit** voor elke klant!
2. **Jouw Master Klantenconfig (`confidential/00_MASTER_KLANTENCONFIG.xlsx`):**
   - Je kunt klantvoorwaarden aanpassen via het scherm van de Studio, **óf** gewoon rechtstreeks in Excel door **`confidential/00_MASTER_KLANTENCONFIG.xlsx`** te openen!
   - In dit Excel-bestand vind je per klant duidelijke kolommen voor:
     - **`Contactpersoon (Voornaam)`** (bv. `Anja en Mario`)
     - **`E-mailadres`** (bv. `anmabeton@hotmail.be`)
     - **`Aanspreking Mail`** (bv. `Goedemiddag Anja, / Goedemiddag Mario` — een `/` zorgt automatisch voor een nieuwe regel in de e-mail)
     - **Alle kortingspercentages (`% Korting RP/SP`, `% Poorten <10`, `% Poorten >=10`, `% Band`, etc.)** en **Actieve Tabbladen**.
   - Sla je `00_MASTER_KLANTENCONFIG.xlsx` op in Excel (`Ctrl + S`)? Dan leest de Studio jouw wijzigingen bij de volgende klik automatisch in!

---

## Stap 3: Hoe gebruik je de Studio bij een prijswijziging?

1. **Opstarten:**
   - Dubbelklik op **`Start_VDM_Studio_Windows.bat`**.
   - *Automatische Updates:* Bij het opstarten controleert de tool op de achtergrond (zonder dat je iets hoeft te doen) of Robin een verbetering heeft klaargezet. Enkel als er daadwerkelijk een nieuwe update is binnengehaald, zie je bovenaan een groene melding.
   - Je internetbrowser (Edge of Chrome) opent automatisch het **VDM Prijslijst Studio** dashboard. *(Let op: laat het kleine zwarte venstertje op de achtergrond openstaan zolang je in de Studio werkt).*
2. **Ingangsdatum & Mailtekst kiezen:**
   - Kies bovenaan bij **Ingangsdatum** de startdatum van de nieuwe prijzen (bv. `01/06/2026`). De tool past dan automatisch overal `1 juni 2026`, `1/6/26` én de bestandsnamen aan.
   - Wil je de tekst van de e-mail aanpassen? Klik rechtsboven op de blauwe tegel **✉️ Markdown Mailer ↗**. Daar kan je de mailtekst aanpassen en direct rechts een voorbeeld per klant bekijken.
3. **Genereren (1 klik):**
   - Klik in de eerste tegel (**Prijslijsten & Outlook Mails**) op de groene knop **`⚡ Genereer`**.
   - Na **~10 seconden** zijn alle 74 prijslijsten, alle 74 Outlook-conceptmails én het Navision-bestand klaar in **`confidential/output/`**!
4. **Controleren & Verzenden in Outlook:**
   - Klik op **`📂 Open Output Map`** (opent automatisch `confidential/output/`).
   - In **`01_Prijslijsten_Klanten`** staan alle schone Excel-prijslijsten.
   - In **`02_Outlook_Conceptmails`** staan alle 74 e-mails klaar. Dubbelklik op een `.eml`-bestand: **Outlook opent direct de volledige e-mail** (met juiste ontvanger, onderwerp, persoonlijke aanspreking, Vandermaesen-logo én de juiste klant-Excel al als bijlage toegevoegd). Je hoeft enkel nog op **Verzenden** te klikken!
