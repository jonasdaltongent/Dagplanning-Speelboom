# Handleiding voor de leraar — Les 7: Een dagplanning in een tabel

**Vak:** Toegepaste Informatica
**Doelgroep:** klas 3MWb — 2de graad Maatschappij en welzijn, dubbele finaliteit
**Lesduur:** 1 × 50 minuten: **10 minuten instructie + 40 minuten keuzewerktijd**
**Context:** De Speelboom, een fictieve buitenschoolse opvang (groep De Eekhoorns, 6 tot 9 jaar), de namiddag van woensdag 2 december
**Waar:** lokaal 18 (Windows 11)
**Lesdag:** maandag 23 november 2026, 9de lesuur (15.25–16.15)
**Kernleerplandoel:** `BV2_04.02` (courante functionaliteiten gebruiken om digitale inhouden te creëren: een tabel invoegen en opmaken) · ondersteunend `BK2_01.02.03` · contextueel `BK2_02.14.01`
**Opdracht in Classroom:** *Les 7 — Een dagplanning in een tabel*, deadline vrijdag 27 november 2026, 20.00 uur

<!-- Beginvragen (stap 1 van CONTEXT.md), beantwoord door Jonas op 19-11-2026:
     klas 3MWb · les 7 van het lessenplan (Tekstverwerking: structuur en tabellen) · maandag 23 november 2026,
     9de uur (15.25–16.15) · gewoon in de klas in lokaal 18 op Windows 11 (Google Documenten in Chrome) ·
     vorige les (les 6, Samenwerken, netiquette en phishing, ma 16-11) is gegeven: lesstart met ophalen uit les 6 ·
     kern: zelf een tabel maken (invoegen, aanvullen, opmaken); stijlen (Kop 1, Kop 2) vallen weg en schuiven naar
     les 10, de activiteitenfiche (eerste les met de regel van 19-11-2026: hoogstens één nieuwe handeling) ·
     de dagplanning komt in het werkdocument zelf, geen apart document en geen zip. -->

## 1. Inhoud van het pakket

```text
W12 - Les 07 - MW - Een dagplanning in een tabel/
├── index.html              de lespagina: startpagina in beelden · route · één stap · checklist (4 stappen, 6 taken)
├── presentatie.html        6 dia's voor het bord; dia 1 is de vaste startdia, dia 6 heeft kloktijden uit de planning
├── css/style.css           opmaak van de lespagina
├── css/slides.css          opmaak van de dia's
├── js/script.js, js/slides.js   interactie (zelftest + codewoord EEKHOORN in stap 4, kloktijden)
├── assets/                 logo's De Speelboom en GO! Dalton Gent, lettertypes (OFL), mascotte, video/inleveren.mp4
├── werkdocument/
│   ├── Dagplanning-Eekhoorns.docx   het werkdocument dat de leerling INLEVERT (een kopie per leerling)
│   ├── maak_werkdocumenten.py       genereert Dagplanning-Eekhoorns.docx (python-docx, Pillow)
│   └── huisstijl.py, go-dalton-gent.png   de huisstijl van het werkdocument
├── lesvoorbereiding.md     de lesvoorbereiding (28 punten + bronnenonderzoek)
├── dalton-lesfiche.html    de Dalton-lesfiche (instructie 10 + KWT 40): openen, Kopieer de fiche, plakken
├── lesdoelen.json          de doelen voor je jaaroverzicht
├── classroom.json          de opdracht voor de Classroom-koppeling
└── README.md               deze handleiding
```

Geen zip, geen gedeeld document en geen losse bestanden: de leerling maakt de dagplanning **in** het werkdocument, onder de regel
*▶ Maak hier je dagplanning* (geen vak: de leerling typt de titel en voegt zelf de tabel in).

### 1b. Hoe de lespagina werkt

Dit is de **eerste les volgens de regels van 19-11-2026**: één kern (een tabel), één product, drie delen, één nieuwe handeling, de
uitleg in het werkdocument, zes dia's.

- **Startpagina**: *Wat heb je nodig?* in drie beelden (de opdracht in Classroom, deze lespagina, je werkdocument
  *Dagplanning-Eekhoorns*), *Dit maak je vandaag* (een opgemaakte dagplanning van een **andere dag**, als tekening), *Na deze les
  kan je* en *Vast?* (*1. Het infokader in je werkdocument, of Toon stapjes · 2. Je buur · 3. Je leraar*).
- **Vier stappen**: 1 *Van briefje naar tabel*, 2 *Een rij erbij en opmaken*, 3 *Waarom een tabel?* (alleen *Maak deel 3 in je
  werkdocument.*), 4 *Inleveren*. Vette regels van hoogstens 8 woorden; *Toon stapjes* toont een tekening van het scherm met
  een rood kader (het raster van **Invoegen** › **Tabel**, het rechtsklikmenu, de werkbalk met **Achtergrondkleur**). Alle beelden
  zijn tekeningen in HTML en CSS.
- **Geen theoriekaart en geen infokader op de pagina**: de uitleg (*Een tabel*, *Tabelkaart De Speelboom*) staat in het werkdocument,
  bij het deel waar de leerling ze nodig heeft.
- **Stap 4 (*Inleveren*)**: *Wat moet je doen?* is vergrendeld achter drie meerkeuzevragen. Zijn ze alle drie juist, dan toont de
  groene balk het codewoord **EEKHOORN** (dat de leerling bovenaan deel 4 typt) en klapt *Wat moet je doen?* open, met onderaan het
  filmpje `assets/video/inleveren.mp4`.
- **Checklist**: 6 taken (één per deel, plus zelftest en ingeleverd). **Geen AI-hulp** (`data-ai-hulp="uit"`).

## 2. Klaarzetten

### Stap 1 — Publiceren via GitHub Pages

Repository [`jonasdaltongent/Dagplanning-Speelboom`](https://github.com/jonasdaltongent/Dagplanning-Speelboom), Pages op branch
`main`, map `/ (root)`:
- **Lespagina:** <https://jonasdaltongent.github.io/Dagplanning-Speelboom/>
- **Presentatie:** <https://jonasdaltongent.github.io/Dagplanning-Speelboom/presentatie.html>

### Stap 2 — De opdracht in Classroom, met de koppeling

De opdracht staat als **concept** klaar in `3MWb - Informatica` en is **ingepland** (20-11-2026): Classroom publiceert ze zelf op
**maandag 23 november 2026 om 14.55 uur**, een half uur vóór het 9de lesuur van 15.25 uur.

| | Opdracht: **Les 7 — Een dagplanning in een tabel** |
|---|---|
| **Onderwerp** | Module 2 — Functionele documenten voor de praktijk |
| **Bijlage 1** | de link naar de lespagina (`Lespagina · Les 7 · Een dagplanning in een tabel · De Speelboom`) |
| **Bijlage 2** | `Dagplanning-Eekhoorns` (Google-document) — **Een kopie maken voor elke leerling** |
| **Punten** | 20 |
| **Deadline** | vrijdag 27 november 2026, 20.00 uur |

De instructietekst bevat geen e-mailadres (er wordt niets gedeeld).

```bash
python3 "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/_tools/zet_opdracht_klaar.py" "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/2026-2027/W12 - Les 07 - MW - Een dagplanning in een tabel" --proef
```

### Stap 3 — Met de hand, als de koppeling niet werkt

Upload `werkdocument/Dagplanning-Eekhoorns.docx` in Drive zelf (**Nieuw** › **Bestand uploaden**; met de instelling *Uploads converteren*
wordt het een Google-document). Maak in Classroom een opdracht met **Bijvoegen** › **Link** (de lespagina) en **Bijvoegen** › **Drive**
(het werkdocument, **Een kopie maken voor elke leerling**), 20 punten, deadline vr 27-11 om 20.00 uur.

### Afvinklijst vóór de les

- [ ] De lespagina en presentatie staan online (HTTP 200).
- [ ] De opdracht staat als **concept** klaar in `3MWb - Informatica` en is ingepland voor ma 23-11 om 14.55 uur.
- [ ] Open het werkdocument in Drive en kijk na dat onder *▶ Maak hier je dagplanning* nog **twee lege regels** staan.
- [ ] Het wachtwoord van de klascomputers op het bord (staat nooit op een dia).
- [ ] Open `presentatie.html` op het bord op **dia 1 (*Zo start je*)** terwijl de leerlingen binnenkomen.
- [ ] Kopieer de Dalton-lesfiche uit `dalton-lesfiche.html` naar je planner.

### 2b. Nagelezen klikpaden

Uit de Nederlandse helppagina's, ruwe tekst, gelezen op 20-11-2026. Waar jouw scherm anders zegt, geldt jouw scherm.

| Handeling | Klikpad / naam | Bron |
|---|---|---|
| Een tabel invoegen | **Invoegen** › **Tabel**, kies hoeveel rijen en kolommen (hoogstens 20 × 20): met de muis over het raster, *4 naar rechts en 7 naar beneden* | [Docs 1696711](https://support.google.com/docs/answer/1696711?hl=nl); het raster is **beschreven**, niet benoemd |
| Een rij toevoegen | rechtsklik op een cel in de tabel › **Rij onder invoegen** (ook **Rij boven**, **Kolom links**, **Kolom rechts invoegen**) | idem; ook `W05 - Les 05 - ORLO` § 2b |
| De hele koprij selecteren | slepen met de muis over de cellen van de rij | beschreven (niet in de help) |
| Celkleur | cellen selecteren › in de werkbalk **Achtergrondkleur** (de help: *Achtergrondkleur of Opvulkleur*); de kleur is **beschreven**: *een lichtgroene kleur* | idem |
| Vet | **Ctrl + B** | [Sneltoetsen Docs 179738](https://support.google.com/docs/answer/179738?hl=nl) |
| Naar de volgende cel | **Tab** | **niet in de help gevonden**: algemene kennis, nog na te kijken |
| Ongedaan maken | **Ctrl + Z** | algemene kennis (Windows) |
| Inleveren | **Inleveren** › venster **Je werk inleveren?** › **Inleveren** | `_afspraken/klikpaden.md` (demoleerling3, 15-10-2026) |
| Computer vergrendelen | **Windows-toets + L** | [Microsoft](https://support.microsoft.com/nl-nl/windows/sneltoetsen-in-windows-dcc61a57-8ff0-cffe-9796-cb9706c75eec) |

## 3. Het verloop van de les

### Instructie en demo (10 minuten, dia 1–6)

- **Dia 1 (0–1'):** *Zo start je.* Wacht tot de meesten de lespagina open hebben.
- **Dia 2 (1–2'):** *Weet je het nog?* Vingervraag over les 6: wat stond er in de eerste rij van de tabel? → **A** (*Taak · Wie · Klaar
  tegen · Klaar?*). Dat is de **koprij**.
- **Dia 3 (2–3'):** *Dit maak je vandaag.* Het briefje van een **andere dag** (maandag 30 november) en de opgemaakte tabel ernaast:
  kijk naar de titel erboven, de lichtgroene koprij en *15.15*. Lees de tekening één keer hardop voor.
- **Dia 4 (3–8'):** de demo in je eigen kopie, alleen het nieuwe: de titel typen, **Enter**, **Invoegen** › **Tabel** (4 × 7), de koprij typen
  met **Tab**, een rij uit zin 1, en rechtsklik › **Rij onder invoegen** met **Ctrl + Z**. **Niet** voorgedaan: de koprij opmaken.
- **Dia 5 (8–9'):** *Controle · iedereen antwoordt.* Waar komt het vieruurtje van 15.15? → **2** (onder 14.30, boven 16.00).
- **Dia 6 (9–10'):** *Zo werk je verder.* Stap 1 (de tabel invoegen, 3') doen ze nu samen: duim omhoog als de tabel er staat.

### Keuzewerktijd (40 minuten)

- **Eerste rondgang (stap 1):** staat er een echte tabel (geen kolommen met spaties of Tab), de titel boven de tabel en niet erin, en
  één ding per cel?
- **Tweede rondgang (stap 2):** staat het vieruurtje tussen 14.30 en 16.00, is de **hele** koprij vet en lichtgroen, en staat er
  overal 15.15 (niet *kwart over drie*)?
- **Laatste minuut:** de exitvraag in deel 4 (*een nieuwe rij of een nieuwe kolom?*), en uitloggen of **Windows + L**. Zeg mondeling dat wie
  niet klaar is, toch inlevert.

## 4. Verbetersleutel

Het briefje en de tabel (deel 1 en 2), op volgorde van het uur, met het vieruurtje erbij:

| Uur | Activiteit | Waar | Wie |
|---|---|---|---|
| 12.00 | Boterhammen eten | refter | Lien |
| 12.45 | Vrij spel | speelplaats | Sam |
| 13.30 | Mijters knutselen | atelier | stagiair |
| 14.30 | Voorlezen | leeshoek | Lien |
| 15.15 | Fruit eten | refter | Sam |
| 16.00 | Buiten spelen | tuin | Sam |
| 17.00 | Ophalen | refter | Lien |

- **Codewoord** bovenaan deel 4: **EEKHOORN** (ontbreekt of fout = −2; een tikfout telt als juist).
- **Deel 1, tabel en koprij (3 punten):** een echte tabel (geen spaties of tabs) met 4 kolommen en de koprij *Uur · Activiteit · Waar · Wie*.
- **Deel 1, de zes activiteiten (6 punten):** per rij 1 punt voor het juiste uur, de activiteit, de plaats en de naam, één ding per cel; één
  fout in de rij = 0,5 punt. Een andere formulering (*Eten*, *Knutselen: mijters*) is goed; de plaats en de naam moeten kloppen.
- **Deel 2, het vieruurtje (2 punten):** een eigen rij *15.15 · Fruit eten · refter · Sam*, tussen 14.30 en 16.00 (onderaan of bovenaan = 0).
- **Deel 2, elk uur op dezelfde manier (1 punt):** *15.15* (niet *kwart over drie*, *15u15*, *15:15*), en de andere uren zonder *uur*.
- **Deel 1 en 2, de titel (1 punt):** *Dagplanning De Eekhoorns – woensdag 2 december* boven de tabel (niet erin), vet.
- **Deel 2, de koprij opgemaakt (2 punten):** de hele koprij vet (1) en lichtgroen (1); maar één cel = 0,5.
- **Deel 2, geen lege rijen (1 punt):** geen lege rij in of onder de tabel.
- **Deel 3, vraag 2 (2 punten):** rij **15.15** (0,5), kolommen **Waar** en **Wie** (0,5), antwoord *in de refter, met Sam* (1). Wie het
  vieruurtje vergat, antwoordt *14.30, leeshoek, Lien*: fout, en een teken dat deel 2 niet af is.
- **Deel 3, vraag 3 (1 punt):** wat er bij Sam misging (de woorden na *en voetbal* schoven op, de kolommen staan niet meer onder
  elkaar) **én** waarom dat in een tabel niet gebeurt (elke kolom blijft op haar plaats). Niet: *tabellen zijn professioneler* of *het is
  overzichtelijker* zonder te zeggen wat er bij Sam misliep.
- **Deel 4, exitvraag (1 punt):** *een kolom*, met een reden: *materiaal* is een nieuw soort gegeven dat je bij elke activiteit wilt zien (zoals
  *Waar* en *Wie*). *Een rij* is fout, ook met een uitleg.

Wat niet telt: spelling en hoofdletters in de cellen, hoe de activiteit precies heet, welke lichtgroene kleur, de kolombreedte, of de titel
ook groter is. Aftrek volgens `_afspraken/evaluatie-en-privacy.md` (niets ingeleverd = 0, te laat = −5, codewoord leeg of fout = −2).

### Essentiële fouten — geef hier altijd feedback op

- Kolommen gemaakt met **spaties of tabs** in plaats van een tabel (deel 1 krijgt dan hoogstens de punten van de zes activiteiten).
- Een **hele zin** of twee gegevens in één cel.
- Het vieruurtje **onderaan** in plaats van op volgorde van het uur.
- De **titel in de tabel** getypt, of maar **één cel** van de koprij opgemaakt.

## 5. Schermafbeeldingen (optioneel)

Alle beelden op de lespagina en de dia's zijn tekeningen in HTML en CSS (het raster van **Invoegen** › **Tabel**, het rechtsklikmenu,
de werkbalk met **Achtergrondkleur** en de kleuren, de dagplanning, de opdracht in Classroom). Een echte schermafbeelding kan er later
voor in de plaats komen (`_afspraken/filmpjes-en-schermafbeeldingen.md`); `assets/screenshots/LEESMIJ.md` houdt dat bij.

## 6. Het materiaal opnieuw maken

```bash
cd "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/2026-2027/W12 - Les 07 - MW - Een dagplanning in een tabel/werkdocument" && python3 maak_werkdocumenten.py
```

## 7. Leerplandoelen in je jaaroverzicht

`lesdoelen.json` (klasgroep `3 MW`): `BV2_04.02` (kern) · `BK2_01.02.03` (ondersteunend) · `BK2_02.14.01` (contextueel). De pre-push hook
werkt `Leerplandoelen 2026-2027.xlsx` bij.

## 8. Wat nog moet blijken in de klas

- **De twee lege regels onder *▶ Maak hier je dagplanning*:** blijven ze bestaan na het omzetten naar een Google-document, en past
  de tabel er zonder dat ze tegen de kop van deel 2 komt? (Zie de afvinklijst.)
- **Het raster van Invoegen › Tabel:** de lespagina *beschrijft* het (*4 naar rechts, 7 naar beneden*) en tekent een eigen raster van 8 × 8;
  het echte raster is anders en de help noemt het niet. Zien de leerlingen de maat (*4 × 7*) tijdens het bewegen?
- **De Tab-toets naar de volgende cel** (lespagina, stap 1) staat niet in de help; nakijken in Documenten met een leerlingaccount.
- **De werkbalk:** heet de knop met de verfemmer bij een leerling **Achtergrondkleur**, en welke kleuren zie je in het palet? De lespagina
  beschrijft *een lichtgroene kleur* en tekent de knop als tekening, niet als schermafbeelding.
- **Rechtsklikken op een Windows-computer in lokaal 18** (muis of touchpad): lukt het met het touchpad?
- **De opmaak nakijken na de deadline:** kan de verbetertool *vet* en *lichtgroen* uit het omgezette Google-document lezen? Zo niet, dan
  kijkt Jonas de 3 punten voor titel en koprij zelf na.
- **De tabel zit buiten een vak:** de tool leest dan de tekst van deel 1, maar niet noodzakelijk de opbouw van de tabel (4 × 7).
- **Tijd:** past stap 1 in 15 minuten (3 samen + 12), en stap 2 in 12? Of is het typen van de zes zinnen te veel?
- **De planning van Sam (deel 3):** is de tekening (monospace, *en voetbal* geel) duidelijk genoeg om de verschuiving te zien?
- **Niet goedgekeurd:** `_vakinhoud/tekstverwerking.md` is een concept van 20-11-2026.
- **De zelftest met het codewoord EEKHOORN en het filmpje Inleveren:** werkt het slot in lokaal 18 vlot en leveren alle leerlingen in?
