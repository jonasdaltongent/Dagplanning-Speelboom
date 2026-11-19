#!/usr/bin/env python3
"""
maak_werkdocumenten.py
Genereert het werkdocument voor les 7 "Een dagplanning in een tabel" (De Speelboom, klas 3MWb):

  Dagplanning-Eekhoorns.docx   het werkdocument dat de leerling INLEVERT (wordt in Classroom een Google-document,
                               een kopie voor elke leerling)

Gebruik:  python3 maak_werkdocumenten.py     (vereist python-docx en Pillow; huisstijl.py en go-dalton-gent.png
          staan ernaast)

Geen zip en geen losse bestanden: de dagplanning zelf maakt de leerling IN dit werkdocument, onder de regel
"▶ Maak hier je dagplanning" van deel 1.

LET OP bij aanpassen:
 - Alle namen en adressen zijn verzonnen; het adres eindigt op .example (RFC 2606). Geen telefoonnummers.
 - De dagplanning staat NIET in een vak: de leerling typt de titel en voegt zelf de tabel in. Onder de regel
   "▶ Maak hier je dagplanning" staan twee lege regels (een tabel mag niet tegen de volgende kop staan). Controleer na het
   uploaden in Drive dat die lege regels er nog zijn.
 - Het voorbeeld (zin 1 als rij, en vraag 1 van deel 3) staat al ingevuld, in het grijs. Laat het ingevuld.
 - Elk antwoord staat in een eigen vak, en elk vak hoort bij één criterium van lesvoorbereiding § 22
   (_afspraken/verbeteren.md, *Een les die je kan verbeteren*). Wijzig je de vakken, wijzig dan § 14 en § 22.
 - De punten in deel 4 (3 · 6 · 2 · 1 · 1 · 2 · 1 · 2 · 1 · 1 = 20) komen overeen met lesvoorbereiding § 22.
 - Het codewoord (EEKHOORN) staat bovenaan deel 4. De lespagina toont het pas als de zelftest juist is.
 - De planning van Sam (deel 3) is een tekening (Pillow, Menlo): dezelfde namiddag met spaties, eerst netjes en dan nadat
   Sam "en voetbal" bijtypte. De tekening moet monospace zijn: alleen dan schuiven de kolommen zichtbaar op.
 - Het bericht van Lien in deel 2 is een tekening in een tabel (ontvangen()); er is niets om op te klikken.
"""
import io
import os

from PIL import Image, ImageDraw, ImageFont

import huisstijl as hs

HERE = os.path.dirname(os.path.abspath(__file__))
UIT = os.path.join(HERE, "Dagplanning-Eekhoorns.docx")

B = hs.BREEDTE
GRIJS_VLAK = "F1F3F4"
GRIJS_RAND = "9AA0A6"
TITEL_PLANNING = "Dagplanning De Eekhoorns – woensdag 2 december"

# Het briefje van Lien (lesvoorbereiding § 4): zes zinnen, telkens "Om [uur] [wat] [waar], met [wie]".
BRIEFJE = [
    "Om 12.00 uur eten de kinderen boterhammen in de refter, met Lien.",
    "Om 12.45 uur is er vrij spel op de speelplaats, met Sam.",
    "Om 13.30 uur knutselen de kinderen mijters in het atelier, met de stagiair.",
    "Om 14.30 uur lezen we voor in de leeshoek, met Lien.",
    "Om 16.00 uur spelen de kinderen buiten in de tuin, met Sam.",
    "Om 17.00 uur halen de ouders hun kind op in de refter, met Lien.",
]

# Het tweede briefje (Extra): vrijdag 4 december, het Sinterklaasfeest. Vier zinnen in dezelfde vorm.
BRIEFJE_EXTRA = [
    "Om 13.00 uur versieren de kinderen de zaal in de refter, met Sam.",
    "Om 14.00 uur oefenen de kinderen Sinterklaasliedjes in de leeshoek, met Lien.",
    "Om 15.30 uur komt Sinterklaas op bezoek in de refter, met Lien.",
    "Om 16.30 uur eten de kinderen speculaas in de tuin, met de stagiair.",
]


def samen(rij, van, tot, cm):
    """Voegt de cellen van..tot van een rij samen en geeft de samengevoegde cel zijn volle breedte (anders raden
    Quick Look en Google Documenten die zelf)."""
    cel = rij.cells[van].merge(rij.cells[tot])
    cel.width = hs.Cm(cm)
    return cel


def ontvangen(doc, titel, velden, tekst):
    """Een ontvangen bericht (mail) als tekening: geen vakken, er valt niets in te typen.
    velden: [(label, tekst)] zoals Van en Onderwerp. tekst: lijst van alinea's."""
    aantal = 1 + len(velden) + 1
    t = doc.add_table(rows=aantal, cols=2)
    t.alignment = hs.WD_TABLE_ALIGNMENT.CENTER
    hs.randen(t, buiten=GRIJS_RAND)
    hs.breedtes(t, [2.6, B - 2.6])
    kop = samen(t.rows[0], 0, 1, B)
    hs.schaduw(kop, "F2F6FC")
    hs.celtekst(kop, titel, vet=True, kleur=hs.RGBColor(0x20, 0x21, 0x24), grootte=11)
    kop.vertical_alignment = hs.WD_CELL_VERTICAL_ALIGNMENT.CENTER
    hs.hoogte(t.rows[0], 0.7)
    for i, (label, waarde) in enumerate(velden, start=1):
        links, rechts = t.rows[i].cells
        hs.celtekst(links, label, kleur=hs.GRIJS, grootte=10)
        hs.celtekst(rechts, waarde)
        for c in (links, rechts):
            c.vertical_alignment = hs.WD_CELL_VERTICAL_ALIGNMENT.CENTER
            hs.cel_randen(c, "DADCE0", dikte=6, kanten=("bottom",))
        hs.hoogte(t.rows[i], 0.75)
    r = 1 + len(velden)
    cel = samen(t.rows[r], 0, 1, B)
    hs.marges(cel, boven=0.15, onder=0.15, links=0.3, rechts=0.3)
    cel.text = ""
    for j, alinea in enumerate(tekst):
        p = cel.paragraphs[0] if j == 0 else cel.add_paragraph()
        p.paragraph_format.space_after = hs.Pt(3)
        p.add_run(alinea)
    hs.niet_splitsen(t)
    hs.ruimte(doc, 6)
    return t


def briefje(doc, titel, zinnen):
    """Het briefje van Lien als materiaal: een kopregel (met de letter) en genummerde zinnen. Geen vakken."""
    t = doc.add_table(rows=1 + len(zinnen), cols=2)
    t.alignment = hs.WD_TABLE_ALIGNMENT.CENTER
    hs.randen(t, buiten=GRIJS_RAND, binnen="E3E3E3")
    hs.breedtes(t, [1.0, B - 1.0])
    kop = samen(t.rows[0], 0, 1, B)
    hs.schaduw(kop, "FFF8E1")
    hs.celtekst(kop, titel, vet=True, grootte=11)
    kop.vertical_alignment = hs.WD_CELL_VERTICAL_ALIGNMENT.CENTER
    hs.hoogte(t.rows[0], 0.7)
    for i, zin in enumerate(zinnen, start=1):
        links, rechts = t.rows[i].cells
        for c in (links, rechts):
            c.vertical_alignment = hs.WD_CELL_VERTICAL_ALIGNMENT.CENTER
            hs.schaduw(c, "FFFDF3")
        hs.celtekst(links, str(i), vet=True, kleur=hs.NAVY_DIEP)
        links.paragraphs[0].alignment = hs.WD_ALIGN_PARAGRAPH.CENTER
        hs.celtekst(rechts, zin)
        hs.hoogte(t.rows[i], 0.8)
    hs.niet_splitsen(t)
    hs.ruimte(doc, 6)
    return t


def materiaal_regel(doc, titel, tekst):
    """Een klein materiaalvak met één regel (de titel om over te nemen)."""
    t = doc.add_table(rows=2, cols=1)
    t.alignment = hs.WD_TABLE_ALIGNMENT.CENTER
    hs.randen(t, buiten=GRIJS_RAND, binnen="E3E3E3")
    hs.breedtes(t, [B])
    kop, inhoud = t.rows[0].cells[0], t.rows[1].cells[0]
    hs.schaduw(kop, "F2F6FC")
    hs.celtekst(kop, titel, vet=True, grootte=11)
    hs.hoogte(t.rows[0], 0.7)
    hs.celtekst(inhoud, tekst)
    inhoud.vertical_alignment = hs.WD_CELL_VERTICAL_ALIGNMENT.CENTER
    hs.hoogte(t.rows[1], 0.85)
    hs.niet_splitsen(t)
    hs.ruimte(doc, 6)
    return t


def plaats_voor_tabel(doc, tekst):
    """De plaats waar de leerling zelf de tabel invoegt: GEEN vak. Een regel met een pijl en twee lege regels eronder
    (de tabel mag niet tegen de volgende kop staan). Controleer na het uploaden dat de lege regels er nog zijn."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = hs.Pt(8)
    p.paragraph_format.space_after = hs.Pt(4)
    p.paragraph_format.keep_with_next = True
    hs.stijl(p.add_run("▶  " + tekst), hs.TITEL, 12, vet=True, kleur=hs.NAVY)
    for _ in range(2):
        leeg = doc.add_paragraph()
        leeg.paragraph_format.space_after = hs.Pt(4)
        leeg.paragraph_format.keep_with_next = True
        leeg.add_run(" ")


# ---------- de tekeningen van de planning van Sam ----------
MENLO = "/System/Library/Fonts/Menlo.ttc"
PLANNING_SAM = [
    ("Uur", "Activiteit", "Waar", "Wie"),
    ("12.00", "Boterhammen eten", "refter", "Lien"),
    ("12.45", "Vrij spel", "speelplaats", "Sam"),
    ("13.30", "Mijters knutselen", "atelier", "stagiair"),
    ("14.30", "Voorlezen", "leeshoek", "Lien"),
    ("16.00", "Buiten spelen", "tuin", "Sam"),
    ("17.00", "Ophalen", "refter", "Lien"),
]
KOLOM = (7, 19, 13)  # breedte in tekens van Uur, Activiteit en Waar; Wie is de laatste kolom


def sam_regels(met_voetbal):
    """De regels van Sam: kolommen met spaties. Met 'en voetbal' bijgetypt achter 'Vrij spel' schuift de rest van die regel
    elf tekens op (de spaties blijven staan). Geeft per regel een lijst van (tekst, markeren)."""
    regels = []
    for rij in PLANNING_SAM:
        uur, wat, waar, wie = rij
        regel = [(uur.ljust(KOLOM[0]), False)]
        if met_voetbal and wat == "Vrij spel":
            regel.append((wat, False))
            regel.append((" en voetbal", True))
            regel.append((" " * (KOLOM[1] - len(wat)) + waar.ljust(KOLOM[2]) + wie, False))
        else:
            regel.append((wat.ljust(KOLOM[1]) + waar.ljust(KOLOM[2]) + wie, False))
        regels.append(regel)
    return regels


def teken_sam(met_voetbal):
    """Een tekening van Sams planning in een tekstverwerker, in een monospace-lettertype (Menlo): elk teken even breed, dus
    de kolommen staan onder elkaar zolang niemand iets bijtypt. Geeft de png als bytes."""
    schaal = 2
    px = 26 * schaal
    lettertype = ImageFont.truetype(MENLO, px, index=0)
    vet = ImageFont.truetype(MENLO, px, index=1)
    breed = int(lettertype.getlength("M"))
    tekens = 60  # vaste breedte: beide tekeningen even groot
    w = tekens * breed + 2 * 30 * schaal
    regel_h = int(px * 1.55)
    kop_h = 46 * schaal
    h = kop_h + len(PLANNING_SAM) * regel_h + 30 * schaal
    im = Image.new("RGB", (w, h), "#FFFFFF")
    d = ImageDraw.Draw(im)
    d.rectangle([0, 0, w - 1, kop_h], fill="#E8EAED")
    d.text((30 * schaal, kop_h // 2), "Planning van Sam", font=ImageFont.truetype(MENLO, 20 * schaal, index=1),
           fill="#3C4043", anchor="lm")
    d.rectangle([0, 0, w - 1, h - 1], outline="#9AA0A6", width=2)
    y = kop_h + 18 * schaal
    for i, regel in enumerate(sam_regels(met_voetbal)):
        x = 30 * schaal
        for tekst, markeren in regel:
            breedte_tekst = breed * len(tekst)
            if markeren:
                d.rectangle([x, y - 4 * schaal, x + breedte_tekst, y + regel_h - 10 * schaal], fill="#FFE66D")
            d.text((x, y), tekst, font=vet if i == 0 else lettertype, fill="#1F1F1F")
            x += breedte_tekst
        y += regel_h
    buf = io.BytesIO()
    im.save(buf, "PNG")
    buf.seek(0)
    return buf


def planning_van_sam(doc):
    """Materiaal a van deel 3: de planning van Sam, twee keer als tekening (eerst, en nadat hij 'en voetbal' bijtypte)."""
    t = doc.add_table(rows=5, cols=1)
    t.alignment = hs.WD_TABLE_ALIGNMENT.CENTER
    hs.randen(t, buiten=GRIJS_RAND, binnen="E3E3E3")
    hs.breedtes(t, [B])
    kop = t.rows[0].cells[0]
    hs.schaduw(kop, "F2F6FC")
    hs.celtekst(kop, "a · De planning van Sam, met spaties gemaakt", vet=True, grootte=11)
    hs.hoogte(t.rows[0], 0.7)
    labels = ["Eerst, zoals Sam ze maakte.", "Daarna, nadat Sam “en voetbal” bijtypte."]
    for k, (rij_label, rij_beeld) in enumerate([(1, 2), (3, 4)]):
        hs.celtekst(t.rows[rij_label].cells[0], labels[k], vet=True, kleur=hs.NAVY_DIEP, grootte=11)
        cel = t.rows[rij_beeld].cells[0]
        hs.marges(cel, boven=0.15, onder=0.2, links=0.3, rechts=0.3)
        p = cel.paragraphs[0]
        hs.zonder_ruimte(p)
        p.add_run().add_picture(teken_sam(k == 1), width=hs.Cm(B - 1.0))
    hs.niet_splitsen(t)
    hs.ruimte(doc, 6)
    return t


# De criteria van lesvoorbereiding § 22 (samen 20): 3 · 6 · 2 · 1 · 1 · 2 · 1 · 2 · 1 · 1
CRITERIA = [
    ("Deel 1: een echte tabel met 4 kolommen en de koprij Uur · Activiteit · Waar · Wie.", 3),
    ("Deel 1: de zes activiteiten staan erin, met het juiste uur, de juiste plaats en de juiste naam, één ding per cel.", 6),
    ("Deel 2: het vieruurtje (15.15 · Fruit eten · refter · Sam) staat tussen 14.30 en 16.00.", 2),
    ("Deel 2: elk uur staat op dezelfde manier, zoals 15.15.", 1),
    ("Deel 1 en 2: de titel staat boven de tabel, in het vet.", 1),
    ("Deel 2: de hele koprij is vet en lichtgroen.", 2),
    ("Deel 2: er is geen lege rij in of onder de tabel.", 1),
    ("Deel 3: vraag 2 heeft de rij, de kolommen en het antwoord.", 2),
    ("Deel 3: vraag 3 zegt wat er bij Sam misging, en waarom dat in een tabel niet gebeurt.", 1),
    ("Deel 4: de exitvraag is beantwoord.", 1),
]


def werkdocument():
    """In de huisstijl van 19-11-2026 (huisstijl.py): per deel de opdracht, het infokader, het materiaal met letters en de
    vragen met cijfers. Wat de leerling maakt, en alles wat hij daarvoor nodig heeft. Waar hij klikt, staat op de lespagina."""
    doc = hs.nieuw_document()
    doc.core_properties.title = "Dagplanning De Eekhoorns - werkdocument"
    doc.core_properties.author = "Toegepaste Informatica"
    hs.kopblok(doc, "DAGPLANNING DE EEKHOORNS", punten=20)

    hs.kader(doc, "Stage in De Speelboom", [
        ("Jouw rol:", "stagiair in groep De Eekhoorns (6 tot 9 jaar) van de buitenschoolse opvang De Speelboom."),
        ("Je mentor:", "Lien Verhoeven."),
    ])

    # ---- Deel 1: van briefje naar tabel (de kern)
    hs.deel(doc, 1, "Van briefje naar tabel")
    hs.opdracht(doc, "Typ de titel. Zet het briefje van Lien eronder in een tabel. Gebruik het infokader.")
    hs.infokader(doc, "Een tabel", [
        ("Rij:", "de vakjes naast elkaar. Een rij ligt van links naar rechts."),
        ("Kolom:", "de vakjes onder elkaar. Een kolom staat van boven naar onder."),
        ("Cel:", "één vakje."),
        ("Koprij:", "de eerste rij. Ze zegt wat er in elke kolom staat."),
        ("In elke cel:", "één ding: een uur, een activiteit, een plaats of een naam."),
    ])
    briefje(doc, "a · Het briefje van Lien", BRIEFJE)
    materiaal_regel(doc, "b · De titel om over te nemen", TITEL_PLANNING)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = hs.Pt(4)
    p.paragraph_format.space_after = hs.Pt(4)
    p.paragraph_format.keep_with_next = True
    p.add_run("c · Zo wordt zin 1 een rij").bold = True
    hs.tabel(doc, ["Uur", "Activiteit", "Waar", "Wie"], [
        ("12.00", "Boterhammen eten", "refter", "Lien"),
    ], [3.0, 6.0, 4.0, B - 13.0], voorbeeld_rij=1)
    hs.tekst(doc, "Je tabel heeft 4 kolommen en 7 rijen: de koprij en zes activiteiten.", klein=True)
    plaats_voor_tabel(doc, "Maak hier je dagplanning")

    # ---- Deel 2: een rij erbij en opmaken
    hs.deel(doc, 2, "Een rij erbij en opmaken")
    hs.opdracht(doc, "Zet het vieruurtje op de juiste plaats. Maak je dagplanning op zoals de tabelkaart.")
    hs.infokader(doc, "Tabelkaart De Speelboom", [
        ("Titel:", "boven de tabel, in het vet."),
        ("Koprij:", "vet en lichtgroen."),
        ("Uur:", "altijd zo: 15.15."),
        ("In elke cel:", "één ding."),
        ("Lege rijen:", "geen."),
    ])
    ontvangen(doc, "a · Een mail van Lien", [
        ("Van", "Lien Verhoeven <lien.verhoeven@speelboom.example>"),
        ("Onderwerp", "Het vieruurtje"),
    ], ["Oei, ik vergat het vieruurtje! Om kwart over drie eten de kinderen fruit in de refter, met Sam. Zet je het op de "
        "juiste plaats? En maak de planning op zoals de tabelkaart: ik hang ze maandag aan de muur. Groetjes, Lien"])

    # ---- Deel 3: waarom een tabel? (geen klikken)
    hs.deel(doc, 3, "Waarom een tabel?")
    hs.opdracht(doc, "Gebruik je eigen dagplanning. Beantwoord vraag 2 en 3.")
    planning_van_sam(doc)
    hs.vraag(doc, "1.  Een papa belt om 12.50 uur: “Waar is mijn dochter nu?”")
    hs.tabel(doc, ["Rij", "Kolom", "Antwoord"], [("12.45", "Waar", "op de speelplaats")],
             [3.4, 3.4, B - 6.8], invul=(0, 1, 2), voorbeeld_rij=1)
    hs.vraag(doc, "2.  Een mama belt om 15.20 uur: “Waar is mijn zoon nu, en wie is bij hem?” Zoek het in je tabel. "
                  "Schrijf de rij, de kolommen en het antwoord.")
    hs.tabel(doc, ["Rij", "Kolommen", "Antwoord"], [("", "", "")], [3.4, 5.0, B - 8.4], invul=(0, 1, 2), rijhoogte=1.2)
    hs.vraag(doc, "3.  Kijk naar de planning van Sam (a). Wat ging er mis toen hij een woord bijtypte? Waarom gebeurt dat "
                  "niet in jouw tabel? Schrijf twee zinnen.")
    hs.antwoordvak(doc, 2.2)

    # ---- Deel 4: klaar
    hs.deel(doc, 4, "Klaar?", nieuwe_pagina=True)
    hs.opdracht(doc, "Typ het codewoord. Kruis aan wat klopt. Beantwoord de exitvraag.")
    hs.vraag(doc, "Het codewoord van de zelftest (stap 4 op de lespagina):")
    hs.antwoordvak(doc, 0.9)
    hs.controletabel(doc, CRITERIA)
    hs.vraag(doc, "Exitvraag: Lien wil ook noteren welk materiaal je nodig hebt (lijm, schaar, boeken …). Maak je daarvoor "
                  "een nieuwe rij of een nieuwe kolom? Waarom? Eén zin.")
    hs.antwoordvak(doc, 1.6)

    # ---- Extra
    hs.deel(doc, None, "Extra — niet verplicht", nieuwe_pagina=True)
    hs.opdracht(doc, "Maak van het tweede briefje van Lien een tweede dagplanning. Volg dezelfde tabelkaart.")
    briefje(doc, "a · Het tweede briefje van Lien: vrijdag 4 december, het Sinterklaasfeest", BRIEFJE_EXTRA)
    plaats_voor_tabel(doc, "Maak hier je tweede dagplanning")

    doc.save(UIT)


if __name__ == "__main__":
    werkdocument()
    print("Gemaakt: Dagplanning-Eekhoorns.docx")
