
import io
import zipfile
from typing import Dict, List

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation


STATUS_OPTIONS = [
    "Dodáno",
    "Částečně dodáno",
    "Chybí",
    "Není k dispozici",
    "Irelevantní",
]

RELEVANCE_EXPLANATION = {
    "P": "🔴 Povinné",
    "D": "🟡 Doporučené",
    "V": "🟢 Volitelné",
    "-": "⚪ Irelevantní",
}


BLOCK_DEFINITIONS = {
    1: "Zadání a organizace",
    2: "Fotodokumentace",
    3: "Energetická data",
    4: "Stavební a objektové technické podklady",
    5: "Technické podklady technologií a energetických zařízení",
    6: "Provozní informace",
    7: "Ekonomika, finance a dotační souvislosti",
}


BLOCK_README_TEXTS = {
    1: (
        "Do této části patří zejména smluvní a zadávací dokumenty, přehled řešených objektů, "
        "interní schvalovací proces, pravidla veřejných zakázek a harmonogram projektu."
    ),
    2: (
        "Do této části patří fotodokumentace objektů, technických místností, střech, technologií "
        "a problematických míst."
    ),
    3: (
        "Do této části patří energetická data, zejména faktury, smlouvy, průběhová data, odečty "
        "a přehledy odběrných míst."
    ),
    4: (
        "Do této části patří stavební a objektové technické podklady, například dokumentace stávajícího "
        "a plánovaného stavu, PENB, obvodový plášť, střechy a statické posudky."
    ),
    5: (
        "Do této části patří technické podklady technologií a energetických zařízení, například vytápění, "
        "chlazení, vzduchotechnika, příprava teplé vody, FVE, trafostanice, MaR / EnMS a osvětlení."
    ),
    6: (
        "Do této části patří informace o reálném provozu objektu, jeho obsazenosti, provozních problémech "
        "a plánovaných změnách."
    ),
    7: (
        "Do této části patří informace o investicích, rozpočtových možnostech a případných dotacích "
        "nebo podpůrných programech."
    ),
}


MASTER_ITEMS: List[Dict[str, str | int]] = [
    {
        "id": "1",
        "block": 1,
        "item_no": 1,
        "name": "Smluvní a zadávací dokumenty",
        "description": "Smlouva, objednávka, nabídka, zápis z úvodního jednání a další dokumenty, ze kterých je patrné zadání služby, očekávaný výstup, rozsah a cíle.",
    },
    {
        "id": "2",
        "block": 1,
        "item_no": 2,
        "name": "Přehled řešených objektů / lokalit",
        "description": "Vyplněný soubor PREHLED_OBJEKTU.xlsx se seznamem řešených objektů nebo lokalit ve schválené struktuře.",
    },
    {
        "id": "3",
        "block": 1,
        "item_no": 3,
        "name": "Schvalovací proces a interní termíny",
        "description": "Termíny rady, zastupitelstva, interní schvalovací kroky, kompetence jednotlivých orgánů a vazba rozhodování na projekt.",
    },
    {
        "id": "4",
        "block": 1,
        "item_no": 4,
        "name": "Pravidla veřejných zakázek",
        "description": "Směrnice o zadávání veřejných zakázek, limity pro přímý nákup, poptávku, veřejné výběrové řízení a související kompetence.",
    },
    {
        "id": "5",
        "block": 1,
        "item_no": 5,
        "name": "Harmonogram projektu a klíčová rozhodnutí",
        "description": "Harmonogram zakázky, návazné termíny a zásadní rozhodnutí v průběhu projektu.",
    },
    {
        "id": "6",
        "block": 2,
        "item_no": 1,
        "name": "Fotodokumentace",
        "description": "Fotografie objektů, technických místností, střech, technologií, vad a problematických míst.",
    },
    {
        "id": "7",
        "block": 3,
        "item_no": 1,
        "name": "Faktury za elektřinu",
        "description": "Faktury za dodávku a distribuci elektřiny za jednotlivá odběrná místa nebo objekty, ideálně za posledních 24–36 měsíců.",
    },
    {
        "id": "8",
        "block": 3,
        "item_no": 2,
        "name": "Smlouvy a ceníky elektřiny",
        "description": "Smlouvy s dodavatelem elektřiny, ceníky, dodatky a informace o produktu nebo cenovém mechanismu.",
    },
    {
        "id": "9",
        "block": 3,
        "item_no": 3,
        "name": "15min profily elektřiny",
        "description": "Průběhová data spotřeby nebo výroby elektřiny v 15min kroku, zejména pro hlavní odběry a výrobny.",
    },
    {
        "id": "10",
        "block": 3,
        "item_no": 4,
        "name": "Seznam EAN a parametry odběrných míst elektřiny",
        "description": "Seznam EAN, zatřídění VN / NN, distribuční sazba, velikost jištění, rezervovaný příkon nebo kapacita a další relevantní parametry odběrných míst.",
    },
    {
        "id": "11",
        "block": 3,
        "item_no": 5,
        "name": "Faktury za zemní plyn",
        "description": "Faktury za dodávku a distribuci zemního plynu, ideálně za posledních 24–36 měsíců.",
    },
    {
        "id": "12",
        "block": 3,
        "item_no": 6,
        "name": "Smlouvy na dodávku plynu",
        "description": "Smlouvy s dodavatelem plynu, dodatky a související ceníkové nebo produktové informace.",
    },
    {
        "id": "13",
        "block": 3,
        "item_no": 7,
        "name": "Odečty plynu / distribuční portál",
        "description": "Odečty plynu, exporty z distribučního portálu nebo jiná dostupná provozní data k odběru plynu.",
    },
    {
        "id": "14",
        "block": 3,
        "item_no": 8,
        "name": "Faktury za teplo",
        "description": "Faktury za teplo nebo CZT, ideálně za posledních 24–36 měsíců.",
    },
    {
        "id": "15",
        "block": 3,
        "item_no": 9,
        "name": "Smlouvy na dodávku tepla",
        "description": "Smlouvy o dodávce tepla nebo CZT a související smluvní podmínky.",
    },
    {
        "id": "16",
        "block": 3,
        "item_no": 10,
        "name": "Odečty tepla / distribuční portál",
        "description": "Odečty tepla, exporty z portálu dodavatele nebo jiná dostupná data k odběru tepla.",
    },
    {
        "id": "17",
        "block": 3,
        "item_no": 11,
        "name": "Faktury za vodu",
        "description": "Faktury za vodu a stočné za relevantní objekty.",
    },
    {
        "id": "18",
        "block": 3,
        "item_no": 12,
        "name": "Mapování odběrných míst k objektům",
        "description": "Přehledová tabulka, ve které je každé odběrné místo jednoznačně přiřazeno ke konkrétnímu objektu.",
    },
    {
        "id": "19",
        "block": 3,
        "item_no": 13,
        "name": "Vlastní výroba elektřiny – hodinová nebo průběhová data",
        "description": "Data o vlastní výrobě elektřiny, zejména z FVE nebo jiných vlastních zdrojů, ideálně v hodinovém nebo jiném průběhovém kroku.",
    },
    {
        "id": "20",
        "block": 4,
        "item_no": 1,
        "name": "Dokumentace stávajícího stavu objektů",
        "description": "Pasporty, projektová dokumentace skutečného stavu, technické zprávy, půdorysy, řezy, pohledy a další podklady ke stávajícímu stavu objektu.",
    },
    {
        "id": "21",
        "block": 4,
        "item_no": 2,
        "name": "Projektová dokumentace plánovaného stavu",
        "description": "Studie, projektová dokumentace připravovaných úprav, návrhy rekonstrukcí a další podklady k plánovanému stavu.",
    },
    {
        "id": "22",
        "block": 4,
        "item_no": 3,
        "name": "PENB, energetické audity a další energetické dokumenty",
        "description": "PENB, starší energetické audity, energetické posudky a další obdobné energetické dokumenty.",
    },
    {
        "id": "23",
        "block": 4,
        "item_no": 4,
        "name": "Obvodový plášť a zateplení",
        "description": "Podklady k fasádě, zateplení, skladbám konstrukcí a realizovaným úpravám obvodového pláště.",
    },
    {
        "id": "24",
        "block": 4,
        "item_no": 5,
        "name": "Výplně otvorů",
        "description": "Podklady k oknům, dveřím, vratům a jejich základním parametrům, stáří nebo rozsahu výměn.",
    },
    {
        "id": "25",
        "block": 4,
        "item_no": 6,
        "name": "Střechy a střešní konstrukce",
        "description": "Podklady k typu střechy, skladbě, stavu, ploše a omezením případných zásahů.",
    },
    {
        "id": "26",
        "block": 4,
        "item_no": 7,
        "name": "Statické posouzení střechy",
        "description": "Statické posudky a další podklady důležité pro zásahy do střechy, zejména pro FVE nebo jiné dodatečné zatížení.",
    },
    {
        "id": "27",
        "block": 4,
        "item_no": 8,
        "name": "Stavební stav a poruchy",
        "description": "Informace o vadách, poruchách, zatékání, vlhkosti nebo jiných známých problémech stavební části objektu.",
    },
    {
        "id": "28",
        "block": 5,
        "item_no": 1,
        "name": "Vytápění",
        "description": "Zdroje tepla, rozvody, schémata, regulace, servis, revize a podklady k předávací stanici tepla.",
    },
    {
        "id": "29",
        "block": 5,
        "item_no": 2,
        "name": "Chlazení",
        "description": "Zdroje chladu, chladicí zařízení, technické listy, servis, revize a provozní problémy chlazení.",
    },
    {
        "id": "30",
        "block": 5,
        "item_no": 3,
        "name": "Vzduchotechnika",
        "description": "Vzduchotechnické jednotky, technické listy, servis, revize a provozní problémy VZT.",
    },
    {
        "id": "31",
        "block": 5,
        "item_no": 4,
        "name": "Příprava teplé vody",
        "description": "Podklady ke zdrojům TUV, zásobníkům, cirkulaci, měření a provozním režimům přípravy teplé vody.",
    },
    {
        "id": "32",
        "block": 5,
        "item_no": 5,
        "name": "Záložní zdroje",
        "description": "Dieselagregáty, UPS nebo jiné záložní zdroje včetně základních parametrů a provozních informací.",
    },
    {
        "id": "33",
        "block": 5,
        "item_no": 6,
        "name": "FVE a vlastní zdroje elektřiny",
        "description": "Technické podklady k FVE a dalším vlastním zdrojům elektřiny, včetně smlouvy o připojení výrobny a jednopólového schématu, pokud existují.",
    },
    {
        "id": "34",
        "block": 5,
        "item_no": 7,
        "name": "Trafostanice",
        "description": "Technické podklady, revize a základní parametry trafostanic a souvisejících zařízení.",
    },
    {
        "id": "35",
        "block": 5,
        "item_no": 8,
        "name": "MaR / EnMS",
        "description": "Podklady k měření a regulaci nebo energetickému managementu, zejména popis systému, exporty, trendy, seznam měřených veličin a přístupy.",
    },
    {
        "id": "36",
        "block": 5,
        "item_no": 9,
        "name": "Voda",
        "description": "Technický stav rozvodů vody, úniky, poruchy a další technické podklady k vodnímu hospodářství objektu.",
    },
    {
        "id": "37",
        "block": 5,
        "item_no": 10,
        "name": "Osvětlení",
        "description": "Přehled svítidel a základní podklady k systému osvětlení objektu.",
    },
    {
        "id": "38",
        "block": 5,
        "item_no": 11,
        "name": "PBŘ",
        "description": "Požárně-bezpečnostní řešení objektu a související omezení pro technologie nebo stavební zásahy.",
    },
    {
        "id": "39",
        "block": 5,
        "item_no": 12,
        "name": "Významné spotřebiče",
        "description": "Technologické celky nebo zařízení s významnou spotřebou energie, důležité pro energetickou bilanci objektu.",
    },
    {
        "id": "40",
        "block": 6,
        "item_no": 1,
        "name": "Provozní režim objektu",
        "description": "Provozní doby objektu, pracovní dny, víkendy, směnnost, sezónnost a základní režim fungování objektu.",
    },
    {
        "id": "41",
        "block": 6,
        "item_no": 2,
        "name": "Obsazenost a způsob využití",
        "description": "Počet uživatelů, způsob využití objektu, intenzita provozu a případné sezónní výkyvy.",
    },
    {
        "id": "42",
        "block": 6,
        "item_no": 3,
        "name": "Uživatelské a provozní problémy",
        "description": "Stížnosti uživatelů a opakované provozní problémy, například chlad, přehřívání, hluk, zápach, vlhkost nebo průvan.",
    },
    {
        "id": "43",
        "block": 6,
        "item_no": 4,
        "name": "Plánované změny a rekonstrukce",
        "description": "Připravované stavební úpravy, změny využití, stěhování nebo další plánované zásahy s dopadem na objekt.",
    },
    {
        "id": "44",
        "block": 6,
        "item_no": 5,
        "name": "Doplňující provozní informace",
        "description": "Další provozní informace, které mají význam pro pochopení skutečného fungování objektu nebo areálu.",
    },
    {
        "id": "45",
        "block": 7,
        "item_no": 1,
        "name": "Historie investic",
        "description": "Přehled významných investic do objektu, technologií nebo energetických opatření v posledních letech.",
    },
    {
        "id": "46",
        "block": 7,
        "item_no": 2,
        "name": "Rozpočtové možnosti a finanční rámec",
        "description": "Rozpočtové možnosti klienta, investiční rámec, CAPEX / OPEX limity a další finanční omezení.",
    },
    {
        "id": "47",
        "block": 7,
        "item_no": 3,
        "name": "Dotace a podpůrné programy",
        "description": "Relevantní dotační programy, podmínky podpory a případně rozpracované nebo podané žádosti.",
    },
]


SERVICES = {
    "Energetický audit": {"short": "EA"},
    "Plán energetického auditu": {"short": "PEA"},
    "Studie potenciálu energetických úspor": {"short": "SPU"},
    "Energetický posudek": {"short": "EP"},
    "Průkaz energetické náročnosti budovy": {"short": "PENB"},
    "Energetická studie": {"short": "ES"},
    "Technickoekonomická studie FVE a bateriového úložiště": {"short": "FVE"},
    "Studie komunitní energetiky": {"short": "KE"},
    "Energetický management": {"short": "EM"},
    "Místní energetická koncepce": {"short": "MEK"},
    "Studie proveditelnosti": {"short": "SP"},
    "Správce procesu přípravy projektu Design&Build / EPC / D&B": {"short": "DB"},
}


RELEVANCE_MATRIX: Dict[str, Dict[str, str]] = {
    "1": {"EA":"P","PEA":"P","SPU":"P","EP":"P","PENB":"D","ES":"P","FVE":"P","KE":"P","EM":"P","MEK":"P","SP":"P","DB":"P"},
    "2": {"EA":"P","PEA":"P","SPU":"P","EP":"P","PENB":"P","ES":"P","FVE":"P","KE":"P","EM":"D","MEK":"P","SP":"P","DB":"D"},
    "3": {"EA":"D","PEA":"P","SPU":"D","EP":"V","PENB":"-","ES":"D","FVE":"D","KE":"D","EM":"D","MEK":"D","SP":"D","DB":"P"},
    "4": {"EA":"V","PEA":"D","SPU":"V","EP":"-","PENB":"-","ES":"D","FVE":"D","KE":"V","EM":"V","MEK":"V","SP":"D","DB":"P"},
    "5": {"EA":"D","PEA":"P","SPU":"D","EP":"V","PENB":"-","ES":"D","FVE":"D","KE":"D","EM":"D","MEK":"D","SP":"P","DB":"P"},
    "6": {"EA":"V","PEA":"V","SPU":"V","EP":"V","PENB":"P","ES":"V","FVE":"P","KE":"V","EM":"V","MEK":"V","SP":"V","DB":"V"},
    "7": {"EA":"P","PEA":"P","SPU":"P","EP":"P","PENB":"-","ES":"P","FVE":"P","KE":"P","EM":"P","MEK":"D","SP":"D","DB":"-"},
    "8": {"EA":"D","PEA":"D","SPU":"D","EP":"D","PENB":"-","ES":"P","FVE":"P","KE":"P","EM":"D","MEK":"D","SP":"D","DB":"-"},
    "9": {"EA":"D","PEA":"D","SPU":"D","EP":"D","PENB":"-","ES":"P","FVE":"P","KE":"P","EM":"D","MEK":"V","SP":"V","DB":"-"},
    "10": {"EA":"P","PEA":"P","SPU":"P","EP":"P","PENB":"-","ES":"P","FVE":"P","KE":"P","EM":"P","MEK":"D","SP":"D","DB":"-"},
    "11": {"EA":"P","PEA":"P","SPU":"P","EP":"D","PENB":"-","ES":"P","FVE":"-","KE":"-","EM":"D","MEK":"D","SP":"D","DB":"-"},
    "12": {"EA":"D","PEA":"D","SPU":"D","EP":"D","PENB":"-","ES":"P","FVE":"-","KE":"-","EM":"D","MEK":"D","SP":"D","DB":"-"},
    "13": {"EA":"D","PEA":"D","SPU":"D","EP":"D","PENB":"-","ES":"D","FVE":"-","KE":"-","EM":"D","MEK":"V","SP":"V","DB":"-"},
    "14": {"EA":"P","PEA":"P","SPU":"P","EP":"D","PENB":"-","ES":"P","FVE":"-","KE":"-","EM":"D","MEK":"D","SP":"D","DB":"-"},
    "15": {"EA":"D","PEA":"D","SPU":"D","EP":"D","PENB":"-","ES":"P","FVE":"-","KE":"-","EM":"D","MEK":"D","SP":"D","DB":"-"},
    "16": {"EA":"D","PEA":"D","SPU":"D","EP":"D","PENB":"-","ES":"D","FVE":"-","KE":"-","EM":"D","MEK":"V","SP":"V","DB":"-"},
    "17": {"EA":"D","PEA":"D","SPU":"D","EP":"-","PENB":"-","ES":"D","FVE":"-","KE":"-","EM":"D","MEK":"V","SP":"V","DB":"-"},
    "18": {"EA":"P","PEA":"P","SPU":"P","EP":"P","PENB":"-","ES":"P","FVE":"P","KE":"P","EM":"P","MEK":"D","SP":"D","DB":"-"},
    "19": {"EA":"V","PEA":"V","SPU":"V","EP":"V","PENB":"-","ES":"D","FVE":"P","KE":"P","EM":"D","MEK":"D","SP":"D","DB":"-"},
    "20": {"EA":"P","PEA":"P","SPU":"P","EP":"D","PENB":"P","ES":"P","FVE":"D","KE":"D","EM":"V","MEK":"D","SP":"D","DB":"D"},
    "21": {"EA":"D","PEA":"P","SPU":"D","EP":"D","PENB":"-","ES":"P","FVE":"P","KE":"D","EM":"V","MEK":"D","SP":"P","DB":"P"},
    "22": {"EA":"D","PEA":"P","SPU":"P","EP":"D","PENB":"P","ES":"P","FVE":"D","KE":"D","EM":"D","MEK":"D","SP":"D","DB":"V"},
    "23": {"EA":"P","PEA":"P","SPU":"P","EP":"D","PENB":"P","ES":"P","FVE":"V","KE":"V","EM":"-","MEK":"D","SP":"D","DB":"-"},
    "24": {"EA":"P","PEA":"P","SPU":"P","EP":"D","PENB":"P","ES":"P","FVE":"V","KE":"V","EM":"-","MEK":"D","SP":"D","DB":"-"},
    "25": {"EA":"P","PEA":"P","SPU":"P","EP":"D","PENB":"P","ES":"P","FVE":"P","KE":"D","EM":"-","MEK":"D","SP":"D","DB":"-"},
    "26": {"EA":"V","PEA":"V","SPU":"V","EP":"-","PENB":"-","ES":"D","FVE":"P","KE":"D","EM":"-","MEK":"V","SP":"D","DB":"-"},
    "27": {"EA":"P","PEA":"P","SPU":"P","EP":"D","PENB":"D","ES":"P","FVE":"D","KE":"D","EM":"V","MEK":"D","SP":"D","DB":"D"},
    "28": {"EA":"P","PEA":"P","SPU":"P","EP":"P","PENB":"D","ES":"P","FVE":"-","KE":"-","EM":"D","MEK":"D","SP":"D","DB":"D"},
    "29": {"EA":"D","PEA":"D","SPU":"D","EP":"D","PENB":"V","ES":"P","FVE":"-","KE":"-","EM":"V","MEK":"V","SP":"D","DB":"D"},
    "30": {"EA":"P","PEA":"P","SPU":"P","EP":"D","PENB":"D","ES":"P","FVE":"-","KE":"-","EM":"V","MEK":"V","SP":"D","DB":"D"},
    "31": {"EA":"D","PEA":"D","SPU":"D","EP":"D","PENB":"D","ES":"D","FVE":"-","KE":"-","EM":"V","MEK":"V","SP":"D","DB":"-"},
    "32": {"EA":"V","PEA":"V","SPU":"V","EP":"V","PENB":"-","ES":"D","FVE":"D","KE":"D","EM":"V","MEK":"V","SP":"D","DB":"D"},
    "33": {"EA":"V","PEA":"V","SPU":"V","EP":"V","PENB":"-","ES":"D","FVE":"P","KE":"P","EM":"D","MEK":"D","SP":"D","DB":"D"},
    "34": {"EA":"V","PEA":"V","SPU":"V","EP":"V","PENB":"-","ES":"D","FVE":"P","KE":"P","EM":"V","MEK":"V","SP":"D","DB":"D"},
    "35": {"EA":"D","PEA":"D","SPU":"D","EP":"D","PENB":"-","ES":"P","FVE":"D","KE":"D","EM":"P","MEK":"D","SP":"D","DB":"D"},
    "36": {"EA":"V","PEA":"V","SPU":"V","EP":"-","PENB":"-","ES":"D","FVE":"-","KE":"-","EM":"D","MEK":"V","SP":"V","DB":"-"},
    "37": {"EA":"D","PEA":"D","SPU":"D","EP":"-","PENB":"-","ES":"D","FVE":"-","KE":"-","EM":"V","MEK":"V","SP":"D","DB":"-"},
    "38": {"EA":"V","PEA":"V","SPU":"V","EP":"V","PENB":"-","ES":"D","FVE":"P","KE":"D","EM":"-","MEK":"-","SP":"D","DB":"D"},
    "39": {"EA":"D","PEA":"D","SPU":"D","EP":"D","PENB":"-","ES":"P","FVE":"P","KE":"P","EM":"D","MEK":"D","SP":"D","DB":"D"},
    "40": {"EA":"P","PEA":"P","SPU":"P","EP":"D","PENB":"D","ES":"P","FVE":"D","KE":"P","EM":"P","MEK":"D","SP":"D","DB":"V"},
    "41": {"EA":"P","PEA":"P","SPU":"P","EP":"D","PENB":"P","ES":"P","FVE":"D","KE":"P","EM":"P","MEK":"D","SP":"D","DB":"V"},
    "42": {"EA":"D","PEA":"D","SPU":"D","EP":"D","PENB":"D","ES":"P","FVE":"V","KE":"D","EM":"D","MEK":"V","SP":"D","DB":"V"},
    "43": {"EA":"D","PEA":"P","SPU":"D","EP":"D","PENB":"-","ES":"P","FVE":"D","KE":"D","EM":"D","MEK":"D","SP":"P","DB":"P"},
    "44": {"EA":"V","PEA":"V","SPU":"V","EP":"V","PENB":"-","ES":"D","FVE":"V","KE":"D","EM":"D","MEK":"V","SP":"D","DB":"V"},
    "45": {"EA":"D","PEA":"P","SPU":"P","EP":"D","PENB":"-","ES":"P","FVE":"D","KE":"D","EM":"D","MEK":"D","SP":"P","DB":"P"},
    "46": {"EA":"D","PEA":"P","SPU":"P","EP":"D","PENB":"-","ES":"P","FVE":"P","KE":"D","EM":"D","MEK":"D","SP":"P","DB":"P"},
    "47": {"EA":"V","PEA":"D","SPU":"D","EP":"V","PENB":"-","ES":"D","FVE":"D","KE":"D","EM":"D","MEK":"D","SP":"P","DB":"D"},
}


ROOT_README_TEXT = """NÁZEV: Hlavní složka projektu – podklady pro službu DPU ENERGY

K ČEMU TATO SLOŽKA SLOUŽÍ:
Tato složka slouží jako centrální úložiště podkladů pro zpracování vybrané služby.

JAK SLOŽKU POUŽÍVAT:
- Soubory nahrávejte do tematicky odpovídajících hlavních složek 01 až 07.
- Pokud si nejste jistí, kam dokument patří, uložte jej do složky 99_ARCHIV_NEZARAZENO.
- Názvy hlavních složek prosím neměňte.
- Při pojmenování souborů používejte pravidla uvedená v souboru KLIC_POJMENOVANI_SOUBORU.txt.
"""

NAMING_KEY_TEXT = """KLÍČ POJMENOVÁNÍ SOUBORŮ

OBECNÝ FORMÁT:
[OBLAST]_[OBJEKT nebo ODBĚR]_[POPIS]_[OBDOBÍ nebo DATUM]_[VERZE]

DOPORUČENÍ:
- používat pouze písmena bez diakritiky, čísla a podtržítka
- nepoužívat mezery
- datum uvádět ve formátu YYYY-MM nebo YYYY-MM-DD
- verzi uvádět například v01, v02
- pokud není znám objekt, použijte identifikátor odběrného místa nebo obecný popis
"""

BLOCK_FOLDERS = {
    1: "01_ZADANI_A_ORGANIZACE",
    2: "02_FOTODOKUMENTACE",
    3: "03_ENERGETICKA_DATA",
    4: "04_STAVEBNI_A_OBJEKTOVE_TECHNICKE_PODKLADY",
    5: "05_TECHNICKE_PODKLADY_TECHNOLOGII_A_ENERGETICKYCH_ZARIZENI",
    6: "06_PROVOZNI_INFORMACE",
    7: "07_EKONOMIKA_FINANCE_A_DOTACE",
}

BLOCK_READMES = {
    1: """NÁZEV: 01_ZADANI_A_ORGANIZACE

K ČEMU SLOŽKA SLOUŽÍ:
Do této složky patří dokumenty, ze kterých je patrné zadání služby, rozsah řešených objektů, interní schvalovací proces, pravidla veřejných zakázek a harmonogram projektu.

CO TYPICKY NAHRÁT:
- smlouvu, objednávku nebo nabídku
- zápis z úvodního jednání
- přehled řešených objektů / lokalit
- termíny rady a zastupitelstva, pokud mají vazbu na projekt
- interní pravidla veřejných zakázek
- harmonogram projektu

POZNÁMKA:
Kontaktní osoby k jednotlivým objektům vyplňte primárně do souboru PREHLED_OBJEKTU.xlsx.
""",
    2: """NÁZEV: 02_FOTODOKUMENTACE

K ČEMU SLOŽKA SLOUŽÍ:
Do této složky patří fotodokumentace objektů, technických místností, střech, technologií a problematických míst.

CO TYPICKY NAHRÁT:
- fotografie exteriéru objektu
- fotografie technických místností
- fotografie střech
- fotografie vad, poruch nebo problematických detailů
- fotografie technologií
""",
    3: """NÁZEV: 03_ENERGETICKA_DATA

K ČEMU SLOŽKA SLOUŽÍ:
Do této složky patří energetická data, zejména faktury, smlouvy, průběhová data, odečty a přehledy odběrných míst.

CO TYPICKY NAHRÁT:
- faktury za elektřinu, zemní plyn, teplo a vodu
- smlouvy a ceníky elektřiny
- smlouvy na dodávku plynu nebo tepla
- 15min profily elektřiny
- odečty z distribučních portálů
- seznam EAN a parametry odběrných míst
- mapování odběrných míst k objektům
- data o vlastní výrobě elektřiny
""",
    4: """NÁZEV: 04_STAVEBNI_A_OBJEKTOVE_TECHNICKE_PODKLADY

K ČEMU SLOŽKA SLOUŽÍ:
Do této složky patří stavební a objektové technické podklady ke stávajícímu a případně i plánovanému stavu objektu.

CO TYPICKY NAHRÁT:
- pasporty objektů
- projektovou dokumentaci stávajícího stavu
- projektovou dokumentaci plánovaného stavu
- PENB, energetické audity a další energetické dokumenty
- podklady k obvodovému plášti, zateplení a výplním otvorů
- podklady ke střechám a střešním konstrukcím
- statické posouzení střechy
- informace o stavebním stavu a známých poruchách
""",
    5: """NÁZEV: 05_TECHNICKE_PODKLADY_TECHNOLOGII_A_ENERGETICKYCH_ZARIZENI

K ČEMU SLOŽKA SLOUŽÍ:
Do této složky patří technické podklady k vytápění, chlazení, vzduchotechnice, přípravě teplé vody, FVE, trafostanicím, MaR / EnMS a dalším technologiím.

CO TYPICKY NAHRÁT:
- podklady k vytápění a zdrojům tepla
- podklady k chlazení
- podklady ke vzduchotechnice
- podklady k přípravě teplé vody
- podklady k záložním zdrojům
- podklady k FVE a vlastním zdrojům elektřiny
- podklady k trafostanicím
- podklady k MaR / EnMS
- technické podklady k vodě, osvětlení nebo PBŘ
- přehled významných spotřebičů
""",
    6: """NÁZEV: 06_PROVOZNI_INFORMACE

K ČEMU SLOŽKA SLOUŽÍ:
Do této složky patří informace o reálném provozu objektu, jeho obsazenosti, uživatelských problémech a plánovaných změnách.

CO TYPICKY NAHRÁT:
- informace o provozních dobách a režimu objektu
- informace o obsazenosti a způsobu využití
- přehled uživatelských nebo provozních problémů
- informace o plánovaných změnách a rekonstrukcích
- další doplňující provozní informace
""",
    7: """NÁZEV: 07_EKONOMIKA_FINANCE_A_DOTACE

K ČEMU SLOŽKA SLOUŽÍ:
Do této složky patří informace o historii investic, finančních možnostech klienta a případných dotacích nebo podpůrných programech.

CO TYPICKY NAHRÁT:
- přehled minulých investic
- informace o rozpočtových možnostech a finančním rámci
- informace o relevantních dotačních programech
- rozpracované nebo podané dotační žádosti, pokud existují
""",
}


def format_readme_for_ui(text: str) -> str:
    return text.replace("\r\n", "\n").strip()

def parse_query_params() -> Dict[str, str]:
    params = st.query_params
    result = {}
    for key in ["customerName", "projectCode", "projectName"]:
        value = params.get(key, "")
        if isinstance(value, list):
            result[key] = value[0] if value else ""
        else:
            result[key] = value
    return result


def safe_name(text: str) -> str:
    bad = '<>:"/\\|?*'
    out = "".join("_" if ch in bad else ch for ch in str(text).strip())
    out = out.replace(" ", "_")
    return out or "PROJEKT"


def dataframe_to_excel_bytes(
    df: pd.DataFrame,
    sheet_name: str,
    dropdowns: Dict[str, List[str]] | None = None,
    freeze_cell: str = "A2",
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    for c_idx, col in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E2F3")
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")

    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    ws.freeze_panes = freeze_cell
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 30

    for col_cells in ws.columns:
        max_len = 0
        col_letter = col_cells[0].column_letter
        for cell in col_cells:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 45)

    if dropdowns:
        header_map = {cell.value: cell.column for cell in ws[1]}
        for column_name, options in dropdowns.items():
            if column_name not in header_map:
                continue
            col_idx = header_map[column_name]
            col_letter = ws.cell(row=1, column=col_idx).column_letter
            dv = DataValidation(
                type="list",
                formula1=f'"{",".join(options)}"',
                allow_blank=True,
            )
            dv.promptTitle = "Vyberte hodnotu"
            dv.prompt = "Použijte jednu z nabízených možností."
            dv.errorTitle = "Neplatná hodnota"
            dv.error = "Použijte jednu z povolených hodnot."
            ws.add_data_validation(dv)
            dv.add(f"{col_letter}2:{col_letter}500")
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def build_service_dataframe(service_name: str) -> pd.DataFrame:
    service_short = SERVICES[service_name]["short"]
    rows = []

    for item in MASTER_ITEMS:
        relevance = RELEVANCE_MATRIX.get(str(item["id"]), {}).get(service_short, "-")
        if relevance == "-":
            continue

        row = item.copy()
        row["relevance"] = relevance
        row["relevance_text"] = RELEVANCE_EXPLANATION[relevance]
        row["default_selected"] = True
        row["display_number"] = f"{row['block']}.{row['item_no']}"
        row["block_name"] = BLOCK_DEFINITIONS[row["block"]]
        rows.append(row)

    return pd.DataFrame(rows)


def build_checklist_xlsx(service_df: pd.DataFrame, service_name: str) -> bytes:
    export_df = service_df.copy()
    export_df["Stav"] = "Chybí"
    export_df["Poznámka"] = ""

    export_df = export_df[
        [
            "display_number",
            "block_name",
            "name",
            "description",
            "relevance_text",
            "Stav",
            "Poznámka",
        ]
    ].rename(
        columns={
            "display_number": "Číslo položky",
            "block_name": "Blok",
            "name": "Položka",
            "description": "Co se očekává",
            "relevance_text": "Relevance",
        }
    )

    dropdowns = {"Stav": STATUS_OPTIONS}
    return dataframe_to_excel_bytes(export_df, "Checklist", dropdowns=dropdowns, freeze_cell="A2")


def build_prehled_objektu_xlsx(service_df: pd.DataFrame) -> bytes:
    base_headers = [
        "Pořadové číslo objektu",
        "Název objektu",
        "Adresa objektu / parcelní číslo",
        "Vlastník objektu",
        "Provozovatel objektu",
        "Využívaná část objektu (odhad v %)",
        "Jméno a příjmení kontaktu",
        "Funkce",
        "Telefon",
        "E-mail",
        "Majetkoprávní vztah",
        "Památková ochrana / správní omezení",
        "Poznámka",
    ]

    item_columns = [
        f"{row['display_number']} | {row['name']}"
        for _, row in service_df.iterrows()
    ]

    headers = base_headers + item_columns

    rows = []
    for i in range(1, 4):
        rows.append(
            [i, "", "", "", "", "", "", "", "", "", "", "", ""]
            + ["Chybí"] * len(item_columns)
        )

    df = pd.DataFrame(rows, columns=headers)

    dropdown_cols = {
        col: ["Nahráno", "Chybí", "Není k dispozici", "Irelevantní"]
        for col in item_columns
    }

    return dataframe_to_excel_bytes(
        df,
        "Přehled_objektů",
        dropdowns=dropdown_cols,
        freeze_cell="D2",
    )


def build_relevance_matrix_xlsx() -> bytes:
    rows = []
    service_shorts = [cfg["short"] for cfg in SERVICES.values()]

    short_to_name = {cfg["short"]: name for name, cfg in SERVICES.items()}

    for item in MASTER_ITEMS:
        row = {
            "ID": str(item["id"]),
            "Číslo položky": f"{item['block']}.{item['item_no']}",
            "Blok": BLOCK_DEFINITIONS[item["block"]],
            "Položka": item["name"],
        }
        for short in service_shorts:
            row[short] = RELEVANCE_MATRIX.get(str(item["id"]), {}).get(short, "-")
        rows.append(row)

    df = pd.DataFrame(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = "Matrice relevance"

    for c_idx, col in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E2F3")
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")

    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_cells in ws.columns:
        max_len = 0
        col_letter = col_cells[0].column_letter
        for cell in col_cells:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 35)

    legend = wb.create_sheet("Legenda")
    legend_rows = [
        ("P", "Povinné"),
        ("D", "Doporučené"),
        ("V", "Volitelné"),
        ("-", "Irelevantní"),
    ]
    legend["A1"] = "Značka"
    legend["B1"] = "Význam"
    legend["A1"].font = Font(bold=True)
    legend["B1"].font = Font(bold=True)
    for i, (a, b) in enumerate(legend_rows, start=2):
        legend[f"A{i}"] = a
        legend[f"B{i}"] = b

    services_ws = wb.create_sheet("Služby")
    services_ws["A1"] = "Zkratka"
    services_ws["B1"] = "Služba"
    services_ws["A1"].font = Font(bold=True)
    services_ws["B1"].font = Font(bold=True)
    for i, short in enumerate(service_shorts, start=2):
        services_ws[f"A{i}"] = short
        services_ws[f"B{i}"] = short_to_name[short]

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def build_zip_package(
    customer_name: str,
    project_code: str,
    project_name: str,
    checklist_bytes: bytes,
    prehled_objektu_bytes: bytes,
    selected_df: pd.DataFrame,
) -> bytes:
    project_folder = safe_name(f"{project_code}_{project_name}_{customer_name}_PODKLADY_DPU_ENERGY")

    def write_dir(zf: zipfile.ZipFile, path: str) -> None:
        directory = path if path.endswith("/") else f"{path}/"
        zf.writestr(directory, "")

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        write_dir(zf, project_folder)
        zf.writestr(f"{project_folder}/README.txt", ROOT_README_TEXT)
        zf.writestr(f"{project_folder}/KLIC_POJMENOVANI_SOUBORU.txt", NAMING_KEY_TEXT)

        for block_no, folder in BLOCK_FOLDERS.items():
            write_dir(zf, f"{project_folder}/{folder}")
            zf.writestr(f"{project_folder}/{folder}/README.txt", BLOCK_READMES[block_no])

        for _, row in selected_df.iterrows():
            block_folder = BLOCK_FOLDERS[int(row["block"])]
            item_folder = safe_name(f"{row['display_number']}_{row['name']}")
            write_dir(zf, f"{project_folder}/{block_folder}/{item_folder}")

        zf.writestr(f"{project_folder}/01_ZADANI_A_ORGANIZACE/CHECKLIST_PODKLADU.xlsx", checklist_bytes)
        zf.writestr(f"{project_folder}/01_ZADANI_A_ORGANIZACE/PREHLED_OBJEKTU.xlsx", prehled_objektu_bytes)

        write_dir(zf, f"{project_folder}/99_ARCHIV_NEZARAZENO")

    output.seek(0)
    return output.getvalue()


st.set_page_config(page_title="Podklady pro studie DPU ENERGY", layout="wide")
st.title("Podklady pro studie DPU ENERGY")

query = parse_query_params()

col1, col2, col3 = st.columns(3)
with col1:
    customer_name = st.text_input("Název zákazníka", value=query.get("customerName", ""))
with col2:
    project_code = st.text_input("Číslo projektu v Caflou", value=query.get("projectCode", ""))
with col3:
    project_name = st.text_input("Název projektu v Caflou", value=query.get("projectName", ""))

service_name = st.selectbox("Typ služby", list(SERVICES.keys()))
service_cfg = SERVICES[service_name]

service_df = build_service_dataframe(service_name)

st.subheader("Seznam podkladů pro danou službu")
st.caption("Odškrtni položky, které po klientovi z nějakého důvodu nechceš požadovat.")

selected_rows = []

for block_no in sorted(service_df["block"].unique()):
    block_df = service_df[service_df["block"] == block_no].copy()
    block_name = BLOCK_DEFINITIONS[block_no]
    block_help = format_readme_for_ui(BLOCK_READMES.get(block_no, BLOCK_README_TEXTS.get(block_no, "")))

    with st.container(border=True):
        st.markdown(f"### {block_no}. {block_name}")

        if block_help:
            intro_line = block_help.split("\n", 1)[0].strip()
            if intro_line:
                st.caption(intro_line)

            with st.expander("Co sem patří"):
                st.markdown(block_help.replace("\n", "  \n"))

        for i, row in block_df.iterrows():
            label = f"{row['display_number']} | {row['name']} | {row['relevance_text']}"
            checked = st.checkbox(label, value=True, key=f"chk_{i}")
            if checked:
                selected_rows.append(i)

selected_df = service_df.loc[selected_rows].copy()

if st.button("Zobrazit náhled checklistu"):
    if selected_df.empty:
        st.warning("Nemáš vybranou žádnou položku.")
    else:
        preview_df = selected_df[
            ["display_number", "block_name", "name", "description", "relevance_text"]
        ].rename(
            columns={
                "display_number": "Číslo položky",
                "block_name": "Blok",
                "name": "Položka",
                "description": "Co se očekává",
                "relevance_text": "Relevance",
            }
        )
        st.dataframe(preview_df, use_container_width=True, height=650)

if selected_df.empty:
    st.warning("Vyber alespoň jednu položku seznamu.")
    st.stop()

checklist_bytes = build_checklist_xlsx(selected_df, service_name)
prehled_objektu_bytes = build_prehled_objektu_xlsx(service_df)
zip_bytes = build_zip_package(
    customer_name=customer_name,
    project_code=project_code,
    project_name=project_name,
    checklist_bytes=checklist_bytes,
    prehled_objektu_bytes=prehled_objektu_bytes,
    selected_df=selected_df,
)

st.subheader("Stažení výstupů")

st.download_button(
    "Stáhnout CHECKLIST_PODKLADU.xlsx",
    data=checklist_bytes,
    file_name=f"CHECKLIST_PODKLADU_{safe_name(project_code)}_{safe_name(SERVICES[service_name]['short'])}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.download_button(
    "Stáhnout PREHLED_OBJEKTU.xlsx",
    data=prehled_objektu_bytes,
    file_name=f"PREHLED_OBJEKTU_{safe_name(project_code)}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


st.download_button(
    "Stáhnout ZIP se složkovou strukturou",
    data=zip_bytes,
    file_name=f"{safe_name(project_code)}_PODKLADY_DPU_ENERGY.zip",
    mime="application/zip",
)