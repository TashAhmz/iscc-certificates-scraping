"""ISCC certificate to GST asset matching with confirmed direct matches.

Version 6 refreshes every confirmed target against the September 2026 GST,
filters the GoldenSource input to Started Up assets by default, updates renamed
Asset Identifiers, and expands the confirmed registry using the latest supplied
ISCC certificate database. Exact company/city rules and confirmed address rules
are evaluated before fuzzy matching and are restricted to Processing Unit rows.

By default the returned DataFrame keeps only ``Asset_Identifier`` and
``Match_Found`` from the matching output. Set ``include_match_diagnostics=True``
to retain the full audit trail during testing.

Version 4 keeps the precision guardrails from version 3, while adding a
dedicated deterministic coverage pass for HEFA, co-processing and biodiesel
processing-unit certificates.
It was revised after auditing the previous matcher, where location words such
as Rotterdam, Hamburg, Shanghai, Jiangsu and Guangdong were sometimes treated
as company-name evidence.

Main safety rules
-----------------
1. Country is a hard gate.
2. A GST company is resolved before its assets are ranked.
3. City, territory and country tokens are removed from fuzzy company-token
   evidence.
4. Only authoritative exact-name and safe substring matches can normally be
   assigned automatically. Acronyms, company forms, asset-derived names and
   distinctive-token matches are review-only unless explicitly overridden.
5. Processing-unit certificates may auto-match with strong company and
   location evidence. Trader/office/non-processing certificates require exact
   company evidence, exact location evidence and a single GST site.
6. General certificates keep the conservative phase/expansion review rule.
7. Target HEFA, co-processing and biodiesel processing units use a second,
   technology-first coverage pass. A different-technology GST row may identify
   the same physical site only where company and location evidence are exact.
   If a site is established but only phased identifiers exist, the base asset
   is preferred and otherwise Phase 1 (or the first stable identifier) is
   selected deterministically.

Typical use
-----------
    from asset_matching import match_assets_to_gst

    df = match_assets_to_gst(
        iscc_df=df,
        gst_df=GST_ASSETS,
    )

    # During matcher calibration/testing:
    test_df = match_assets_to_gst(
        iscc_df=df,
        gst_df=GST_ASSETS,
        include_match_diagnostics=True,
    )
"""

from __future__ import annotations

import re
import unicodedata
from collections import Counter, defaultdict
from typing import Any, Iterable

import numpy as np
import pandas as pd

try:
    from thefuzz import fuzz
except ImportError:  # pragma: no cover - useful when only rapidfuzz is installed
    from rapidfuzz import fuzz

try:
    from mappings import LEGAL_SUFFIXES
except ImportError:  # Allows the module to be tested independently.
    LEGAL_SUFFIXES = set()


MATCHER_VERSION = "6.0.0-current-started-assets-2026-09"


COUNTRY_MATCH_ALIASES = {
    "usa": "united states",
    "u s a": "united states",
    "us": "united states",
    "u s": "united states",
    "united states of america": "united states",
    "the united states of america": "united states",
    "uk": "united kingdom",
    "u k": "united kingdom",
    "great britain": "united kingdom",
    "the united kingdom": "united kingdom",
    "republic of korea": "south korea",
    "korea republic of": "south korea",
    "the republic of korea": "south korea",
    "republic of ireland": "ireland",
    "russian federation": "russia",
    "turkiye": "turkey",
    "czech republic": "czechia",
    "the united arab emirates": "united arab emirates",
    "uae": "united arab emirates",
    "u a e": "united arab emirates",
    "taiwan province of china": "taiwan",
    "taiwan province china": "taiwan",
    "hong kong sar": "hong kong",
    "hong kong special administrative region": "hong kong",
}


# Add only confirmed ownership, abbreviation or historic-name relationships.
# The left side is the ISCC company; the right side must resolve to the GST
# Company/Producer, Company/Producer Short Name, or a recognised GST company
# form. A confirmed alias is eligible for automatic matching if the location
# and site rules are also satisfied.
COMPANY_ALIAS_OVERRIDES: dict[str, str] = {
    # "phillips 66 company": "p66",
    # "diamond green diesel": "dgd",
    # "bsgf company limited": "bbgi",
    # "shandong sanju bioenergy": "shangdong shangjia",
}


# Use this only where a certificate-holder-to-asset relationship has been
# manually confirmed and textual matching is not sufficient.
MANUAL_ASSET_OVERRIDES: dict[tuple[str, str], str] = {
    # ("hawaii renewables llc", "united states"): "Par Pacific Kapolei",
    # ("lianyungang jiaao enproenergy", "china"): "JiaaoBP Zhejiang",
}

# Known same-city/name collisions that must never be auto-assigned. Confirmed
# direct rules still take precedence where a relationship has been validated.
BLOCKED_ISCC_ASSET_MATCHES: tuple[tuple[str, str], ...] = (
    ("Molinos Agro S.A.", "Viterra San Lorenzo"),
    ("Bunge Zrt.", "Novoal (Bunge) Bruck an der Mur"),
    ("Irish Biofuels Production Ltd", "Green Biofuels Ireland Wexford"),
)


# Exact processing-unit company/city matches curated from the supplied ISCC
# processing-unit list and the supplied started-up GST asset list. Each
# target was validated against the current GST Asset Identifier column.
#
# Tuple fields: (ISCC company, ISCC city, GST Asset Identifier,
#                required processing terms, audit reason)
CONFIRMED_STARTED_ASSET_MATCHES: tuple[
    tuple[str, str, str, tuple[str, ...], str], ...
] = (
    ('ABID Biotreibstoffe GmbH', 'Hohenau An Der March', 'ABID Hohenau', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Aceites Del Sur Coosur', 'Tarancon (Cuenca)', 'Acesur (Enersur) Cuenca', (), 'CONFIRMED_GROUP_SITE'),
    ('Aceites Manuelita S.A.', 'Meta', 'Manuelita ', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('ACOR', 'Valladolid', 'ACOR Olmedo', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Adesso Bioproducts AS', 'Gamle Fredrikstad', 'Adesso Bioproducts  Fredrikstad', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Adesso BioProducts AB', 'Ödsmål', 'Adesso Stenungsund', (), 'CONFIRMED_SITE_ALIAS'),
    ('ADM Agri-Industries Company', 'Lloydminster', 'ADM Lloydminster', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('ADM Hamburg', 'Hamburg', 'ADM Hamburg', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('ADM Oilseeds Germany', 'Mainz', 'ADM Mainz', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('AEKYUNG CHEMICAL CO.', 'Ulsan', 'Aekyung Ulsan', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Agropodnik', 'Dobronín', 'Agropodnik  Jihlava', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('AI ENERGY PUBLIC COMPANY   LIMITED', 'Samut Sakhon', 'A I Energy Samutsakorn', (), 'CONFIRMED_SITE_ALIAS'),
    ('Alcoholes del Uruguay SA', 'Montevideo', 'Alcoholes del Uruguay Capurro', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Alpha Biofuels (S) Pte Ltd', 'Singapore', 'Alpha Biofuels Singapore', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('American Greenfuels', 'Ct New Haven', 'American GreenFuels New Haven', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Anhui Yisheng New Energy Co.', 'Chuzhou City', 'Yisheng Chuzhou', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('ASB Biodiesel (Hong Kong)   Limited', 'N. T.', 'ASB Hong Kong', (), 'CONFIRMED_SITE_ALIAS'),
    ('Astra Bioplant Ltd.', 'Ruse', 'Astra  Ruse', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Aves Enerji Yag Ve Gida San.A.Ş', 'Mersin', 'Aves AS Mesrin', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('BBGI BIODIESEL COMPANY LIMITED', 'Phra Nakhon Si Ayutthaya', 'BBGI Company Bang Pa-in', (), 'CONFIRMED_SITE_ALIAS'),
    ('BE8 S.A.', 'Estrada Da Fruteira SN° Lote AB', 'Be8 Marialva', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('BE8 S.A.', 'Km SN', 'Be8 Passo Fundo', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('BINATURAL BAHIA Ltda.', 'Via De Penetração Iv - Lote', 'Binatural Bahia Simões Filho', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('BIO D TECHNOLOGY FZCO', 'Dubai', 'BioD Technology ', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Bio-Venta', 'Ventspils', 'Bio-Venta Ventspils', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Biocom Cuenca S.L.U.', 'Cuenca', 'Biocom Energia Cuenca', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Biocom Energía S.L.U.', 'Valencia', 'Biocom Energía Valencia (Algemesi)', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Biodiesel S.A. (ΒΙΟΝΤΗΖΕΛ ΕΠΕ)', 'Assiros', 'Biodiesel A.E. (Viodizel) Assiros', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Biodiesel Argent BV', 'An Amsterdam', 'Argent Energy Amsterdam', (), 'CONFIRMED_SITE_ALIAS'),
    ('Biodiesel Industries Australia   PTY Limited', 'Rutherford', 'Biodiesel Industries Australia Rutherford', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Biodiesel Kärnten GmbH', 'Arnoldstein', 'BioDiesel Kaernten Arnoldstein', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Biodiesel Süd GmbH', 'Bleiburg', 'Bio Oil Bleiburg', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Bioenergy Development Group', 'Memphis', 'Bioenergy Development Group Memphis', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Bioport SA', 'Baltar', 'Bioportdiesel Baltar', (), 'CONFIRMED_SITE_ALIAS'),
    ('Biotech Energy (Pvt) Limited', 'NA Lahore', 'Biotech Energy Sheikhupura', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Biotrading 2007 SL', 'Sevilla', 'Biotrading Sevilla', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Blue Whale Bioenergy   (Zhejiang) Co.', 'Jiaxing', 'Blue Whale Zhapu', (), 'CONFIRMED_SITE_ALIAS'),
    ('BP Energía España S.A.U.', 'Castellón', 'BP Castellón de la Plana', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('BP Europa SE', 'Lingen', 'BP Lingen', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('BP Raffinaderij Rotterdam BV', 'Na Europoort Rotterdam', 'BP Rotterdam', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Braya Renewable Fuels   (Newfoundland) LP', 'Come By Chance', 'Braya Come By Chance Phase 1', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('BREMFIELD SDN. BHD.', 'Pulau Indah', 'Bremfield Sdn Bhd (Mewah) Pulau Indah', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('BSGF company limited', 'Bangkok', 'Bangchak Bangkok', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Bunge Alimentos', 'Nova Mutum', 'Bunge Nova Mutum', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('BUNGE ARGENTINA SA', 'San Lorenzo', 'Bunge San Lorenzo', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Bunge Argentina S.A.', 'Santa Fé', 'Bunge San Lorenzo', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Bunge Deutschland GmbH', 'Mannheim', 'Bunge Mannheim', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Bunge Romania S.R.L.', 'Buzau', 'Bunge Buzău', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Cargill GmbH', 'Frankfurt A.M.', 'Cargill Frankfurt', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Cargill NV', 'Gent', 'Cargill Ghent', (), 'CONFIRMED_SITE_ALIAS'),
    ('Cargill NV', 'Moervaartkaai', 'Cargill Ghent', (), 'CONFIRMED_SITE_ALIAS'),
    ('Cargill S.A.C.I.', 'Villa Gobernador Galvez - Santa Fe', 'Cargill Rosario', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('CARGILL NOVOS HORIZONTES LTDA', 'E', 'Cargill Anápolis', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Carotino SDN BHD', 'Pasir Gudang', 'Carotino Pasir Gudang', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('CEPSA Bioenergía San Roque SL', 'San Roque (Cádiz)', 'Moeve (Cepsa) San Roque', (), 'CONFIRMED_SITE_ALIAS'),
    ('Champway Technology Limited', 'N.T.', 'Champway Hong Kong', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Changzhou City Jintan District Weige Biological Technology Co.', 'Changzhou', 'Topnotch - Changzhou City Jintan District Weige Biological Changzhou', (), 'CONFIRMED_COMPANY_SITE'),
    ('Chant Oil Co.', 'New Taipei City', 'Chant Oil New Taipei City', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Chongqing Yubang New Energy Technology Co.', 'Jiangjin District ', 'Chongqing Yubang Chongqing', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Circular Energy Company   Limited (Branch 1)', 'Khlong Luang District', 'Circular Energy Khlong Kluang', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('COFCO International Brasil', 'Rua B', 'COFCO Rondonópolis', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('College Biofuels Unlimited Company', 'Co. Meath', 'CollegeGroup Nobber', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('College Proteins ULC', 'Co. Meath', 'CollegeGroup Nobber', (), 'CONFIRMED_COMPANY_ADDRESS_SITE'),
    ('College Proteins ULC', 'Nobber', 'CollegeGroup Nobber', (), 'CONFIRMED_GROUP_SITE'),
    ('Daka ecoMotion AS', 'Dk Løsning', 'Daka ecoMotion  Losning', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('DB TARIMSAL ENERJI SAN VE TİCARET A.Ş.', 'TorbaliİZmi̇R', 'DB Tarımsal Enerji Torbalı', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Dezhou Rongguang   Bio-Technology Co.', 'Dezhou City', 'Dezhou Rongguang Dezhou', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Diamond Green Diesel', 'Norco', 'DGD Norco', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Diamond Green Diesel', 'Texas', 'DGD Port Arthur', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('DOBLE L BIOENERGIAS SA', 'Sa Pereira', 'Double L Bioenergías Sa Pereira', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('DP LUBRIFICANTI', 'Aprilia', 'DP Lubrificanti SRL Aprilia', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('DS DANSUK CO.', 'Gyeonggi-Do', 'DS Dansuk Gyeonggi-do', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('DS DANSUK CO.', 'Pyeongtaek-Si', 'DS Dansuk Pyeongtaek-si', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('DS DANSUK CO.', 'Siheung', 'DS Dansuk Gyeonggi-do', (), 'CONFIRMED_SITE_ALIAS'),
    ('DUAL DUARTE ALBUQUERQUE COMERCIO E INDUSTRIA', 'Rodovia Br Sn - Km', 'Duarte Albuquerque Pedra Preta', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('EA Bio Innovation Co.', 'Rayong Province', 'EA Bio Innovation  Map Ta Phut', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('ECO AND SOLUTIONS CO.', 'Jeollabuk-Do', 'Eco Solutions Jeonju', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('ECO Biochemical Technology   (Zhangjiagang) Company Limited', 'Zhangjiagang', 'EcoCeres Zhangjiagang', (), 'CONFIRMED_SITE_ALIAS'),
    ('ECO FOX', 'Vasto (Ch)', 'ECO Fox Vasto', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Eco Fuels Danube GmbH', 'Krems', 'Bio Oil Krems an der Donau', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('EcoCeres Renewable Fuels Sdn.   Bhd.', 'Johor', 'EcoCeres Pasir Gudang', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Ecodiesel Colombia S.A.', 'Santander', 'Ecodiesel Barrancabermeja', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('ecoMotion GmbH', 'Lünen', 'ecoMotion Lünen', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('ecoMotion GmbH', 'Malchin', 'ecoMotion Malchin', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('ECOMOTION BIODIESEL', 'Montmeló', 'ecoMotion Barcelona', (), 'CONFIRMED_SITE_ALIAS'),
    ('Ecoson B.V.', 'Nm Son En Breugel', 'EcoSon Son en Breugel', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Emax Solutions Co.', 'Jeollabuk-Do', 'Emax Solution Suncheon', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Emax Solutions Co.', 'Jeollanam-Do', 'Emax Solution Suncheon', (), 'CONFIRMED_SITE_ALIAS'),
    ('ENAP Refinerías', 'Concón', 'ENAP Concon', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('ENERFUEL', 'Sines', 'Enerfuel – GALP  Sines', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Energy Absolute Public Company   Limited', 'Amphoe Kabinburi', 'Energy absolute Prachinburi', (), 'CONFIRMED_SITE_ALIAS'),
    ('Eni Spa', 'Taranto', 'ENI Taranto', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Equinor Refining Norway AS', 'Mongstad', 'Equinor Mongstad', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Estener', 'Le Havre', 'Estener (Saria Industries) Le Havre', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Everfuel Production Fredercia AS', 'Fredericia', 'Everfuel - HySynergy Fredericia', (), 'CONFIRMED_PROJECT_SITE'),
    ('EXPLORA S.A.', 'S Puerto General San Martín - Santa Fe', 'Explora San Lorenzo', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('F.J. Sánchez Sucesores', 'Carboneras (Almeria)', 'F.J. Sanchez Carboneras', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Fábrica Torrejana', 'Riachos', 'Fabrica Torrejana Riachos', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('FGV Biotechnologies Sdn Bhd', 'Kuantan', 'FGV Biotechnologies Kuantan', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Future Prelude Sdn Bhd', 'Pelabuhan Klang', 'Future Prelude Port Klang', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Fytoenergeia S.A.', 'Serres', 'New Energy (FytoEnergia) Serres', (), 'CONFIRMED_SITE_ALIAS'),
    ('Genting Biorefinery Sdn Bhd', 'Lahad Datu', 'Genting Biodiesel Lahad Datu', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('GF ENERGY S.A.', 'Korinth', 'GF Energy Corinth', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Green Biofuels Ireland Ltd', 'Co. Wexford New Ross', 'Green Biofuels Ireland Wexford ', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Green Edible Oil Sdn Bhd', 'Sabah', 'Green Edible Oil Lahad Datu', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Green Fuel Extremadura S.A.', 'Los Santos De Maimona', 'Green Fuel Extremadura Los Santos de Maimona', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Greenergy Biofuels Amsterdam   BV', 'Ah Amsterdam', 'Greenergy Amsterdam', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Greenergy Biofuels Teesside   Limited', 'Middlesbrough', 'Greenergy Seal Sands', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('GRI Energy Co.', 'Gyeonggi-Do', 'GRI Energy Gyeonggi-do', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('GS Bio', 'Jeollanam-Do Yeosu-Si', 'GS Bio Yeosu', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Guangzhou Hongtai New Energy Technology Co.', 'Guangzhou City', 'Guangdong Hongtai ', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Gulf Lubes Malaysia', 'Pulau Indah', 'Gulf Lubes Malaysia Klang', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Gunvor Biodiesel Berantevilla   SL', 'Berantevilla', 'Gunvor  Berantevilla', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Gunvor España SL', 'Palos De La Frontera (Huelva)', 'Gunvor Huelva', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Hainan Huanyu New Energy Co.', 'Haikou', 'Hainan Huanyu (Haixin) Lingao', (), 'CONFIRMED_COMPANY_ADDRESS_SITE'),
    ('Hawaii Renewables LLC', 'Komohana', 'Par Pacific Kapolei', (), 'CONFIRMED_OWNERSHIP_SITE'),
    ('HD Hyundai Chemical Co.', 'Chungcheongnam-Do', 'Hyundai Daesan', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('HD Hyundai Oilbank', 'Chungcheongnam-Do', 'Hyundai Daesan', (), 'CONFIRMED_COMPANY_REGION'),
    ('Hebei Jingu Plasticizer Co. Ltd', 'Xinji', 'Hebei Jingu Xinji', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Hebei Jingu Recycling   Resources Development Co.', 'Xinji City', 'Hebei Jingu Xinji', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Hebei Longhai Bioenergy Co.', 'Handan', 'Hebei Longhai Handan', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Hellenic Biopetroleum S.A.', 'Kilkis', 'Hellenic Petroleum (EL.VI) Kilkis', (), 'CONFIRMED_RENAMED_SITE'),
    ('HENAN JUNHENG Industrial Group   biotechnology company.', 'Puyang', 'Henan Junheng Puyang Phase 1', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('HOLBORN Europa Raffinerie GmbH', 'Hamburg', 'Tamoil/Oilinvest Hamburg', (), 'CONFIRMED_SITE_ALIAS'),
    ('Ina Industrija nafte d.d.', 'Kostrena', 'INA/Chevron Rijeka', (), 'CONFIRMED_SITE_ALIAS'),
    ('INDIAN OIL CORPORATION LIMITED', 'Panipat', 'Indian Oil Panipat', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Iniciativas Bioenergeticas', 'La Rioja', 'Iniciativas Bioenergeticas  Calahorra\xa0', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Innoltek', 'Saint-Jean-Sur-Richelieu', 'Innoltek St-Jean-Sur-Richelieu', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('ITAL BI-OIL S.R.L.', 'Monopoli', 'Ital Bi Oil SRL Monopoli', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('JBS\xa0 SA', 'Lins', 'Jbs Lins', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('JBS SA', 'SN', 'Jbs Campo Verde', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('JC Chemical Co.', 'Ulsan', 'JC Chemical Ulsan', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Jiangxi Zunchuang New Energy Co.', 'Shangrao City', 'Jiangxi Zunchuang Dexing', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Jiujiang Oasis Energy Technology Co.', 'Jiujiang City', 'Jiujiang Oasis Energy  Juljiang', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Jiujiang Oasis Environmental Technology Co.', 'Jiujiang City', 'Jiujiang Oasis Energy  Juljiang', (), 'CONFIRMED_COMPANY_ADDRESS_SITE'),
    ('Kaleesuwari Refinery &   Industry Private Limited', 'Andhra Pradesh', 'Kaleesuwari Refinery and Industry Kakinada', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Kanola Biofuels B.V.', 'Lexmond', 'Kanola/Dutch Biofuels Rotterdam', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('La Nivernaise de Raffinage', 'Premery', 'La Nivernaise de Raffinage SAS Prémery', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('LanzaJet', 'Soperton', 'LanzaJet Soperton', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('LDC Argentina S.A.', 'Santa Fe', 'Ldc General Lagos', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('LianYunGang Jiaao Enproenergy   Co.', 'Lianyungang City', 'Jiaao/BP Zhejiang', (), 'CONFIRMED_SITE_ALIAS'),
    ('Linares Biodiesel Technology', 'Linares', 'LiBiTech Linares', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Longyan Zhuoyue New Energy Co.', 'Longyan', 'Longyan Longyan/Xiamen', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Longyan Zhuoyue New Energy Biotechnology Co.', 'Longyan City', 'Longyan Xiamen', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Lootah Biofuels LLC', 'Dubai', 'Lootah Biofuels Dubai', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Louis Dreyfus Company   Wittenberg', 'Lutherstadt Wittenberg', 'Louis Dreyfus Wittenberg', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Maoming Hongyu Energy   Technology Co.', 'Maoming City', 'Maoming Hongyu  Maoming', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Masol Cartagena Biofuel S.L.', 'Cartagena (Murcia)', 'Musim Mas Cartagena', (), 'CONFIRMED_GROUP_SITE'),
    ('Masol Continental Biofuel Srl', 'Livorno', 'Musim Mas Livorno', (), 'CONFIRMED_SITE_ALIAS'),
    ('Masol Iberia Biofuel S.L.', 'Castellon De La Plana', 'Musim Mas Castellón de la Plana', (), 'CONFIRMED_SITE_ALIAS'),
    ('Masol Iberia Biofuel S.L.', 'Ferrol', 'Musim Mas Ferrol', (), 'CONFIRMED_SITE_ALIAS'),
    ('Mercuria Biofuels Brunsbüttel   GmbH & Co. KG', 'Brunsbüttel', 'Mercuria (Vesta) Brunsbüttel', (), 'CONFIRMED_SITE_ALIAS'),
    ('Meroco', 'Leopoldov', 'Meroco Envien Bratislava', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('MERT YAĞ ASFALT MADENCİLİK VE   LASTİK GERI DÖNÜŞÜM SAN. TİC.LTD.ŞTİ', 'Bor-Ni̇Ğde', 'Mert Oil Biodiesel Niğde', (), 'CONFIRMED_SITE_ALIAS'),
    ('Mewah-Oils SDN BHD', 'Selangor', 'Bremfield Sdn Bhd (Mewah) Pulau Indah', (), 'CONFIRMED_COMPANY_ADDRESS_SITE'),
    ('Mil Oil Hellas', 'Serres', 'Miloil Hellas Serres', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('MINERVA S.A.', 'Zona Rural SN', 'Minerva Palmeiras de Goiás', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Moeve', 'Palos De La Frontera (Huelva)', 'Moeve (Cepsa) Huelva ', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Moeve', 'San Roque (Cádiz)', 'Moeve (Cepsa) San Roque', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Moeve Chemicals', 'Cádiz', 'Moeve (Cepsa) San Roque', (), 'CONFIRMED_INTEGRATED_SITE'),
    ('Moeve Chemicals', 'Huelva', 'Moeve (Cepsa) Huelva ', (), 'CONFIRMED_INTEGRATED_SITE'),
    ('MOL Downstream Private Company   Limited by Shares', 'Százhalombatta', 'MOL Budapest', (), 'CONFIRMED_SITE_ALIAS'),
    ('Montana Renewables', 'Great Falls', 'Montana Great Falls Phase 1', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Münzer Bioindustrie GmbH', 'Vienna', 'Munzer Wien', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Münzer Paltental GmbH', 'Gaishorn Am See', 'Munzer Glashorn', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Natural Energy West GmbH', 'Marl', 'NEW Marl', (), 'CONFIRMED_SITE_ALIAS'),
    ('Neste Oyj', 'Kulloo', 'Neste Porvoo Phase 1', (), 'CONFIRMED_SITE_ALIAS'),
    ('Neste Oyj', 'Rotterdam', 'Neste Rotterdam Phase 1', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Neste Components B.V.', 'Rotterdam', 'Neste Rotterdam Phase 1', (), 'CONFIRMED_SITE_ALIAS'),
    ('Neste Singapore Pte Ltd', 'Singapore', 'Neste Singapore Phase 1', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Neutral Fuels', 'Dubai', 'Neutral Fuels Dubai', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Neutral Fuels', 'Uae', 'Neutral Fuels Abu Dhabi', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('New biodiesel Co.', 'Surat Thani', 'New Biodiesel Surat Thani', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('New Rise Renewables Reno', 'Peru Dr.', 'XCF/New Rise Reno', (), 'CONFIRMED_SITE_ALIAS'),
    ('Nexsol (Malaysia) Sdn Bhd', 'Kawasan Perindustrian Tanjung Langsat Pasir Gudang', 'Wilmar Pasir Gudang', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('NINGBO JIESEN GREEN FUEL CO.', 'Fenghua District Haiyan Village', 'Ningbo Jiesen Ningbo', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('NORD-ESTER', 'Dunkerque', 'Nord Ester Dunkirk', (), 'CONFIRMED_SITE_ALIAS'),
    ('Oberösterreichische Biodiesel   - Bulgarien GmbH', 'Rousse', 'Astra  Ruse', (), 'CONFIRMED_SITE_ALIAS'),
    ('OLEOPLAN S.A. OLEOS VEGETAIS   PLANALTO', 'Veranópolis', 'Oleoplan Veranópolis', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('OLFAR SA ALIMENTO E ENERGIA', 'Bairro Village Porto Real', 'Olfar Porto Real', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('OLFAR SA ALIMENTO E ENERGIA', 'Trecho Azinópolis SN', 'Olfar Porangatu', (), 'CONFIRMED_SITE_ALIAS'),
    ('Olleco', 'Liverpool', 'Olleco Liverpool', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Oman Blending Services LLC', 'Samail', 'Oman Blending Services Samail Industrial Estate', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('OMV Deutschland Operations   GmbH & Co. KG', 'Burghausen', 'OMV Burghausen', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('OMV Downstream GmbH', 'Schwechat', 'OMV Schwechat', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('OMV PETROM', 'Brazi', 'OMV Ploiesti Phase 1', (), 'CONFIRMED_SITE_ALIAS'),
    ('ORLEN S.A.', 'Płock', 'PKN Orlen HEFA Plock', ('hefa', 'hvo'), 'CONFIRMED_TECHNOLOGY_SITE'),
    ('ORLEN S.A.', 'Płock', 'PKN Orlen Plock', ('co processing', 'co-processing'), 'CONFIRMED_TECHNOLOGY_SITE'),
    ('ORLEN Południe S.A.', 'Trzebinia', 'PKN Orlen Trzebinia', (), 'CONFIRMED_SITE_ALIAS'),
    ('ORLEN Unipetrol RPA s.r.o.', 'Litvinov', 'Unipetrol Litvinov', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Padang Raya Cakrawala', 'Kecamatan Lubuk Begalung -', 'Apical Padang', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('PETROBRAS BIOCOMBUSTIVEL SA -   Candeias', 'Candeias', 'Petrobras Candeias', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('PETROBRAS BIOCOMBUSTIVEL SA -   Montes Claros', 'Montes Claros', 'Petrobras Montes Claros', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Petrogal', 'Sines', 'Galp Sines', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Petróleo Brasileiro SA PETROBRAS', 'Rodovia Washington Luiz SN Br SN', 'Petrobras Duque de Caxias', (), 'CONFIRMED_REFINERY_ADDRESS'),
    ('PGEO Bioproducts Sdn. Bhd.', 'Pasir Gudang', 'Wilmar Pasir Gudang', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Phillips 66 Limited', 'North Lincolnshire', 'P66 Hull', (), 'CONFIRMED_SITE_ALIAS'),
    ('Phillips 66 Company', 'Rodeo', 'P66 Rodeo Phase 1', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Potencial Agro', 'Av. Eduardo Pedro Hammerschmidt', 'Potencial Biodiesel Lapa', (), 'CONFIRMED_COMPANY_ADDRESS_SITE'),
    ('Preol', 'Lovosice', 'Agrofert (Preol) Lovosice', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Primagra', 'Milín', 'Agrofert Milín', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Prio Bio SA', 'Gafanha Da Nazaré', 'Prio Aveiro', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Prisma Comercial Exportadora   de Oleoquímicos Ltda', 'Sumaré', 'Prisma Comercial Exportadora De Oleoquimicos Sumaré', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('PT Energi Unggul Persada', 'Jl Raya Sungai Limau Rt Rw Kelurahan Sungai Limau Kecamatan Sungai Kunyit', 'Gama Mempawah', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('PT Energi Unggul Persada', 'Jl Tanjung Meranggas Segendis Rt Kelurahan Bontang Lestari   Kecamatan Bontang Selatan', 'Gama Bontang', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('PT Intibenua Perkasatama', 'Kecamatan Sungai Sembilan -', 'Musim Mas Medan', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('PT Kutai Refinery Nusantara', 'Kelurahan Kariangau', 'Apical Balikpapan', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('PT. LDC Indonesia', 'Kec. Panjang Km Lk Ii', 'LDC Indonesia Bandar Lampung (Panjang)', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('PT Multimas Nabati Asahan', 'Kab. Serang Km', 'Wilmar Serang (Kramatwatu/Bojonegara)', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('PT Musim Mas', 'Belawan I -', 'Musim Mas Medan', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('PT Musim Mas', 'Mabar -', 'Musim Mas Medan', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('PT. Musim Mas', 'Medan', 'Musim Mas Medan', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('PT Musim Mas', 'Nongsa -', 'Musim Mas Batam (Kabil–Nongsa)', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('PT Pertamina Patra Niaga', 'Kec. Cilacap Tengah', 'Pertamina Cilacap Phase 1', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('PT Sari Dumai Sejati', 'Dumai', 'Apical Dumai', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('PT Sinarmas Bio Energy', 'Kecamatan Tarumajaya Kawasan Industri Marunda Center Blok D', 'Sinarmas Bio Energy Marunda Center (Bekasi/Jakarta area)', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('PT SMART Tbk', 'Kabupaten Kotabaru', 'SMART Tbk Tarjun, Kotabaru (South Kalimantan)', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('PT Sukajadi Sawit Mekar', 'Kecamatan Mentaya Hilir Utara -', 'Musim Mas Kotawaringin Timur (Bagendang)', (), 'CONFIRMED_SITE_ALIAS'),
    ('PT Wilmar Bioenergi Indonesia', 'Kota Dumai', 'Wilmar Dumai', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('PT Wilmar Nabati Indonesia', 'Kecamatan Kebomas', 'Wilmar Gresik', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('PT Wilmar Nabati Indonesia', 'Kota Dumai', 'Wilmar Dumai', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('PTP Argent B.V.', 'Amsterdam', 'Argent Energy Amsterdam', (), 'CONFIRMED_SITE_ALIAS'),
    ('RapSol GmbH', 'Lübz', 'Rapsol Lübz', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('REG Geismar', 'Geismar', 'REG Geismar Phase 1', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('REG Seneca', 'Seneca', 'REG Seneca Seneca', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Repsol Fuels S.A.U.', 'A Coruña', 'Repsol A Coruna', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Repsol Fuels S.A.U.', 'Cartagena (Murcia)', 'Repsol Cartagena', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Repsol Fuels S.A.U.', 'Puertollano (Ciudad Real)', 'Repsol Puertollano', (), 'CONFIRMED_SITE_ALIAS'),
    ('Repsol Materials', 'Puertollano', 'Repsol Puertollano', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('REVO International Inc.', 'Midorigahama', 'Revo International Tahara-Shi', (), 'CONFIRMED_SITE_ALIAS'),
    ('ROSSI Biofuel Bioüzemanyag   Gyártó és Kereskedelmi Zrt.', 'Komárom', 'Envien (Rossie Mol) Komárom', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Sabio fuels s.r.l', 'Castenedolo (Bs)', 'Sabio Castenedolo', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('SAFFAIRE SKY ENERGY LLC', 'Nishi-Ku', 'JGC/Cosmo (Saffaire Sky Energy) Osaka', (), 'CONFIRMED_SITE_ALIAS'),
    ('SAIPOL', 'Bassens', 'Avril (Saipol) Bassens', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('SAIPOL', 'Grand-Couronne', 'Avril (Saipol) Grand-Couronne', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('SAIPOL', 'Le Meriot', 'Avril (Saipol) Le Mériot', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('SAIPOL', 'Sète', 'Avril (Saipol) Sète', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('SARAS S.P.A. O IN FORMA ESTESA SARAS S.P.A. - RAFFINERIE SARDE', 'Sarroch', 'Saras Sarroch', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('SD Guthrie International   Biodiesel Sdn Bhd', 'Selangor', 'Sime Darby Oil Biodiesel Carey Island', (), 'CONFIRMED_SITE_ALIAS'),
    ('Seaboard Energy Kansas', 'Hugoton', 'SE Hugoton', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Seaboard Energy Missouri', 'St. Joseph', 'Seaboard Energy Missouri St. Joseph', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Seaboard Energy Missouri', 'Stockyards Expy', 'Seaboard Energy Missouri St. Joseph', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Seaboard Energy Oklahoma LLC', 'Guymon', 'Seaboard Energy Oklahoma Guymon', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Seara Alimentos Ltda', 'Comerciante Lauro Guilherme Guths Sn', 'Seara Alimentos Mafra', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Shandong Baoshun Chemical   Technology Co.', 'Heze', 'Shandong Baoshun Heze (Haixin) Heze', (), 'CONFIRMED_SITE_ALIAS'),
    ('Shandong Haike Chemical Co.', 'Dongying City', 'Haike Dongying', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Shandong Huidong New Energy   Co.', 'Dongying City', 'Shandong Huidong New Energy  Dongying', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Shandong Sanju Bioenergy Co.', 'Rizhao City', 'Sanju Rizhao Shandong (Haixin) Rizhao', (), 'CONFIRMED_SITE_ALIAS'),
    ('Shandong Sanju Bioenergy Co.', 'Rizhao City，Shandong Province', 'Sanju Rizhao Shandong (Haixin) Rizhao', (), 'CONFIRMED_SITE_ALIAS'),
    ('Shandong Shangjia Renewable   Resources Technology Co.', 'Dongying City', 'Shangdong Shangjia Unknown', (), 'CONFIRMED_SITE_ALIAS'),
    ('Shanghai Zhongqi Environment Technology Co.', 'Fengxian District', 'Shanghai Zhongqi Shanghai', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Sinopec Ningbo Zhenhai   Refining & Chemical Co.', 'Ningbo City', 'Sinopec Ningbo', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Sinopec Ningbo Zhenhai   Refining & Chemical Co.', 'Zhenhai District No.', 'Sinopec Ningbo', (), 'CONFIRMED_SITE_ALIAS'),
    ('SK eco prime Co.', 'Ulsan', 'SK ECOPrime Ulsan', (), 'CONFIRMED_SITE_ALIAS'),
    ('SK Energy Co.', 'Ulsan', 'SK Ulsan', (), 'CONFIRMED_COMPANY_SITE'),
    ('SLOVNAFT', 'Bratislava', 'MOL/Slovnaft Bratislava', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Sovena Oilseeds Portugal', 'Palença De Baixo', 'Sovena Almada', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('SPC Biodiesel Sdn Bhd', 'Lahad Datu', 'SPC Biodiesel Lahad Datu', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('St. Bernard Renewables LLC', 'West St. Bernard Highway', 'St Bernard (PBF/ENI) Chalmette', (), 'CONFIRMED_COMPANY_ADDRESS_SITE'),
    ('St1 Sverige AB', 'Gothenburg', 'ST1/SCA Gothenburg', (), 'CONFIRMED_JV_SITE'),
    ('Sunoil Bio Fuels B.V.', 'At Emmen', 'Sunoil  Emmen', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Tangshan Jinlihai Biodiesel   Co.', 'Tangshan City', 'Tangshan Jinlihai(via Golden Base) Tangshan', (), 'CONFIRMED_SITE_ALIAS'),
    ('Temperatior s.r.o.', 'Liberec Vi-Rochlice', 'Temperatior Liberec', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('TotalEnergies Fluids', 'Oudalle', 'Total Oudalle', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('TotalEnergies Petrochemicals   & Refining', 'Antwerp', 'Total Antwerp', (), 'CONFIRMED_SITE_ALIAS'),
    ('TotalEnergies Petrochemicals   & Refining', 'Antwerpen', 'Total Antwerp', (), 'CONFIRMED_SITE_ALIAS'),
    ('TotalEnergies Petrochemicals   France', 'Gonfreville L’Orcher', 'Total Oudalle', (), 'CONFIRMED_SITE_ALIAS'),
    ('TotalEnergies Raffinage France', 'Gonfreville LOrcher', 'Total Oudalle', (), 'CONFIRMED_SITE_ALIAS'),
    ('TotalEnergies Raffinage France', 'Plateforme De La Mède', 'Total La Mede', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('TotalEnergies Refinery Antwerp', 'Antwerpen Haven', 'Total Antwerp', (), 'CONFIRMED_SITE_ALIAS'),
    ('TPG Oil & Gas Sdn Bhd', 'Johor', 'KPN - TPG OIL & GAS Pasir Gudang', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('UAB MESTILLA', 'Klaipėda', 'Mestilla Klaipėda', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('UPM-Kymmene Oyj', 'Lappeenranta', 'UPM Lappeenranta', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Vance Bioenergy Sdn Bhd', 'Kompleks Perindustrian Tanjung Langsat', 'Vance Bioenergy Pasir Gudang/Tanjung Langsat', (), 'CONFIRMED_SITE_ALIAS'),
    ('Vance Bioenergy Sdn Bhd', 'Pasir Gudang', 'Vance Bioenergy Pasir Gudang/Tanjung Langsat', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('VAROPreem Sverige AB', 'Gothenburg', 'Varo (Preem) Gothenburg Phase 1', (), 'CONFIRMED_SITE_ALIAS'),
    ('VAROPreem Sverige AB', 'Lysekil', 'Varo (Preem) Lysekil Phase 1', (), 'CONFIRMED_SITE_ALIAS'),
    ('Verasuwan Co.', 'Amphoe Mueang Samut Sakhon', 'Verasuwan Samutsakorn', (), 'CONFIRMED_COMPANY_ADDRESS_SITE'),
    ('VERBIO Bitterfeld GmbH', 'Bitterfeld-Wolfen', 'Verbio Bitterfeld', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('VERBIO Diesel Canada   Corporation', 'Welland', 'Welland Welland', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Viterra Botlek B.V.', 'Ks Rotterdam-Botlek', 'Viterra (Glencore) Rotterdam', (), 'CONFIRMED_SITE_ALIAS'),
    ('Wakud International LLC', 'Barka', 'Wakud International Khazaen Economic City', (), 'CONFIRMED_CURRENT_GST_ISCC_MATCH'),
    ('Wenzhou Zhongke New Energy   Technology Co.', 'Wenzhou City', 'Wenzhou Zhongke Wenzhou', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Xiamen Zhuoyue Biomass Energy   Co.', 'Xiamen', 'Longyan Xiamen', (), 'HIGH_CONFIDENCE_COMPANY_SITE'),
    ('Zhejiang Jiaao Enproenergy Co.', 'Zhejiang Province', 'Zhejiang Jiaao, ZJJA Tongxiang', (), 'CONFIRMED_COMPANY_REGION_SITE'),
    ('ADM do Brasil - Rondonopolis', 'Av. Senador Atilio Fontana', 'Adm Do Brasil Rondonópolis', (), 'CONFIRMED_COMPANY_ADDRESS_SITE'),
)


# ---------------------------------------------------------------------------
# Confirmed address-based matches for rows whose parsed City is blank or poor
# ---------------------------------------------------------------------------

CONFIRMED_STARTED_ADDRESS_MATCHES: tuple[
    tuple[str, str, tuple[str, ...], str, tuple[str, ...], str], ...
] = (
    ('Argent Energy Limited', 'United Kingdom', ('ellesmere port', 'oil sites road', 'stanlow'), 'Argent Energy Stanlow ', ('biodiesel',), 'CONFIRMED_ADDRESS_SITE'),
    ('North Atlantic Energies', 'France', ('port jerome', 'port-jerome', 'boulevard kennedy'), 'Exxon Port-Jérôme-sur-Seine', ('co processing', 'co-processing'), 'CONFIRMED_RENAMED_SITE'),
    ('ExxonMobil Chemical France', 'France', ('port jerome', 'port-jerome', 'avenue du president kennedy'), 'Exxon Port-Jérôme-sur-Seine', (), 'CONFIRMED_INTEGRATED_SITE'),
    ('PT. Sari Dumai Oleo', 'Indonesia', ('dumai', 'sungai sembilan'), 'Apical Dumai', ('biodiesel',), 'CONFIRMED_ADDRESS_SITE'),
)


# ---------------------------------------------------------------------------
# Target processing-unit coverage
# ---------------------------------------------------------------------------

# Only rows that are Processing Units and contain one or more of these plant
# types enter the higher-recall pass. HVO alone is intentionally not included;
# the certificate must explicitly say HEFA, co-processing or biodiesel.
TARGET_PROCESSING_CATEGORIES = frozenset(
    {"HEFA", "CO_PROCESSING", "BIODIESEL"}
)

# Confirmed naming/ownership relationships used only by the target-plant pass.
# They do not relax matching for traders, offices or unrelated technologies.
TARGET_COMPANY_ALIAS_OVERRIDES: dict[str, str] = {
    "phillips 66 company": "p66",
    "phillips 66 limited": "p66",
    "neste components": "neste",
    "neste singapore": "neste",
    "new rise renewables reno": "xcf new rise",
    "natural energy west": "new",
    "viterra botlek": "viterra glencore",
    "mercuria biofuels brunsbuttel": "mercuria vesta",
    "masol continental biofuel": "musim mas",
    "masol iberia biofuel": "musim mas",
    "pt sari dumai oleo": "apical",
    "lukoil neftohim burgas": "litasco",
    "preol": "agrofert preol",
    "bbgi biodiesel company": "bbgi company",
    "mol downstream private company limited by shares": "mol",
    "omv downstream": "omv",
    "omv petrom": "omv",
    "hellenic biopetroleum": "hellenic petroleum",
    "genting biorefinery": "genting biodiesel",
    "oberosterreichische biodiesel bulgarien": "astra",
    "shandong sanju bioenergy": "sanju rizhao shandong haixin",
    "shandong baoshun chemical technology": "shandong baoshun heze haixin",
    "saudi aramco total refining petrochemical company satorp": "aramco total",
    "hanwha totalenergies petrochemical": "hanwha total",
    "hawaii renewables": "par pacific",
}

# Location-aware direct matches for known difficult records. These are checked
# only for the three target processing technologies. Location terms prevent a
# company with multiple sites from being forced to the wrong facility.
TARGET_SITE_OVERRIDES: tuple[dict[str, Any], ...] = (
    {"company": "olfar", "country": "brazil", "locations": ("porangatu", "azinopolis"), "asset": "Olfar Porangatu"},
    {"company": "olfar", "country": "brazil", "locations": ("porto real",), "asset": "Olfar Porto Real"},
    {"company": "henan junheng", "country": "china", "locations": ("puyang",), "asset": "Henan Junheng Puyang Phase 1"},
    {"company": "neste", "country": "netherlands", "locations": ("rotterdam", "maasvlakte", "botlek"), "asset": "Neste Rotterdam Phase 1"},
    {"company": "neste", "country": "finland", "locations": ("porvoo", "kulloo"), "asset": "Neste Porvoo Phase 1"},
    {"company": "neste", "country": "singapore", "locations": ("singapore", "tuas"), "asset": "Neste Singapore Phase 1"},
    {"company": "phillips 66", "country": "united states", "locations": ("rodeo",), "asset": "P66 Rodeo Phase 1"},
    {"company": "phillips 66", "country": "united states", "locations": ("old ocean", "sweeny"), "asset": "P66 Sweeny"},
    {"company": "phillips 66", "country": "united kingdom", "locations": ("hull", "north lincolnshire"), "asset": "P66 Hull"},
    {"company": "new rise renewables", "country": "united states", "locations": ("reno", "mccarran", "peru dr"), "asset": "XCF/New Rise Reno"},
    {"company": "shandong baoshun", "country": "china", "locations": ("heze", "juye"), "asset": "Shandong Baoshun Heze (Haixin) Heze"},
    {"company": "shandong sanju", "country": "china", "locations": ("rizhao", "ju county"), "asset": "Sanju Rizhao Shandong (Haixin) Rizhao"},
    {"company": "totalenergies", "country": "germany", "locations": ("leuna",), "asset": "Total Leuna"},
    {"company": "totalenergies", "country": "belgium", "locations": ("antwerp", "antwerpen"), "asset": "Total Antwerp"},
    {"company": "totalenergies", "country": "france", "locations": ("oudalle", "gonfreville"), "asset": "Total Oudalle"},
    {"company": "totalenergies", "country": "france", "locations": ("donges",), "asset": "Total Donges"},
    {"company": "totalenergies", "country": "france", "locations": ("la mede", "mede", "chateauneuf les martigues"), "asset": "Total La Mede"},
    {"company": "bp", "country": "netherlands", "locations": ("rotterdam", "europoort"), "asset": "BP Rotterdam"},
    {"company": "bp", "country": "germany", "locations": ("lingen",), "asset": "BP Lingen"},
    {"company": "bp", "country": "germany", "locations": ("gelsenkirchen",), "asset": "BP Gelsenkirchen"},
    {"company": "bp", "country": "spain", "locations": ("castellon",), "asset": "BP Castellón de la Plana"},
    {"company": "omv", "country": "austria", "locations": ("schwechat",), "asset": "OMV Schwechat"},
    {"company": "omv", "country": "romania", "locations": ("ploiesti", "petrobrazi", "brazi"), "asset": "OMV Ploiesti Phase 1"},
    {"company": "mol downstream", "country": "hungary", "locations": ("szazhalombatta",), "asset": "MOL Budapest"},
    {"company": "lukoil neftohim burgas", "country": "bulgaria", "locations": ("burgas",), "asset": "Litasco Burgas"},
    {"company": "preol", "country": "czechia", "locations": ("lovosice",), "asset": "Agrofert (Preol) Lovosice"},
    {"company": "meroco", "country": "slovakia", "locations": ("leopoldov", "bratislava"), "asset": "Meroco Envien Bratislava"},
    {"company": "adesso", "country": "sweden", "locations": ("odsmal", "stenungsund"), "asset": "Adesso Stenungsund"},
    {"company": "cargill", "country": "belgium", "locations": ("gent", "ghent", "moervaartkaai"), "asset": "Cargill Ghent"},
    {"company": "cargill", "country": "argentina", "locations": ("rosario", "villa gobernador galvez", "santa fe"), "asset": "Cargill Rosario"},
    {"company": "natural energy west", "country": "germany", "locations": ("marl",), "asset": "NEW Marl"},
    {"company": "masol continental", "country": "italy", "locations": ("livorno",), "asset": "Musim Mas Livorno"},
    {"company": "masol iberia", "country": "spain", "locations": ("castellon",), "asset": "Musim Mas Castellón de la Plana"},
    {"company": "masol iberia", "country": "spain", "locations": ("ferrol",), "asset": "Musim Mas Ferrol"},
    {"company": "nord ester", "country": "france", "locations": ("dunkerque", "dunkirk"), "asset": "Nord Ester Dunkirk"},
    {"company": "viterra botlek", "country": "netherlands", "locations": ("rotterdam", "botlek"), "asset": "Viterra (Glencore) Rotterdam"},
    {"company": "mercuria biofuels", "country": "germany", "locations": ("brunsbuttel",), "asset": "Mercuria (Vesta) Brunsbüttel"},
    {"company": "pt sukajadi sawit mekar", "country": "indonesia", "locations": ("mentaya", "bagendang", "kotawaringin"), "asset": "Musim Mas Kotawaringin Timur (Bagendang)"},
    {"company": "pt intibenua perkasatama", "country": "indonesia", "locations": ("sungai sembilan", "dumai", "medan"), "asset": "Musim Mas Medan"},
    {"company": "sk eco prime", "country": "south korea", "locations": ("ulsan",), "asset": "SK ECOPrime Ulsan"},
    {"company": "emax solutions", "country": "south korea", "locations": ("jeollanam", "jeollabuk", "suncheon"), "asset": "Emax Solution Suncheon"},
    {"company": "aekyung chemical", "country": "south korea", "locations": ("ulsan",), "asset": "Aekyung Ulsan"},
    {"company": "pt sari dumai oleo", "country": "indonesia", "locations": ("dumai",), "asset": "Apical Dumai"},
    {"company": "bbgi biodiesel", "country": "thailand", "locations": ("bang pa in", "ayutthaya"), "asset": "BBGI Company Bang Pa-in"},
    {"company": "gunvor", "country": "spain", "locations": ("huelva", "palos de la frontera"), "asset": "Gunvor Huelva"},
    {"company": "munzer", "country": "austria", "locations": ("vienna", "wien"), "asset": "Munzer Wien"},
    {"company": "tangshan jinlihai", "country": "china", "locations": ("tangshan",), "asset": "Tangshan Jinlihai(via Golden Base) Tangshan"},
    {"company": "new biodiesel", "country": "thailand", "locations": ("surat thani",), "asset": "New Biodiesel Surat Thani"},
    {"company": "alcoholes del uruguay", "country": "uruguay", "locations": ("montevideo", "capurro"), "asset": "Alcoholes del Uruguay Capurro"},
    {"company": "vance bioenergy", "country": "malaysia", "locations": ("tanjung langsat", "pasir gudang"), "asset": "Vance Bioenergy Pasir Gudang/Tanjung Langsat"},
    {"company": "ds dansuk", "country": "south korea", "locations": ("pyeongtaek",), "asset": "DS Dansuk Pyeongtaek-si"},
    {"company": "ds dansuk", "country": "south korea", "locations": ("gyeonggi", "siheung"), "asset": "DS Dansuk Gyeonggi-do"},
    {"company": "neutral fuels", "country": "united arab emirates", "locations": ("dubai",), "asset": "Neutral Fuels Dubai"},
    {"company": "neutral fuels", "country": "united arab emirates", "locations": ("abu dhabi", "mafraq"), "asset": "Neutral Fuels Abu Dhabi"},
    {"company": "hawaii renewables", "country": "united states", "locations": ("kapolei", "komohana"), "asset": "Kapolei"},
    {"company": "hanwha totalenergies", "country": "south korea", "locations": ("daesan",), "asset": "Hanwha/Total Daesan"},
    {"company": "tupras", "country": "turkey", "locations": ("izmir", "aliaga"), "asset": "Tupras Izmir Phase 1"},
    {"company": "cpc", "country": "taiwan", "locations": ("taoyuan",), "asset": "CPC Taiwan Taoyuan"},
    {"company": "cepsa", "country": "spain", "locations": ("huelva",), "asset": "Moeve (Cepsa) Huelva "},
    {"company": "cepsa", "country": "spain", "locations": ("san roque", "algeciras"), "asset": "Moeve (Cepsa) San Roque"},
    {"company": "saudi aramco total", "country": "saudi arabia", "locations": ("jubail",), "asset": "Aramco/Total Jubail"},
    {"company": "satorp", "country": "saudi arabia", "locations": ("jubail",), "asset": "Aramco/Total Jubail"},
    {"company": "argent", "country": "netherlands", "locations": ("amsterdam",), "asset": "Argent Energy Amsterdam"},
    {"company": "argent energy", "country": "united kingdom", "locations": ("stanlow",), "asset": "Argent Energy Stanlow "},
    {"company": "argent energy", "country": "united kingdom", "locations": ("motherwell",), "asset": "Argent Energy Motherwell"},
    {"company": "montana renewables", "country": "united states", "locations": ("great falls",), "asset": "Montana Great Falls Phase 1"},
    {"company": "panjin penyao bioenergy", "country": "china", "locations": ("panjin",), "asset": "Panjin Pengyao Bioenergy Panjin"},
    {"company": "eco biochemical technology zhangjiagang", "country": "china", "locations": ("zhangjiagang",), "asset": "EcoCeres Zhangjiagang"},
    {"company": "shandong zhonghai fine chemical industry", "country": "china", "locations": ("binzhou",), "asset": "Zhongdiyou Binzhou"},
    {"company": "orlen poludnie", "country": "poland", "locations": ("trzebinia",), "asset": "PKN Orlen Trzebinia"},
    {"company": "taoyuan refinery", "country": "taiwan", "locations": ("taoyuan",), "asset": "CPC Taiwan Taoyuan"},
    {"company": "mangalore refinery and petrochemicals", "country": "india", "locations": ("mangalore", "mangaluru"), "asset": "M11/Indian Oil Mangalore"},
    {"company": "varopreem", "country": "sweden", "locations": ("lysekil",), "asset": "Varo (Preem) Lysekil Phase 1"},
    {"company": "varopreem", "country": "sweden", "locations": ("gothenburg",), "asset": "Varo (Preem) Gothenburg Phase 1"},
    {"company": "bio oils huelva", "country": "spain", "locations": ("huelva", "palos de la frontera"), "asset": "Moeve (Cepsa)/Apical Huelva Phase 1"},
    {"company": "petroleos del norte", "country": "spain", "locations": ("muskiz", "vizcaya"), "asset": "Repsol - Petronor Refinery Phase 1 Muskiz"},
    {"company": "blue whale bioenergy", "country": "china", "locations": ("jiaxing", "zhapu"), "asset": "Blue Whale Zhapu"},
    {"company": "hebei feitian future energy technology", "country": "china", "locations": ("xinji",), "asset": "Hebei Feitian Xinji"},
    {"company": "shandong huidong new energy", "country": "china", "locations": ("dongying", "kenli"), "asset": "Shandong Huidong New Energy  Dongying"},
    {"company": "lianyungang jiaao enproenergy", "country": "china", "locations": ("lianyungang", "guanyun"), "asset": "Jiaao/BP Zhejiang"},
    {"company": "sinopec ningbo zhenhai", "country": "china", "locations": ("ningbo", "zhenhai"), "asset": "Sinopec Ningbo"},
    {"company": "oberosterreichische biodiesel bulgarien", "country": "bulgaria", "locations": ("rousse", "ruse"), "asset": "Astra  Ruse"},
    {"company": "mil oil hellas", "country": "greece", "locations": ("serres",), "asset": "Miloil Hellas Serres"},
    {"company": "shell nederland raffinaderij", "country": "netherlands", "locations": ("rotterdam", "vondelingenplaat"), "asset": "Shell Rotterdam"},
    {"company": "kanola biofuels", "country": "netherlands", "locations": ("lexmond", "rotterdam"), "asset": "Kanola/Dutch Biofuels Rotterdam"},
    {"company": "argent energy", "country": "united kingdom", "locations": (), "asset": "Argent Energy Stanlow "},
    {"company": "saffaire sky energy", "country": "japan", "locations": ("sakai", "nishi ku", "chikko shinmachi"), "asset": "JGC/Cosmo (Saffaire Sky Energy) Osaka"},
    {"company": "repsol fuels", "country": "spain", "locations": ("cartagena", "murcia"), "asset": "Repsol Cartagena"},
    {"company": "repsol fuels", "country": "spain", "locations": ("puertollano",), "asset": "Repsol Puertollano"},
    {"company": "repsol fuels", "country": "spain", "locations": ("a coruna", "coruna"), "asset": "Repsol A Coruna"},
    {"company": "moeve", "country": "spain", "locations": ("san roque", "cadiz"), "asset": "Moeve (Cepsa) San Roque"},
    {"company": "biodiesel karnten", "country": "austria", "locations": ("arnoldstein",), "asset": "BioDiesel Kaernten Arnoldstein"},
    {"company": "revo international", "country": "japan", "locations": ("tahara", "midorigahama"), "asset": "Revo International Tahara-Shi"},
    {"company": "ina industrija nafte", "country": "croatia", "locations": ("kostrena", "urinj", "rijeka"), "asset": "INA/Chevron Rijeka"},
    {"company": "abu dhabi oil refining", "country": "united arab emirates", "locations": ("ruwais", "ruways", "al ruwais"), "asset": "Adnoc/ENI/OMV Ruwais"},
    {"company": "adnoc refining", "country": "united arab emirates", "locations": ("ruwais", "ruways", "al ruwais"), "asset": "Adnoc/ENI/OMV Ruwais"},
    {"company": "holborn europa raffinerie", "country": "germany", "locations": ("hamburg", "moorburger"), "asset": "Tamoil/Oilinvest Hamburg"},
    {"company": "repsol fuels", "country": "spain", "locations": ("la pobla de mafumet", "tarragona"), "asset": "Repsol - Tarragona Tarragona"},
    {"company": "aceites manuelita", "country": "colombia", "locations": ("meta", "yaguarito", "san carlos de guaroa"), "asset": "Manuelita"},
    {"company": "bunge argentina", "country": "argentina", "locations": ("san lorenzo", "puerto general san martin", "pgsm", "santa fe"), "asset": "Bunge San Lorenzo"},
    {"company": "energy absolute", "country": "thailand", "locations": ("kabinburi", "prachinburi", "tambon nongkee"), "asset": "Energy absolute Prachinburi"},
    {"company": "ea bio innovation", "country": "thailand", "locations": ("rayong", "map ta phut", "map taput"), "asset": "EA Bio Innovation  Map Ta Phut"},
    {"company": "explora", "country": "argentina", "locations": ("san lorenzo", "puerto general san martin", "santa fe"), "asset": "Explora San Lorenzo"},
    {"company": "future prelude", "country": "malaysia", "locations": ("pelabuhan klang", "port klang", "pulau indah"), "asset": "Future Prelude Port Klang"},
    {"company": "pt multimas nabati asahan", "country": "indonesia", "locations": ("serang", "kramatwatu", "bojonegara"), "asset": "Wilmar Serang (Kramatwatu/Bojonegara)"},
    {"company": "adm oilseeds germany", "country": "germany", "locations": ("mainz", "dammweg"), "asset": "ADM Mainz"},
    {"company": "adm hamburg", "country": "germany", "locations": ("hamburg", "nippoldstr"), "asset": "ADM Hamburg"},
    {"company": "bio d", "country": "colombia", "locations": ("facatativa", "cundinamarca", "mansilla"), "asset": "Bio D Facatativa"},
    {"company": "tpg oil gas", "country": "malaysia", "locations": ("pasir gudang", "tanjung langsat", "johor"), "asset": "KPN - TPG OIL & GAS Pasir Gudang"},
    {"company": "asb biodiesel", "country": "hong kong", "locations": ("hong kong", "tseung kwan o", "new territories", "n t"), "asset": "ASB Hong Kong"},
    {"company": "ldc argentina", "country": "argentina", "locations": ("general lagos", "santa fe", "ruta prov"), "asset": "Ldc General Lagos"},
    {"company": "pt smart", "country": "indonesia", "locations": ("tarjun", "kotabaru", "kalimantan"), "asset": "SMART Tbk Tarjun, Kotabaru (South Kalimantan)"},
    {"company": "pt sinarmas bio energy", "country": "indonesia", "locations": ("marunda", "bekasi", "tarumajaya"), "asset": "Sinarmas Bio Energy Marunda Center (Bekasi/Jakarta area)"},
    {"company": "pt ldc indonesia", "country": "indonesia", "locations": ("bandar lampung", "panjang", "way lunix", "way lunik"), "asset": "LDC Indonesia Bandar Lampung (Panjang)"},
    {"company": "ecomotion biodiesel", "country": "spain", "locations": ("montmelo", "barcelona"), "asset": "ecoMotion Barcelona"},
    {"company": "fytoenergeia", "country": "greece", "locations": ("serres", "paralimnio"), "asset": "New Energy (FytoEnergia) Serres"},
    {"company": "ai energy", "country": "thailand", "locations": ("samut sakhon", "samutsakorn", "khlong maduea"), "asset": "A I Energy Samutsakorn"},
    {"company": "bio d technology", "country": "united arab emirates", "locations": ("dubai", "jebel ali"), "asset": "BioD Technology"},
    {"company": "sd guthrie international biodiesel", "country": "malaysia", "locations": ("carey island", "kuala langat", "selangor"), "asset": "Sime Darby Oil Biodiesel Carey Island"},
    {"company": "turkiye petrol rafinerileri anonim sirketi", "country": "turkey", "locations": ("izmir", "aliaga"), "asset": "Tupras Izmir Phase 1"},
    {"company": "mert yag asfalt madencilik ve lastik geri donusum", "country": "turkey", "locations": ("nigde", "bor"), "asset": "Mert Oil Biodiesel Niğde"},
    {"company": "bioport", "country": "portugal", "locations": ("baltar", "parada"), "asset": "Bioportdiesel Baltar"},
    {"company": "shandong shangjia renewable resources technology", "country": "china", "locations": ("dongying",), "asset": "Shangdong Shangjia Unknown"},
)

# Target-specific thresholds. They are intentionally more permissive than the
# general matcher, but still require company evidence. A city match on its own
# can never create a target asset assignment.
TARGET_MATCH_CONFIG: dict[str, float] = {
    "min_authoritative_company": 90.0,
    "min_review_company": 86.0,
    "min_fuzzy_company": 88.0,
    "min_location_for_review_method": 72.0,
    "min_location_for_distinctive": 90.0,
    "min_different_site_margin": 8.0,
    "company_weight": 0.58,
    "location_weight": 0.37,
    "technology_bonus": 5.0,
}


# These methods are allowed to produce an automatic match. Methods based on
# acronyms, token similarity or company text extracted only from Asset
# Identifier are deliberately review-only.
SAFE_AUTO_COMPANY_METHODS = frozenset(
    {
        "MANUAL_COMPANY_ALIAS",
        "EXACT_COMPANY_NAME",
        "COMPANY_SUBSTRING",
    }
)

SAFE_AUTO_ASSET_COMPANY_METHODS = SAFE_AUTO_COMPANY_METHODS

EXACT_COMPANY_METHODS = frozenset(
    {
        "MANUAL_COMPANY_ALIAS",
        "EXACT_COMPANY_NAME",
    }
)

EXACT_LOCATION_METHODS = frozenset(
    {
        "GST_CITY_IN_CERTIFICATE_ADDRESS",
        "ASSET_LOCATION_IN_CERTIFICATE_ADDRESS",
    }
)

REVIEW_ONLY_COMPANY_METHODS = frozenset(
    {
        "ASSET_COMPANY_EXACT",
        "ASSET_COMPANY_SUBSTRING",
        "COMPANY_PREFIX_RELATION",
        "ASSET_COMPANY_PREFIX_RELATION",
        "COMPANY_FORM_MATCH",
        "ASSET_COMPANY_FORM_MATCH",
        "DISTINCTIVE_COMPANY_TOKENS",
    }
)


MATCH_CONFIG: dict[str, float] = {
    # Company resolution
    "min_company_score": 86.0,
    "min_company_margin": 8.0,
    "strong_company_score": 90.0,

    # Site resolution
    "strong_location_score": 90.0,
    "exact_location_score": 95.0,
    "location_conflict_score": 55.0,
    "min_site_margin": 10.0,

    # Non-processing/trader/office certificates are intentionally stricter.
    "non_processing_min_location": 98.0,
    "non_processing_min_site_margin": 15.0,

    # Used only for ranking assets after the company has been locked.
    "asset_company_weight": 0.50,
    "location_weight": 0.50,
}


_EMPTY_VALUES = {
    "",
    "nan",
    "none",
    "null",
    "unknown",
    "n/a",
    "na",
    "n.a",
    "n.a.",
    "-",
}

_EXCEL_ERROR_VALUES = {
    "#VALUE!",
    "#N/A",
    "#REF!",
    "#DIV/0!",
    "#NAME?",
    "#NUM!",
    "#NULL!",
}

_COMMON_LEGAL_SUFFIXES = {
    "ltd",
    "limited",
    "llc",
    "inc",
    "incorporated",
    "corp",
    "corporation",
    "co",
    "company",
    "plc",
    "gmbh",
    "bv",
    "b v",
    "nv",
    "n v",
    "sa",
    "s a",
    "sau",
    "s a u",
    "sas",
    "sasu",
    "spa",
    "s p a",
    "srl",
    "s r l",
    "sro",
    "s r o",
    "as",
    "a s",
    "sl",
    "s l",
    "pte",
    "pte ltd",
    "pty",
    "pty ltd",
    "sdn bhd",
    "bhd",
    "oyj",
    "ab",
    "ag",
    "kg",
    "kgaa",
    "lp",
    "l p",
    "ulc",
    "zrt",
    "kft",
}


# These words are common across many businesses and cannot be the only reason
# two companies are considered related.
_COMPANY_GENERIC_TOKENS = {
    "and",
    "agricultural",
    "agriculture",
    "agro",
    "agroindustrial",
    "agroindustria",
    "bio",
    "bioenergy",
    "biochemical",
    "biotechnology",
    "builder",
    "builders",
    "chemical",
    "chemicals",
    "component",
    "components",
    "construction",
    "development",
    "engineering",
    "environment",
    "environmental",
    "energy",
    "energies",
    "food",
    "foods",
    "fuel",
    "fuels",
    "future",
    "global",
    "green",
    "group",
    "holding",
    "holdings",
    "industrial",
    "industries",
    "industry",
    "international",
    "logistics",
    "material",
    "materials",
    "new",
    "oil",
    "oils",
    "plant",
    "power",
    "product",
    "products",
    "production",
    "recycling",
    "refinery",
    "refining",
    "renewable",
    "renewables",
    "resource",
    "resources",
    "service",
    "services",
    "solution",
    "solutions",
    "system",
    "systems",
    "technology",
    "technologies",
    "trading",
    "waste",
}

_LOCATION_NOISE_TOKENS = {
    "city",
    "province",
    "district",
    "county",
    "municipality",
    "prefecture",
    "state",
    "region",
    "shi",
    "no",
    "na",
}

_ASSET_NOISE_TOKENS = {
    "asset",
    "facility",
    "hefa",
    "hvo",
    "plant",
    "project",
    "rd",
    "refinery",
    "saf",
    "site",
    "unit",
}

_ASSET_VARIANT_RE = re.compile(
    r"\b(?:"
    r"phase\s*(?:\d+|[ivx]+)"
    r"|saf\s+expansion"
    r"|rd\s+expansion"
    r"|expansion"
    r"|project\s*\d*"
    r"|unit\s*\d+"
    r")\b",
    flags=re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Basic normalisation
# ---------------------------------------------------------------------------


def _safe_match_text(value: Any) -> str:
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    text = str(value).strip()
    if not text:
        return ""
    if text.upper() in _EXCEL_ERROR_VALUES:
        return ""
    if text.lower() in _EMPTY_VALUES:
        return ""
    return text


_CHAR_TRANSLITERATION = str.maketrans(
    {
        # Characters that Unicode NFKD does not reliably decompose to ASCII.
        # These occur frequently in European company and site names.
        "ł": "l",
        "Ł": "L",
        "đ": "d",
        "Đ": "D",
        "ð": "d",
        "Ð": "D",
        "þ": "th",
        "Þ": "Th",
        "æ": "ae",
        "Æ": "AE",
        "œ": "oe",
        "Œ": "OE",
        "ø": "o",
        "Ø": "O",
        "ı": "i",
        "İ": "I",
        "ß": "ss",
    }
)


def _strip_accents(value: str) -> str:
    value = value.translate(_CHAR_TRANSLITERATION)
    return "".join(
        char
        for char in unicodedata.normalize("NFKD", value)
        if not unicodedata.combining(char)
    )


def _normalize_match_text(value: Any) -> str:
    text = _safe_match_text(value)
    if not text:
        return ""

    text = _strip_accents(text.lower())
    text = text.replace("&amp;", " and ").replace("&", " and ")
    text = text.replace("，", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _build_legal_suffixes() -> list[str]:
    suffixes = set(_COMMON_LEGAL_SUFFIXES)

    for suffix in LEGAL_SUFFIXES:
        normalized = _normalize_match_text(suffix)
        # One-character values are too dangerous for token removal.
        if len(normalized) >= 2:
            suffixes.add(normalized)

    return sorted(suffixes, key=len, reverse=True)


_MATCH_LEGAL_SUFFIXES = _build_legal_suffixes()


def _normalize_company(value: Any) -> str:
    text = _normalize_match_text(value)
    if not text:
        return ""

    for suffix in _MATCH_LEGAL_SUFFIXES:
        text = re.sub(rf"\b{re.escape(suffix)}\b", " ", text)

    text = re.sub(r"\bthe\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _normalize_location(value: Any) -> str:
    text = _normalize_match_text(value)
    if not text:
        return ""

    return " ".join(
        token
        for token in text.split()
        if token not in _LOCATION_NOISE_TOKENS
    )


def _normalize_country(value: Any) -> str:
    text = _normalize_match_text(value)
    if not text:
        return ""
    return COUNTRY_MATCH_ALIASES.get(text, text)


def _is_blocked_company_asset(company: Any, asset_identifier: Any) -> bool:
    company_norm = _normalize_company(company)
    asset_norm = _normalize_match_text(asset_identifier)
    return any(
        company_norm == _normalize_company(blocked_company)
        and asset_norm == _normalize_match_text(blocked_asset)
        for blocked_company, blocked_asset in BLOCKED_ISCC_ASSET_MATCHES
    )


def _normalised_tokens(value: Any) -> list[str]:
    return [
        token
        for token in _normalize_match_text(value).split()
        if token and not token.isdigit()
    ]


def _location_token_set(value: Any) -> set[str]:
    return {
        token
        for token in _normalize_location(value).split()
        if len(token) >= 3 and not token.isdigit()
    }


def _company_tokens(
    value: Any,
    *,
    location_tokens: set[str] | None = None,
    meaningful_only: bool = False,
) -> list[str]:
    excluded_locations = location_tokens or set()

    tokens = [
        token
        for token in _normalize_company(value).split()
        if (
            len(token) >= 2
            and not token.isdigit()
            and token not in excluded_locations
        )
    ]

    if meaningful_only:
        # Do not fall back to generic tokens. The previous fallback allowed
        # phrases such as "renewable energy" to become company evidence.
        return [
            token
            for token in tokens
            if token not in _COMPANY_GENERIC_TOKENS
        ]

    return tokens


def _company_forms(
    value: Any,
    *,
    location_tokens: set[str] | None = None,
) -> set[str]:
    """Return conservative compact/acronym company forms.

    Two-character acronyms are deliberately excluded because the audit found
    collisions such as GE, JC and KB. Three-character forms such as DGD, P66
    and ST1 remain useful as review candidates or confirmed aliases.
    """

    normalized = _normalize_company(value)
    if not normalized:
        return set()

    excluded_locations = location_tokens or set()
    raw_tokens = [
        token
        for token in normalized.split()
        if token not in excluded_locations
    ]

    alpha_tokens = [
        token
        for token in raw_tokens
        if token.isalpha() and token not in _COMPANY_GENERIC_TOKENS
    ]
    digit_tokens = [token for token in raw_tokens if token.isdigit()]
    alphanumeric_tokens = [
        token
        for token in raw_tokens
        if any(char.isalpha() for char in token)
        and any(char.isdigit() for char in token)
        and token not in _COMPANY_GENERIC_TOKENS
    ]

    forms: set[str] = set()

    meaningful_compact = "".join(alpha_tokens + alphanumeric_tokens + digit_tokens)
    if len(meaningful_compact) >= 5 or (
        len(meaningful_compact) >= 3
        and any(char.isdigit() for char in meaningful_compact)
    ):
        forms.add(meaningful_compact)

    if alpha_tokens:
        acronym = "".join(token[0] for token in alpha_tokens) + "".join(digit_tokens)
        if len(acronym) >= 3:
            forms.add(acronym)

    # Preserve an existing short alphanumeric brand such as P66 or ST1.
    for token in alphanumeric_tokens:
        if len(token) >= 3:
            forms.add(token)

    return forms


# ---------------------------------------------------------------------------
# GST preparation
# ---------------------------------------------------------------------------


def _asset_company_text(
    asset_identifier: Any,
    city: Any,
    territory: Any,
    country_location_tokens: set[str] | None = None,
) -> str:
    asset = _normalize_company(asset_identifier)
    if not asset:
        return ""

    location_tokens = (
        _location_token_set(city)
        | _location_token_set(territory)
        | (country_location_tokens or set())
    )

    keep: list[str] = []
    for token in asset.split():
        if token.isdigit():
            continue
        if token in location_tokens:
            continue
        if token in _ASSET_NOISE_TOKENS:
            continue
        if token in {"phase", "expansion"}:
            continue
        keep.append(token)

    text = " ".join(keep)
    text = _ASSET_VARIANT_RE.sub(" ", text)
    return re.sub(r"\s+", " ", text).strip()


def _asset_location_text(asset_identifier: Any, asset_company_norm: Any) -> str:
    asset = _normalize_location(asset_identifier)
    company_tokens = set(_company_tokens(asset_company_norm))

    keep: list[str] = []
    for token in asset.split():
        if token.isdigit():
            continue
        if token in company_tokens:
            continue
        if token in _ASSET_NOISE_TOKENS:
            continue
        if token in {"phase", "expansion"}:
            continue
        keep.append(token)

    text = " ".join(keep)
    text = _ASSET_VARIANT_RE.sub(" ", text)
    return re.sub(r"\s+", " ", text).strip()


def _asset_site_key(asset_identifier: Any) -> str:
    text = _normalize_match_text(asset_identifier)
    text = _ASSET_VARIANT_RE.sub(" ", text)
    text = re.sub(r"\b(?:phase|expansion)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _is_variant_asset(asset_identifier: Any) -> bool:
    return bool(_ASSET_VARIANT_RE.search(_normalize_match_text(asset_identifier)))


def _build_location_tokens_by_country(
    gst: pd.DataFrame,
) -> dict[str, set[str]]:
    location_tokens_by_country: dict[str, set[str]] = {}

    for country, group in gst.groupby("__country_norm", sort=False):
        if not country:
            continue

        tokens = set(_normalize_country(country).split())
        for column in ["__city_norm", "__pbi_city_norm", "__territory_norm"]:
            if column not in group.columns:
                continue
            for value in group[column]:
                tokens.update(_location_token_set(value))

        # Short generic direction words are not useful as location exclusions.
        tokens -= {"north", "south", "east", "west"}
        location_tokens_by_country[country] = tokens

    return location_tokens_by_country


def _prepare_gst_asset_matching(
    gst_df: pd.DataFrame,
    allowed_statuses: Iterable[str] | None = ("Started Up",),
) -> tuple[pd.DataFrame, dict[str, set[str]]]:
    """Prepare GST rows for matching.

    By default only GoldenSource rows whose ``Probability of success`` is
    ``Started Up`` are eligible. Pass ``allowed_statuses=None`` to use every GST
    row, or pass another iterable of status labels where required.
    """

    required = {
        "Asset Identifier",
        "Company/Producer",
        "Company/Producer Short Name",
        "Territory",
    }
    missing = required - set(gst_df.columns)
    if missing:
        raise KeyError(f"Missing GST asset matching columns: {missing}")

    source = gst_df.copy()
    if allowed_statuses is not None:
        if "Probability of success" not in source.columns:
            raise KeyError(
                "GST status filtering is enabled, but the GoldenSource column "
                "'Probability of success' is missing. Pass allowed_statuses=None "
                "only when the input has already been intentionally filtered."
            )

        allowed = {
            _normalize_match_text(status)
            for status in allowed_statuses
            if _normalize_match_text(status)
        }
        if not allowed:
            raise ValueError("allowed_statuses must contain at least one status")

        source = source[
            source["Probability of success"]
            .apply(_normalize_match_text)
            .isin(allowed)
        ].copy()
        if source.empty:
            raise ValueError(
                "No GST rows remain after filtering Probability of success to: "
                + ", ".join(sorted(allowed_statuses))
            )

    possible_country_columns = ["COUNTRY_NM", "PBI Country", "Country"]
    available_country_columns = [
        column for column in possible_country_columns if column in source.columns
    ]
    if not available_country_columns:
        raise KeyError(
            "No usable GST country column found. Expected one of: "
            "'COUNTRY_NM', 'PBI Country', 'Country'."
        )

    columns = [
        "Asset Identifier",
        "Company/Producer",
        "Company/Producer Short Name",
        "Territory",
    ]
    for optional in [
        "City",
        "PBI City",
        "Production Technology",
        "Probability of success",
    ]:
        if optional in source.columns:
            columns.append(optional)
    columns.extend(available_country_columns)
    columns = list(dict.fromkeys(columns))

    gst = source[columns].copy()
    if "City" not in gst.columns:
        gst["City"] = ""
    if "PBI City" not in gst.columns:
        gst["PBI City"] = ""
    if "Production Technology" not in gst.columns:
        gst["Production Technology"] = ""

    gst["Asset Identifier"] = (
        gst["Asset Identifier"].fillna("").astype(str).str.strip()
    )
    gst = gst[gst["Asset Identifier"].ne("")].copy()

    def get_country(row: pd.Series) -> str:
        for column in possible_country_columns:
            if column not in row.index:
                continue
            value = _safe_match_text(row[column])
            if value:
                return value
        return ""

    gst["__country_raw"] = gst.apply(get_country, axis=1)
    gst["__country_norm"] = gst["__country_raw"].apply(_normalize_country)
    gst["__producer_norm"] = gst["Company/Producer"].apply(_normalize_company)
    gst["__short_norm"] = gst["Company/Producer Short Name"].apply(
        _normalize_company
    )
    gst["__city_norm"] = gst["City"].apply(_normalize_location)
    gst["__pbi_city_norm"] = gst["PBI City"].apply(_normalize_location)
    gst["__territory_norm"] = gst["Territory"].apply(_normalize_location)
    gst["__asset_norm"] = gst["Asset Identifier"].apply(_normalize_match_text)
    gst["__target_technologies"] = gst["Production Technology"].apply(
        _gst_target_technologies
    )

    location_tokens_by_country = _build_location_tokens_by_country(gst)

    gst["__asset_company_norm"] = gst.apply(
        lambda row: _asset_company_text(
            row["Asset Identifier"],
            row["City"],
            row["Territory"],
            location_tokens_by_country.get(row["__country_norm"], set()),
        ),
        axis=1,
    )
    gst["__asset_location_norm"] = gst.apply(
        lambda row: _asset_location_text(
            row["Asset Identifier"],
            row["__asset_company_norm"],
        ),
        axis=1,
    )
    gst["__site_key"] = gst["Asset Identifier"].apply(_asset_site_key)
    gst["__is_variant"] = gst["Asset Identifier"].apply(_is_variant_asset)

    gst["__company_key"] = np.where(
        gst["__short_norm"].ne(""),
        gst["__short_norm"],
        np.where(
            gst["__producer_norm"].ne(""),
            gst["__producer_norm"],
            gst["__asset_company_norm"],
        ),
    )

    return gst, location_tokens_by_country



# ---------------------------------------------------------------------------
# Confirmed started-asset direct matches
# ---------------------------------------------------------------------------


def _build_confirmed_started_match_index(
    gst: pd.DataFrame,
    rules: tuple[tuple[str, str, str, tuple[str, ...], str], ...],
    *,
    strict: bool = True,
) -> dict[tuple[str, str], list[dict[str, Any]]]:
    """Build and validate the exact ISCC company/city override index.

    The source rules are intentionally raw display strings so they remain easy
    to audit. They are normalised only when the index is built. Every target is
    resolved to the exact spelling currently present in the GST. A missing GST
    target raises an error by default rather than silently producing a stale or
    incorrect hardcoded match.
    """

    asset_rows_by_norm: dict[str, pd.Series] = {}
    for _, asset_row in gst.iterrows():
        asset_norm = _safe_match_text(asset_row.get("__asset_norm", ""))
        if asset_norm:
            asset_rows_by_norm.setdefault(asset_norm, asset_row)

    index: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    missing_assets: list[str] = []

    for company, city, asset, processing_terms, reason in rules:
        company_norm = _normalize_company(company)
        city_norm = _normalize_location(city)
        asset_norm = _normalize_match_text(asset)

        # Exact direct matches must contain both identifiers. Blank-location
        # company rules are deliberately rejected because they can map the
        # same company to the wrong site.
        if not company_norm or not city_norm or not asset_norm:
            continue

        asset_row = asset_rows_by_norm.get(asset_norm)
        if asset_row is None:
            missing_assets.append(asset)
            continue

        normalized_terms = tuple(
            dict.fromkeys(
                term_norm
                for term in processing_terms
                if (term_norm := _normalize_match_text(term))
            )
        )

        entry = {
            "asset": _safe_match_text(asset_row["Asset Identifier"]),
            "asset_row": asset_row,
            "processing_terms": normalized_terms,
            "reason": reason,
            "source_company": company,
            "source_city": city,
        }
        index[(company_norm, city_norm)].append(entry)

    if missing_assets and strict:
        unique_missing = sorted(set(missing_assets))
        sample = ", ".join(unique_missing[:10])
        more = "" if len(unique_missing) <= 10 else f" (+{len(unique_missing) - 10} more)"
        raise KeyError(
            "Confirmed started-asset rules reference GST Asset Identifiers "
            f"that are no longer present: {sample}{more}"
        )

    # Reject conflicting generic rules for the same exact company/city key.
    conflicts: list[str] = []
    for key, entries in index.items():
        generic_assets = {
            _normalize_match_text(entry["asset"])
            for entry in entries
            if not entry["processing_terms"]
        }
        if len(generic_assets) > 1:
            conflicts.append(f"{key[0]} | {key[1]}")

    if conflicts:
        raise ValueError(
            "Conflicting generic confirmed matches for: "
            + "; ".join(conflicts[:10])
        )

    return index


def _find_confirmed_started_asset_match(
    *,
    company: Any,
    city: Any,
    scope: Any,
    processing_unit_type: Any,
    is_processing: bool,
    match_index: dict[tuple[str, str], list[dict[str, Any]]],
) -> dict[str, Any] | None:
    """Return an exact curated match for one Processing Unit row, if present."""

    if not is_processing:
        return None

    key = (_normalize_company(company), _normalize_location(city))
    candidates = match_index.get(key, [])
    if not candidates:
        return None

    processing_blob = _normalize_match_text(
        f"{_safe_match_text(scope)} {_safe_match_text(processing_unit_type)}"
    )

    # Technology-specific rules take precedence over generic company/site
    # rules. This is currently used for ORLEN Płock, where HEFA/HVO and
    # co-processing have separate GST Asset Identifiers.
    specific: list[tuple[int, int, dict[str, Any]]] = []
    for position, candidate in enumerate(candidates):
        terms = candidate["processing_terms"]
        matched_terms = [term for term in terms if term in processing_blob]
        if matched_terms:
            # Prefer more evidence and then the longest explicit term. The
            # source order is retained as a deterministic final tiebreaker.
            strength = (len(matched_terms) * 100) + max(map(len, matched_terms))
            specific.append((strength, -position, candidate))

    if specific:
        specific.sort(key=lambda item: (item[0], item[1]), reverse=True)
        return specific[0][2]

    generic = [candidate for candidate in candidates if not candidate["processing_terms"]]
    if not generic:
        return None

    unique_assets = {
        _normalize_match_text(candidate["asset"])
        for candidate in generic
    }
    if len(unique_assets) != 1:
        return None

    return generic[0]


def _apply_confirmed_started_asset_output(
    output: dict[str, Any],
    confirmed_match: dict[str, Any],
    *,
    company_raw: Any,
    city_raw: Any,
) -> None:
    """Populate the standard diagnostic object for a confirmed direct match."""

    asset_row = confirmed_match["asset_row"]
    asset = confirmed_match["asset"]
    display_company = (
        _safe_match_text(asset_row.get("Company/Producer Short Name", ""))
        or _safe_match_text(asset_row.get("Company/Producer", ""))
    )
    evidence = (
        f"{_safe_match_text(company_raw)} | {_safe_match_text(city_raw)}"
    )

    output.update(
        {
            "Asset_Identifier": asset,
            "Suggested_Asset_Identifier": asset,
            "Match_Found": 1,
            "Match_Status": "Matched",
            "Match_Confidence": "High",
            "Match_Method": "CONFIRMED_STARTED_ASSET_OVERRIDE",
            "Auto_Match_Eligible": 1,
            "Matched_GST_Company": display_company,
            "Best_GST_Company_Candidate": display_company,
            "Company_Match_Method": "CONFIRMED_STARTED_ASSET_OVERRIDE",
            "Company_Match_Evidence": evidence,
            "Matched_Territory": _safe_match_text(asset_row.get("Territory", "")),
            "Company_Score": 100.0,
            "Company_Score_Margin": 100.0,
            "Asset_Company_Score": 100.0,
            "Asset_Company_Match_Method": "CONFIRMED_STARTED_ASSET_OVERRIDE",
            "Asset_Company_Match_Evidence": evidence,
            "Location_Score": 100.0,
            "Location_Match_Method": "CONFIRMED_STARTED_ASSET_OVERRIDE",
            "Overall_Score": 100.0,
            "Score_Margin": 100.0,
            "Candidate_Count": 1,
            "Candidate_Site_Count": 1,
            "Matched_Site_Key": _safe_match_text(asset_row.get("__site_key", "")),
            "Runner_Up_Asset": "",
            "Runner_Up_Score": 0.0,
            "Review_Reason": "",
            "Target_Force_Match_Applied": 1,
            "Phase_Fallback_Applied": int(bool(asset_row.get("__is_variant", False))),
            "Target_Coverage_Status": "CONFIRMED_STARTED_ASSET_OVERRIDE",
        }
    )


def _build_confirmed_address_match_index(
    gst: pd.DataFrame,
    rules: tuple[
        tuple[str, str, tuple[str, ...], str, tuple[str, ...], str], ...
    ],
    *,
    strict: bool = True,
) -> dict[tuple[str, str], list[dict[str, Any]]]:
    """Validate and index confirmed company/country/address mappings."""

    asset_rows_by_norm: dict[str, pd.Series] = {}
    for _, asset_row in gst.iterrows():
        asset_norm = _safe_match_text(asset_row.get("__asset_norm", ""))
        if asset_norm:
            asset_rows_by_norm.setdefault(asset_norm, asset_row)

    index: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    missing_assets: list[str] = []

    for company, country, locations, asset, processing_terms, reason in rules:
        company_norm = _normalize_company(company)
        country_norm = _normalize_country(country)
        asset_norm = _normalize_match_text(asset)
        if not company_norm or not country_norm or not asset_norm:
            continue

        asset_row = asset_rows_by_norm.get(asset_norm)
        if asset_row is None:
            missing_assets.append(asset)
            continue

        location_terms = tuple(
            dict.fromkeys(
                value
                for term in locations
                if (value := _normalize_location(term))
            )
        )
        if not location_terms:
            continue

        required_terms = tuple(
            dict.fromkeys(
                value
                for term in processing_terms
                if (value := _normalize_match_text(term))
            )
        )
        index[(company_norm, country_norm)].append(
            {
                "asset": _safe_match_text(asset_row["Asset Identifier"]),
                "asset_row": asset_row,
                "locations": location_terms,
                "processing_terms": required_terms,
                "reason": reason,
            }
        )

    if missing_assets and strict:
        unique_missing = sorted(set(missing_assets))
        raise KeyError(
            "Confirmed address rules reference Started Up GST Asset Identifiers "
            "that are no longer present: " + ", ".join(unique_missing[:15])
        )

    return index


def _find_confirmed_address_asset_match(
    *,
    company: Any,
    country: Any,
    city: Any,
    certificate_holder: Any,
    scope: Any,
    processing_unit_type: Any,
    is_processing: bool,
    match_index: dict[tuple[str, str], list[dict[str, Any]]],
) -> dict[str, Any] | None:
    if not is_processing:
        return None

    key = (_normalize_company(company), _normalize_country(country))
    candidates = match_index.get(key, [])
    if not candidates:
        return None

    location_blob = _normalize_location(
        f"{_safe_match_text(city)} {_safe_match_text(certificate_holder)}"
    )
    processing_blob = _normalize_match_text(
        f"{_safe_match_text(scope)} {_safe_match_text(processing_unit_type)}"
    )

    eligible: list[tuple[int, int, dict[str, Any]]] = []
    for position, candidate in enumerate(candidates):
        matched_locations = [
            term for term in candidate["locations"] if _contains_phrase(term, location_blob)
        ]
        if not matched_locations:
            continue

        required_terms = candidate["processing_terms"]
        if required_terms and not any(term in processing_blob for term in required_terms):
            continue

        strength = max(map(len, matched_locations)) + (100 if required_terms else 0)
        eligible.append((strength, -position, candidate))

    if not eligible:
        return None

    eligible.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return eligible[0][2]


# ---------------------------------------------------------------------------
# Company matching
# ---------------------------------------------------------------------------


def _build_company_records(
    gst_subset: pd.DataFrame,
    location_tokens: set[str],
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    valid = gst_subset[gst_subset["__company_key"].ne("")]

    for company_key, group in valid.groupby("__company_key", sort=False):
        authoritative_variants = {
            value
            for column in ["__producer_norm", "__short_norm"]
            for value in group[column]
            if value
        }
        asset_variants = {
            value
            for value in group["__asset_company_norm"]
            if value and value not in authoritative_variants
        }
        all_variants = authoritative_variants | asset_variants

        short_names = [
            _safe_match_text(value)
            for value in group["Company/Producer Short Name"]
            if _safe_match_text(value)
        ]
        producer_names = [
            _safe_match_text(value)
            for value in group["Company/Producer"]
            if _safe_match_text(value)
        ]

        display_name = (
            short_names[0]
            if short_names
            else producer_names[0]
            if producer_names
            else company_key
        )

        records.append(
            {
                "company_key": company_key,
                "authoritative_variants": tuple(sorted(authoritative_variants)),
                "asset_variants": tuple(sorted(asset_variants)),
                "all_variants": tuple(sorted(all_variants)),
                "authoritative_forms": set().union(
                    *(
                        _company_forms(value, location_tokens=location_tokens)
                        for value in authoritative_variants
                    )
                )
                if authoritative_variants
                else set(),
                "asset_forms": set().union(
                    *(
                        _company_forms(value, location_tokens=location_tokens)
                        for value in asset_variants
                    )
                )
                if asset_variants
                else set(),
                "tokens": set().union(
                    *(
                        set(
                            _company_tokens(
                                value,
                                location_tokens=location_tokens,
                                meaningful_only=True,
                            )
                        )
                        for value in all_variants
                    )
                )
                if all_variants
                else set(),
                "display_name": display_name,
            }
        )

    return records


def _build_token_document_frequency(
    records: Iterable[dict[str, Any]],
) -> Counter[str]:
    frequencies: Counter[str] = Counter()
    for record in records:
        frequencies.update(set(record["tokens"]))
    return frequencies


def _build_company_search_index(
    records: list[dict[str, Any]],
) -> dict[str, Any]:
    """Build a lightweight retrieval index before expensive scoring."""

    variant_to_indices: dict[str, set[int]] = defaultdict(set)
    token_to_indices: dict[str, set[int]] = defaultdict(set)
    prefix_to_indices: dict[str, set[int]] = defaultdict(set)
    form_to_indices: dict[str, set[int]] = defaultdict(set)

    for index, record in enumerate(records):
        for variant in record["all_variants"]:
            variant_to_indices[variant].add(index)

        for token in record["tokens"]:
            token_to_indices[token].add(index)
            if len(token) >= 5:
                prefix_to_indices[token[:3]].add(index)

        for form in record["authoritative_forms"] | record["asset_forms"]:
            form_to_indices[form].add(index)

    return {
        "records": records,
        "variant_to_indices": variant_to_indices,
        "token_to_indices": token_to_indices,
        "prefix_to_indices": prefix_to_indices,
        "form_to_indices": form_to_indices,
    }


def _candidate_company_indices(
    query_norm: str,
    search_index: dict[str, Any],
    location_tokens: set[str],
    alias_target: str = "",
) -> set[int]:
    """Retrieve plausible companies without scanning the entire country."""

    candidate_indices: set[int] = set()

    for value in [query_norm, alias_target]:
        if not value:
            continue

        candidate_indices.update(
            search_index["variant_to_indices"].get(value, set())
        )

        for token in _company_tokens(
            value,
            location_tokens=location_tokens,
            meaningful_only=True,
        ):
            candidate_indices.update(
                search_index["token_to_indices"].get(token, set())
            )
            if len(token) >= 5:
                candidate_indices.update(
                    search_index["prefix_to_indices"].get(token[:3], set())
                )

        for form in _company_forms(
            value,
            location_tokens=location_tokens,
        ):
            candidate_indices.update(
                search_index["form_to_indices"].get(form, set())
            )

    return candidate_indices


def _token_relation(left: str, right: str) -> tuple[float, str]:
    if left == right:
        return 100.0, "EXACT_TOKEN"

    minimum_length = min(len(left), len(right))
    if minimum_length >= 5 and (
        left.startswith(right) or right.startswith(left)
    ):
        return 96.0, "PREFIX_TOKEN"

    if any(char.isdigit() for char in left + right):
        if minimum_length >= 3 and (
            left.startswith(right) or right.startswith(left)
        ):
            return 94.0, "ALPHANUMERIC_PREFIX"

    if minimum_length >= 5:
        score = float(fuzz.ratio(left, right))
        if score >= 90.0:
            return score, "FUZZY_TOKEN"

    return 0.0, ""


def _is_informative_token(
    token: str,
    token_df: Counter[str],
    company_count: int,
) -> bool:
    if len(token) < 4 and not any(char.isdigit() for char in token):
        return False
    if token in _COMPANY_GENERIC_TOKENS:
        return False

    rare_limit = max(3, int(round(company_count * 0.01)))
    return token_df.get(token, 0) <= rare_limit


def _safe_substring_relation(
    query_norm: str,
    candidate_norm: str,
    location_tokens: set[str],
) -> tuple[bool, str]:
    """Return a safe whole-token/prefix subset relationship.

    This intentionally does not concatenate all tokens. Concatenation allowed
    cross-boundary errors such as ``massara`` + ``s`` looking like ``saras``.
    """

    query_all = _company_tokens(
        query_norm,
        location_tokens=location_tokens,
        meaningful_only=False,
    )
    candidate_all = _company_tokens(
        candidate_norm,
        location_tokens=location_tokens,
        meaningful_only=False,
    )
    query_distinctive = _company_tokens(
        query_norm,
        location_tokens=location_tokens,
        meaningful_only=True,
    )
    candidate_distinctive = _company_tokens(
        candidate_norm,
        location_tokens=location_tokens,
        meaningful_only=True,
    )

    if not query_distinctive or not candidate_distinctive:
        return False, ""

    def sequence_contains(shorter: list[str], longer: list[str]) -> bool:
        if len(shorter) > len(longer):
            return False
        return any(
            longer[index : index + len(shorter)] == shorter
            for index in range(len(longer) - len(shorter) + 1)
        )

    # Exact whole-token phrase containment, e.g. "ecoceres" within
    # "ecoceres renewable fuels".
    if sequence_contains(candidate_all, query_all):
        if any(len(token) >= 4 for token in candidate_distinctive):
            return True, " ".join(candidate_all)

    if sequence_contains(query_all, candidate_all):
        if any(len(token) >= 4 for token in query_distinctive):
            return True, " ".join(query_all)

    # Conservative token-prefix relation, e.g. Total -> TotalEnergies or
    # Jiaao -> JiaaoBP. Every token in the shorter distinctive side must have
    # an exact or long-prefix counterpart.
    if len(query_distinctive) <= len(candidate_distinctive):
        shorter = query_distinctive
        longer = candidate_distinctive
    else:
        shorter = candidate_distinctive
        longer = query_distinctive

    relations: list[tuple[str, str]] = []
    for short_token in shorter:
        match = next(
            (
                long_token
                for long_token in longer
                if (
                    short_token == long_token
                    or (
                        min(len(short_token), len(long_token)) >= 5
                        and (
                            short_token.startswith(long_token)
                            or long_token.startswith(short_token)
                        )
                    )
                )
            ),
            "",
        )
        if not match:
            return False, ""
        relations.append((short_token, match))

    if not relations:
        return False, ""

    if not any(min(len(left), len(right)) >= 4 for left, right in relations):
        return False, ""

    evidence = ", ".join(f"{left}~{right}" for left, right in relations)
    return True, evidence


def _score_distinctive_tokens(
    query_norm: str,
    record: dict[str, Any],
    token_df: Counter[str],
    company_count: int,
    location_tokens: set[str],
) -> tuple[float, str]:
    query_tokens = _company_tokens(
        query_norm,
        location_tokens=location_tokens,
        meaningful_only=True,
    )
    candidate_tokens = set(record["tokens"])

    if not query_tokens or not candidate_tokens:
        return 0.0, ""

    matched: list[tuple[str, str, float, str]] = []
    anchors: list[tuple[str, str, float, str]] = []

    for query_token in query_tokens:
        best_candidate = ""
        best_score = 0.0
        best_method = ""

        for candidate_token in candidate_tokens:
            score, method = _token_relation(query_token, candidate_token)
            if score > best_score:
                best_candidate = candidate_token
                best_score = score
                best_method = method

        if best_score >= 90.0:
            match = (query_token, best_candidate, best_score, best_method)
            matched.append(match)

            informative = _is_informative_token(
                query_token,
                token_df,
                company_count,
            ) or _is_informative_token(
                best_candidate,
                token_df,
                company_count,
            )

            # A spelling-only fuzzy hit cannot be the sole company anchor.
            if informative and best_method in {
                "EXACT_TOKEN",
                "PREFIX_TOKEN",
                "ALPHANUMERIC_PREFIX",
            }:
                anchors.append(match)

    if not anchors:
        return 0.0, ""

    coverage = len(matched) / max(len(query_tokens), 1)
    anchor_strength = max(item[2] for item in anchors) / 100.0
    # Review-only method: keep its score below authoritative substring matches.
    score = min(80.0 + (10.0 * coverage) + (2.0 * anchor_strength), 92.0)

    evidence = ", ".join(
        f"{left}~{right}" for left, right, _, _ in matched
    )
    return score, evidence


def _score_company_record(
    query_norm: str,
    record: dict[str, Any],
    token_df: Counter[str],
    company_count: int,
    location_tokens: set[str],
    alias_target: str = "",
) -> tuple[float, str, str]:
    if not query_norm:
        return 0.0, "", ""

    authoritative_variants = record["authoritative_variants"]
    asset_variants = record["asset_variants"]
    all_variants = record["all_variants"]

    if alias_target:
        target_forms = _company_forms(
            alias_target,
            location_tokens=location_tokens,
        )
        if (
            alias_target == record["company_key"]
            or alias_target in all_variants
            or bool(
                target_forms
                & (record["authoritative_forms"] | record["asset_forms"])
            )
        ):
            return 100.0, "MANUAL_COMPANY_ALIAS", alias_target

    # Only producer and short-name fields are authoritative enough for the
    # safest automatic methods.
    if query_norm in authoritative_variants:
        return 100.0, "EXACT_COMPANY_NAME", query_norm

    for variant in authoritative_variants:
        is_related, evidence = _safe_substring_relation(
            query_norm,
            variant,
            location_tokens,
        )
        if is_related:
            if "~" in evidence:
                return 95.0, "COMPANY_PREFIX_RELATION", evidence
            return 97.0, "COMPANY_SUBSTRING", evidence or variant

    # Exact/substring agreement with text extracted only from Asset Identifier
    # is useful for suggestions (e.g. subsidiary names) but requires review or
    # an explicit alias before automatic assignment.
    if query_norm in asset_variants:
        return 96.0, "ASSET_COMPANY_EXACT", query_norm

    for variant in asset_variants:
        is_related, evidence = _safe_substring_relation(
            query_norm,
            variant,
            location_tokens,
        )
        if is_related:
            if "~" in evidence:
                return 92.0, "ASSET_COMPANY_PREFIX_RELATION", evidence
            return 94.0, "ASSET_COMPANY_SUBSTRING", evidence or variant

    query_forms = _company_forms(
        query_norm,
        location_tokens=location_tokens,
    )

    shared_authoritative_forms = query_forms & record["authoritative_forms"]
    if shared_authoritative_forms:
        return (
            93.0,
            "COMPANY_FORM_MATCH",
            sorted(shared_authoritative_forms)[0],
        )

    shared_asset_forms = query_forms & record["asset_forms"]
    if shared_asset_forms:
        return (
            91.0,
            "ASSET_COMPANY_FORM_MATCH",
            sorted(shared_asset_forms)[0],
        )

    score, evidence = _score_distinctive_tokens(
        query_norm=query_norm,
        record=record,
        token_df=token_df,
        company_count=company_count,
        location_tokens=location_tokens,
    )
    if score > 0:
        return score, "DISTINCTIVE_COMPANY_TOKENS", evidence

    return 0.0, "", ""


# ---------------------------------------------------------------------------
# Location matching
# ---------------------------------------------------------------------------


def _address_without_company(certificate_holder: Any, company_name: Any) -> str:
    holder = _normalize_location(certificate_holder)
    company = _normalize_location(company_name)

    if not holder:
        return ""
    if company and holder.startswith(company):
        return holder[len(company) :].strip()

    return holder


def _contains_phrase(needle: str, haystack: str) -> bool:
    if not needle or not haystack:
        return False
    return bool(re.search(rf"\b{re.escape(needle)}\b", haystack))


def _location_similarity(left: str, right: str) -> float:
    if not left or not right:
        return 0.0
    if left == right:
        return 100.0
    if _contains_phrase(left, right) or _contains_phrase(right, left):
        return 100.0

    scores = [
        float(fuzz.ratio(left, right)),
        float(fuzz.token_set_ratio(left, right)),
    ]
    if min(len(left), len(right)) >= 5:
        scores.append(min(float(fuzz.partial_ratio(left, right)), 90.0))
    return max(scores)


def _score_location(
    iscc_company: Any,
    iscc_city: Any,
    certificate_holder: Any,
    gst_city: Any,
    gst_pbi_city: Any,
    gst_territory: Any,
    gst_asset_location_norm: Any,
) -> tuple[float, str]:
    address = _address_without_company(certificate_holder, iscc_company)
    iscc_city_norm = _normalize_location(iscc_city)
    gst_city_norm = _normalize_location(gst_city)
    pbi_city_norm = _normalize_location(gst_pbi_city)
    territory_norm = _normalize_location(gst_territory)
    asset_location_norm = _normalize_location(gst_asset_location_norm)

    candidates: list[tuple[float, str]] = []

    for candidate_city, label in [
        (gst_city_norm, "GST_CITY_IN_CERTIFICATE_ADDRESS"),
        (pbi_city_norm, "PBI_CITY_IN_CERTIFICATE_ADDRESS"),
    ]:
        if candidate_city and _contains_phrase(candidate_city, address):
            candidates.append((100.0, label))

    if asset_location_norm and _contains_phrase(asset_location_norm, address):
        candidates.append((98.0, "ASSET_LOCATION_IN_CERTIFICATE_ADDRESS"))

    # Province/state evidence helps ranking but cannot by itself satisfy the
    # strong location threshold used for automatic matching.
    if territory_norm and _contains_phrase(territory_norm, address):
        candidates.append((75.0, "GST_TERRITORY_IN_CERTIFICATE_ADDRESS"))

    for candidate_city, label in [
        (gst_city_norm, "ISCC_CITY_TO_GST_CITY"),
        (pbi_city_norm, "ISCC_CITY_TO_PBI_CITY"),
    ]:
        if iscc_city_norm and candidate_city:
            candidates.append(
                (
                    min(_location_similarity(iscc_city_norm, candidate_city), 95.0),
                    label,
                )
            )

    if iscc_city_norm and asset_location_norm:
        candidates.append(
            (
                min(
                    _location_similarity(iscc_city_norm, asset_location_norm),
                    92.0,
                ),
                "ISCC_CITY_TO_ASSET_LOCATION",
            )
        )

    if iscc_city_norm and territory_norm:
        candidates.append(
            (
                min(_location_similarity(iscc_city_norm, territory_norm), 70.0),
                "ISCC_CITY_TO_GST_TERRITORY",
            )
        )

    if not candidates:
        return 0.0, ""

    return max(candidates, key=lambda item: item[0])


# ---------------------------------------------------------------------------
# Match orchestration
# ---------------------------------------------------------------------------



def _target_processing_categories(
    scope: Any,
    processing_unit_type: Any,
) -> frozenset[str]:
    """Return target plant categories explicitly present on a processing unit."""

    if not _is_processing_unit(scope, processing_unit_type):
        return frozenset()

    text = _normalize_match_text(processing_unit_type)
    categories: set[str] = set()
    if re.search(r"\bhefa\s+plant\b", text):
        categories.add("HEFA")
    if re.search(r"\bco\s*processing\s+plant\b", text):
        categories.add("CO_PROCESSING")
    if re.search(r"\bbiodiesel\s+plant\b", text):
        categories.add("BIODIESEL")
    return frozenset(categories)


def _gst_target_technologies(value: Any) -> frozenset[str]:
    text = _normalize_match_text(value)
    categories: set[str] = set()
    if "hefa" in text:
        categories.add("HEFA")
    if re.search(r"\bco\s*pro\b", text) or "co processing" in text:
        categories.add("CO_PROCESSING")
    if "biodiesel" in text:
        categories.add("BIODIESEL")
    return frozenset(categories)


def _target_technology_compatible(
    asset_technologies: frozenset[str],
    certificate_categories: frozenset[str],
) -> bool:
    return bool(asset_technologies & certificate_categories)


def _phase_number(value: Any) -> int:
    text = _normalize_match_text(value)
    match = re.search(r"\bphase\s+(\d+)\b", text)
    return int(match.group(1)) if match else 999


def _target_asset_preference_key(asset_identifier: Any) -> tuple[Any, ...]:
    """Stable site-variant preference: base asset, Phase 1, then lexical."""

    raw = _safe_match_text(asset_identifier)
    normalized = _normalize_match_text(raw)
    is_variant = _is_variant_asset(raw)
    phase = _phase_number(raw)
    is_expansion = int("expansion" in normalized)
    return (int(is_variant), phase, is_expansion, normalized)


def _choose_target_site_asset(
    site_assets: list[dict[str, Any]],
) -> tuple[dict[str, Any], bool]:
    """Choose one deterministic identifier where a site has phases/variants."""

    chosen = min(
        site_assets,
        key=lambda item: _target_asset_preference_key(item["asset"]),
    )
    phase_fallback = len(site_assets) > 1 and not any(
        not item.get("is_variant", False) for item in site_assets
    )
    return chosen, phase_fallback


def _company_rule_matches(query_norm: str, rule_company: Any) -> bool:
    rule_norm = _normalize_company(rule_company)
    if not query_norm or not rule_norm:
        return False
    if query_norm == rule_norm:
        return True
    query_tokens = set(_company_tokens(query_norm, meaningful_only=True))
    rule_tokens = set(_company_tokens(rule_norm, meaningful_only=True))
    if rule_tokens and rule_tokens.issubset(query_tokens):
        return True
    if query_tokens and query_tokens.issubset(rule_tokens):
        return True
    return float(fuzz.token_set_ratio(query_norm, rule_norm)) >= 94.0


def _find_target_site_override(
    company_norm: str,
    country_norm: str,
    city: Any,
    certificate_holder: Any,
    gst: pd.DataFrame,
    rules: tuple[dict[str, Any], ...],
) -> tuple[pd.Series | None, str]:
    location_blob = _normalize_location(
        f"{_safe_match_text(city)} {_safe_match_text(certificate_holder)}"
    )

    for rule in rules:
        if _normalize_country(rule.get("country", "")) != country_norm:
            continue
        if not _company_rule_matches(company_norm, rule.get("company", "")):
            continue

        location_terms = tuple(rule.get("locations", ()))
        if location_terms and not any(
            _normalize_location(term) in location_blob
            for term in location_terms
            if _normalize_location(term)
        ):
            continue

        asset = _safe_match_text(rule.get("asset", ""))
        rows = gst[gst["Asset Identifier"].eq(asset)]
        if rows.empty:
            continue
        return rows.iloc[0], "TARGET_SITE_OVERRIDE"

    return None, ""


def _target_fuzzy_company_score(
    query_norm: str,
    variants: Iterable[str],
    location_tokens: set[str],
) -> tuple[float, str]:
    """Broader spelling score requiring a non-generic company anchor."""

    query_tokens = _company_tokens(
        query_norm,
        location_tokens=location_tokens,
        meaningful_only=True,
    )
    if not query_tokens:
        return 0.0, ""

    best_score = 0.0
    best_evidence = ""
    for variant in variants:
        candidate_tokens = _company_tokens(
            variant,
            location_tokens=location_tokens,
            meaningful_only=True,
        )
        if not candidate_tokens:
            continue

        anchors: list[str] = []
        for left in query_tokens:
            for right in candidate_tokens:
                relation, _ = _token_relation(left, right)
                if relation >= 94.0:
                    anchors.append(f"{left}~{right}")
                    break
        if not anchors:
            continue

        ratio = max(
            float(fuzz.ratio(query_norm, variant)),
            float(fuzz.token_set_ratio(query_norm, variant)),
        )
        coverage = len(anchors) / max(
            min(len(query_tokens), len(candidate_tokens)),
            1,
        )
        score = min((0.82 * ratio) + (18.0 * min(coverage, 1.0)), 96.0)
        if score > best_score:
            best_score = score
            best_evidence = ", ".join(anchors)

    return best_score, best_evidence


def _target_asset_record(
    asset_row: pd.Series,
    location_tokens: set[str],
) -> dict[str, Any]:
    authoritative_variants = tuple(
        value
        for value in [
            asset_row["__producer_norm"],
            asset_row["__short_norm"],
        ]
        if value
    )
    asset_variants = tuple(
        value
        for value in [asset_row["__asset_company_norm"]]
        if value and value not in authoritative_variants
    )
    all_variants = tuple(dict.fromkeys(authoritative_variants + asset_variants))

    return {
        "company_key": asset_row["__company_key"],
        "authoritative_variants": authoritative_variants,
        "asset_variants": asset_variants,
        "all_variants": all_variants,
        "authoritative_forms": set().union(
            *(
                _company_forms(value, location_tokens=location_tokens)
                for value in authoritative_variants
            )
        )
        if authoritative_variants
        else set(),
        "asset_forms": set().union(
            *(
                _company_forms(value, location_tokens=location_tokens)
                for value in asset_variants
            )
        )
        if asset_variants
        else set(),
        "tokens": set().union(
            *(
                set(
                    _company_tokens(
                        value,
                        location_tokens=location_tokens,
                        meaningful_only=True,
                    )
                )
                for value in all_variants
            )
        )
        if all_variants
        else set(),
        "display_name": (
            _safe_match_text(asset_row["Company/Producer Short Name"])
            or _safe_match_text(asset_row["Company/Producer"])
        ),
    }


def _target_candidate_is_credible(
    company_method: str,
    company_score: float,
    company_evidence: str,
    location_score: float,
    unique_company_sites: int,
    config: dict[str, float],
) -> bool:
    authoritative = {
        "MANUAL_COMPANY_ALIAS",
        "EXACT_COMPANY_NAME",
        "COMPANY_SUBSTRING",
    }
    review_supported = {
        "COMPANY_PREFIX_RELATION",
        "COMPANY_FORM_MATCH",
        "ASSET_COMPANY_EXACT",
        "ASSET_COMPANY_SUBSTRING",
        "ASSET_COMPANY_PREFIX_RELATION",
        "ASSET_COMPANY_FORM_MATCH",
    }

    evidence_tokens = {
        token
        for token in _normalize_match_text(company_evidence).split()
        if (
            len(token) >= 4
            and token not in _COMPANY_GENERIC_TOKENS
            and token not in _LOCATION_NOISE_TOKENS
        )
    }

    if company_method in authoritative:
        if company_score < config["min_authoritative_company"]:
            return False
        # A substring based only on a generic word such as "biodiesel" is not
        # company evidence. This blocks cross-company plant assignments.
        if company_method == "COMPANY_SUBSTRING" and not evidence_tokens:
            return False
        if location_score >= 80.0:
            return True
        # Exact/confirmed company plus one GST site is usable only when the
        # certificate contains no usable location evidence at all. A low but
        # non-zero location score is treated as a conflict, not as permission
        # to force the sole technology-compatible site.
        return (
            company_method in {"EXACT_COMPANY_NAME", "MANUAL_COMPANY_ALIAS"}
            and unique_company_sites == 1
            and location_score == 0.0
        )

    if company_method in review_supported:
        return (
            company_score >= config["min_review_company"]
            and location_score >= config["min_location_for_review_method"]
        )

    if company_method == "DISTINCTIVE_COMPANY_TOKENS":
        meaningful_evidence = any(
            token
            and token not in _COMPANY_GENERIC_TOKENS
            and len(token) >= 4
            for relation in company_evidence.split(",")
            for token in relation.strip().split("~")
        )
        return (
            meaningful_evidence
            and company_score >= config["min_review_company"]
            and location_score >= config["min_location_for_distinctive"]
        )

    if company_method == "TARGET_FUZZY_COMPANY":
        return (
            bool(evidence_tokens)
            and company_score >= config["min_fuzzy_company"]
            and location_score >= config["min_location_for_distinctive"]
        )

    return False


def _target_cross_technology_candidate_is_credible(
    company_method: str,
    company_score: float,
    company_evidence: str,
    location_score: float,
) -> bool:
    """Allow a site-level GST record where the GST technology differs.

    Some refineries are represented in the GST by a hydrogen, e-fuels or
    another project row rather than a dedicated HEFA/co-processing/biodiesel
    row. For the target processing-unit pass, that row can still identify the
    physical site, but only with strong company evidence and an exact location.
    """

    allowed_methods = {
        "MANUAL_COMPANY_ALIAS",
        "EXACT_COMPANY_NAME",
        "COMPANY_SUBSTRING",
        "ASSET_COMPANY_EXACT",
        "ASSET_COMPANY_SUBSTRING",
        "COMPANY_PREFIX_RELATION",
        "ASSET_COMPANY_PREFIX_RELATION",
    }
    if company_method not in allowed_methods or company_score < 95.0:
        return False

    evidence_tokens = {
        token
        for token in _normalize_match_text(company_evidence).split()
        if (
            len(token) >= 4
            and token not in _COMPANY_GENERIC_TOKENS
            and token not in _LOCATION_NOISE_TOKENS
        )
    }
    if company_method not in {"MANUAL_COMPANY_ALIAS", "EXACT_COMPANY_NAME"}:
        if not evidence_tokens:
            return False

    # Technology mismatch is acceptable only for an effectively exact site.
    return location_score >= 95.0


def _apply_target_processing_coverage(
    result: pd.DataFrame,
    gst: pd.DataFrame,
    location_tokens_by_country: dict[str, set[str]],
    token_df: Counter[str],
    company_count: int,
    *,
    target_company_alias_overrides: dict[str, str] | None = None,
    target_site_overrides: tuple[dict[str, Any], ...] | None = None,
    target_config: dict[str, float] | None = None,
) -> pd.DataFrame:
    """Fill credible HEFA/co-processing/biodiesel Processing Unit matches."""

    final = result.copy()
    cfg = TARGET_MATCH_CONFIG.copy()
    if target_config:
        cfg.update(target_config)

    alias_source = (
        TARGET_COMPANY_ALIAS_OVERRIDES
        if target_company_alias_overrides is None
        else target_company_alias_overrides
    )
    aliases = {
        _normalize_company(source): _normalize_company(target)
        for source, target in alias_source.items()
        if _normalize_company(source) and _normalize_company(target)
    }
    rules = TARGET_SITE_OVERRIDES if target_site_overrides is None else target_site_overrides

    target_columns = {
        "Target_Processing_Categories": "",
        "Target_Processing_Unit": 0,
        "Target_Force_Match_Applied": 0,
        "Phase_Fallback_Applied": 0,
        "Target_Coverage_Status": "NOT_TARGET",
    }
    for column, default in target_columns.items():
        final[column] = default

    assets_by_country = {
        country: group.copy()
        for country, group in gst.groupby("__country_norm", sort=False)
        if country
    }
    target_cache: dict[tuple[Any, ...], dict[str, Any] | None] = {}
    target_candidate_bundle_cache: dict[
        tuple[str, tuple[str, ...]], list[dict[str, Any]]
    ] = {}
    target_company_score_cache: dict[
        tuple[str, str, tuple[str, ...], str], list[dict[str, Any]]
    ] = {}

    for index, row in final.iterrows():
        categories = _target_processing_categories(
            row.get("Scope", ""),
            row.get("Processing_Unit_Type", ""),
        )
        if not categories:
            continue

        final.at[index, "Target_Processing_Categories"] = ", ".join(
            sorted(categories)
        )
        final.at[index, "Target_Processing_Unit"] = 1

        try:
            already_matched = int(float(row.get("Match_Found", 0) or 0)) == 1
        except (TypeError, ValueError):
            already_matched = str(row.get("Match_Found", "")).strip() == "1"
        if already_matched and _safe_match_text(row.get("Asset_Identifier", "")):
            if row.get("Match_Method", "") == "CONFIRMED_STARTED_ASSET_OVERRIDE":
                final.at[index, "Target_Force_Match_Applied"] = 1
                final.at[index, "Target_Coverage_Status"] = (
                    "CONFIRMED_STARTED_ASSET_OVERRIDE"
                )
            else:
                final.at[index, "Target_Coverage_Status"] = "ALREADY_MATCHED"
            continue

        company_raw = row.get("Company_Name", "")
        city_raw = row.get("City", "")
        country_raw = row.get("Country", "")
        certificate_holder = row.get("Certificate_Holder", "")
        company_norm = _normalize_company(company_raw)
        country_norm = _normalize_country(country_raw)
        location_blob = _normalize_location(
            f"{_safe_match_text(city_raw)} {_safe_match_text(certificate_holder)}"
        )

        cache_key = (
            company_norm,
            country_norm,
            location_blob,
            tuple(sorted(categories)),
        )
        if cache_key in target_cache:
            cached = target_cache[cache_key]
            if cached is None:
                final.at[index, "Target_Coverage_Status"] = "NO_PLAUSIBLE_GST_ASSET"
                continue
            selected = cached.copy()
        else:
            selected: dict[str, Any] | None = None
            override_row, override_method = _find_target_site_override(
                company_norm=company_norm,
                country_norm=country_norm,
                city=city_raw,
                certificate_holder=certificate_holder,
                gst=gst,
                rules=rules,
            )
            if override_row is not None:
                selected = {
                    "asset": override_row["Asset Identifier"],
                    "territory": _safe_match_text(override_row["Territory"]),
                    "site_key": override_row["__site_key"],
                    "is_variant": bool(override_row["__is_variant"]),
                    "company": (
                        _safe_match_text(override_row["Company/Producer Short Name"])
                        or _safe_match_text(override_row["Company/Producer"])
                    ),
                    "company_score": 100.0,
                    "company_method": override_method,
                    "company_evidence": override_row["Asset Identifier"],
                    "location_score": 100.0,
                    "location_method": override_method,
                    "overall_score": 100.0,
                    "score_margin": 100.0,
                    "candidate_count": 1,
                    "site_count": 1,
                    "runner_asset": "",
                    "runner_score": 0.0,
                    "phase_fallback": bool(override_row["__is_variant"]),
                    "coverage_status": "DIRECT_OVERRIDE",
                }
            elif company_norm and country_norm in assets_by_country:
                category_key = tuple(sorted(categories))
                bundle_key = (country_norm, category_key)
                country_location_tokens = location_tokens_by_country.get(
                    country_norm,
                    set(),
                )

                if bundle_key not in target_candidate_bundle_cache:
                    country_assets = assets_by_country[country_norm]

                    # Start with every GST asset in the country. Technology-
                    # compatible rows receive the normal target pass. Rows for
                    # a different technology are retained only as a strict
                    # site-level fallback, requiring exact company/location
                    # evidence later in the scoring loop.
                    bundles: list[dict[str, Any]] = []
                    if not country_assets.empty:
                        company_site_counts = country_assets.groupby(
                            "__company_key"
                        )["__site_key"].nunique().to_dict()
                        for _, prepared_asset_row in country_assets.iterrows():
                            technology_match = _target_technology_compatible(
                                prepared_asset_row["__target_technologies"],
                                categories,
                            )
                            bundles.append(
                                {
                                    "asset_row": prepared_asset_row,
                                    "record": _target_asset_record(
                                        prepared_asset_row,
                                        country_location_tokens,
                                    ),
                                    "unique_company_sites": int(
                                        company_site_counts.get(
                                            prepared_asset_row["__company_key"],
                                            0,
                                        )
                                    ),
                                    "technology_match": technology_match,
                                }
                            )
                    target_candidate_bundle_cache[bundle_key] = bundles

                bundles = target_candidate_bundle_cache[bundle_key]
                if bundles:
                    alias_target = aliases.get(company_norm, "")
                    company_score_key = (
                        company_norm,
                        country_norm,
                        category_key,
                        alias_target,
                    )
                    if company_score_key not in target_company_score_cache:
                        scored_bundles: list[dict[str, Any]] = []
                        for bundle in bundles:
                            record = bundle["record"]
                            company_score, company_method, company_evidence = (
                                _score_company_record(
                                    query_norm=company_norm,
                                    record=record,
                                    token_df=token_df,
                                    company_count=company_count,
                                    location_tokens=country_location_tokens,
                                    alias_target=alias_target,
                                )
                            )
                            fuzzy_score, fuzzy_evidence = (
                                _target_fuzzy_company_score(
                                    company_norm,
                                    record["all_variants"],
                                    country_location_tokens,
                                )
                            )
                            if fuzzy_score > company_score:
                                company_score = fuzzy_score
                                company_method = "TARGET_FUZZY_COMPANY"
                                company_evidence = fuzzy_evidence
                            if company_score <= 0:
                                continue
                            scored_bundles.append(
                                {
                                    **bundle,
                                    "company_score": company_score,
                                    "company_method": company_method,
                                    "company_evidence": company_evidence,
                                }
                            )
                        target_company_score_cache[company_score_key] = scored_bundles

                    candidate_items: list[dict[str, Any]] = []
                    for scored_bundle in target_company_score_cache[
                        company_score_key
                    ]:
                        asset_row = scored_bundle["asset_row"]
                        if _is_blocked_company_asset(
                            company_raw, asset_row["Asset Identifier"]
                        ):
                            continue
                        record = scored_bundle["record"]
                        company_score = scored_bundle["company_score"]
                        company_method = scored_bundle["company_method"]
                        company_evidence = scored_bundle["company_evidence"]

                        location_score, location_method = _score_location(
                            iscc_company=company_raw,
                            iscc_city=city_raw,
                            certificate_holder=certificate_holder,
                            gst_city=asset_row["City"],
                            gst_pbi_city=asset_row["PBI City"],
                            gst_territory=asset_row["Territory"],
                            gst_asset_location_norm=asset_row[
                                "__asset_location_norm"
                            ],
                        )
                        technology_match = bool(
                            scored_bundle.get("technology_match", False)
                        )
                        if technology_match:
                            credible = _target_candidate_is_credible(
                                company_method=company_method,
                                company_score=company_score,
                                company_evidence=company_evidence,
                                location_score=location_score,
                                unique_company_sites=scored_bundle[
                                    "unique_company_sites"
                                ],
                                config=cfg,
                            )
                        else:
                            credible = (
                                _target_cross_technology_candidate_is_credible(
                                    company_method=company_method,
                                    company_score=company_score,
                                    company_evidence=company_evidence,
                                    location_score=location_score,
                                )
                            )
                        if not credible:
                            continue

                        overall_score = (
                            company_score * cfg["company_weight"]
                            + location_score * cfg["location_weight"]
                            + (cfg["technology_bonus"] if technology_match else 0.0)
                        )
                        candidate_items.append(
                            {
                                "asset": asset_row["Asset Identifier"],
                                "territory": _safe_match_text(
                                    asset_row["Territory"]
                                ),
                                "site_key": asset_row["__site_key"],
                                "is_variant": bool(asset_row["__is_variant"]),
                                "company": record["display_name"],
                                "company_score": company_score,
                                "company_method": company_method,
                                "company_evidence": company_evidence,
                                "location_score": location_score,
                                "location_method": location_method,
                                "overall_score": overall_score,
                                "technology_match": technology_match,
                            }
                        )

                    if candidate_items:
                        by_site: dict[str, list[dict[str, Any]]] = defaultdict(list)
                        for item in candidate_items:
                            by_site[item["site_key"]].append(item)

                        site_winners: list[dict[str, Any]] = []
                        for site_key, site_items in by_site.items():
                            best_scored = max(
                                site_items,
                                key=lambda item: (
                                    item["overall_score"],
                                    item["location_score"],
                                    item["company_score"],
                                ),
                            )
                            chosen, phase_fallback = _choose_target_site_asset(
                                site_items
                            )
                            # Keep the chosen identifier but retain the best
                            # evidence scores for the site.
                            chosen = chosen.copy()
                            for field in [
                                "company",
                                "company_score",
                                "company_method",
                                "company_evidence",
                                "location_score",
                                "location_method",
                                "overall_score",
                                "technology_match",
                            ]:
                                chosen[field] = best_scored[field]
                            chosen["phase_fallback"] = phase_fallback
                            site_winners.append(chosen)

                        site_winners.sort(
                            key=lambda item: (
                                item["overall_score"],
                                item["location_score"],
                                item["company_score"],
                                -_phase_number(item["asset"]),
                            ),
                            reverse=True,
                        )
                        top = site_winners[0]
                        runner = site_winners[1] if len(site_winners) > 1 else None
                        margin = (
                            top["overall_score"] - runner["overall_score"]
                            if runner
                            else 100.0
                        )

                        # Multiple company sites require location or a clear
                        # score margin. One unique credible site can be used
                        # even when the parsed City is imperfect.
                        decisive = (
                            len(site_winners) == 1
                            or top["location_score"] >= 90.0
                            or margin >= cfg["min_different_site_margin"]
                        )
                        if decisive:
                            selected = {
                                **top,
                                "score_margin": margin,
                                "candidate_count": len(candidate_items),
                                "site_count": len(site_winners),
                                "runner_asset": runner["asset"] if runner else "",
                                "runner_score": runner["overall_score"] if runner else 0.0,
                                "coverage_status": (
                                    "TARGET_TECH_COMPANY_SITE"
                                    if top.get("technology_match", False)
                                    else "TARGET_SITE_COMPANY_LOCATION"
                                ),
                            }

            target_cache[cache_key] = selected.copy() if selected else None

        if selected is None:
            final.at[index, "Target_Coverage_Status"] = "NO_PLAUSIBLE_GST_ASSET"
            # Make the reason explicit without replacing useful prior details.
            prior_reason = _safe_match_text(final.at[index, "Review_Reason"])
            if not prior_reason:
                final.at[index, "Review_Reason"] = (
                    "Target processing unit has no credible compatible GST asset"
                )
            continue

        confidence = "High" if (
            selected["company_score"] >= 97.0
            and selected["location_score"] >= 90.0
        ) else "Medium"
        method = (
            "TARGET_PROCESSING_DIRECT_OVERRIDE"
            if selected["coverage_status"] == "DIRECT_OVERRIDE"
            else "TARGET_PROCESSING_COMPANY_LOCATION"
        )
        if selected.get("phase_fallback"):
            method += "_PHASE_FALLBACK"

        final.at[index, "Asset_Identifier"] = selected["asset"]
        final.at[index, "Suggested_Asset_Identifier"] = selected["asset"]
        final.at[index, "Match_Found"] = 1
        final.at[index, "Match_Status"] = "Matched"
        final.at[index, "Match_Confidence"] = confidence
        final.at[index, "Match_Method"] = method
        final.at[index, "Auto_Match_Eligible"] = 1
        final.at[index, "Matched_GST_Company"] = selected["company"]
        final.at[index, "Best_GST_Company_Candidate"] = selected["company"]
        final.at[index, "Company_Match_Method"] = selected["company_method"]
        final.at[index, "Company_Match_Evidence"] = selected["company_evidence"]
        final.at[index, "Matched_Territory"] = selected["territory"]
        final.at[index, "Company_Score"] = round(selected["company_score"], 1)
        final.at[index, "Company_Score_Margin"] = round(
            selected["score_margin"], 1
        )
        final.at[index, "Asset_Company_Score"] = round(
            selected["company_score"], 1
        )
        final.at[index, "Asset_Company_Match_Method"] = selected[
            "company_method"
        ]
        final.at[index, "Asset_Company_Match_Evidence"] = selected[
            "company_evidence"
        ]
        final.at[index, "Location_Score"] = round(selected["location_score"], 1)
        final.at[index, "Location_Match_Method"] = selected["location_method"]
        final.at[index, "Overall_Score"] = round(selected["overall_score"], 1)
        final.at[index, "Score_Margin"] = round(selected["score_margin"], 1)
        final.at[index, "Candidate_Count"] = selected["candidate_count"]
        final.at[index, "Candidate_Site_Count"] = selected["site_count"]
        final.at[index, "Matched_Site_Key"] = selected["site_key"]
        final.at[index, "Runner_Up_Asset"] = selected["runner_asset"]
        final.at[index, "Runner_Up_Score"] = round(selected["runner_score"], 1)
        final.at[index, "Review_Reason"] = ""
        final.at[index, "Target_Force_Match_Applied"] = 1
        final.at[index, "Phase_Fallback_Applied"] = int(
            bool(selected.get("phase_fallback"))
        )
        final.at[index, "Target_Coverage_Status"] = selected["coverage_status"]

    return final

def _is_processing_unit(scope: Any, processing_unit_type: Any = "") -> bool:
    scope_norm = _normalize_match_text(scope)
    if "processing unit" in scope_norm:
        return True

    # Use Processing_Unit_Type only as a conservative secondary signal where
    # the scraped Scope is absent. A populated type does not override an
    # explicitly non-processing scope.
    if not scope_norm:
        processing_norm = _normalize_match_text(processing_unit_type)
        return bool(processing_norm)

    return False


def _base_output(is_processing: bool) -> dict[str, Any]:
    return {
        "Matcher_Version": MATCHER_VERSION,
        "Asset_Identifier": None,
        "Suggested_Asset_Identifier": None,
        "Match_Found": 0,
        "Match_Status": "No Match",
        "Match_Confidence": "",
        "Match_Method": "",
        "Auto_Match_Eligible": 0,
        "Is_Processing_Unit": int(is_processing),
        "Matched_GST_Company": "",
        "Best_GST_Company_Candidate": "",
        "Company_Match_Method": "",
        "Company_Match_Evidence": "",
        "Matched_Territory": "",
        "Company_Score": 0.0,
        "Company_Score_Margin": 0.0,
        "Asset_Company_Score": 0.0,
        "Asset_Company_Match_Method": "",
        "Asset_Company_Match_Evidence": "",
        "Location_Score": 0.0,
        "Location_Match_Method": "",
        "Overall_Score": 0.0,
        "Score_Margin": 0.0,
        "Candidate_Count": 0,
        "Candidate_Site_Count": 0,
        "Matched_Site_Key": "",
        "Runner_Up_Asset": "",
        "Runner_Up_Score": 0.0,
        "Review_Reason": "",
        "Target_Processing_Categories": "",
        "Target_Processing_Unit": 0,
        "Target_Force_Match_Applied": 0,
        "Phase_Fallback_Applied": 0,
        "Target_Coverage_Status": "NOT_TARGET",
    }


def match_assets_to_gst(
    iscc_df: pd.DataFrame,
    gst_df: pd.DataFrame,
    company_alias_overrides: dict[str, str] | None = None,
    manual_asset_overrides: dict[tuple[str, str], str] | None = None,
    config: dict[str, float] | None = None,
    gst_statuses: Iterable[str] | None = ("Started Up",),
    force_target_processing_units: bool = True,
    target_company_alias_overrides: dict[str, str] | None = None,
    target_site_overrides: tuple[dict[str, Any], ...] | None = None,
    target_config: dict[str, float] | None = None,
    use_confirmed_started_asset_matches: bool = True,
    confirmed_started_asset_matches: tuple[
        tuple[str, str, str, tuple[str, ...], str], ...
    ] | None = None,
    strict_confirmed_match_validation: bool = True,
    confirmed_started_address_matches: tuple[
        tuple[str, str, tuple[str, ...], str, tuple[str, ...], str], ...
    ] | None = None,
    include_match_diagnostics: bool = False,
) -> pd.DataFrame:
    """Match ISCC certificate rows to GST Asset Identifiers.

    The function does not overwrite ``Company_Name``. By default, only GST
    rows marked ``Started Up`` are eligible; use ``gst_statuses=None`` to opt out.
    Exact company-and-city
    rules in ``CONFIRMED_STARTED_ASSET_MATCHES`` are applied first, but only to
    Processing Unit rows. Remaining rows use the conservative matcher and the
    optional HEFA/co-processing/biodiesel coverage pass.

    By default, the returned DataFrame contains the original source columns
    plus only ``Asset_Identifier`` and ``Match_Found`` from matching. Set
    ``include_match_diagnostics=True`` to retain all matcher audit columns.
    """

    required_iscc = {"Company_Name", "City", "Country"}
    missing = required_iscc - set(iscc_df.columns)
    if missing:
        raise KeyError(f"Missing ISCC asset matching columns: {missing}")

    cfg = MATCH_CONFIG.copy()
    if config:
        cfg.update(config)

    alias_source = (
        COMPANY_ALIAS_OVERRIDES
        if company_alias_overrides is None
        else company_alias_overrides
    )
    aliases = {
        _normalize_company(source): _normalize_company(target)
        for source, target in alias_source.items()
        if _normalize_company(source) and _normalize_company(target)
    }

    override_source = (
        MANUAL_ASSET_OVERRIDES
        if manual_asset_overrides is None
        else manual_asset_overrides
    )
    asset_overrides = {
        (_normalize_company(company), _normalize_country(country)): asset
        for (company, country), asset in override_source.items()
        if _normalize_company(company) and _normalize_country(country)
    }

    gst, location_tokens_by_country = _prepare_gst_asset_matching(
        gst_df, allowed_statuses=gst_statuses
    )

    confirmed_rule_source = (
        CONFIRMED_STARTED_ASSET_MATCHES
        if confirmed_started_asset_matches is None
        else confirmed_started_asset_matches
    )
    confirmed_match_index = (
        _build_confirmed_started_match_index(
            gst,
            confirmed_rule_source,
            strict=strict_confirmed_match_validation,
        )
        if use_confirmed_started_asset_matches
        else {}
    )

    address_rule_source = (
        CONFIRMED_STARTED_ADDRESS_MATCHES
        if confirmed_started_address_matches is None
        else confirmed_started_address_matches
    )
    confirmed_address_match_index = (
        _build_confirmed_address_match_index(
            gst,
            address_rule_source,
            strict=strict_confirmed_match_validation,
        )
        if use_confirmed_started_asset_matches
        else {}
    )

    result = iscc_df.copy()

    result["Company_City"] = [
        " ".join(
            item
            for item in [
                _safe_match_text(company),
                _safe_match_text(city),
            ]
            if item
        )
        for company, city in zip(result["Company_Name"], result["City"])
    ]

    assets_by_country = {
        country: group.copy()
        for country, group in gst.groupby("__country_norm", sort=False)
        if country
    }

    companies_by_country = {
        country: _build_company_records(
            group,
            location_tokens_by_country.get(country, set()),
        )
        for country, group in assets_by_country.items()
    }
    company_search_indexes = {
        country: _build_company_search_index(records)
        for country, records in companies_by_country.items()
    }

    all_records = [
        record
        for country_records in companies_by_country.values()
        for record in country_records
    ]
    token_df = _build_token_document_frequency(all_records)
    company_count = len(all_records)

    # Repeated certificate holders are common across annual certificate
    # renewals, so cache company resolution by normalized company and country.
    company_resolution_cache: dict[
        tuple[str, str, str], list[dict[str, Any]]
    ] = {}

    output_rows: list[dict[str, Any]] = []
    row_result_cache: dict[tuple[Any, ...], dict[str, Any]] = {}

    def append_output(cache_key: tuple[Any, ...], value: dict[str, Any]) -> None:
        row_result_cache[cache_key] = value.copy()
        output_rows.append(value)

    for row in result.itertuples(index=False):
        company_raw = getattr(row, "Company_Name", "")
        city_raw = getattr(row, "City", "")
        country_raw = getattr(row, "Country", "")
        certificate_holder = getattr(row, "Certificate_Holder", "")
        scope_raw = getattr(row, "Scope", "")
        processing_type_raw = getattr(row, "Processing_Unit_Type", "")

        company_norm = _normalize_company(company_raw)
        country_norm = _normalize_country(country_raw)
        is_processing = _is_processing_unit(scope_raw, processing_type_raw)
        output = _base_output(is_processing)

        match_cache_key = (
            company_norm,
            _normalize_location(city_raw),
            country_norm,
            _address_without_company(certificate_holder, company_raw),
            is_processing,
            _normalize_match_text(processing_type_raw),
        )
        cached_output = row_result_cache.get(match_cache_key)
        if cached_output is not None:
            output_rows.append(cached_output.copy())
            continue

        confirmed_match = _find_confirmed_started_asset_match(
            company=company_raw,
            city=city_raw,
            scope=scope_raw,
            processing_unit_type=processing_type_raw,
            is_processing=is_processing,
            match_index=confirmed_match_index,
        )
        if confirmed_match is not None:
            _apply_confirmed_started_asset_output(
                output,
                confirmed_match,
                company_raw=company_raw,
                city_raw=city_raw,
            )
            append_output(match_cache_key, output)
            continue

        confirmed_address_match = _find_confirmed_address_asset_match(
            company=company_raw,
            country=country_raw,
            city=city_raw,
            certificate_holder=certificate_holder,
            scope=scope_raw,
            processing_unit_type=processing_type_raw,
            is_processing=is_processing,
            match_index=confirmed_address_match_index,
        )
        if confirmed_address_match is not None:
            _apply_confirmed_started_asset_output(
                output,
                confirmed_address_match,
                company_raw=company_raw,
                city_raw=city_raw,
            )
            output["Match_Method"] = "CONFIRMED_STARTED_ADDRESS_OVERRIDE"
            output["Company_Match_Method"] = "CONFIRMED_STARTED_ADDRESS_OVERRIDE"
            output["Asset_Company_Match_Method"] = "CONFIRMED_STARTED_ADDRESS_OVERRIDE"
            output["Location_Match_Method"] = "CONFIRMED_STARTED_ADDRESS_OVERRIDE"
            output["Target_Coverage_Status"] = "CONFIRMED_STARTED_ADDRESS_OVERRIDE"
            append_output(match_cache_key, output)
            continue

        if not company_norm:
            output["Review_Reason"] = "No usable company name"
            append_output(match_cache_key, output)
            continue

        if not country_norm:
            output["Review_Reason"] = "No usable certificate country"
            append_output(match_cache_key, output)
            continue

        if country_norm not in assets_by_country:
            output["Review_Reason"] = (
                f"Certificate country not found in GST: {country_raw}"
            )
            append_output(match_cache_key, output)
            continue

        country_assets = assets_by_country[country_norm]
        country_location_tokens = location_tokens_by_country.get(
            country_norm,
            set(),
        )

        # ------------------------------------------------------------------
        # Exact manual asset override
        # ------------------------------------------------------------------
        override_asset = asset_overrides.get((company_norm, country_norm), "")
        if override_asset:
            override_rows = country_assets[
                country_assets["Asset Identifier"].eq(override_asset)
            ]
            if not override_rows.empty:
                override_row = override_rows.iloc[0]
                display_company = (
                    _safe_match_text(
                        override_row["Company/Producer Short Name"]
                    )
                    or _safe_match_text(override_row["Company/Producer"])
                )
                output.update(
                    {
                        "Asset_Identifier": override_asset,
                        "Suggested_Asset_Identifier": override_asset,
                        "Match_Found": 1,
                        "Match_Status": "Matched",
                        "Match_Confidence": "High",
                        "Match_Method": "MANUAL_ASSET_OVERRIDE",
                        "Auto_Match_Eligible": 1,
                        "Matched_GST_Company": display_company,
                        "Best_GST_Company_Candidate": display_company,
                        "Company_Match_Method": "MANUAL_ASSET_OVERRIDE",
                        "Company_Match_Evidence": override_asset,
                        "Company_Score": 100.0,
                        "Company_Score_Margin": 100.0,
                        "Asset_Company_Score": 100.0,
                        "Asset_Company_Match_Method": "MANUAL_ASSET_OVERRIDE",
                        "Asset_Company_Match_Evidence": override_asset,
                        "Location_Score": 100.0,
                        "Location_Match_Method": "MANUAL_ASSET_OVERRIDE",
                        "Overall_Score": 100.0,
                        "Score_Margin": 100.0,
                        "Candidate_Count": 1,
                        "Candidate_Site_Count": 1,
                        "Matched_Site_Key": override_row["__site_key"],
                        "Matched_Territory": _safe_match_text(
                            override_row["Territory"]
                        ),
                    }
                )
                append_output(match_cache_key, output)
                continue

        # ------------------------------------------------------------------
        # Resolve exactly one GST company. Location is not used here.
        # ------------------------------------------------------------------
        records = companies_by_country.get(country_norm, [])
        search_index = company_search_indexes.get(country_norm, {})
        alias_target = aliases.get(company_norm, "")
        cache_key = (company_norm, country_norm, alias_target)

        if cache_key in company_resolution_cache:
            scored_companies = company_resolution_cache[cache_key]
        else:
            candidate_indices = _candidate_company_indices(
                query_norm=company_norm,
                search_index=search_index,
                location_tokens=country_location_tokens,
                alias_target=alias_target,
            )

            scored_companies: list[dict[str, Any]] = []
            for record_index in candidate_indices:
                record = records[record_index]
                score, method, evidence = _score_company_record(
                    query_norm=company_norm,
                    record=record,
                    token_df=token_df,
                    company_count=company_count,
                    location_tokens=country_location_tokens,
                    alias_target=alias_target,
                )
                if score > 0:
                    scored_companies.append(
                        {
                            **record,
                            "score": score,
                            "method": method,
                            "evidence": evidence,
                        }
                    )

            scored_companies.sort(
                key=lambda item: (
                    item["score"],
                    item["method"] in SAFE_AUTO_COMPANY_METHODS,
                ),
                reverse=True,
            )
            company_resolution_cache[cache_key] = scored_companies

        if not scored_companies:
            output["Review_Reason"] = (
                "No company candidate with non-location company evidence"
            )
            append_output(match_cache_key, output)
            continue

        best_company = scored_companies[0]
        second_company = (
            scored_companies[1] if len(scored_companies) > 1 else None
        )
        company_margin = (
            best_company["score"] - second_company["score"]
            if second_company
            else 100.0
        )

        output.update(
            {
                "Best_GST_Company_Candidate": best_company["display_name"],
                "Company_Match_Method": best_company["method"],
                "Company_Match_Evidence": best_company["evidence"],
                "Company_Score": round(best_company["score"], 1),
                "Company_Score_Margin": round(company_margin, 1),
            }
        )

        if best_company["score"] < cfg["min_company_score"]:
            output["Review_Reason"] = "Best company score below threshold"
            append_output(match_cache_key, output)
            continue

        if (
            second_company
            and second_company["score"] >= cfg["min_company_score"]
            and company_margin < cfg["min_company_margin"]
        ):
            output["Match_Status"] = "Review"
            output["Match_Confidence"] = "Review"
            output["Review_Reason"] = (
                "Ambiguous company match; assets were not ranked"
            )
            append_output(match_cache_key, output)
            continue

        accepted_company_key = best_company["company_key"]
        asset_candidates = country_assets[
            country_assets["__company_key"].eq(accepted_company_key)
        ].copy()

        if asset_candidates.empty:
            output["Review_Reason"] = "Company resolved but no GST assets found"
            append_output(match_cache_key, output)
            continue

        output["Matched_GST_Company"] = best_company["display_name"]

        # ------------------------------------------------------------------
        # Rank only assets belonging to the locked GST company.
        # ------------------------------------------------------------------
        scored_assets: list[dict[str, Any]] = []

        for _, asset_row in asset_candidates.iterrows():
            if _is_blocked_company_asset(
                company_raw, asset_row["Asset Identifier"]
            ):
                continue
            authoritative_variants = tuple(
                value
                for value in [
                    asset_row["__producer_norm"],
                    asset_row["__short_norm"],
                ]
                if value
            )
            asset_only_variants = tuple(
                value
                for value in [asset_row["__asset_company_norm"]]
                if value and value not in authoritative_variants
            )
            all_variants = tuple(
                dict.fromkeys(authoritative_variants + asset_only_variants)
            )

            asset_record = {
                "company_key": asset_row["__company_key"],
                "authoritative_variants": authoritative_variants,
                "asset_variants": asset_only_variants,
                "all_variants": all_variants,
                "authoritative_forms": set().union(
                    *(
                        _company_forms(
                            value,
                            location_tokens=country_location_tokens,
                        )
                        for value in authoritative_variants
                    )
                )
                if authoritative_variants
                else set(),
                "asset_forms": set().union(
                    *(
                        _company_forms(
                            value,
                            location_tokens=country_location_tokens,
                        )
                        for value in asset_only_variants
                    )
                )
                if asset_only_variants
                else set(),
                "tokens": set().union(
                    *(
                        set(
                            _company_tokens(
                                value,
                                location_tokens=country_location_tokens,
                                meaningful_only=True,
                            )
                        )
                        for value in all_variants
                    )
                )
                if all_variants
                else set(),
                "display_name": best_company["display_name"],
            }

            (
                asset_company_score,
                asset_company_method,
                asset_company_evidence,
            ) = _score_company_record(
                query_norm=company_norm,
                record=asset_record,
                token_df=token_df,
                company_count=company_count,
                location_tokens=country_location_tokens,
                alias_target=alias_target,
            )

            location_score, location_method = _score_location(
                iscc_company=company_raw,
                iscc_city=city_raw,
                certificate_holder=certificate_holder,
                gst_city=asset_row["City"],
                gst_pbi_city=asset_row["PBI City"],
                gst_territory=asset_row["Territory"],
                gst_asset_location_norm=asset_row["__asset_location_norm"],
            )

            overall_score = (
                asset_company_score * cfg["asset_company_weight"]
                + location_score * cfg["location_weight"]
            )

            scored_assets.append(
                {
                    "asset": asset_row["Asset Identifier"],
                    "territory": _safe_match_text(asset_row["Territory"]),
                    "site_key": asset_row["__site_key"],
                    "is_variant": bool(asset_row["__is_variant"]),
                    "asset_company_score": asset_company_score,
                    "asset_company_method": asset_company_method,
                    "asset_company_evidence": asset_company_evidence,
                    "location_score": location_score,
                    "location_method": location_method,
                    "overall_score": overall_score,
                }
            )

        if not scored_assets:
            output["Review_Reason"] = (
                "All candidate assets were excluded by confirmed block rules"
            )
            append_output(match_cache_key, output)
            continue

        scored_assets.sort(
            key=lambda item: (
                item["overall_score"],
                item["location_score"],
                item["asset_company_score"],
                not item["is_variant"],
            ),
            reverse=True,
        )

        deduplicated_assets: list[dict[str, Any]] = []
        seen_assets: set[str] = set()
        for item in scored_assets:
            if item["asset"] in seen_assets:
                continue
            seen_assets.add(item["asset"])
            deduplicated_assets.append(item)
        scored_assets = deduplicated_assets

        top = scored_assets[0]
        top_site_assets = [
            item
            for item in scored_assets
            if item["site_key"] == top["site_key"]
        ]
        base_assets = [
            item for item in top_site_assets if not item["is_variant"]
        ]

        # Prefer a generic/base Asset Identifier over phase/expansion rows.
        if len(base_assets) == 1:
            top = base_assets[0]

        runner = next(
            (
                item
                for item in scored_assets
                if item["asset"] != top["asset"]
            ),
            None,
        )
        score_margin = (
            top["overall_score"] - runner["overall_score"]
            if runner
            else 100.0
        )
        unique_site_count = len(
            {item["site_key"] for item in scored_assets if item["site_key"]}
        )

        output.update(
            {
                "Suggested_Asset_Identifier": top["asset"],
                "Matched_Territory": top["territory"],
                "Asset_Company_Score": round(
                    top["asset_company_score"],
                    1,
                ),
                "Asset_Company_Match_Method": top[
                    "asset_company_method"
                ],
                "Asset_Company_Match_Evidence": top[
                    "asset_company_evidence"
                ],
                "Location_Score": round(top["location_score"], 1),
                "Location_Match_Method": top["location_method"],
                "Overall_Score": round(top["overall_score"], 1),
                "Score_Margin": round(score_margin, 1),
                "Candidate_Count": len(scored_assets),
                "Candidate_Site_Count": unique_site_count,
                "Matched_Site_Key": top["site_key"],
                "Runner_Up_Asset": runner["asset"] if runner else "",
                "Runner_Up_Score": round(runner["overall_score"], 1)
                if runner
                else 0.0,
            }
        )

        # Multiple phase/expansion identifiers for one site remain a review.
        same_site_variants = len(top_site_assets) > 1 and not base_assets
        if same_site_variants:
            output["Match_Status"] = "Review"
            output["Match_Confidence"] = "Review"
            output["Review_Reason"] = (
                "Site resolved, but phase/expansion Asset Identifier is ambiguous"
            )
            append_output(match_cache_key, output)
            continue

        if (
            runner
            and runner["site_key"] != top["site_key"]
            and score_margin < cfg["min_site_margin"]
        ):
            output["Match_Status"] = "Review"
            output["Match_Confidence"] = "Review"
            output["Review_Reason"] = (
                "Multiple sites for the same company; location evidence is not decisive"
            )
            append_output(match_cache_key, output)
            continue

        company_method_is_safe = (
            best_company["method"] in SAFE_AUTO_COMPANY_METHODS
        )
        asset_company_method_is_safe = (
            top["asset_company_method"] in SAFE_AUTO_ASSET_COMPANY_METHODS
        )
        strong_company = best_company["score"] >= cfg["strong_company_score"]
        strong_asset_company = (
            top["asset_company_score"] >= cfg["min_company_score"]
        )
        strong_location = (
            top["location_score"] >= cfg["strong_location_score"]
        )
        exact_location = (
            top["location_score"] >= cfg["exact_location_score"]
            and top["location_method"] in EXACT_LOCATION_METHODS
        )
        location_conflict = (
            bool(top["site_key"])
            and top["location_score"] < cfg["location_conflict_score"]
        )

        # ------------------------------------------------------------------
        # Conservative automatic assignment rules
        # ------------------------------------------------------------------
        processing_auto_match = (
            is_processing
            and company_method_is_safe
            and asset_company_method_is_safe
            and strong_company
            and strong_asset_company
            and strong_location
        )

        non_processing_auto_match = (
            not is_processing
            and best_company["method"] in EXACT_COMPANY_METHODS
            and top["asset_company_method"] in EXACT_COMPANY_METHODS
            and strong_company
            and strong_asset_company
            and exact_location
            and top["location_score"] >= cfg["non_processing_min_location"]
            and unique_site_count == 1
            and score_margin >= cfg["non_processing_min_site_margin"]
        )

        if processing_auto_match or non_processing_auto_match:
            confidence = (
                "High"
                if (
                    best_company["method"] in EXACT_COMPANY_METHODS
                    and top["asset_company_method"] in EXACT_COMPANY_METHODS
                    and exact_location
                )
                else "Medium"
            )

            output.update(
                {
                    "Asset_Identifier": top["asset"],
                    "Match_Found": 1,
                    "Match_Status": "Matched",
                    "Match_Confidence": confidence,
                    "Match_Method": (
                        "SAFE_PROCESSING_COMPANY_AND_LOCATION"
                        if processing_auto_match
                        else "SAFE_NON_PROCESSING_EXACT_COMPANY_SITE"
                    ),
                    "Auto_Match_Eligible": 1,
                }
            )

        else:
            output["Match_Status"] = "Review"
            output["Match_Confidence"] = "Review"

            if not company_method_is_safe:
                output["Review_Reason"] = (
                    "Company match method is review-only: "
                    f"{best_company['method']}"
                )
            elif not asset_company_method_is_safe:
                output["Review_Reason"] = (
                    "Specific asset company evidence is review-only: "
                    f"{top['asset_company_method']}"
                )
            elif location_conflict:
                output["Review_Reason"] = (
                    "Company resolved, but certificate and GST locations conflict"
                )
            elif not is_processing:
                output["Review_Reason"] = (
                    "Non-processing certificate requires exact company evidence, "
                    "exact location evidence and one unique GST site"
                )
            elif not strong_company:
                output["Review_Reason"] = (
                    "Company evidence is below the automatic threshold"
                )
            elif not strong_asset_company:
                output["Review_Reason"] = (
                    "Specific asset company evidence is below the automatic threshold"
                )
            elif not strong_location:
                output["Review_Reason"] = (
                    "Company resolved, but location evidence is insufficient"
                )
            else:
                output["Review_Reason"] = (
                    "Insufficient evidence for automatic asset assignment"
                )

        append_output(match_cache_key, output)

    match_results = pd.DataFrame(output_rows, index=result.index)

    # Make rerunning the matcher idempotent: replace prior matching columns.
    for column in match_results.columns:
        if column in result.columns:
            result = result.drop(columns=[column])

    combined = pd.concat([result, match_results], axis=1)

    if force_target_processing_units:
        combined = _apply_target_processing_coverage(
            result=combined,
            gst=gst,
            location_tokens_by_country=location_tokens_by_country,
            token_df=token_df,
            company_count=company_count,
            target_company_alias_overrides=target_company_alias_overrides,
            target_site_overrides=target_site_overrides,
            target_config=target_config,
        )

    if include_match_diagnostics:
        return combined

    # Production output: preserve original non-matching columns and retain only
    # the two requested matcher results. This also removes diagnostic columns
    # left in an input DataFrame from a previous test run.
    matcher_generated_columns = set(_base_output(False)) | {"Company_City"}
    source_columns = [
        column
        for column in iscc_df.columns
        if column not in matcher_generated_columns
    ]
    output_columns = source_columns + ["Asset_Identifier", "Match_Found"]
    output_columns = list(dict.fromkeys(output_columns))
    return combined.loc[:, output_columns].copy()


def summarize_match_results(df: pd.DataFrame) -> pd.DataFrame:
    """Return a status/method summary from a diagnostic test run.

    Call ``match_assets_to_gst(..., include_match_diagnostics=True)`` before
    using this helper.
    """

    required = {"Match_Status", "Company_Match_Method"}
    missing = required - set(df.columns)
    if missing:
        raise KeyError(f"Missing matcher output columns: {missing}")

    summary = (
        df.groupby(
            ["Match_Status", "Company_Match_Method"],
            dropna=False,
        )
        .size()
        .rename("Rows")
        .reset_index()
        .sort_values(
            ["Match_Status", "Rows"],
            ascending=[True, False],
        )
        .reset_index(drop=True)
    )
    return summary
