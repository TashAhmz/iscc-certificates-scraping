# Create a dictionary that maps each Scope* abbreviation to a broader facility category
FACILITY_GROUPING_MAP = {
    
    # Non-liquid biofuels assets
    "BG": "Non-liquid biofuels assets",
    "BM": "Non-liquid biofuels assets",
    "LNGT": "Non-liquid biofuels assets",
    "LP": "Non-liquid biofuels assets",
    # Add here

    # LCF Trading liquid biofuels assets
    "BP": "LCF Trading liquid biofuels assets",
    "CPP": "LCF Trading liquid biofuels assets",
    "EP": "LCF Trading liquid biofuels assets",
    "ET": "LCF Trading liquid biofuels assets",
    "MT": "LCF Trading liquid biofuels assets",
    "HEFA": "LCF Trading liquid biofuels assets",
    "HVO": "LCF Trading liquid biofuels assets",
    # Add here

    # Feedstock for Biofuels Companies
    "COF": "Feedstock for Biofuels Companies", 
    "COP": "Feedstock for Biofuels Companies",
    "CP": "Feedstock for Biofuels Companies",
    "FA": "Feedstock for Biofuels Companies",
    "FG": "Feedstock for Biofuels Companies",
    "FP": "Feedstock for Biofuels Companies",
    "ISHC": "Feedstock for Biofuels Companies",
    "MP": "Feedstock for Biofuels Companies",
    "OM": "Feedstock for Biofuels Companies",
    "OT": "Feedstock for Biofuels Companies",
    "PM": "Feedstock for Biofuels Companies",
    "PO": "Feedstock for Biofuels Companies",
    "PU": "Feedstock for Biofuels Companies",
    "PYP": "Feedstock for Biofuels Companies",
    "SM": "Feedstock for Biofuels Companies",
    "TW": "Feedstock for Biofuels Companies",
    # Add here
    
    # Other Processing Plants/Units
    "COMP": "Other Processing Plants/Units",
    "CV": "Other Processing Plants/Units",
    "ML": "Other Processing Plants/Units",
    "MRP": "Other Processing Plants/Units",
    "PP": "Other Processing Plants/Units",
    "RE": "Other Processing Plants/Units",
    "SC": "Other Processing Plants/Units",
    "CR": "Other Processing Plants/Units",
    "FPR": "Other Processing Plants/Units",
    "FSA": "Other Processing Plants/Units",
    "WH": "Other Processing Plants/Units",
    # Add here

    # Transport & Logistics Companies
    "LC": "Transport & Logistics Companies",
    "TC": "Transport & Logistics Companies",
    # Add here
    
    # Electricity Assets
    "IPEL": "Electricity Assets",
    "IPES": "Electricity Assets",
    "IPER": "Electricity Assets",
    "IPEM": "Electricity Assets",
    # Add here
    
    # Chemicals Assets
    "PWP": "Chemicals Assets",
    "SCP": "Chemicals Assets",
    # Add here
    
    # Trading Companies
    "PoC-TR": "Trading Companies",
    "TR": "Trading Companies",
    "TRS": "Trading Companies"
    # Add here
}

CERTIFICATE_TYPE_MAP = {
    "EU ISCC": "Mandated",
    "ISCC PLUS": "Voluntary",
    "ISCC JAPAN": "Japan FIT",
    "ISCC CORSIA": "Aviation",
    "ISCC CFC": "Carbon Footprint Certification",
    "DE B BLE BM": "Legacy"
}

# Manual corrections for country names that are incomplete or inconsistent
MANUAL_COUNTRY_OVERRIDES = {
    "Republic of": "South Korea", 
    "Province of China": "Taiwan",
    "Bolivarian Republic of": "Venezuela",
    "Ireland": "Republic of Ireland",
    "Viet Nam": "Vietnam",
    "Macedonia": "North Macedonia",
    "Russian Federation": "Russia",
    "Côte d'Ivoire": "Ivory Coast",
    "The Netherlands": "Netherlands",
    "British": "British Virgin Islands",
    "Islamic Republic of": "Iran",
    "Lettland": "Latvia",
    "Libyan Arab Jamahiriya": "Libya",
    "Reunion": "Réunion",
    "Swaziland": "Eswatini",
    "United Republic of": "Tanzania",
    "Slowakei": "Slovakia",
    "Hong Kong": "China"
}

COLUMN_MAP = {
    "cert_status": "Status_Code",
    "cert_number": "Certificate_ID",
    "cert_owner": "Certificate_Holder",
    "cert_scope": "Scope",
    "cert_processingunittype": "Processing_Unit_Type",
    "cert_in_put": "Raw_Material",
    "cert_add_on": "Add-ons** /CTS",
    "cert_products": "Products",
    "cert_valid_from": "Valid_From",
    "cert_valid_until": "Valid_Until",
    "cert_suspended_date": "Suspended",
    "cert_issuer": "Issuing_CB",
    "cert_map": "Map",
    "cert_file": "Certificate",
    "cert_audit": "Audit_Report"
}

STATUS_MAP = {
    1: "Valid",
    5: "Expired",
    10: "Expired",
    12: "Terminated",
    13: "Withdrawn",
    15: "Suspended",
    20: "Expired",
    21: "Expired"
}

SCOPE_DESCRIPTIONS = {
    "BFO": "Biomarine fuel operator",
    "BG": "Biogas plant",
    "BM": "Biomethane plant",
    "BP": "Biodiesel plant",
    "COF": "Central Office (Group of farms/plantations)",
    "COMP": "Compounding plant",
    "COP": "Central Office (Group of Points of Origin)",
    "CP": "Collecting Point (for waste/residue material not grown/harvested on farms/plantations)",
    "CPP": "Co-Processing plant",
    "CR": "Crushing plant",
    "CV": "Converter",
    "EL": "Electrolyser",
    "EP": "Ethanol plant",
    "ET": "ETBE plant",
    "FA": "Farm / Plantation",
    "FG": "First Gathering Point (for biomass grown/harvested on farms/plantations)",
    "FP": "Food processing plant",
    "FPR": "Final Product Refinement",
    "FSA": "Forest sourcing area",
    "HEFA": "HEFA plant",
    "HVO": "HVO plant",
    "ISHC": "Central Office for Independent Smallholders",
    "IPEL": "Installation producing energy (electricity, heating or cooling) from bioliquids",
    "IPES": "Installation producing energy (electricity, heating or cooling) from solid biomass",
    "IPER": "Installation producing energy (electricity, heating or cooling) from raw biogas",
    "IPEM": "Installation producing energy (electricity, heating or cooling) from biomethane",
    "LC": "Logistic Center",
    "LNGT": "LNG terminal",
    "LP": "Liquefaction Plant",
    "ML": "Methanol plant",
    "MP": "Melting plant",
    "MRP": "Mechanical Recycling Plant",
    "MT": "MTBE plant",
    "OM": "Oil mill",
    "OT": "Other conversion unit",
    "PM": "Pulp mill",
    "PO": "Point of Origin",
    "PoC-TR": "Proof of Compliance Trader",
    "PP": "Polymerization plant",
    "PU": "Processing Unit",
    "PWP": "Plastic Waste Processor",
    "PYP": "Pyrolysis plant",
    "RE": "Refinery",
    "SC": "Cracker",
    "SCP": "Speciality Chemical Plant",
    "SM": "Sugar mill",
    "TC": "Transport Company",
    "TR": "Trader",
    "TRS": "Trader with storage",
    "TW": "Treatment plant for waste/ residues",
    "WH": "Warehouse",
    "WR36": "acc. to 36th BImSchV (double counting of biofuels in Germany)"
}
