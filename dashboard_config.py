"""
Classification config for Vahan dashboard.
Maps raw keys from master_sheet to dashboard categories.
"""

# Maker Wise: Passenger Vehicles | EV-PV | 2W | CV | Tractors
MAKER_CLASSIFICATION = {
    "Passenger Vehicles": [
        "MARUTI SUZUKI INDIA LTD",
        "MAHINDRA & MAHINDRA LIMITED",
        "TATA MOTORS PASSENGER VEHICLES LTD",
        "HYUNDAI MOTOR INDIA LTD",
        "TOYOTA KIRLOSKAR MOTOR PVT LTD",
        "KIA INDIA PRIVATE LIMITED",
        "SKODA AUTO VOLKSWAGEN INDIA PVT LTD",
        "HONDA CARS INDIA LTD",
        "JSW MG MOTOR INDIA PVT LTD",
        "RENAULT INDIA PVT LTD",
        "NISSAN MOTOR INDIA PVT LTD",
        "STELLANTIS AUTOMOBILES INDIA PVT LTD",
        "VINFAST AUTO INDIA PVT LTD",
        "BMW INDIA PVT LTD",
        "MERCEDES-BENZ INDIA PVT LTD",
        "ISUZU MOTORS INDIA PVT LTD",
        "VOLVO AUTO INDIA PVT LTD",
    ],
    "EV-PV": [
        "TATA PASSENGER ELECTRIC MOBILITY LTD",
        "MAHINDRA LAST MILE MOBILITY LTD",
        "MAHINDRA ELECTRIC AUTOMOBILE LTD",
        "TESLA INDIA MOTORS AND ENERGY PVT LTD",
        "BYD INDIA PRIVATE LIMITED",
    ],
    "2W": [
        "HERO MOTOCORP LTD",
        "HONDA MOTORCYCLE AND SCOOTER INDIA (P) LTD",
        "TVS MOTOR COMPANY LTD",
        "BAJAJ AUTO LTD",
        "SUZUKI MOTORCYCLE INDIA PVT LTD",
        "ROYAL-ENFIELD (UNIT OF EICHER LTD)",
        "INDIA YAMAHA MOTOR PVT LTD",
        "ATHER ENERGY LTD",
        "OLA ELECTRIC TECHNOLOGIES PVT LTD",
        "CLASSIC LEGENDS PVT LTD",
        "RIVER MOBILITY PVT LTD",
        "INDIA KAWASAKI MOTORS PVT LTD",
        "BGAUSS AUTO PRIVATE LIMITED",
        "TRIUMPH MOTORCYCLES (INDIA) PVT LTD",
    ],
    "CV": [
        "SML ISUZU LTD",
        "TATA MOTORS LTD",
        "ASHOK LEYLAND LTD",
        "VE COMMERCIAL VEHICLES LTD",
        "SWITCH MOBILITY AUTOMOTIVE LTD",
        "FORCE MOTORS LIMITED",
        "DAIMLER INDIA COMMERCIAL VEHICLES PVT. LTD",
        "VE COMMERCIAL VEHICLES LTD (VOLVO BUSES DIVISION)",
    ],
    "Tractors": [
        "MAHINDRA & MAHINDRA LIMITED (TRACTOR)",
        "MAHINDRA & MAHINDRA LIMITED (SWARAJ DIVISION)",
        "ESCORTS KUBOTA LIMITED (AGRI MACHINERY GROUP)",
        "EICHER TRACTORS",
        "JOHN DEERE INDIA  PVT LTD(TRACTOR DEVISION)",
        "TAFE LIMITED",
        "INTERNATIONAL TRACTORS LIMITED",
        "CNH INDUSTRIAL (INDIA) PVT LTD",
    ],
}

# Statewise (selected states for dashboard)
STATES = [
    "UTTAR PRADESH",
    "MAHARASHTRA",
    "TAMIL NADU",
    "GUJARAT",
    "KARNATAKA",
    "MADHYA PRADESH",
    "RAJASTHAN",
    "BIHAR",
    "WEST BENGAL",
    "HARYANA",
    "ANDHRA PRADESH",
    "ODISHA",
    "KERALA",
]

# Fuel Wise (selected fuel types for dashboard)
FUELS = [
    "PETROL/ETHANOL",
    "PETROL",
    "DIESEL",
    "PURE EV",
    "PETROL(E20)/CNG",
    "ELECTRIC(BOV)",
    "CNG ONLY",
    "PETROL/CNG",
    "PETROL(E20)/HYBRID",
    "STRONG HYBRID EV",
    "PETROL/HYBRID",
    "DIESEL/HYBRID",
]


def normalize_key(s):
    """Normalize for matching: strip, upper, collapse spaces."""
    import re
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s).strip().upper())


def get_maker_to_category():
    """Return dict: normalized_maker_name -> category (Passenger Vehicles, EV-PV, 2W, CV, Tractors)."""
    mapping = {}
    for category, makers in MAKER_CLASSIFICATION.items():
        for m in makers:
            mapping[normalize_key(m)] = category
    return mapping


def get_maker_category_for_key(key, maker_to_cat):
    """Return category for a raw Key from Excel, or None if not in classification."""
    n = normalize_key(key)
    if n in maker_to_cat:
        return maker_to_cat[n]
    # Try partial match (key contains one of our maker names)
    for maker_n, cat in maker_to_cat.items():
        if maker_n in n or n in maker_n:
            return cat
    return None
