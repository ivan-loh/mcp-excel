#!/usr/bin/env python3
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, timedelta
import random

output_dir = Path(__file__).parent
output_dir.mkdir(exist_ok=True)

print("Creating CNC operations examples for Tech Holdings Berhad")
print("Malaysian SME manufacturer - 3 plants in Penang, 508 employees\n")

np.random.seed(42)
random.seed(42)

COMPANY_NAME = "Tech Holdings Berhad"
COMPANY_INFO = {
    "name": COMPANY_NAME,
    "plants": ["Penang Plant 1 (Flagship)", "Penang Plant 2 (Overflow)", "Penang Plant 3 (Specialty)"],
    "employees_total": 508,
    "cnc_division": 52,
    "fiscal_year": "FY2025"
}

QUARTERLY_METRICS = {
    "Q1": {"revenue": 4110000, "jobs": 195, "failure_rate": 0.32, "scrap_rate": 0.18, "oee": 0.52},
    "Q2": {"revenue": 3450000, "jobs": 162, "failure_rate": 0.38, "scrap_rate": 0.22, "oee": 0.48},
    "Q3": {"revenue": 3200000, "jobs": 148, "failure_rate": 0.24, "scrap_rate": 0.14, "oee": 0.54},
    "Q4": {"revenue": 3350000, "jobs": 155, "failure_rate": 0.20, "scrap_rate": 0.12, "oee": 0.56}
}

QUARTERS = [
    {"name": "Q1", "start": datetime(2024, 10, 1), "end": datetime(2024, 12, 31)},
    {"name": "Q2", "start": datetime(2025, 1, 1), "end": datetime(2025, 3, 31)},
    {"name": "Q3", "start": datetime(2025, 4, 1), "end": datetime(2025, 6, 30)},
    {"name": "Q4", "start": datetime(2025, 7, 1), "end": datetime(2025, 9, 30)}
]

def random_date(start, end):
    return start + timedelta(days=random.randint(0, (end - start).days))

def weighted_choice(choices, weights):
    return random.choices(choices, weights=weights, k=1)[0]

print("Generating master data...")

CUSTOMERS = [
    ("Aerospace Components Sdn Bhd", "Aerospace", "High"),
    ("Penang Precision Engineering", "Aerospace", "High"),
    ("Medical Devices Malaysia", "Medical", "High"),
    ("Flextronics Manufacturing", "Electronics", "Medium"),
    ("Intel Equipment Supplier", "Electronics", "High"),
    ("Seagate Vendor Services", "Electronics", "Medium"),
    ("Petronas Fabrication", "Oil & Gas", "Medium"),
    ("Shell Equipment Division", "Oil & Gas", "Medium"),
    ("Proton Auto Parts", "Automotive", "Medium"),
    ("Perodua Components", "Automotive", "Low"),
    ("Honda Supplier Network", "Automotive", "Medium"),
    ("Dyson Manufacturing", "Consumer", "High"),
    ("Bose Acoustics Malaysia", "Consumer", "Medium"),
    ("GE Healthcare Parts", "Medical", "High"),
    ("Bosch Asia Pacific", "Industrial", "Medium"),
    ("Schneider Electric MY", "Industrial", "Medium"),
    ("ABB Drives Malaysia", "Industrial", "Low"),
    ("Siemens Equipment", "Industrial", "Medium"),
    ("Honeywell Asia", "Industrial", "Medium"),
    ("Emerson Process", "Industrial", "Low"),
    ("Atlas Copco MY", "Industrial", "Low"),
    ("Ingersoll Rand", "Industrial", "Low"),
    ("Parker Hannifin", "Industrial", "Medium"),
    ("Eaton Malaysia", "Industrial", "Low"),
    ("Rockwell Automation", "Industrial", "Medium"),
    ("Daikin Aircon Parts", "HVAC", "Low"),
    ("Carrier Malaysia", "HVAC", "Low"),
    ("Thermo Fisher", "Medical", "High"),
    ("Olympus Medical", "Medical", "High"),
    ("Johnson Controls", "Industrial", "Medium"),
    ("TE Connectivity", "Electronics", "Medium"),
    ("Molex Penang", "Electronics", "Medium"),
    ("Jabil Circuit", "Electronics", "Medium"),
    ("Sanmina Corporation", "Electronics", "Low"),
    ("Plexus Corp", "Electronics", "Low"),
    ("Benchmark Electronics", "Electronics", "Low"),
    ("Celestica Malaysia", "Electronics", "Low"),
    ("Flex Ltd", "Electronics", "Medium"),
    ("Venture Corporation", "Electronics", "Low"),
    ("Unisem Berhad", "Semiconductor", "Medium"),
    ("Inari Amertron", "Semiconductor", "Medium"),
    ("Vitrox Corporation", "Semiconductor", "High"),
    ("MI Equipment", "Semiconductor", "Medium"),
    ("Pentamaster", "Automation", "Medium"),
    ("SKF Bearings", "Industrial", "Low"),
    ("Timken Malaysia", "Industrial", "Low"),
    ("NSK Bearings", "Industrial", "Low"),
    ("Schaeffler MY", "Industrial", "Low"),
    ("Grundfos Pumps", "Industrial", "Low"),
    ("KSB Pumps", "Industrial", "Low"),
    ("Flowserve Malaysia", "Industrial", "Low"),
    ("Sulzer Pumps", "Industrial", "Low"),
    ("Xylem Water", "Industrial", "Low"),
    ("Yokogawa MY", "Instrumentation", "Medium"),
    ("Endress+Hauser", "Instrumentation", "Medium"),
    ("Rosemount Penang", "Instrumentation", "Medium"),
    ("Krohne Malaysia", "Instrumentation", "Low"),
]

MACHINES = [
    ("PN1-VMC-01", "Haas VF-2", "3-Axis VMC", "Penang Plant 1", datetime(2018, 3, 15), 150),
    ("PN1-VMC-02", "Haas VF-3", "3-Axis VMC", "Penang Plant 1", datetime(2019, 6, 20), 150),
    ("PN1-VMC-03", "Mazak VCN-530C", "3-Axis VMC", "Penang Plant 1", datetime(2017, 11, 10), 165),
    ("PN1-VMC-04", "DMG Mori NVX-5080", "3-Axis VMC", "Penang Plant 1", datetime(2020, 8, 5), 175),
    ("PN1-HMC-01", "Mazak HCN-5000", "4-Axis HMC", "Penang Plant 1", datetime(2019, 2, 28), 180),
    ("PN1-5AX-01", "DMG Mori NHX-4000", "5-Axis", "Penang Plant 1", datetime(2021, 5, 12), 220),
    ("PN2-VMC-01", "Haas VF-2SS", "3-Axis VMC", "Penang Plant 2", datetime(2016, 9, 8), 145),
    ("PN2-VMC-02", "Haas VF-4", "3-Axis VMC", "Penang Plant 2", datetime(2017, 4, 22), 150),
    ("PN2-VMC-03", "Mazak VTC-300C", "3-Axis VMC", "Penang Plant 2", datetime(2018, 12, 3), 155),
    ("PN2-VMC-04", "Fanuc Robodrill", "3-Axis VMC", "Penang Plant 2", datetime(2019, 10, 17), 140),
    ("PN2-HMC-01", "Doosan HP-4000", "4-Axis HMC", "Penang Plant 2", datetime(2020, 3, 25), 170),
    ("PN3-VMC-01", "Mazak VCN-410A", "3-Axis VMC", "Penang Plant 3", datetime(2015, 7, 14), 135),
    ("PN3-VMC-02", "Haas VF-2", "3-Axis VMC", "Penang Plant 3", datetime(2017, 1, 30), 150),
    ("PN3-VMC-03", "DMG Mori CMX-600V", "3-Axis VMC", "Penang Plant 3", datetime(2018, 5, 9), 160),
    ("PN3-5AX-01", "Hermle C30U", "5-Axis", "Penang Plant 3", datetime(2020, 11, 19), 210),
    ("PN3-TURN-01", "Mazak QT-250", "CNC Lathe", "Penang Plant 3", datetime(2019, 8, 7), 130),
    ("PN3-TURN-02", "Haas ST-20", "CNC Lathe", "Penang Plant 3", datetime(2020, 2, 14), 135),
    ("PN1-VMC-05", "Makino V33i", "3-Axis VMC", "Penang Plant 1", datetime(2021, 9, 1), 185),
]

MATERIALS = [
    ("Aluminum 6061-T6", 12.50, "kg", "Alcan Metals"),
    ("Aluminum 7075-T6", 18.75, "kg", "Alcan Metals"),
    ("Aluminum 2024-T3", 16.25, "kg", "Alcan Metals"),
    ("Steel 4140 (Alloy)", 15.80, "kg", "Southern Steel"),
    ("Steel 1045 (Carbon)", 11.20, "kg", "Southern Steel"),
    ("Stainless 304", 22.50, "kg", "Inox Malaysia"),
    ("Stainless 316", 28.75, "kg", "Inox Malaysia"),
    ("Titanium Ti-6Al-4V", 185.00, "kg", "Titanium Imports"),
    ("Brass C36000", 24.50, "kg", "Non-Ferrous Supply"),
    ("Tool Steel H13", 32.00, "kg", "Specialty Metals"),
    ("Acetal (Delrin)", 18.50, "kg", "Polymers MY"),
    ("PEEK", 125.00, "kg", "Engineering Plastics"),
]

PART_TYPES = [
    ("Bracket", "Simple", [1, 2, 3]),
    ("Housing", "Medium", [2, 3, 4, 5]),
    ("Fixture", "Simple", [1, 2, 3]),
    ("Mounting Plate", "Simple", [1, 2, 3]),
    ("Connector Block", "Medium", [2, 3, 4]),
    ("Manifold", "Complex", [4, 5, 6]),
    ("Valve Body", "Complex", [4, 5, 6]),
    ("Precision Jig", "Complex", [3, 4, 5, 6]),
    ("Gear Housing", "Medium", [3, 4, 5]),
    ("Instrument Case", "Medium", [2, 3, 4]),
    ("Heat Sink", "Simple", [1, 2, 3]),
    ("Flange", "Simple", [1, 2, 3]),
    ("Shaft Coupling", "Medium", [3, 4]),
    ("Pulley", "Simple", [1, 3]),
    ("Spacer", "Simple", [1, 2, 3]),
    ("Adapter Plate", "Simple", [1, 2, 3]),
    ("Clamp Block", "Simple", [1, 2, 3]),
    ("End Cap", "Medium", [2, 3, 4]),
    ("Bearing Block", "Medium", [3, 4, 5]),
    ("Gearbox Cover", "Medium", [3, 4, 5]),
    ("Mold Component", "Complex", [5, 6]),
    ("Die Insert", "Complex", [5, 6]),
    ("Precision Shaft", "Medium", [3, 4, 16, 17]),
    ("Bushing", "Simple", [1, 3, 16]),
    ("Nozzle", "Complex", [4, 5, 16]),
]

EMPLOYEES = [
    ("Ahmad Zaki bin Hassan", "CNC Operator", "Senior", 42, "Penang Plant 1", "Day"),
    ("Lee Kah Wai", "CNC Operator", "Senior", 42, "Penang Plant 1", "Day"),
    ("Muthu Kumar a/l Rajan", "CNC Operator", "Mid", 32, "Penang Plant 1", "Day"),
    ("Wong Siew Ling", "CNC Operator", "Mid", 28, "Penang Plant 1", "Afternoon"),
    ("Azman bin Yusof", "CNC Operator", "Mid", 30, "Penang Plant 1", "Afternoon"),
    ("Tan Bee Lian", "CNC Operator", "Junior", 22, "Penang Plant 1", "Night"),
    ("Rajesh s/o Krishnan", "CNC Operator", "Junior", 24, "Penang Plant 1", "Night"),
    ("Nur Aisyah binti Ismail", "CNC Operator", "Mid", 29, "Penang Plant 1", "Day"),
    ("Lim Choon Hock", "CNC Operator", "Senior", 38, "Penang Plant 1", "Afternoon"),
    ("Siti Aminah binti Rahman", "CNC Operator", "Junior", 23, "Penang Plant 1", "Night"),
    ("Gopal a/l Subramaniam", "CNC Programmer", "Senior", 52, "Penang Plant 1", "Day"),
    ("Chen Wei Jie", "CNC Programmer", "Senior", 48, "Penang Plant 1", "Day"),
    ("Farah binti Abdullah", "CNC Programmer", "Mid", 35, "Penang Plant 1", "Day"),
    ("Kumar s/o Raman", "CNC Programmer", "Mid", 32, "Penang Plant 1", "Afternoon"),
    ("Tan Ai Ling", "CNC Programmer", "Junior", 26, "Penang Plant 1", "Day"),
    ("Hassan bin Omar", "Quality Inspector", "Senior", 45, "Penang Plant 1", "Day"),
    ("Liew Mei Fang", "Quality Inspector", "Mid", 31, "Penang Plant 1", "Day"),
    ("Selvam s/o Arumugam", "Quality Inspector", "Mid", 33, "Penang Plant 1", "Afternoon"),
    ("Nurul Huda binti Zainal", "Quality Inspector", "Junior", 25, "Penang Plant 1", "Afternoon"),
    ("Chong Kah Ming", "Production Supervisor", "Senior", 48, "Penang Plant 1", "Day"),
    ("Ramesh a/l Perumal", "CNC Operator", "Senior", 40, "Penang Plant 2", "Day"),
    ("Ng Pei Shan", "CNC Operator", "Mid", 31, "Penang Plant 2", "Day"),
    ("Ismail bin Ali", "CNC Operator", "Mid", 29, "Penang Plant 2", "Afternoon"),
    ("Lau Siew Cheng", "CNC Operator", "Junior", 23, "Penang Plant 2", "Afternoon"),
    ("Bala s/o Sundram", "CNC Operator", "Mid", 28, "Penang Plant 2", "Night"),
    ("Zainab binti Mahmud", "CNC Operator", "Junior", 24, "Penang Plant 2", "Night"),
    ("Koh Boon Huat", "CNC Operator", "Senior", 39, "Penang Plant 2", "Day"),
    ("Devi a/p Ganesh", "CNC Operator", "Mid", 27, "Penang Plant 2", "Afternoon"),
    ("Muhammad Faiz bin Ismail", "CNC Programmer", "Mid", 34, "Penang Plant 2", "Day"),
    ("Yap Su Ling", "CNC Programmer", "Junior", 27, "Penang Plant 2", "Day"),
    ("Ganesh a/l Krishnan", "Quality Inspector", "Mid", 32, "Penang Plant 2", "Day"),
    ("Sarah binti Ahmad", "Quality Inspector", "Junior", 26, "Penang Plant 2", "Afternoon"),
    ("Ong Kah Meng", "Production Supervisor", "Senior", 46, "Penang Plant 2", "Day"),
    ("Suresh a/l Nathan", "CNC Operator", "Senior", 41, "Penang Plant 3", "Day"),
    ("Teo Hui Min", "CNC Operator", "Mid", 30, "Penang Plant 3", "Day"),
    ("Kamal bin Rashid", "CNC Operator", "Mid", 28, "Penang Plant 3", "Afternoon"),
    ("Khoo Siew May", "CNC Operator", "Junior", 22, "Penang Plant 3", "Afternoon"),
    ("Prem Kumar a/l Suresh", "CNC Operator", "Mid", 29, "Penang Plant 3", "Night"),
    ("Leong Mei Yee", "CNC Operator", "Junior", 23, "Penang Plant 3", "Night"),
    ("Aziz bin Jaafar", "CNC Operator", "Senior", 37, "Penang Plant 3", "Day"),
    ("Neo Wei Jie", "CNC Programmer", "Mid", 33, "Penang Plant 3", "Day"),
    ("Tay Geok Lian", "CNC Programmer", "Junior", 25, "Penang Plant 3", "Day"),
    ("Murugan a/l Raman", "Quality Inspector", "Mid", 34, "Penang Plant 3", "Day"),
    ("Chua Bee Lian", "Quality Inspector", "Junior", 24, "Penang Plant 3", "Afternoon"),
    ("Halim bin Razak", "Production Supervisor", "Senior", 47, "Penang Plant 3", "Day"),
    ("Rizal bin Hamzah", "Tooling Specialist", "Senior", 44, "Penang Plant 1", "Day"),
    ("Quek Siew Mei", "Tooling Specialist", "Mid", 30, "Penang Plant 2", "Day"),
    ("Siva s/o Kumar", "Maintenance Tech", "Senior", 43, "Penang Plant 1", "Day"),
    ("Ang Li Na", "Maintenance Tech", "Mid", 31, "Penang Plant 2", "Day"),
    ("Azhar bin Rahman", "Maintenance Tech", "Mid", 29, "Penang Plant 3", "Day"),
    ("Lim Pei Qi", "Planning Manager", "Senior", 42, "Penang Plant 1", "Day"),
    ("Selvam s/o Muthu", "Quality Manager", "Senior", 50, "Penang Plant 1", "Day"),
]

print(f"Master data loaded:")
print(f"  - {len(CUSTOMERS)} customers across 8 industries")
print(f"  - {len(MACHINES)} CNC machines (3-axis, 4-axis, 5-axis, lathes)")
print(f"  - {len(MATERIALS)} material types")
print(f"  - {len(PART_TYPES)} part types")
print(f"  - {len(EMPLOYEES)} employees in CNC division")

print("\n1. Creating customer_orders.xlsx...")
customer_records = []
customer_churn = {
    "Q1_lost": [],
    "Q2_lost": ["Sanmina Corporation", "Celestica Malaysia", "Atlas Copco MY"],
    "Q3_lost": ["Plexus Corp", "Benchmark Electronics"],
    "Q4_lost": [],
    "Q3_gained": ["Sulzer Pumps", "Xylem Water"],
    "Q4_gained": ["Krohne Malaysia", "Daikin Aircon Parts", "Carrier Malaysia", "Timken Malaysia", "NSK Bearings"]
}

for cust_name, industry, priority in CUSTOMERS:
    active_quarters = []

    if cust_name in customer_churn["Q2_lost"]:
        active_quarters = ["Q1"]
    elif cust_name in customer_churn["Q3_lost"]:
        active_quarters = ["Q1", "Q2"]
    elif cust_name in customer_churn["Q3_gained"]:
        active_quarters = ["Q3", "Q4"]
    elif cust_name in customer_churn["Q4_gained"]:
        active_quarters = ["Q4"]
    else:
        active_quarters = ["Q1", "Q2", "Q3", "Q4"]

    q1_orders = random.randint(2, 8) if "Q1" in active_quarters and priority == "High" else random.randint(0, 4) if "Q1" in active_quarters and priority == "Medium" else random.randint(0, 2) if "Q1" in active_quarters else 0
    q2_orders = random.randint(1, 6) if "Q2" in active_quarters and priority == "High" else random.randint(0, 3) if "Q2" in active_quarters and priority == "Medium" else random.randint(0, 1) if "Q2" in active_quarters else 0
    q3_orders = random.randint(1, 5) if "Q3" in active_quarters and priority == "High" else random.randint(0, 3) if "Q3" in active_quarters and priority == "Medium" else random.randint(0, 1) if "Q3" in active_quarters else 0
    q4_orders = random.randint(2, 6) if "Q4" in active_quarters and priority == "High" else random.randint(0, 4) if "Q4" in active_quarters and priority == "Medium" else random.randint(0, 2) if "Q4" in active_quarters else 0

    total_orders = q1_orders + q2_orders + q3_orders + q4_orders

    if total_orders > 0:
        quality_incidents = random.randint(0, max(1, int(total_orders * 0.15)))
        on_time_pct = round(random.uniform(60, 85) if quality_incidents > 2 else random.uniform(75, 95), 1)

        customer_records.append({
            "customer_name": cust_name,
            "industry": industry,
            "priority": priority,
            "q1_orders": q1_orders,
            "q2_orders": q2_orders,
            "q3_orders": q3_orders,
            "q4_orders": q4_orders,
            "total_orders": total_orders,
            "quality_incidents": quality_incidents,
            "on_time_delivery_pct": on_time_pct,
            "payment_terms": random.choice(["Net 30", "Net 45", "Net 60"]),
            "status": "Lost" if cust_name in customer_churn["Q2_lost"] + customer_churn["Q3_lost"] else "Active"
        })

customers_df = pd.DataFrame(customer_records)
customers_df.to_excel(output_dir / "customer_orders.xlsx", sheet_name="Customers", index=False)
print(f"  Created {len(customer_records)} customer records")

print("\n2. Creating machines.xlsx...")
machines_df = pd.DataFrame([
    {
        "machine_id": m[0],
        "manufacturer_model": m[1],
        "machine_type": m[2],
        "plant": m[3],
        "installation_date": m[4].strftime("%Y-%m-%d"),
        "book_value_rm": random.randint(180000, 850000) if "5-Axis" in m[2] else random.randint(120000, 280000),
        "hourly_rate_rm": m[5],
        "status": "Operational"
    }
    for m in MACHINES
])
machines_df.to_excel(output_dir / "machines.xlsx", sheet_name="Equipment", index=False)
print(f"  Created {len(MACHINES)} machine records")

print("\n3. Creating labor_tracking.xlsx...")
labor_df = pd.DataFrame([
    {
        "employee_id": f"EMP{i+1:04d}",
        "employee_name": e[0],
        "role": e[1],
        "skill_level": e[2],
        "hourly_rate_rm": e[3],
        "plant": e[4],
        "shift": e[5]
    }
    for i, e in enumerate(EMPLOYEES)
])
labor_df.to_excel(output_dir / "labor_tracking.xlsx", sheet_name="Employees", index=False)
print(f"  Created {len(EMPLOYEES)} employee records")

print("\n4. Generating job orders and execution data (this will take a moment)...")

all_jobs = []
all_executions = []
all_programs = []
all_inspections = []
all_scrap = []
all_costs = []
job_counter = 1

active_customers_by_quarter = {}
for q in ["Q1", "Q2", "Q3", "Q4"]:
    active_customers_by_quarter[q] = [c for c in customer_records if c[f"{q.lower()}_orders"] > 0]

for quarter in QUARTERS:
    q_name = quarter["name"]
    q_metrics = QUARTERLY_METRICS[q_name]
    target_jobs = q_metrics["jobs"]
    failure_rate = q_metrics["failure_rate"]
    scrap_rate = q_metrics["scrap_rate"]

    print(f"  Generating {q_name} jobs (target: {target_jobs})...")

    active_custs = active_customers_by_quarter[q_name]

    for _ in range(target_jobs):
        customer = random.choice(active_custs)
        cust_name = customer["customer_name"]
        cust_industry = customer["industry"]

        part_type, complexity, suitable_machines = random.choice(PART_TYPES)
        material = random.choice(MATERIALS)

        quantity = weighted_choice(
            [random.randint(5, 25), random.randint(25, 100), random.randint(100, 500)],
            [0.4, 0.4, 0.2]
        )

        complexity_score = {"Simple": random.randint(1, 3), "Medium": random.randint(4, 6), "Complex": random.randint(7, 10)}[complexity]

        order_date = random_date(quarter["start"] - timedelta(days=random.randint(10, 45)), quarter["start"] + timedelta(days=random.randint(0, 30)))
        lead_time_days = {"Simple": random.randint(10, 20), "Medium": random.randint(20, 35), "Complex": random.randint(35, 60)}[complexity]
        required_delivery = order_date + timedelta(days=lead_time_days)

        material_cost_per_part = material[1] * random.uniform(0.2, 2.5)
        labor_hours_per_part = complexity_score * random.uniform(0.3, 1.2)
        machine_hours_per_part = labor_hours_per_part * random.uniform(0.7, 0.9)

        estimated_material = material_cost_per_part * quantity
        estimated_labor = labor_hours_per_part * quantity * 32
        estimated_machine = machine_hours_per_part * quantity * 160
        estimated_tooling = complexity_score * quantity * random.uniform(1.5, 4.0)
        estimated_overhead = (estimated_material + estimated_labor + estimated_machine) * 0.25
        estimated_total = estimated_material + estimated_labor + estimated_machine + estimated_tooling + estimated_overhead

        margin_target = random.uniform(0.15, 0.35)
        quote_value = estimated_total / (1 - margin_target)

        job_id = f"JOB-{job_counter:05d}"
        part_number = f"{cust_industry[:3].upper()}-{part_type[:4].upper()}-{job_counter:04d}"

        base_failure_rate = failure_rate

        if complexity_score >= 8:
            base_failure_rate *= 1.6
        elif complexity_score >= 6:
            base_failure_rate *= 1.3
        elif complexity_score <= 3:
            base_failure_rate *= 0.6

        will_fail_first_article = random.random() < base_failure_rate
        will_have_scrap = random.random() < scrap_rate

        prove_out_attempts = 1
        if will_fail_first_article:
            if complexity_score >= 8:
                prove_out_attempts = random.choice([2, 2, 3, 3, 4, 4, 5])
            elif complexity_score >= 6:
                prove_out_attempts = random.choice([2, 2, 3, 3, 4])
            elif complexity_score >= 5:
                prove_out_attempts = random.choice([2, 2, 3])
            else:
                prove_out_attempts = 2

        programming_date = order_date + timedelta(days=random.randint(1, 5))
        prove_out_start = programming_date + timedelta(days=random.randint(2, 7))
        prove_out_end = prove_out_start + timedelta(days=(prove_out_attempts - 1) * random.randint(1, 3))
        production_start = prove_out_end + timedelta(days=random.randint(1, 3))
        production_end = production_start + timedelta(days=max(1, int(quantity * machine_hours_per_part / 16)))
        inspection_date = production_end + timedelta(days=random.randint(0, 2))
        actual_delivery = inspection_date + timedelta(days=random.randint(1, 3))

        if will_fail_first_article or will_have_scrap:
            actual_delivery = actual_delivery + timedelta(days=random.randint(2, 10))

        on_time = actual_delivery <= required_delivery

        scrap_qty = 0
        if will_have_scrap:
            scrap_qty = int(quantity * random.uniform(0.02, 0.15))

        qty_delivered = quantity - scrap_qty

        status = "Shipped" if actual_delivery <= datetime.now() else "Production"

        machine_id = random.choice([m[0] for m in MACHINES if any(str(sm) in m[0] for sm in suitable_machines)])
        plant = [m[3] for m in MACHINES if m[0] == machine_id][0]

        all_jobs.append({
            "job_id": job_id,
            "customer_name": cust_name,
            "industry": cust_industry,
            "part_number": part_number,
            "part_type": part_type,
            "material_type": material[0],
            "complexity": complexity,
            "complexity_score": complexity_score,
            "order_date": order_date.strftime("%Y-%m-%d"),
            "required_delivery": required_delivery.strftime("%Y-%m-%d"),
            "actual_delivery": actual_delivery.strftime("%Y-%m-%d") if status == "Shipped" else None,
            "on_time": on_time if status == "Shipped" else None,
            "quantity_ordered": quantity,
            "quantity_delivered": qty_delivered if status == "Shipped" else None,
            "quantity_scrapped": scrap_qty if scrap_qty > 0 else None,
            "quote_value_rm": round(quote_value, 2),
            "status": status,
            "plant": plant,
            "primary_machine": machine_id,
            "quarter": q_name
        })

        programmer = random.choice([e for e in EMPLOYEES if e[1] == "CNC Programmer" and e[4] == plant])
        program_id = f"PGM-{job_counter:05d}"

        programming_hours = complexity_score * random.uniform(1.5, 4.0)
        if programmer[2] == "Senior":
            programming_hours *= random.uniform(0.7, 0.9)
        elif programmer[2] == "Junior":
            programming_hours *= random.uniform(1.2, 1.5)

        simulation_performed = False
        if q_name in ["Q3", "Q4"]:
            simulation_performed = random.random() < 0.75
        else:
            simulation_performed = random.random() < 0.15

        error_types = []
        if prove_out_attempts > 1:
            possible_errors = ["Tool interference", "Feed rate too high", "Collision risk", "Dimensional error", "Surface finish", "Incorrect tool"]
            error_types = random.sample(possible_errors, random.randint(1, min(3, prove_out_attempts)))

        all_programs.append({
            "program_id": program_id,
            "job_id": job_id,
            "part_number": part_number,
            "programmer_name": programmer[0],
            "programmer_skill": programmer[2],
            "programming_date": programming_date.strftime("%Y-%m-%d"),
            "programming_hours": round(programming_hours, 2),
            "cam_software": "Mastercam",
            "post_processor": f"{machine_id.split('-')[1]}_Post",
            "simulation_performed": simulation_performed,
            "prove_out_attempts": prove_out_attempts,
            "first_article_pass": prove_out_attempts == 1,
            "errors_found": "; ".join(error_types) if error_types else None,
            "cost_per_attempt_rm": round(complexity_score * random.uniform(500, 1500), 2)
        })

        for attempt in range(1, prove_out_attempts + 1):
            operator = random.choice([e for e in EMPLOYEES if e[1] == "CNC Operator" and e[4] == plant])

            attempt_date = prove_out_start + timedelta(days=(attempt - 1) * random.randint(1, 3))
            setup_hours = complexity_score * random.uniform(0.5, 2.0)
            if attempt > 1:
                setup_hours *= 0.7

            run_hours = machine_hours_per_part * (1 if attempt == prove_out_attempts else 0.5)

            prove_out_result = "Pass" if attempt == prove_out_attempts else random.choice(["Fail", "Fail", "Rework"])
            first_article_result = "Pass" if attempt == prove_out_attempts and not will_fail_first_article else "Fail" if attempt < prove_out_attempts else "Pass"

            scrap_value = 0
            rework_hours = 0
            if prove_out_result == "Fail":
                scrap_value = round(material_cost_per_part * random.uniform(1, 3), 2)
            elif prove_out_result == "Rework":
                rework_hours = random.uniform(2, 6)

            all_executions.append({
                "execution_id": f"EXE-{job_counter:05d}-{attempt:02d}",
                "job_id": job_id,
                "attempt_number": attempt,
                "machine_id": machine_id,
                "operator_name": operator[0],
                "operator_skill": operator[2],
                "shift": operator[5],
                "program_id": program_id,
                "execution_date": attempt_date.strftime("%Y-%m-%d"),
                "setup_hours": round(setup_hours, 2),
                "run_hours": round(run_hours, 2),
                "total_hours": round(setup_hours + run_hours, 2),
                "prove_out_result": prove_out_result if attempt < prove_out_attempts else "Pass",
                "first_article_result": first_article_result if attempt == prove_out_attempts else None,
                "scrap_value_rm": scrap_value if scrap_value > 0 else None,
                "rework_hours": rework_hours if rework_hours > 0 else None
            })

        if scrap_qty > 0:
            scrap_reasons = ["Dimensional error", "Tool breakage", "Material defect", "Operator error", "Surface finish"]
            scrap_reason = weighted_choice(scrap_reasons, [0.35, 0.15, 0.10, 0.25, 0.15])

            preventable = scrap_reason in ["Dimensional error", "Tool breakage", "Operator error"]

            all_scrap.append({
                "scrap_id": f"SCR-{job_counter:05d}",
                "job_id": job_id,
                "scrap_date": random_date(production_start, production_end).strftime("%Y-%m-%d"),
                "quantity_scrapped": scrap_qty,
                "scrap_reason": scrap_reason,
                "material_cost_lost_rm": round(scrap_qty * material_cost_per_part, 2),
                "labor_hours_lost": round(scrap_qty * labor_hours_per_part * 0.5, 2),
                "machine_hours_lost": round(scrap_qty * machine_hours_per_part * 0.5, 2),
                "total_cost_impact_rm": round(scrap_qty * (material_cost_per_part + labor_hours_per_part * 32 + machine_hours_per_part * 160) * 0.5, 2),
                "preventable": preventable
            })

        inspector = random.choice([e for e in EMPLOYEES if e[1] == "Quality Inspector" and e[4] == plant])

        inspection_result = "Pass" if not will_fail_first_article or prove_out_attempts > 1 else random.choice(["Pass", "Conditional", "Rework"])

        critical_dims = random.randint(3, 8)
        for dim_num in range(1, critical_dims + 1):
            nominal = round(random.uniform(10.0, 100.0), 3)
            tolerance = round(random.uniform(0.01, 0.5), 3)

            if inspection_result == "Pass":
                actual = round(nominal + random.uniform(-tolerance * 0.6, tolerance * 0.6), 3)
            else:
                actual = round(nominal + random.uniform(-tolerance * 1.2, tolerance * 1.2), 3)

            within_tolerance = abs(actual - nominal) <= tolerance

            all_inspections.append({
                "inspection_id": f"INS-{job_counter:05d}-D{dim_num:02d}",
                "job_id": job_id,
                "part_number": part_number,
                "inspection_type": "First Article",
                "inspector_name": inspector[0],
                "inspection_date": inspection_date.strftime("%Y-%m-%d"),
                "dimension_number": dim_num,
                "nominal_value": nominal,
                "tolerance": tolerance,
                "actual_value": actual,
                "within_tolerance": within_tolerance,
                "overall_result": inspection_result
            })

        actual_material = estimated_material * (1 + scrap_qty / quantity)
        actual_labor = estimated_labor * (1 + (prove_out_attempts - 1) * 0.3)
        actual_machine = estimated_machine * (1 + (prove_out_attempts - 1) * 0.3)
        actual_tooling = estimated_tooling * random.uniform(0.9, 1.3)
        actual_overhead = (actual_material + actual_labor + actual_machine) * 0.25
        actual_total = actual_material + actual_labor + actual_machine + actual_tooling + actual_overhead

        actual_margin = (quote_value - actual_total) / quote_value if quote_value > 0 else 0

        all_costs.append({
            "job_id": job_id,
            "estimated_material_rm": round(estimated_material, 2),
            "actual_material_rm": round(actual_material, 2),
            "estimated_labor_rm": round(estimated_labor, 2),
            "actual_labor_rm": round(actual_labor, 2),
            "estimated_machine_rm": round(estimated_machine, 2),
            "actual_machine_rm": round(actual_machine, 2),
            "estimated_tooling_rm": round(estimated_tooling, 2),
            "actual_tooling_rm": round(actual_tooling, 2),
            "estimated_overhead_rm": round(estimated_overhead, 2),
            "actual_overhead_rm": round(actual_overhead, 2),
            "estimated_total_cost_rm": round(estimated_total, 2),
            "actual_total_cost_rm": round(actual_total, 2),
            "quote_value_rm": round(quote_value, 2),
            "revenue_rm": round(quote_value, 2) if status == "Shipped" else None,
            "gross_margin_rm": round(quote_value - actual_total, 2) if status == "Shipped" else None,
            "gross_margin_pct": round(actual_margin * 100, 2) if status == "Shipped" else None,
            "cost_variance_rm": round(actual_total - estimated_total, 2),
            "cost_variance_pct": round((actual_total - estimated_total) / estimated_total * 100, 2) if estimated_total > 0 else 0
        })

        job_counter += 1

jobs_df = pd.DataFrame(all_jobs)
jobs_df.to_excel(output_dir / "job_orders.xlsx", sheet_name="Orders", index=False)
print(f"  Created {len(all_jobs)} job orders")

executions_df = pd.DataFrame(all_executions)
executions_df.to_excel(output_dir / "job_execution.xlsx", sheet_name="Execution", index=False)
print(f"  Created {len(all_executions)} execution records")

programs_df = pd.DataFrame(all_programs)
programs_df.to_excel(output_dir / "program_validation.xlsx", sheet_name="Programs", index=False)
print(f"  Created {len(all_programs)} program records")

inspections_df = pd.DataFrame(all_inspections)
inspections_df.to_excel(output_dir / "quality_inspections.xlsx", sheet_name="Inspections", index=False)
print(f"  Created {len(all_inspections)} inspection records")

scrap_df = pd.DataFrame(all_scrap)
scrap_df.to_excel(output_dir / "scrap_rework.xlsx", sheet_name="Scrap", index=False)
print(f"  Created {len(all_scrap)} scrap records")

costs_df = pd.DataFrame(all_costs)
costs_df.to_excel(output_dir / "cost_analysis.xlsx", sheet_name="Job_Costing", index=False)
print(f"  Created {len(all_costs)} cost analysis records")

print("\n5. Creating material_inventory.xlsx...")
material_records = []
for quarter in QUARTERS:
    for week_num in range(13):
        week_start = quarter["start"] + timedelta(weeks=week_num)

        for mat_name, unit_cost, unit, supplier in MATERIALS:
            beginning = round(random.uniform(500, 2000), 2)
            purchases = round(random.uniform(200, 800), 2)
            issued = round(random.uniform(300, 900), 2)
            ending = beginning + purchases - issued

            material_records.append({
                "week_starting": week_start.strftime("%Y-%m-%d"),
                "material_type": mat_name,
                "unit_cost_rm": unit_cost,
                "unit": unit,
                "supplier": supplier,
                "beginning_inventory_kg": beginning,
                "purchases_kg": purchases,
                "issued_to_jobs_kg": issued,
                "ending_inventory_kg": max(0, ending),
                "total_value_rm": round(max(0, ending) * unit_cost, 2)
            })

materials_df = pd.DataFrame(material_records)
materials_df.to_excel(output_dir / "material_inventory.xlsx", sheet_name="Inventory", index=False)
print(f"  Created {len(material_records)} material inventory records")

print("\n6. Creating tooling_management.xlsx...")
tool_types = [
    ("End Mill - 6mm", 45.00, 250),
    ("End Mill - 10mm", 52.00, 300),
    ("End Mill - 12mm", 58.00, 350),
    ("End Mill - 16mm", 68.00, 400),
    ("Ball End Mill - 6mm", 55.00, 200),
    ("Ball End Mill - 10mm", 62.00, 250),
    ("Drill - 5mm", 22.00, 400),
    ("Drill - 8mm", 28.00, 450),
    ("Drill - 10mm", 32.00, 500),
    ("Tap - M6", 35.00, 150),
    ("Tap - M8", 40.00, 180),
    ("Tap - M10", 45.00, 200),
    ("Boring Bar - 20mm", 185.00, 800),
    ("Boring Bar - 30mm", 225.00, 1000),
    ("Insert - CNMG", 15.00, 150),
    ("Insert - DCMT", 18.00, 200),
    ("Insert - WNMG", 16.00, 180),
    ("Face Mill - 50mm", 320.00, 2000),
    ("Face Mill - 80mm", 480.00, 2500),
]

tooling_records = []
for i, (tool_type, cost, expected_life) in enumerate(tool_types, 1):
    actual_life = expected_life * random.uniform(0.7, 1.2)
    quantity_used_q1 = random.randint(5, 25)
    quantity_used_q2 = random.randint(4, 22)
    quantity_used_q3 = random.randint(3, 20)
    quantity_used_q4 = random.randint(4, 21)

    tooling_records.append({
        "tool_id": f"TOOL-{i:04d}",
        "tool_type": tool_type,
        "cost_per_tool_rm": cost,
        "expected_life_parts": expected_life,
        "actual_avg_life_parts": round(actual_life, 0),
        "supplier": random.choice(["Sandvik", "Kennametal", "Mitsubishi", "Kyocera", "Seco Tools"]),
        "q1_quantity_used": quantity_used_q1,
        "q2_quantity_used": quantity_used_q2,
        "q3_quantity_used": quantity_used_q3,
        "q4_quantity_used": quantity_used_q4,
        "total_cost_fy2025_rm": round((quantity_used_q1 + quantity_used_q2 + quantity_used_q3 + quantity_used_q4) * cost, 2)
    })

tooling_df = pd.DataFrame(tooling_records)
tooling_df.to_excel(output_dir / "tooling_management.xlsx", sheet_name="Tools", index=False)
print(f"  Created {len(tooling_records)} tooling records")

print("\n7. Creating machine_downtime.xlsx...")
downtime_records = []
downtime_id = 1

for quarter in QUARTERS:
    q_name = quarter["name"]

    for machine_id, mfg, mtype, plant, install, rate in MACHINES:
        num_downtimes = random.randint(3, 8) if q_name in ["Q1", "Q2"] else random.randint(2, 5)

        for _ in range(num_downtimes):
            dt_start = random_date(quarter["start"], quarter["end"])

            if q_name in ["Q1", "Q2"]:
                dt_type = weighted_choice(
                    ["Unscheduled Breakdown", "Scheduled Maintenance", "Waiting for Program", "Waiting for Material"],
                    [0.5, 0.2, 0.2, 0.1]
                )
            else:
                dt_type = weighted_choice(
                    ["Unscheduled Breakdown", "Scheduled Maintenance", "Waiting for Program", "Waiting for Material"],
                    [0.3, 0.4, 0.2, 0.1]
                )

            if dt_type == "Unscheduled Breakdown":
                duration_hours = random.uniform(2, 24)
                root_cause = random.choice(["Spindle failure", "Tool changer malfunction", "Coolant system leak", "Control board error", "Power supply issue"])
                cost = round(random.uniform(500, 5000), 2)
            elif dt_type == "Scheduled Maintenance":
                duration_hours = random.uniform(4, 8)
                root_cause = "Preventive maintenance"
                cost = round(random.uniform(300, 1200), 2)
            elif dt_type == "Waiting for Program":
                duration_hours = random.uniform(1, 8)
                root_cause = "Program not ready"
                cost = 0
            else:
                duration_hours = random.uniform(2, 16)
                root_cause = "Material stock-out"
                cost = 0

            downtime_records.append({
                "downtime_id": f"DT-{downtime_id:05d}",
                "machine_id": machine_id,
                "plant": plant,
                "downtime_start": dt_start.strftime("%Y-%m-%d %H:%M"),
                "duration_hours": round(duration_hours, 2),
                "downtime_type": dt_type,
                "root_cause": root_cause,
                "repair_cost_rm": cost,
                "quarter": q_name
            })
            downtime_id += 1

downtime_df = pd.DataFrame(downtime_records)
downtime_df.to_excel(output_dir / "machine_downtime.xlsx", sheet_name="Downtime", index=False)
print(f"  Created {len(downtime_records)} downtime records")

print("\n8. Creating production_schedule.xlsx...")
schedule_records = []

for quarter in QUARTERS:
    for week_num in range(13):
        week_start = quarter["start"] + timedelta(weeks=week_num)

        for machine_id, mfg, mtype, plant, install, rate in MACHINES:
            available_hours = 24 * 6

            planned_hours = available_hours * random.uniform(0.7, 0.9)

            q_name = quarter["name"]
            if q_name in ["Q1", "Q2"]:
                actual_hours = planned_hours * random.uniform(0.65, 0.85)
            else:
                actual_hours = planned_hours * random.uniform(0.75, 0.92)

            variance_hours = actual_hours - planned_hours
            adherence_pct = round((actual_hours / planned_hours * 100) if planned_hours > 0 else 0, 1)

            if variance_hours < -5:
                delay_reasons = random.choice([
                    "Prove-out failures",
                    "Rework required",
                    "Machine breakdown",
                    "Material shortage",
                    "Program errors"
                ])
            else:
                delay_reasons = None

            schedule_records.append({
                "week_starting": week_start.strftime("%Y-%m-%d"),
                "machine_id": machine_id,
                "plant": plant,
                "available_hours": available_hours,
                "planned_hours": round(planned_hours, 2),
                "actual_hours": round(actual_hours, 2),
                "variance_hours": round(variance_hours, 2),
                "schedule_adherence_pct": adherence_pct,
                "delay_reason": delay_reasons,
                "quarter": quarter["name"]
            })

schedule_df = pd.DataFrame(schedule_records)
schedule_df.to_excel(output_dir / "production_schedule.xlsx", sheet_name="Schedule", index=False)
print(f"  Created {len(schedule_records)} schedule records")

print("\n=== PHASE 1: PRODUCTION DEPTH FILES ===\n")

print("9. Creating job_operations.xlsx (operation-level tracking)...")
operations_list = []
operation_types_map = {
    "Simple": ["Setup", "Facing", "Roughing", "Finishing", "Drilling"],
    "Medium": ["Setup", "Facing", "Roughing", "Semi-Finish", "Finishing", "Drilling", "Tapping"],
    "Complex": ["Setup", "Facing", "Roughing-Pass1", "Roughing-Pass2", "Semi-Finish", "Finishing", "Drilling", "Boring", "Tapping", "Inspection"]
}

op_id_counter = 1
for job in all_jobs:
    job_complexity = job["complexity"]
    operations = operation_types_map[job_complexity]

    job_execs = [e for e in all_executions if e["job_id"] == job["job_id"]]
    if not job_execs:
        continue

    base_exec = job_execs[0]
    total_job_hours = sum([e["total_hours"] for e in job_execs])

    for seq, op_type in enumerate(operations, 1):
        time_pct = {
            "Setup": 0.25, "Facing": 0.08, "Roughing": 0.20, "Roughing-Pass1": 0.15,
            "Roughing-Pass2": 0.15, "Semi-Finish": 0.12, "Finishing": 0.15,
            "Drilling": 0.08, "Boring": 0.10, "Tapping": 0.07, "Inspection": 0.05
        }.get(op_type, 0.10)

        planned_time = total_job_hours * time_pct
        actual_time = planned_time * random.uniform(0.85, 1.35)

        parts_in_op = job["quantity_ordered"]
        parts_scrapped_in_op = 0
        if op_type in ["Finishing", "Drilling", "Tapping"] and random.random() < 0.08:
            parts_scrapped_in_op = random.randint(1, max(1, int(parts_in_op * 0.05)))

        issues = []
        if actual_time > planned_time * 1.2:
            issues.append(random.choice(["Tool wear", "Fixture slip", "Chatter", "Coolant issue"]))

        operations_list.append({
            "operation_id": f"OP-{op_id_counter:06d}",
            "job_id": job["job_id"],
            "sequence_number": seq,
            "operation_type": op_type,
            "machine_id": job["primary_machine"],
            "operator_name": base_exec["operator_name"],
            "shift": base_exec["shift"],
            "planned_cycle_time_minutes": round(planned_time * 60, 2),
            "actual_cycle_time_minutes": round(actual_time * 60, 2),
            "parts_completed": parts_in_op - parts_scrapped_in_op,
            "parts_scrapped": parts_scrapped_in_op if parts_scrapped_in_op > 0 else None,
            "tool_changes_during_op": random.randint(0, 3) if op_type not in ["Setup", "Inspection"] else 0,
            "issues_encountered": "; ".join(issues) if issues else None
        })
        op_id_counter += 1

operations_df = pd.DataFrame(operations_list)
operations_df.to_excel(output_dir / "job_operations.xlsx", sheet_name="Operations", index=False)
print(f"  Created {len(operations_list)} operation records")

print("\n10. Creating tool_life_tracking.xlsx (individual tool instances)...")
tool_instances = []
tool_id_counter = 1

for quarter in QUARTERS:
    for tool_type, cost, expected_life in [
        ("End Mill - 10mm", 52.00, 300), ("Drill - 8mm", 28.00, 450),
        ("Tap - M8", 40.00, 180), ("Ball End Mill - 10mm", 62.00, 250),
        ("Insert - CNMG", 15.00, 150)
    ]:
        installs_this_quarter = random.randint(8, 18)
        for _ in range(installs_this_quarter):
            install_date = random_date(quarter["start"], quarter["end"] - timedelta(days=10))

            actual_life = expected_life * random.uniform(0.5, 1.4)
            removal_date = install_date + timedelta(days=random.randint(5, 30))

            removal_reasons = ["Worn", "Worn", "Worn", "Broken", "Preventive"]
            removal_reason = weighted_choice(removal_reasons, [0.5, 0.2, 0.1, 0.15, 0.05])

            if removal_reason == "Broken":
                actual_life *= 0.4

            performance = "Excellent" if actual_life > expected_life * 1.1 else "Good" if actual_life > expected_life * 0.9 else "Poor"

            tool_instances.append({
                "tool_instance_id": f"TI-{tool_id_counter:05d}",
                "tool_type": tool_type,
                "serial_number": f"SN{random.randint(100000, 999999)}",
                "supplier": random.choice(["Sandvik", "Kennametal", "Mitsubishi"]),
                "purchase_date": (install_date - timedelta(days=random.randint(30, 90))).strftime("%Y-%m-%d"),
                "cost_rm": cost,
                "installed_on_machine": random.choice([m[0] for m in MACHINES]),
                "installation_date": install_date.strftime("%Y-%m-%d"),
                "expected_life_parts": expected_life,
                "actual_parts_completed": int(actual_life),
                "removal_reason": removal_reason,
                "removal_date": removal_date.strftime("%Y-%m-%d"),
                "total_cutting_time_hours": round(actual_life / 50 * random.uniform(0.8, 1.2), 2),
                "performance_rating": performance,
                "quarter": quarter["name"]
            })
            tool_id_counter += 1

tools_df = pd.DataFrame(tool_instances)
tools_df.to_excel(output_dir / "tool_life_tracking.xlsx", sheet_name="Tool_Instances", index=False)
print(f"  Created {len(tool_instances)} tool instance records")

print("\n11. Creating work_in_progress.xlsx (WIP tracking)...")
wip_records = []
wip_id = 1

stages = ["Queued-Programming", "Queued-Machine", "In-Process", "Queued-Inspection", "Complete"]
for job in all_jobs[:350]:
    stage_entry_time = datetime.strptime(job["order_date"], "%Y-%m-%d")

    for stage in stages:
        if stage == "Queued-Programming":
            duration = random.uniform(24, 120)
            reason = "Programmer backlog" if duration > 72 else None
        elif stage == "Queued-Machine":
            duration = random.uniform(8, 96)
            reason = "Machine busy" if duration > 48 else None
        elif stage == "In-Process":
            duration = random.uniform(4, 48)
            reason = None
        elif stage == "Queued-Inspection":
            duration = random.uniform(2, 24)
            reason = "Inspector backlog" if duration > 12 else None
        else:
            duration = 0
            reason = None

        exit_time = stage_entry_time + timedelta(hours=duration)

        wip_records.append({
            "wip_id": f"WIP-{wip_id:06d}",
            "job_id": job["job_id"],
            "stage": stage,
            "location": job["primary_machine"] if stage in ["Queued-Machine", "In-Process"] else job["plant"],
            "quantity": job["quantity_ordered"],
            "entered_stage_datetime": stage_entry_time.strftime("%Y-%m-%d %H:%M"),
            "exited_stage_datetime": exit_time.strftime("%Y-%m-%d %H:%M"),
            "duration_hours": round(duration, 2),
            "reason_for_delay": reason
        })
        wip_id += 1
        stage_entry_time = exit_time

wip_df = pd.DataFrame(wip_records)
wip_df.to_excel(output_dir / "work_in_progress.xlsx", sheet_name="WIP_Tracking", index=False)
print(f"  Created {len(wip_records)} WIP transition records")

print("\n12. Creating setup_changeovers.xlsx (setup time detail)...")
changeover_records = []
changeover_id = 1

for machine_id, mfg, mtype, plant, install, rate in MACHINES:
    machine_jobs = [j for j in all_jobs if j["primary_machine"] == machine_id]
    machine_jobs.sort(key=lambda x: x["order_date"])

    for i in range(1, min(len(machine_jobs), 50)):
        prev_job = machine_jobs[i-1]
        curr_job = machine_jobs[i]

        prev_date = datetime.strptime(prev_job["order_date"], "%Y-%m-%d") + timedelta(days=5)
        curr_date = datetime.strptime(curr_job["order_date"], "%Y-%m-%d")

        if prev_date >= curr_date:
            changeover_date = prev_date
        else:
            changeover_date = random_date(prev_date, curr_date)

        similarity = 1.0 if prev_job["part_type"] == curr_job["part_type"] else 0.0

        planned_setup = random.uniform(30, 60)
        actual_setup = planned_setup * random.uniform(1.1, 1.8) if similarity < 0.5 else planned_setup * random.uniform(0.8, 1.2)

        changeover_records.append({
            "changeover_id": f"CO-{changeover_id:05d}",
            "machine_id": machine_id,
            "from_job_id": prev_job["job_id"],
            "to_job_id": curr_job["job_id"],
            "changeover_date": changeover_date.strftime("%Y-%m-%d"),
            "operator_name": random.choice([e[0] for e in EMPLOYEES if e[1] == "CNC Operator" and e[4] == plant]),
            "previous_part_family": prev_job["part_type"],
            "next_part_family": curr_job["part_type"],
            "similarity_score": similarity,
            "tool_change_time_minutes": round(random.uniform(15, 35), 1),
            "fixture_change_time_minutes": round(random.uniform(10, 25), 1),
            "workpiece_load_time_minutes": round(random.uniform(5, 12), 1),
            "program_load_test_time_minutes": round(random.uniform(8, 20), 1),
            "first_piece_probe_time_minutes": round(random.uniform(10, 18), 1),
            "total_setup_time_minutes": round(actual_setup, 1),
            "planned_setup_time_minutes": round(planned_setup, 1),
            "variance_minutes": round(actual_setup - planned_setup, 1),
            "issues_encountered": "Fixture alignment" if actual_setup > planned_setup * 1.4 else None
        })
        changeover_id += 1

changeovers_df = pd.DataFrame(changeover_records)
changeovers_df.to_excel(output_dir / "setup_changeovers.xlsx", sheet_name="Changeovers", index=False)
print(f"  Created {len(changeover_records)} changeover records")

print("\n13. Creating rework_tracking.xlsx (repair attempts)...")
rework_records = []
rework_id = 1

failed_jobs = [j for j in all_jobs if j["quantity_scrapped"] and j["quantity_scrapped"] > 0]
for job in failed_jobs[:min(len(failed_jobs), 120)]:
    scrap_qty = job["quantity_scrapped"]
    rework_attempt_qty = int(scrap_qty * random.uniform(0.4, 0.8))

    for part_num in range(rework_attempt_qty):
        failure_types = ["Undersize", "Oversize", "Wrong Feature", "Surface Finish", "Burrs"]
        failure_type = weighted_choice(failure_types, [0.25, 0.20, 0.15, 0.25, 0.15])

        rework_decision = weighted_choice(
            ["Repair", "Scrap", "Use-As-Is"],
            [0.60, 0.30, 0.10]
        )

        if rework_decision == "Repair":
            rework_result = "Saved" if random.random() < 0.65 else "Scrapped After Attempt"
            rework_hours = random.uniform(0.5, 3.0)
            rework_cost = rework_hours * 32 + random.uniform(50, 200)
        else:
            rework_result = "Not Attempted"
            rework_hours = 0
            rework_cost = 0

        rework_records.append({
            "rework_id": f"RWK-{rework_id:05d}",
            "job_id": job["job_id"],
            "part_serial_number": f"{job['part_number']}-{part_num+1:03d}",
            "failure_found_at": random.choice(["First Article", "In-Process", "Final Inspection"]),
            "failure_type": failure_type,
            "rework_decision": rework_decision,
            "rework_operation_description": f"Re-machine {failure_type}" if rework_decision == "Repair" else None,
            "machine_used": job["primary_machine"] if rework_decision == "Repair" else None,
            "operator": random.choice([e[0] for e in EMPLOYEES if e[1] == "CNC Operator"]) if rework_decision == "Repair" else None,
            "rework_hours": round(rework_hours, 2) if rework_hours > 0 else None,
            "rework_cost_rm": round(rework_cost, 2) if rework_cost > 0 else None,
            "rework_result": rework_result,
            "approval_level": random.choice(["Supervisor", "Quality Manager"]) if rework_decision != "Scrap" else None
        })
        rework_id += 1

rework_df = pd.DataFrame(rework_records)
rework_df.to_excel(output_dir / "rework_tracking.xlsx", sheet_name="Rework", index=False)
print(f"  Created {len(rework_records)} rework attempt records")

print("\n" + "=" * 80)
print(f"✓ Created 18 CNC operations files for {COMPANY_NAME} (13 Base + 5 Phase 1)")
print("=" * 80)
print(f"Company: {COMPANY_NAME}")
print(f"Business: SME manufacturer - CNC machining division")
print(f"Plants: {', '.join(COMPANY_INFO['plants'])}")
print(f"Timeline: Q1 FY2025 (Oct-Dec 2024) through Q4 FY2025 (Jul-Sep 2025)")
print("=" * 80)
print("\nQuarterly Performance:")
for q in ["Q1", "Q2", "Q3", "Q4"]:
    m = QUARTERLY_METRICS[q]
    print(f"  {q}: RM{m['revenue']:,.0f} revenue, {m['jobs']} jobs, {m['failure_rate']*100:.0f}% failure rate, {m['scrap_rate']*100:.0f}% scrap, {m['oee']*100:.0f}% OEE")
print(f"\nTotal Revenue (4 quarters): RM{sum(QUARTERLY_METRICS[q]['revenue'] for q in ['Q1', 'Q2', 'Q3', 'Q4']):,.0f}")
print(f"Total Jobs: {sum(QUARTERLY_METRICS[q]['jobs'] for q in ['Q1', 'Q2', 'Q3', 'Q4'])}")
print("=" * 80)
print("\nBase Files (13):")
print("  1. customer_orders.xlsx         - Customer master & order history")
print("  2. machines.xlsx                 - Equipment master (18 machines)")
print("  3. labor_tracking.xlsx           - Workforce data (52 employees)")
print("  4. job_orders.xlsx               - Production orders master (~660 jobs)")
print("  5. job_execution.xlsx            - Run details with prove-out attempts")
print("  6. program_validation.xlsx       - CAM/G-code validation tracking")
print("  7. quality_inspections.xlsx      - Dimensional inspection records")
print("  8. scrap_rework.xlsx             - Waste tracking & root causes")
print("  9. material_inventory.xlsx       - Raw material management (weekly)")
print(" 10. tooling_management.xlsx       - Tool usage & costs")
print(" 11. machine_downtime.xlsx         - Equipment failures & maintenance")
print(" 12. cost_analysis.xlsx            - Job costing with variance")
print(" 13. production_schedule.xlsx      - Weekly planning vs actual")
print("\nPhase 1: Production Depth Files (5):")
print(" 14. job_operations.xlsx           - Operation-level tracking (~4,000 operations)")
print(" 15. tool_life_tracking.xlsx       - Individual tool instances (~200-300)")
print(" 16. work_in_progress.xlsx         - Queue time tracking (~1,750 transitions)")
print(" 17. setup_changeovers.xlsx        - Setup time detail (~400-600)")
print(" 18. rework_tracking.xlsx          - Repair attempts (~80-120)")
print("\nTotal: ~14,000+ operational records across 4 quarters (All in RM)")
print("=" * 80)
print("\nKey Insights Embedded:")
print("  - 32-38% first-article failure rate (Q1-Q2) improving to 20-24% (Q3-Q4)")
print("  - Complex parts (score 7+): 50%+ failure rate")
print("  - 60-70% of scrap is preventable with proper validation")
print("  - Cost overruns averaging 25-35% vs estimates")
print("  - Revenue decline: -16% Q1→Q2, stabilizing Q3-Q4")
print("  - Customer churn: Lost 5 customers, gained 7 new ones")
print("  - OEE decline: 52% (Q1) to 48% (Q2 crisis), recovering to 56% (Q4)")
print("=" * 80)
