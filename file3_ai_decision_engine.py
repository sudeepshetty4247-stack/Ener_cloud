"""
EnerCloud - File 3: AI Decision Engine
Reads live city data from Firebase, runs Solar/Battery/Grid logic,
and writes decisions back to the cloud every 6 seconds.

SETUP:
    pip install firebase-admin
    Same serviceAccountKey.json and DATABASE_URL as File 1.

USAGE:
    Run this in a SECOND terminal while file1_iot_simulator.py is running.
    python file3_ai_decision_engine.py
"""

import firebase_admin
from firebase_admin import credentials, db
import time
import datetime

# ─── STEP 1: Replace with your Firebase Realtime Database URL ───────────────
DATABASE_URL = "https://enercloud-ed207-default-rtdb.firebaseio.com/"
# ────────────────────────────────────────────────────────────────────────────

# Thresholds for decision logic
SOLAR_PRIORITY_THRESHOLD = 50      # kW — use solar if above this
BATTERY_HIGH_SOC = 70              # % — battery well-charged, use it
BATTERY_LOW_SOC = 25               # % — battery low, avoid discharge
GRID_SURGE_DEMAND = 1000           # kW — demand is very high, alert

CITY_NAMES = ["Mumbai", "Delhi", "Chennai", "Bangalore", "Hyderabad"]

# Initialize Firebase (only once)
if not firebase_admin._apps:
    cred = credentials.Certificate("serviceAccountKey.json")
    firebase_admin.initialize_app(cred, {"databaseURL": DATABASE_URL})


def decide_energy_source(reading):
    """
    Core AI decision logic.
    Determines the best energy source: Solar, Battery, Grid, or Hybrid.

    Rules (priority order):
    1. If solar >= demand → use Solar (excess charges battery)
    2. If solar > threshold AND battery >= high SOC → Solar + Battery hybrid
    3. If solar > threshold but battery low → Solar + Grid hybrid
    4. If solar below threshold but battery high → Battery only
    5. If night / solar zero AND battery charged → Battery
    6. Otherwise → Grid (last resort)
    """
    demand = reading.get("demand_kw", 0)
    solar = reading.get("solar_kw", 0)
    battery_soc = reading.get("battery_soc_pct", 0)
    temp = reading.get("temperature_c", 30)

    # Calculate solar coverage percentage
    solar_coverage = (solar / demand * 100) if demand > 0 else 0

    # Decision logic
    if solar >= demand:
        source = "Solar"
        reason = f"Solar fully covers demand ({solar:.0f} ≥ {demand:.0f} kW). Excess charges battery."
        efficiency = 98
        co2_saved = round(demand * 0.82, 1)  # kg CO2 saved vs coal

    elif solar >= SOLAR_PRIORITY_THRESHOLD and battery_soc >= BATTERY_HIGH_SOC:
        source = "Solar + Battery"
        reason = f"Solar partial ({solar:.0f} kW) + Battery at {battery_soc:.0f}% covers deficit."
        efficiency = 94
        co2_saved = round((solar + (demand - solar) * 0.5) * 0.82, 1)

    elif solar >= SOLAR_PRIORITY_THRESHOLD and battery_soc < BATTERY_LOW_SOC:
        source = "Solar + Grid"
        reason = f"Solar partial ({solar:.0f} kW), battery low ({battery_soc:.0f}%). Grid fills deficit."
        efficiency = 87
        co2_saved = round(solar * 0.82, 1)

    elif solar < SOLAR_PRIORITY_THRESHOLD and battery_soc >= BATTERY_HIGH_SOC:
        source = "Battery"
        reason = f"Low solar ({solar:.0f} kW), battery well-charged ({battery_soc:.0f}%). Using battery."
        efficiency = 91
        co2_saved = round(demand * 0.6 * 0.82, 1)

    elif solar == 0 and battery_soc >= BATTERY_HIGH_SOC:
        source = "Battery"
        reason = f"Nighttime / no solar. Battery at {battery_soc:.0f}%. Running on stored energy."
        efficiency = 91
        co2_saved = round(demand * 0.55 * 0.82, 1)

    else:
        source = "Grid"
        reason = f"Low solar ({solar:.0f} kW) + low battery ({battery_soc:.0f}%). Falling back to grid."
        efficiency = 80
        co2_saved = 0

    # Heat warning
    alert = None
    if temp > 40:
        alert = f"⚠️ High temperature ({temp}°C) — risk of transformer overload"
    if demand > GRID_SURGE_DEMAND:
        alert = f"🚨 Demand surge ({demand:.0f} kW) — consider load shedding"

    return {
        "recommended_source": source,
        "reason": reason,
        "solar_coverage_pct": round(solar_coverage, 1),
        "system_efficiency_pct": efficiency,
        "co2_saved_kg": co2_saved,
        "alert": alert,
        "decision_time": datetime.datetime.now().isoformat(),
    }


def process_city(city):
    """Read latest data for one city, run AI, write decision back."""
    ref = db.reference(f"cities/{city}")

    # Read latest reading
    latest = ref.child("latest").get()
    if not latest:
        print(f"  [{city}] No data yet — waiting for IoT simulator...")
        return

    # Run decision engine
    decision = decide_energy_source(latest)

    # Write decision back to Firebase
    ref.child("decision").set(decision)

    # Also log to decision history
    ref.child("decision_history").push(decision)

    # Print summary
    alert_str = f" | ⚠️ {decision['alert']}" if decision["alert"] else ""
    print(
        f"  {city:12s} → {decision['recommended_source']:20s} | "
        f"Solar: {decision['solar_coverage_pct']:5.1f}% | "
        f"CO₂ saved: {decision['co2_saved_kg']:6.1f} kg{alert_str}"
    )


def main():
    print("=" * 70)
    print("  EnerCloud AI Decision Engine — Reading Firebase, writing decisions")
    print("=" * 70)
    print(f"  Database: {DATABASE_URL}")
    print("  Press Ctrl+C to stop.\n")

    cycle = 1
    while True:
        print(f"─── Decision Cycle {cycle} ──────────────────────────────────────────────")
        for city in CITY_NAMES:
            try:
                process_city(city)
            except Exception as e:
                print(f"  [{city}] Error: {e}")

        print(f"  ✓ Decisions written. Waiting 6 seconds...\n")
        cycle += 1
        time.sleep(6)


if __name__ == "__main__":
    main()
