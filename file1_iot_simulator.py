"""
EnerCloud - File 1: IoT Simulator
Simulates smart meter readings for 5 cities and sends them to Firebase every 5 seconds.

SETUP:
    pip install firebase-admin
    Download your Firebase service account key JSON from:
    Firebase Console → Project Settings → Service Accounts → Generate new private key
    Save it as 'serviceAccountKey.json' in the same folder as this script.

USAGE:
    python file1_iot_simulator.py
"""

import firebase_admin
from firebase_admin import credentials, db
import time
import random
import datetime

# ─── STEP 1: Replace with your Firebase Realtime Database URL ───────────────
DATABASE_URL = "https://enercloud-ed207-default-rtdb.firebaseio.com/"
# ────────────────────────────────────────────────────────────────────────────

# Initialize Firebase
cred = credentials.Certificate("serviceAccountKey.json")
firebase_admin.initialize_app(cred, {"databaseURL": DATABASE_URL})

# 5 cities with base energy profiles
CITIES = {
    "Mumbai": {
        "base_demand": 850,
        "solar_capacity": 320,
        "battery_capacity": 200,
        "grid_limit": 600,
    },
    "Delhi": {
        "base_demand": 920,
        "solar_capacity": 380,
        "battery_capacity": 250,
        "grid_limit": 700,
    },
    "Chennai": {
        "base_demand": 680,
        "solar_capacity": 410,
        "battery_capacity": 180,
        "grid_limit": 500,
    },
    "Bangalore": {
        "base_demand": 760,
        "solar_capacity": 290,
        "battery_capacity": 220,
        "grid_limit": 550,
    },
    "Hyderabad": {
        "base_demand": 700,
        "solar_capacity": 350,
        "battery_capacity": 190,
        "grid_limit": 520,
    },
}


def get_solar_output(city, hour):
    """Simulate solar output based on time of day (peak at noon)."""
    capacity = CITIES[city]["solar_capacity"]
    if 6 <= hour <= 18:
        # Bell curve centered at noon
        peak_factor = 1 - abs(hour - 12) / 6
        solar = capacity * peak_factor * random.uniform(0.85, 1.0)
    else:
        solar = 0.0
    return round(solar, 2)


def get_demand(city, hour):
    """Simulate energy demand with morning/evening peaks."""
    base = CITIES[city]["base_demand"]
    # Morning peak at 8am, evening peak at 8pm
    if 7 <= hour <= 10:
        factor = random.uniform(1.1, 1.3)
    elif 18 <= hour <= 22:
        factor = random.uniform(1.15, 1.35)
    elif 0 <= hour <= 5:
        factor = random.uniform(0.6, 0.75)
    else:
        factor = random.uniform(0.85, 1.05)
    return round(base * factor, 2)


def simulate_reading(city):
    """Generate one smart meter reading for a city."""
    now = datetime.datetime.now()
    hour = now.hour

    demand = get_demand(city, hour)
    solar = get_solar_output(city, hour)
    battery_soc = round(random.uniform(20, 95), 1)  # State of charge %
    voltage = round(random.uniform(218, 242), 1)     # Grid voltage (V)
    frequency = round(random.uniform(49.8, 50.2), 2) # Grid frequency (Hz)
    temperature = round(random.uniform(24, 42), 1)   # Ambient temp (°C)

    return {
        "city": city,
        "timestamp": now.isoformat(),
        "unix_time": int(time.time()),
        "demand_kw": demand,
        "solar_kw": solar,
        "battery_soc_pct": battery_soc,
        "grid_voltage_v": voltage,
        "grid_frequency_hz": frequency,
        "temperature_c": temperature,
        "status": "active",
    }


def push_to_firebase(city, reading):
    """Push reading to Firebase under /cities/{city}/latest and /cities/{city}/history."""
    ref = db.reference(f"cities/{city}")

    # Overwrite latest reading (for real-time dashboard)
    ref.child("latest").set(reading)

    # Append to history (for report generation)
    ref.child("history").push(reading)

    print(
        f"  [{reading['timestamp'][11:19]}] {city:12s} | "
        f"Demand: {reading['demand_kw']:6.1f} kW | "
        f"Solar: {reading['solar_kw']:5.1f} kW | "
        f"Battery: {reading['battery_soc_pct']:4.1f}% | "
        f"Temp: {reading['temperature_c']}°C"
    )


def main():
    print("=" * 70)
    print("  EnerCloud IoT Simulator — Sending data to Firebase every 5 seconds")
    print("=" * 70)
    print(f"  Database: {DATABASE_URL}")
    print(f"  Cities  : {', '.join(CITIES.keys())}")
    print("  Press Ctrl+C to stop.\n")

    cycle = 1
    while True:
        print(f"─── Cycle {cycle} ─────────────────────────────────────────────────────")
        for city in CITIES:
            reading = simulate_reading(city)
            push_to_firebase(city, reading)

        print(f"  ✓ All 5 cities pushed. Waiting 5 seconds...\n")
        cycle += 1
        time.sleep(5)


if __name__ == "__main__":
    main()
