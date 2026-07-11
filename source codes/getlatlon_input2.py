import pandas as pd
import os
import pickle
import time
import random
from tqdm import tqdm
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
#from bs4 import BeautifulSoup no more scraping from html
import requests
import re
from rapidfuzz import process

# ---------------------- Config ----------------------
INPUT_FILE = "input2.xlsx"
OUTPUT_FILE = "output_with_lat_lon.xlsx"
CACHE_FILE = "postcode_cache.pkl"
BATCH_SIZE = 100  # for Postcodes.io batch size
# ----------------------------------------------------

# ---------------------- Load Data ----------------------
if os.path.exists(OUTPUT_FILE):
    df = pd.read_excel(OUTPUT_FILE)
    print("Resuming from existing output file...")
else:
    df = pd.read_excel(INPUT_FILE)
    df["lat"] = None
    df["lon"] = None
    print("Starting fresh...")

# Load or initialize cache
if os.path.exists(CACHE_FILE):
    with open(CACHE_FILE, "rb") as f:
        cache = pickle.load(f)
    print(f"Loaded cache with {len(cache)} entries")
else:
    cache = {}

# ---------------------- Normalize Postcodes ----------------------
def normalize_uk_postcode(pc):
    pc = str(pc).upper().strip().replace(" ", "")
    match = re.match(r'^([A-Z]{1,2}\d[A-Z\d]?)(\d[A-Z]{2})$', pc)
    if match:
        return f"{match.group(1)} {match.group(2)}"
    return pc

df['postcode'] = df['postcode'].astype(str).apply(normalize_uk_postcode)

# ---------------------- Postcodes.io lookup ----------------------
def fetch_postcodes_io(postcodes):
    url = "https://api.postcodes.io/postcodes"
    results = {}
    for i in range(0, len(postcodes), BATCH_SIZE):
        batch = postcodes[i:i+BATCH_SIZE]
        try:
            response = requests.post(url, json={"postcodes": batch}, timeout=10)
            data = response.json()
            for item in data['result']:
                pc = item['query']
                if item['result']:
                    lat = item['result']['latitude']
                    lon = item['result']['longitude']
                else:
                    lat = lon = None
                results[pc] = (lat, lon)
        except Exception as e:
            print(f"Error fetching batch {batch}: {e}")
            for pc in batch:
                results[pc] = (None, None)
    return results

to_lookup = [pc for pc in df['postcode'].unique() if pc not in cache and pc]
print(f"Looking up {len(to_lookup)} postcodes via Postcodes.io ...")
batch_results = fetch_postcodes_io(to_lookup)
cache.update(batch_results)


# ---------------------- findthatpostcode.uk JSON lookup ----------------------
def fetch_postcode_from_site(pc, max_retries=3):
    """
    Fetch latitude and longitude for a UK postcode from findthatpostcode.uk JSON API.
    Returns (lat, lon) as floats, or (None, None) if not found.
    """
    pc_url = pc.replace(" ", "+")
    url = f"https://findthatpostcode.uk/postcodes/{pc_url}.json"
    headers = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}

    for attempt in range(1, max_retries + 1):
        try:
            r = requests.get(url, headers=headers, timeout=10)
            print(f"{pc} → status {r.status_code}")

            if r.status_code != 200:
                time.sleep(0.3 + random.random() * 0.5)
                continue

            data = r.json()
            lat = lon = None

            # 1️⃣ Try attributes first
            attrs = data.get("data", {}).get("attributes", {})
            if isinstance(attrs, dict):
                lat = attrs.get("lat") or attrs.get("latitude")
                lon = attrs.get("long") or attrs.get("lon") or attrs.get("longitude")

            # 2️⃣ Fallback to top-level location
            if (lat is None or lon is None) and "location" in data:
                loc = data["location"]
                lat = loc.get("lat")
                lon = loc.get("lon")

            # Return coordinates if found
            if lat is not None and lon is not None:
                return float(lat), float(lon)

        except Exception as e:
            print(f"Attempt {attempt} error for {pc}: {e}")

        # Exponential backoff with jitter
        time.sleep((2 ** attempt) + random.random())

    print(f"findthatpostcode JSON ultimately failed for {pc}")
    return None, None

# ---------------------- findthatpostcode.uk JSON fallback ----------------------
missing = [pc for pc, val in cache.items() if val == (None, None)]
print(f"Trying findthatpostcode JSON for {len(missing)} postcodes...")

for pc in tqdm(missing):
    if not pc:
        continue

    lat, lon = fetch_postcode_from_site(pc, max_retries=3)
    cache[pc] = (lat, lon)

    time.sleep(0.3 + random.random() * 0.5)


# ---------------------- Nominatim fallback with fuzzy retired handling ----------------------
geolocator = Nominatim(user_agent="uk_postcode_geocoder")
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1, swallow_exceptions=True)

# Prepare list of valid postcodes for fuzzy matching
valid_postcodes = [pc for pc, val in cache.items() if val != (None, None)]

def nominatim_lookup(pc):
    # Try multiple query formats
    queries = [
        {"postalcode": pc, "country": "United Kingdom"},
        f"{pc}, United Kingdom",
        f"{pc}, UK",
        pc,
        pc.split(" ")[0]  # outward code centroid
    ]
    for q in queries:
        try:
            location = geocode(q, exactly_one=True)
            if location:
                return location.latitude, location.longitude
        except:
            continue

    # Fuzzy match to nearest valid postcode if still missing
    match, score, _ = process.extractOne(pc, valid_postcodes)
    if score > 80:
        return cache[match]
    return None, None

missing = [pc for pc, val in cache.items() if val == (None, None)]
print(f"Filling {len(missing)} missing postcodes with Nominatim/fuzzy fallback ...")

for pc in tqdm(missing):
    if not pc:
        continue
    lat, lon = nominatim_lookup(pc)
    cache[pc] = (lat, lon)
    time.sleep(1 + random.random())

# ---------------------- Fill dataframe ----------------------
for i, row in df.iterrows():
    pc = row['postcode']
    if pc in cache:
        lat, lon = cache[pc]
        df.at[i, 'lat'] = lat
        df.at[i, 'lon'] = lon

# Save cache and final output
with open(CACHE_FILE, "wb") as f:
    pickle.dump(cache, f)

df.to_excel(OUTPUT_FILE, index=False)
print(f"Done! Saved {OUTPUT_FILE}")