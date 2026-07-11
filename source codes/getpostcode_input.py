import pandas as pd
import logging
import time
import os
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from tqdm import tqdm

logging.getLogger("geopy").setLevel(logging.ERROR)

OUTPUT_FILE = "output_with_uk_postcodes.xlsx"

# Load existing output if it exists (resume)
if os.path.exists(OUTPUT_FILE):
    df = pd.read_excel(OUTPUT_FILE)
    print("Resuming from existing output file...")
else:
    df = pd.read_excel("input.xlsx")
    df["postcode"] = None
    print("Starting fresh...")

print(f"Total rows: {len(df)}")

geolocator = Nominatim(user_agent="latlon_to_uk_postcode")

reverse = RateLimiter(
    geolocator.reverse,
    min_delay_seconds=1,
    max_retries=1,
    swallow_exceptions=True
)

cache = {}

def get_uk_postcode(lat, lon):
    if (
        pd.isna(lat) or pd.isna(lon) or
        lat == 0 or lon == 0 or
        not (-90 <= lat <= 90) or
        not (-180 <= lon <= 180)
    ):
        return None

    key = (round(lat, 5), round(lon, 5))
    if key in cache:
        return cache[key]

    location = reverse(key, exactly_one=True)
    if not location:
        cache[key] = None
        return None

    address = location.raw.get("address", {})
    if address.get("country_code") != "gb":
        cache[key] = None
        return None

    postcode = address.get("postcode")
    cache[key] = postcode
    return postcode


SAVE_INTERVAL = 5  # seconds
last_save = time.time()

for i, row in enumerate(tqdm(df.itertuples(index=False), total=len(df))):

    # Skip rows already processed
    if pd.notna(df.at[i, "postcode"]):
        continue

    df.at[i, "postcode"] = get_uk_postcode(row.lat, row.lon)

    if time.time() - last_save >= SAVE_INTERVAL:
        df.to_excel(OUTPUT_FILE, index=False)
        last_save = time.time()

# Final save
df.to_excel(OUTPUT_FILE, index=False)
print("Done!")
