import pandas as pd
import aiohttp
import asyncio
import os

input_file_path = r"C:\Users\Yee Chen\OneDrive - Spotless Water\powerbirefresh\latlonpairs.xlsx"
output_file_path = r"C:\Users\Yee Chen\OneDrive - Spotless Water\powerbirefresh\latlonpairsWDuration.xlsx"

OSRM_BASE = "http://router.project-osrm.org"
semaphore = asyncio.Semaphore(5)  # lower concurrency for stability

if os.path.exists(output_file_path):
    df = pd.read_excel(output_file_path)
    if "duration_seconds" not in df.columns:
        df["duration_seconds"] = None
    if "duration_minutes" not in df.columns:
        df["duration_minutes"] = None
    print("Resuming from existing output file...")
else:
    df = pd.read_excel(input_file_path)
    df["duration_seconds"] = None
    df["duration_minutes"] = None
    df.to_excel(output_file_path, index=False)
    print("Created new output file...")

async def fetch_duration(session, o_lat, o_lon, d_lat, d_lon, retries=5):
    url = f"{OSRM_BASE}/route/v1/driving/{o_lon},{o_lat};{d_lon},{d_lat}?overview=false"
    for attempt in range(1, retries + 1):
        async with semaphore:
            try:
                async with session.get(url, timeout=20) as r:
                    if r.status == 429:
                        raise Exception("Rate limited")
                    j = await r.json()
                    if j.get("routes") and len(j["routes"]) > 0:
                        duration = j["routes"][0]["duration"]
                        print(f"URL: {url} → Duration: {duration}")
                        return duration
                    else:
                        print(f"No routes found for {url}")
            except Exception as e:
                print(f"Attempt {attempt} failed for {url}: {e}")
        await asyncio.sleep(5 * attempt)
    return None

async def main(batch_size=100, save_interval=15):
    async with aiohttp.ClientSession() as session:
        # Only process rows where duration_seconds is empty
        unprocessed = df[df["duration_seconds"].isna()]
        total_rows = len(df)
        remaining_rows = len(unprocessed)
        print(f"Total rows: {total_rows}, Remaining: {remaining_rows}")

        for start in range(0, remaining_rows, batch_size):
            batch = unprocessed.iloc[start:start + batch_size]
            tasks = [
                fetch_duration(
                    session,
                    row.origin_lat,
                    row.origin_lon,
                    row.dest_lat,
                    row.dest_lon
                )
                for row in batch.itertuples()
            ]
            results = await asyncio.gather(*tasks)
            for idx, duration in zip(batch.index, results):
                df.at[idx, "duration_seconds"] = duration
                df.at[idx, "duration_minutes"] = duration / 60 if duration else None

            df.to_excel(output_file_path, index=False)
            print(f"Saved progress up to row {batch.index[-1]}")
            await asyncio.sleep(save_interval)  # cooldown to avoid throttling

if __name__ == "__main__":
    asyncio.run(main())
    print(f"All done! Data saved to {output_file_path}")
