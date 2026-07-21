import pyodbc
import pandas as pd
import os
import shutil
import xlwings as xw
import sys
import tkinter as tk
from tkinter import messagebox
import aiohttp
import asyncio

def open_gui():

    result = None

    root = tk.Tk()

    root.title("Spotless Call Out Generator")
    root.geometry("600x400")

    poi_rows = []

    # Radius
    tk.Label(root, text="Radius (miles)").pack()
    radius_entry = tk.Entry(root)
    radius_entry.insert(0, "8")
    radius_entry.pack()


    # POI section
    tk.Label(root, text="POIs").pack()

    frame = tk.Frame(root)
    frame.pack()

    headers = ["POI Name", "Latitude", "Longitude"]

    for i, h in enumerate(headers):
        tk.Label(frame, text=h, width=15).grid(row=0, column=i)


    def add_poi_row():

        row = len(poi_rows) + 1

        name = tk.Entry(frame, width=15)
        lat = tk.Entry(frame, width=15)
        lon = tk.Entry(frame, width=15)

        name.grid(row=row, column=0)
        lat.grid(row=row, column=1)
        lon.grid(row=row, column=2)

        poi_rows.append(
            {
                "name": name,
                "lat": lat,
                "lon": lon
            }
        )


    add_poi_row()


    tk.Button(
        root,
        text="Add another POI",
        command=add_poi_row
    ).pack(pady=10)


    def submit():

        try:

            # Validate radius
            try:
                radius_miles = float(radius_entry.get())

                if radius_miles <= 0:
                    raise ValueError("Radius must be greater than zero")

            except ValueError:
                messagebox.showerror(
                    "Input error",
                    "Please enter a valid radius in miles."
                )
                return


            pois = []


            # Validate each POI row
            for row in poi_rows:

                name = row["name"].get().strip()
                lat_text = row["lat"].get().strip()
                lon_text = row["lon"].get().strip()


                if not name:
                    messagebox.showerror(
                        "Input error",
                        "Every POI must have a name."
                    )
                    return


                try:
                    lat = float(lat_text)
                    lon = float(lon_text)

                except ValueError:
                    messagebox.showerror(
                        "Input error",
                        f"Invalid coordinates for {name}."
                    )
                    return


                # Latitude validation
                if lat < -90 or lat > 90:
                    messagebox.showerror(
                        "Input error",
                        f"Latitude for {name} must be between -90 and 90."
                    )
                    return


                # Longitude validation
                if lon < -180 or lon > 180:
                    messagebox.showerror(
                        "Input error",
                        f"Longitude for {name} must be between -180 and 180."
                    )
                    return


                pois.append(
                    (
                        name,
                        lat,
                        lon
                    )
                )


            # Convert miles to metres for SQL
            radius = radius_miles * 1609.34


            print("Radius metres:", radius)
            print("POIs:", pois)


            nonlocal result

            result = radius, pois

            root.destroy()


        except Exception as e:

            messagebox.showerror(
                "Unexpected error",
                str(e)
            )

    tk.Button(
        root,
        text="Generate Call Out List",
        command=submit
    ).pack(pady=20)


    root.mainloop()
    return result


radius, pois = open_gui()

if not radius or not pois:
    print("No input supplied. Exiting.")
    sys.exit()

if getattr(sys, "frozen", False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

output_path = os.path.join(base_dir, "Output")
template_path = os.path.join(base_dir, "EVENT CALL OUT Template.xlsx")
os.makedirs(output_path, exist_ok=True)


#--SQL server read credentials--#
SERVER = ""
DATABASE = ""
USERNAME = ""
PASSWORD = ""

conn = pyodbc.connect(
    f"""
    DRIVER={{ODBC Driver 18 for SQL Server}};
    SERVER={SERVER};
    DATABASE={DATABASE};
    UID={USERNAME};
    PWD={PASSWORD};
    Encrypt=yes;
    TrustServerCertificate=no;
    """
)


header = """
SET NOCOUNT ON;
DECLARE @POIs TABLE
(
    POIName varchar(50),
    Lat float,
    Lon float
);
"""

for name, lat, lon in pois:
    header += (
        f"\nINSERT INTO @POIs VALUES "
        f"('{name}', {lat}, {lon});"
    )

header += f"""

DECLARE @MilesInMetresRadius DECIMAL(18,2) = {radius};

"""


sql = """
WITH Tx AS (
    SELECT
        d.tenantId,
        d.Reference,
        CAST(d.TransactionDate AT TIME ZONE 'UTC' AT TIME ZONE 'GMT Standard Time' AS DATE) AS TransactionDate,
        d.Litres,
        d.KioskEntityId
    FROM spotless_production.dbo.TransactionEntities d
    WHERE d.TransactionType IN (4,6,8,9)
      AND d.IsDeleted = 0
      AND d.Test = 0
),


----STATION USAGE :
StationUsage AS (
    SELECT
        x.tenantId,
        STRING_AGG(CAST(x.Name AS VARCHAR(MAX)), ';') AS [Stations Used]
    FROM (
        SELECT DISTINCT
            d2.tenantId,
            kee2.Name
        FROM spotless_production.dbo.TransactionEntities d2
        JOIN spotless_production.dbo.KioskEntities ke2
            ON ke2.Id = d2.KioskEntityId
        JOIN spotless_production.dbo.LocationEntities kee2
            ON kee2.Id = ke2.LocationEntityId
        WHERE d2.TransactionType IN (4,6,8,9)
          AND d2.IsDeleted = 0
          AND d2.Test = 0
          AND kee2.IsDeleted = 0
    ) x
    GROUP BY x.tenantId
),


----AGGREGATED FACT:
TxAgg AS (
    SELECT
        tenantId,

        COUNT(DISTINCT CASE 
            WHEN TransactionDate >= DATEADD(MONTH,-6,GETDATE())
            THEN Reference END
        ) AS no_of_dispenses,

        AVG(CASE 
            WHEN TransactionDate >= DATEADD(MONTH,-3,GETDATE())
            THEN Litres END
        ) AS avg_litres_3mths,

        AVG(CASE 
            WHEN TransactionDate >= DATEADD(MONTH,-6,GETDATE())
            THEN Litres END
        ) AS avg_litres_6mths,

        SUM(CASE
            WHEN TransactionDate >= DATEADD(MONTH,-6,GETDATE())
             --AND KioskEntityId = 183
            THEN Litres ELSE 0 END
        ) AS litres_kiosk_POI,

        SUM(CASE
            WHEN TransactionDate >= DATEADD(MONTH,-6,GETDATE())
             --AND KioskEntityId <> 183
            THEN Litres ELSE 0 END
        ) AS litres_other_kiosks,

        MAX(TransactionDate) AS last_used_date
    FROM Tx
    GROUP BY tenantId
),



----COMBINED DATASET REMEMBER TO AGG FIRST THEN JOIN PLS:
CombinedData AS (
    
    SELECT
        p.POIName,
        b.[Name] AS [Company Name],
        c.[Name] AS [First Name],
        c.[Surname] AS [Last Name],
        a.telephone AS [Phone Number],
        a.AdminEmailAddress AS [Email],

        CASE a.AccountPrimaryUse
            WHEN 0 THEN 'Window Cleaner'
            WHEN 1 THEN 'Aquarium'
            WHEN 2 THEN 'Valeter'
            WHEN 3 THEN 'Other'
        END AS [Type],

        CASE
            WHEN tx.last_used_date >= DATEADD(MONTH,-3,GETDATE()) THEN 'Active'
            WHEN tx.last_used_date IS NULL THEN 'Never Used'
            ELSE 'Inactive'
        END AS [Status],

        tx.last_used_date AS [Last Used Date],
        CAST(a.CreationTime AT TIME ZONE 'UTC' AT TIME ZONE 'GMT Standard Time' AS DATE) AS [Creation Date],
        a.lat,
        a.lon,
        a.CurrentBalance,
        a.tenantId,
        COUNT(DISTINCT e.cardnumber) AS [no. of keyfobs],
        tx.no_of_dispenses AS [no. of dispenses],
        tx.avg_litres_3mths AS [Avg_Litres (3mths)],
        tx.avg_litres_6mths AS [Avg_Litres (6mths)],
        tx.litres_kiosk_POI AS [Litres (6mths @ Kiosk POI)],
        tx.litres_other_kiosks AS [Litres (6mths @ Other Kiosks)],
        su.[Stations Used],
        
        CASE
        WHEN tx.last_used_date IS NULL 
            THEN 'Never Used'

        WHEN tx.last_used_date < DATEADD(MONTH,-3,GETDATE())
            THEN 'Inactive'

        WHEN DATEDIFF(DAY,
                      CAST(a.CreationTime AS DATE),
                      GETDATE()) < 14
            THEN 'New User'

        WHEN tx.no_of_dispenses / 3.0 >= 20
            THEN 'Daily'

        WHEN tx.no_of_dispenses / 3.0 >= 10
            THEN 'Frequent'

        WHEN tx.no_of_dispenses / 3.0 >= 2
            THEN 'Part-Time'

        ELSE 'Emergency'
    END AS CustomerCategory,

        geography::Point(a.Lat,a.Lon,4326)
        .STDistance(geography::Point(p.Lat,p.Lon,4326)) AS 'DistanceMetres'

    FROM @POIs p
    JOIN spotless_production.dbo.TenantAccountEntities a
        ON a.IsDeleted = 0
       AND a.IsTest = 0
       AND a.Discontinued = 0
       AND a.Lat IS NOT NULL
       AND a.Lat BETWEEN
            p.Lat - (@MilesInMetresRadius / 111000.0)
        AND p.Lat + (@MilesInMetresRadius / 111000.0)
       AND a.Lon BETWEEN
            p.Lon - (@MilesInMetresRadius / (111000.0 * COS(RADIANS(p.Lat))))
        AND p.Lon + (@MilesInMetresRadius / (111000.0 * COS(RADIANS(p.Lat))))
       AND geography::Point(a.Lat, a.Lon, 4326)
            .STDistance(geography::Point(p.Lat, p.Lon, 4326))
            <= @MilesInMetresRadius 
    LEFT JOIN spotless_production.dbo.AbpTenants b ON b.Id = a.tenantId
    LEFT JOIN (
        SELECT *
        FROM (
            SELECT *, ROW_NUMBER() OVER (PARTITION BY tenantId ORDER BY CreationTime DESC) AS rn
            FROM spotless_production.dbo.AbpUsers
            WHERE IsEmailConfirmed = 1 AND IsDeleted = 0
        ) t
        WHERE rn = 1
    ) c ON c.tenantId = a.tenantId
    LEFT JOIN TxAgg tx ON tx.tenantId = a.tenantId
    LEFT JOIN spotless_production.dbo.AccountCardEntities e ON e.TenantId = a.tenantId AND e.IsDeleted = 0
    LEFT JOIN StationUsage su ON su.tenantId = a.tenantId
    where b.IsDeleted = 0 AND b.IsTest = 0 AND c.IsDeleted = 0 AND c.IsEmailConfirmed = 1 AND a.tenantId <> 19068
    --AND su.[Stations Used] LIKE '%Rochester%'

    GROUP BY p.POIName,p.lat, p.lon, b.Name, c.Name, c.Surname, a.telephone, a.AdminEmailAddress, a.AccountPrimaryUse, tx.last_used_date,
             a.CreationTime, a.Lat, a.Lon, a.CurrentBalance, a.tenantId, su.[Stations Used], tx.no_of_dispenses,
             tx.avg_litres_3mths, tx.avg_litres_6mths, tx.litres_kiosk_POI, tx.litres_other_kiosks

    UNION ALL

    -----THE STATIONS UNION BACK IN
    SELECT 
        p.POIName,
        [Name] AS [Company Name],
        NULL AS [First Name],
        NULL AS [Last Name],
        NULL AS [Phone Number],
        CAST(NULL AS VARCHAR(255)) AS [Email],
        'Station' AS [Type],
        'Station' AS [Status],
        NULL AS [Last Used Date],
        CAST(CreationTime AT TIME ZONE 'UTC' AT TIME ZONE 'GMT Standard Time' AS DATE) AS [Creation Date],
        l.Lat,
        l.Lon,
        NULL AS [CurrentBalance],
        CAST(NULL AS INT) AS [TenantId],
        NULL AS [no. of keyfobs],
        NULL AS [no. of dispenses],
        NULL AS [Avg_Litres (3mths)],
        NULL AS [Avg_Litres (6mths)],
        NULL AS [Litres (6mths @ Kiosk POI)],
        NULL AS [Litres (6mths @ Other Kiosks)],
        NULL AS [Stations Used],
        NULL AS CustomerCategory,
        geography::Point(l.Lat,l.Lon,4326)
        .STDistance(geography::Point(p.Lat,p.Lon,4326)) AS 'DistanceMetres'
    FROM @POIs p
    JOIN spotless_production.dbo.LocationEntities l
        ON l.IsDeleted = 0
       AND l.Name <> 'Spotless Test Location'
       AND l.Lat IS NOT NULL
       AND l.Lat BETWEEN
            p.Lat - (@MilesInMetresRadius / 111000.0)
        AND p.Lat + (@MilesInMetresRadius / 111000.0)
       AND l.Lon BETWEEN
            p.Lon - (@MilesInMetresRadius / (111000.0 * COS(RADIANS(p.Lat))))
        AND p.Lon + (@MilesInMetresRadius / (111000.0 * COS(RADIANS(p.Lat))))
       AND geography::Point(l.Lat,l.Lon,4326)
            .STDistance(geography::Point(p.Lat,p.Lon,4326))
            <= @MilesInMetresRadius
)

-- FINAL OUTPUT :

SELECT 
    POIName,
    [Company Name],
    [First Name],
    [Last Name],
    [Phone Number], 
    [Email],
    [Type], 
    [Status],
    [Last Used Date],
    [Creation Date],
    /*DistanceMiles,*/
    DistanceMetres / 450.616 AS [Drive time from POI],
    [lat],
    [lon],
    [CurrentBalance],
    [tenantId],
    [no. of keyfobs],
    CustomerCategory,
    [no. of dispenses] AS [Total no. of dispense (6mths)],
    [Avg_Litres (3mths)],
    [Avg_Litres (6mths)],
    --[Litres (6mths @ Kiosk POI)],
    --[Litres (6mths @ Other Kiosks)],
    [Stations Used]
    
FROM CombinedData
order by [first name], [Drive time from POI] desc


"""
sql = header + sql


def format_phone_number(phone):
    if pd.isna(phone):
        return None
    phone=''.join(c for c in str(phone) if c.isdigit())
    if phone.startswith('44'):
        phone='0'+phone[2:]
    elif phone and not phone.startswith('0'):
        phone='0'+phone
    return phone[:5]+' '+phone[5:] if len(phone)>=11 else phone

def sentence_case(text):
    if pd.isna(text) or text=="NULL":
        return ""
    if not isinstance(text,str):
        return text
    if text.upper()=="NEVER USED":
        return "Never used"
    return text.capitalize()

def process_data(df, station_name):
    print(f"Starting processing: {station_name}")

    print(f"Station: {station_name}")

    #print("  ✓ Input Data loaded")
    
    df["LAST_USED_DT"]=pd.to_datetime(df["Last Used Date"],dayfirst=True,errors="coerce").fillna(pd.Timestamp("1900-01-01"))

    df_processed=pd.DataFrame({
        "COMPANY":df["Company Name"].apply(sentence_case),
        "FIRST NAME":df["First Name"].apply(sentence_case),
        "LAST NAME":df["Last Name"].apply(sentence_case),
        "EMAIL":df["Email"].apply(sentence_case),
        "PHONE":df["Phone Number"].apply(format_phone_number),
        "TYPE":df["Type"].apply(sentence_case),
        "ACTIVE?":df["Status"].apply(sentence_case),
        "LAST USED":df["Last Used Date"].fillna("Never used").apply(sentence_case),
        "CREATION DATE":df["Creation Date"],
        #"MILES FROM STATION":df["Drive time from POI"],
        'MILES FROM STATION': (
            pd.to_numeric(df['Drive time from POI'], errors='coerce') * 8 / 25 #8miles is about 25mins drivetime
        ).round(1),
        #'DRIVE TIME (MINUTES)': df["duration_minutes"],

    })
    df_processed["LAST_USED_DT"]=df["LAST_USED_DT"]

    df_stations=df_processed[df_processed["ACTIVE?"].str.upper()=="STATION"].drop(columns=["LAST_USED_DT"])
    customers=df_processed[df_processed["ACTIVE?"].str.upper()!="STATION"].copy()
    customers=customers.dropna(subset=["PHONE"])

    dupes=customers[customers.duplicated("PHONE",keep=False)]
    deduped=customers.loc[customers.groupby("PHONE")["LAST_USED_DT"].idxmax()].drop(columns=["LAST_USED_DT"]).sort_values("MILES FROM STATION")

    if not dupes.empty:
        dropped=customers.loc[dupes.index]
        dropped=dropped.loc[~dropped.index.isin(deduped.index)].drop(columns=["LAST_USED_DT"])
    else:
        dropped=pd.DataFrame(columns=deduped.columns)

    #dupes_combined=pd.concat([dropped,df_stations],ignore_index=True)
    dupes_combined=dropped.copy()

    out=os.path.join(output_path,f"{station_name.upper()} CALL OUT LIST.xlsx")
    shutil.copy(template_path, out)
    app=xw.App(visible=False)
    try:
        wb=app.books.open(out)
        s=wb.sheets["Call Out List"]
        s["A1"].value=f"{station_name.upper()} CALL OUT LIST"
        s["A2"].value=list(deduped.columns)
        s["A3"].value=deduped.fillna("").values

        d=wb.sheets["Dupes"] if "Dupes" in [x.name for x in wb.sheets] else wb.sheets.add("Dupes")
        d["A1"].value=f"{station_name.upper()} DUPLICATES"
        if not dupes_combined.empty:
            d["A2"].value=list(dupes_combined.columns)
            d["A3"].value=dupes_combined.fillna("").values
        wb.save()
        wb.close()
        print("  ✓ Output workbook created")
    finally:
        app.quit()

    print(f"Finished {station_name}")


#--added duration query--#
OSRM_BASE = "http://router.project-osrm.org"
#semaphore = asyncio.Semaphore(5)


async def fetch_duration(session, semaphore, origin_lat, origin_lon, dest_lat, dest_lon, retries=5):

    url = (
        f"{OSRM_BASE}/route/v1/driving/"
        f"{origin_lon},{origin_lat};{dest_lon},{dest_lat}"
        "?overview=false"
    )

    for attempt in range(1, retries + 1):

        async with semaphore:

            try:
                async with session.get(url, timeout=20) as r:

                    if r.status == 429:
                        raise Exception("Rate limited")

                    data = await r.json()

                    if data.get("routes"):
                        return data["routes"][0]["duration"]

            except Exception as e:
                print(f"Attempt {attempt} failed: {e}")

        await asyncio.sleep(5 * attempt)

    return None



async def calculate_durations(df, poi_lat, poi_lon):
    semaphore = asyncio.Semaphore(5)

    async with aiohttp.ClientSession() as session:
        
        tasks = [
            fetch_duration(
                session,
                semaphore,
                poi_lat,
                poi_lon,
                row.lat,
                row.lon
            )
            for row in df.itertuples()
        ]

        results = await asyncio.gather(*tasks)


    df["duration_minutes"] = [
        round(x / 60, 1) if x else None
        for x in results
    ]
    # Add buffer time
    df["duration_minutes"] = df["duration_minutes"].apply(
        lambda x: round(x + 7, 1) if x >= 5 else round(x + 3, 1)
        if pd.notna(x) else None
    )
    return df
#------------------------#

try :
    print("Running SQL...")

    df = pd.read_sql(sql, conn)

    print(f"{len(df)} rows returned")

    print(df["POIName"].unique())
    print(pois)


    for poi_name, poi_lat, poi_lon in pois:

        station_name = poi_name

        station_df = df[df["POIName"] == station_name].copy()
        
        if station_df.empty:
            print(f"No customers found for {station_name}")
            continue

        # Get this POI's coordinates
        poi_info = next(
            p for p in pois 
            if p[0] == station_name
        )

        poi_lat = poi_info[1]
        poi_lon = poi_info[2]


        # Remove stations
        customer_routes = station_df[
            station_df["Status"] != "Station"
        ][
            ["Email", "lat", "lon"]
        ].copy()


        if not customer_routes.empty:

            customer_routes = asyncio.run(
                calculate_durations(
                    customer_routes,
                    poi_lat,
                    poi_lon
                )
            )

            station_df = station_df.merge(
                customer_routes[["Email", "duration_minutes"]],
                on="Email",
                how="left"
            )

            duration_output = os.path.join(
                output_path,
                f"{station_name} Customer Driving Duration.xlsx"
            )


            customer_routes.to_excel(
                duration_output,
                index=False
            )


            print(
                f"✓ {station_name} duration file created"
            )


        process_data(station_df, station_name)

    print("Processing complete.")
finally :
    conn.close()