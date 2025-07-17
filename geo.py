import requests
import time

def geocode_address(address):
    try:
        url = "https://nominatim.openstreetmap.org/search"
        params = {'q': address, 'format': 'json'}
        r = requests.get(url, params=params, headers={'User-Agent': 'macquarie-app'})
        data = r.json()
        if data:
            return data[0]['lat'], data[0]['lon']
    except Exception as e:
        print("Geolocation failed:", e)
    return None, None

def enrich_geolocation(df):
    unique_addresses = df['address'].dropna().unique()[:5]
    coords = {}
    for addr in unique_addresses:
        coords[addr] = geocode_address(addr)
        time.sleep(1) 
    df['lat'] = df['address'].map(lambda x: coords.get(x, (None, None))[0])
    df['lon'] = df['address'].map(lambda x: coords.get(x, (None, None))[1])
    return df
