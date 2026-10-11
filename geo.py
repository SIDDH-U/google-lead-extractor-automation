"""
Geocoding and grid generation utilities using OpenStreetMap Nominatim.
"""
import math
import time
from typing import Dict, List, Tuple, Any, Optional
import requests
from config import NOMINATIM_BASE_URL, NOMINATIM_USER_AGENT, NOMINATIM_TIMEOUT_SEC, DEFAULT_GRID_STEP_KM


class GeoError(Exception):
    """Exception raised for errors during geocoding or grid generation."""
    pass


def get_city_bounding_box(city_name: str) -> Dict[str, Any]:
    """
    Fetch the bounding box and coordinates for a city using OpenStreetMap Nominatim.
    
    Nominatim returns boundingbox as [south_lat, north_lat, west_lon, east_lon].
    
    Returns:
        dict with keys:
            city: query city name
            display_name: full OSM formatted address/name
            lat_min: float (South latitude)
            lat_max: float (North latitude)
            lon_min: float (West longitude)
            lon_max: float (East longitude)
            center_lat: float
            center_lon: float
            osm_type: str
    """
    if not city_name or not city_name.strip():
        raise GeoError("City name cannot be empty.")

    params = {
        "q": city_name.strip(),
        "format": "json",
        "polygon_geojson": 0,
        "addressdetails": 1,
        "limit": 5,
    }
    headers = {
        "User-Agent": NOMINATIM_USER_AGENT,
        "Accept": "application/json",
    }

    try:
        response = requests.get(
            NOMINATIM_BASE_URL,
            params=params,
            headers=headers,
            timeout=NOMINATIM_TIMEOUT_SEC,
        )
        response.raise_for_status()
        data = response.json()
    except requests.RequestException as e:
        raise GeoError(f"Failed to query OpenStreetMap Nominatim: {e}")

    if not data or len(data) == 0:
        raise GeoError(
            f"Could not find coordinates for city '{city_name}'. Please verify the spelling (e.g., 'Austin, TX' or 'Paris, France')."
        )

    # Prioritize city / town / municipality / county over huge state_districts
    selected_result = data[0]
    for item in data:
        addr_type = item.get("addresstype", "").lower()
        osm_type = item.get("type", "").lower()
        bbox = item.get("boundingbox", [])
        if len(bbox) == 4:
            lat_span = abs(float(bbox[1]) - float(bbox[0]))
            lon_span = abs(float(bbox[3]) - float(bbox[2]))
            # If it's explicitly a city/town/municipality or has a compact urban span (< 0.5 degrees ~ 50 km)
            if addr_type in ("city", "town", "municipality", "borough", "suburb", "county") or (lat_span < 0.6 and lon_span < 0.6):
                selected_result = item
                break

    bbox = selected_result.get("boundingbox")
    if not bbox or len(bbox) < 4:
        raise GeoError(f"Nominatim did not return a valid bounding box for '{city_name}'.")

    # Nominatim bbox is [south, north, west, east]
    lat_south = float(bbox[0])
    lat_north = float(bbox[1])
    lon_west = float(bbox[2])
    lon_east = float(bbox[3])
    center_lat = float(selected_result.get("lat", (lat_south + lat_north) / 2))
    center_lon = float(selected_result.get("lon", (lon_west + lon_east) / 2))

    return {
        "city": city_name.strip(),
        "display_name": selected_result.get("display_name", city_name),
        "lat_min": min(lat_south, lat_north),
        "lat_max": max(lat_south, lat_north),
        "lon_min": min(lon_west, lon_east),
        "lon_max": max(lon_west, lon_east),
        "center_lat": center_lat,
        "center_lon": center_lon,
        "osm_type": selected_result.get("type", "city"),
    }


def generate_grid_points(
    bbox: Dict[str, Any],
    step_km: float = DEFAULT_GRID_STEP_KM
) -> List[Tuple[float, float, int, int]]:
    """
    Split a bounding box into a grid of search points approximately `step_km` apart.
    
    1 degree latitude ~ 110.574 km
    1 degree longitude ~ 111.320 * cos(latitude) km
    
    Returns:
        List of tuples: (lat, lon, grid_x, grid_y)
    """
    lat_min = bbox["lat_min"]
    lat_max = bbox["lat_max"]
    lon_min = bbox["lon_min"]
    lon_max = bbox["lon_max"]
    center_lat = bbox.get("center_lat", (lat_min + lat_max) / 2.0)

    # Convert km step to latitude degrees
    km_per_lat_deg = 110.574
    delta_lat = step_km / km_per_lat_deg

    # Convert km step to longitude degrees at center latitude
    lat_rad = math.radians(center_lat)
    km_per_lon_deg = 111.320 * max(0.01, math.cos(lat_rad))
    delta_lon = step_km / km_per_lon_deg

    grid_points: List[Tuple[float, float, int, int]] = []

    # Calculate number of steps
    lat_span = lat_max - lat_min
    lon_span = lon_max - lon_min

    # Ensure at least 1 point even if bounding box is very small
    num_lat_steps = max(1, math.ceil(lat_span / delta_lat))
    num_lon_steps = max(1, math.ceil(lon_span / delta_lon))

    # Center the grid offsets inside the bounding box
    actual_delta_lat = lat_span / num_lat_steps if num_lat_steps > 1 else 0
    actual_delta_lon = lon_span / num_lon_steps if num_lon_steps > 1 else 0

    if num_lat_steps == 1 and num_lon_steps == 1:
        # Single point at center
        grid_points.append((round(center_lat, 6), round(bbox.get("center_lon", (lon_min + lon_max) / 2.0), 6), 0, 0))
        return grid_points

    for y in range(num_lat_steps + 1):
        cur_lat = lat_min + (y * delta_lat)
        if cur_lat > lat_max + (delta_lat * 0.1):
            break
        # Clamp within reasonable bounds
        cur_lat = min(cur_lat, lat_max)

        for x in range(num_lon_steps + 1):
            cur_lon = lon_min + (x * delta_lon)
            if cur_lon > lon_max + (delta_lon * 0.1):
                break
            cur_lon = min(cur_lon, lon_max)

            grid_points.append((round(cur_lat, 6), round(cur_lon, 6), x, y))

    return grid_points


def calculate_haversine_distance_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """
    Calculate the great circle distance between two points on the earth in km.
    """
    R = 6371.0  # Earth radius in kilometers
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (
        math.sin(dlat / 2.0) ** 2
        + math.cos(math.radians(lat1))
        * math.cos(math.radians(lat2))
        * math.sin(dlon / 2.0) ** 2
    )
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c
