"""
Google Maps CID (Customer ID) extraction and short URL converter.

Google Maps place URLs embed a 64-bit hexadecimal feature identifier (e.g., 0x89c259...:0xd134e199e405a163).
The second hex part represents the CID in hexadecimal.
When converted to an unsigned 64-bit integer (decimal), the short link is:
https://maps.google.com/?cid=<decimal_cid>
"""
import re
from typing import Optional, Tuple
from urllib.parse import urlparse, parse_qs, unquote


# Regex to match 0x...:0x... hex pairs in URLs
HEX_CID_PATTERN = re.compile(r"0x[0-9a-fA-F]+:0x([0-9a-fA-F]+)")
# Regex to match data=!1s0x...:0x... in Google Maps place URLs
DATA_HEX_PATTERN = re.compile(r"!1s0x[0-9a-fA-F]+:0x([0-9a-fA-F]+)")
# Regex to match direct decimal CID parameter
DECIMAL_CID_PATTERN = re.compile(r"[?&](?:cid|ludocid)=(\d+)")
# Regex to match ftid parameter
FTID_PATTERN = re.compile(r"[?&]ftid=0x[0-9a-fA-F]+:0x([0-9a-fA-F]+)")


def hex_to_decimal_cid(hex_str: str) -> Optional[str]:
    """
    Convert a hexadecimal string (e.g., 'd134e199e405a163' or '0xd134e199e405a163')
    to an unsigned 64-bit decimal string.
    """
    if not hex_str:
        return None
    cleaned_hex = hex_str.strip()
    if cleaned_hex.lower().startswith("0x"):
        cleaned_hex = cleaned_hex[2:]
    try:
        decimal_val = int(cleaned_hex, 16)
        # Ensure it is positive unsigned 64-bit
        if decimal_val < 0:
            decimal_val = decimal_val & 0xFFFFFFFFFFFFFFFF
        return str(decimal_val)
    except (ValueError, TypeError):
        return None


def extract_cid_from_url(url: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Extract hex CID and decimal CID from a Google Maps URL.
    
    Returns:
        Tuple: (hex_cid, decimal_cid)
    """
    if not url:
        return None, None

    url = unquote(url)

    # 1. Check for ftid parameter (e.g. ?ftid=0x...:0x<hex>)
    ftid_match = FTID_PATTERN.search(url)
    if ftid_match:
        hex_cid = ftid_match.group(1)
        dec_cid = hex_to_decimal_cid(hex_cid)
        if dec_cid:
            return hex_cid, dec_cid

    # 2. Check for !1s0x...:0x<hex> pattern (standard data string)
    data_match = DATA_HEX_PATTERN.search(url)
    if data_match:
        hex_cid = data_match.group(1)
        dec_cid = hex_to_decimal_cid(hex_cid)
        if dec_cid:
            return hex_cid, dec_cid

    # 3. Check for general 0x...:0x<hex> pattern
    hex_match = HEX_CID_PATTERN.search(url)
    if hex_match:
        hex_cid = hex_match.group(1)
        dec_cid = hex_to_decimal_cid(hex_cid)
        if dec_cid:
            return hex_cid, dec_cid

    # 4. Check for direct decimal CID parameter (cid=... or ludocid=...)
    dec_match = DECIMAL_CID_PATTERN.search(url)
    if dec_match:
        dec_cid = dec_match.group(1)
        # Convert decimal back to hex if needed
        try:
            hex_cid = hex(int(dec_cid))[2:]
        except ValueError:
            hex_cid = None
        return hex_cid, dec_cid

    return None, None


def get_short_cid_url(url_or_cid: str) -> str:
    """
    Given a Google Maps URL or a decimal/hex CID, returns the short link:
    https://maps.google.com/?cid=<decimal_cid>
    
    If no CID can be extracted, returns the cleaned original URL.
    """
    if not url_or_cid:
        return ""

    url_str = str(url_or_cid).strip()

    # If it's already purely decimal digits and looks like a CID
    if url_str.isdigit() and len(url_str) >= 10:
        return f"https://maps.google.com/?cid={url_str}"

    # If it's a hex string (e.g. 0xd134e199e405a163)
    if url_str.startswith("0x") and ":" not in url_str:
        dec_cid = hex_to_decimal_cid(url_str)
        if dec_cid:
            return f"https://maps.google.com/?cid={dec_cid}"

    # Try extracting from URL
    _, dec_cid = extract_cid_from_url(url_str)
    if dec_cid:
        return f"https://maps.google.com/?cid={dec_cid}"

    return url_str
