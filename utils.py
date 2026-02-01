import pandas as pd

def normalize(name: str) -> str:
    return str(name).strip().replace("\u200f", "").replace("\u200e", "")

def to_numeric(s):
    return pd.to_numeric(s, errors="coerce")

def clean_number(x):
    if pd.isna(x):
        return x
    try:
        s = str(x).strip()
        trans = {
            ord('٠'): '0', ord('١'): '1', ord('٢'): '2', ord('٣'): '3',
            ord('٤'): '4', ord('٥'): '5', ord('٦'): '6', ord('٧'): '7',
            ord('٨'): '8', ord('٩'): '9',
            ord('٬'): ',', ord('،'): ',', ord('٫'): '.',
        }
        s = s.translate(trans)
        s = s.replace('%', '')
        s = s.replace(',', '')
        return s
    except Exception:
        return x

def norm_key(s: str) -> str:
    s = normalize(s)
    trans = {
        ord('آ'): 'ا', ord('أ'): 'ا', ord('إ'): 'ا',
        ord('ى'): 'ي', ord('ة'): 'ه', ord('ؤ'): 'و', ord('ئ'): 'ي',
        ord('٠'): '0', ord('١'): '1', ord('٢'): '2', ord('٣'): '3',
        ord('٤'): '4', ord('٥'): '5', ord('٦'): '6', ord('٧'): '7',
        ord('٨'): '8', ord('٩'): '9',
        ord('٬'): ',', ord('،'): ',', ord('٫'): ',',
        ord('ـ'): '',
    }
    s = s.translate(trans).lower()
    return ''.join(ch for ch in s if ch.isalnum())

def match_col(df: pd.DataFrame, aliases: list[str]) -> str | None:
    keys = {norm_key(c): c for c in df.columns}
    for a in aliases:
        k = norm_key(a)
        if k in keys:
            return keys[k]
    return None
