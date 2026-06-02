"""
NJL Catalogue — Production Python Backend Server (FastAPI Optimized)
===================================================================
Optimized for deployment on cloud platforms (Render/Railway) and local execution.
Loads the Parquet database into memory exactly once on startup.
"""

import os
import math
import socket
import numpy as np
import pandas as pd
import uvicorn
from fastapi import FastAPI, Query
from fastapi.responses import HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

# ─── Configuration ────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARQUET_FILE = os.path.join(BASE_DIR, "master_serial_enriched.parquet")
HTML_FILE    = os.path.join(BASE_DIR, "index.html")

# Bind to Render's dynamically assigned PORT, fallback to 8000 locally
PORT = int(os.environ.get("PORT", 8000))
HOST = "0.0.0.0"
MAX_FILTER_RESULTS = 1000

# ─── In-memory state (populated once at startup) ──────────────────────────────
df_global: pd.DataFrame = None
TOTAL_COUNT: int        = 0
SERIAL_CACHE:      list    = []  
HUID_CACHE:        list    = []  
WAREHOUSE_CACHE:   list    = []  
DUP_HUID_CACHE:    list    = []  
ROW_TEXT_CACHE:    object  = None 

# Dropdown filter caches
CATEGORY_CACHE:        list = []
SUBCATEGORY_CACHE:     list = []
PRODUCT_GROUP_CACHE:   list = []
SUB_PROD_GRP_CACHE:    list = []
SKU_STATUS_CACHE:      list = []
VENDOR_CACHE:          list = []

SEARCH_COLS = [
    "SERIALNUMBER", "ITEMID", "VENDACCOUNT", "HUID", "WAREHOUSE", "WAREHOUSE_NAME",
    "Category", "Subcategory", "Product_Group", "Sub_Product_Group", "PWC_SKUSTATUS",
]

SORT_MAP = {
    "serial_asc":   ("SERIALNUMBER",       True),
    "serial_desc":  ("SERIALNUMBER",       False),
    "weight_asc":   ("NETWEIGHT",          True),
    "weight_desc":  ("NETWEIGHT",          False),
    "category":     ("Category",           True),
    "product_group":("Product_Group",      True),
    "avail_desc":   ("AVAILABLE_PHYSICAL", False),
    "avail_asc":    ("AVAILABLE_PHYSICAL", True),
}

# ─── Startup Data Loader ──────────────────────────────────────────────────────
def load_data():
    global df_global, TOTAL_COUNT, SERIAL_CACHE, HUID_CACHE, WAREHOUSE_CACHE, ROW_TEXT_CACHE
    global CATEGORY_CACHE, SUBCATEGORY_CACHE, PRODUCT_GROUP_CACHE, SUB_PROD_GRP_CACHE, SKU_STATUS_CACHE, VENDOR_CACHE, DUP_HUID_CACHE

    if not os.path.exists(PARQUET_FILE):
        raise FileNotFoundError(f"Cannot find '{PARQUET_FILE}'. Check your file pathways.")

    print(f"[startup] Loading {PARQUET_FILE} ...", flush=True)
    df = pd.read_parquet(PARQUET_FILE)

    # Normalize dtypes
    for col in ("GROSSQTY", "NETWEIGHT"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    str_cols = [c for c in df.columns if c not in ("GROSSQTY", "NETWEIGHT")]
    df[str_cols] = df[str_cols].fillna("").astype(str)
    df.replace([np.inf, -np.inf], 0.0, inplace=True)

    df_global = df
    TOTAL_COUNT = len(df)
    print(f"[startup] Loaded {TOTAL_COUNT:,} rows.", flush=True)

    # --- CPU-OPTIMIZED SEARCH INDEX BUILDER ---
    print("[startup] Pre-building search index ...", flush=True)
    search_cols_present = [c for c in SEARCH_COLS if c in df.columns]
    
    if search_cols_present:
        text_series = df[search_cols_present[0]].astype(str).str.lower()
        for col in search_cols_present[1:]:
            text_series = text_series + " " + df[col].astype(str).str.lower()
        ROW_TEXT_CACHE = text_series
    else:
        ROW_TEXT_CACHE = pd.Series([""] * len(df), index=df.index)
        
    print(f"[startup] Search index ready ({len(ROW_TEXT_CACHE):,} rows).", flush=True)
    # ------------------------------------------

    print("[startup] Building filter caches ...", flush=True)
    s_counts = df["SERIALNUMBER"].value_counts().sort_index()
    SERIAL_CACHE[:] = [(k, int(v)) for k, v in s_counts.items() if k]

    huid_map = {}
    for raw in df["HUID"]:
        for part in str(raw).split(","):
            part = part.strip()
            if part:
                huid_map[part] = huid_map.get(part, 0) + 1
    HUID_CACHE[:] = sorted(huid_map.items(), key=lambda x: x[0])

    w_counts = df["WAREHOUSE"].value_counts().sort_index()
    WAREHOUSE_CACHE[:] = [(k, int(v)) for k, v in w_counts.items() if k]

    def build_cache(col):
        if col not in df.columns: return []
        vc = df[col].value_counts().sort_index()
        return [(k, int(v)) for k, v in vc.items() if k and k != 'nan']

    CATEGORY_CACHE[:]      = build_cache('Category')
    SUBCATEGORY_CACHE[:]   = build_cache('Subcategory')
    PRODUCT_GROUP_CACHE[:] = build_cache('Product_Group')
    SUB_PROD_GRP_CACHE[:]  = build_cache('Sub_Product_Group')
    SKU_STATUS_CACHE[:]    = build_cache('PWC_SKUSTATUS')
    VENDOR_CACHE[:]        = build_cache('VENDACCOUNT')

    # Build duplicate HUID cache
    huid_exploded = (
        df[["HUID", "ITEMID"]]
        .assign(HUID=df["HUID"].str.split(","))
        .explode("HUID")
    )
    huid_exploded["HUID"] = huid_exploded["HUID"].str.strip()
    huid_exploded = huid_exploded[huid_exploded["HUID"] != ""]

    grp = huid_exploded.groupby("HUID")
    sku_counts = grp["ITEMID"].nunique()
    row_counts = grp["ITEMID"].count()
    dup_huids = sku_counts[sku_counts > 1].index

    dup_skus = (
        huid_exploded[huid_exploded["HUID"].isin(dup_huids)]
        .groupby("HUID")["ITEMID"]
        .apply(lambda s: sorted(s.unique().tolist()))
    )

    DUP_HUID_CACHE[:] = sorted(
        [
            {
                "huid":      h,
                "skus":      dup_skus[h],
                "sku_count": int(sku_counts[h]),
                "row_count": int(row_counts[h]),
            }
            for h in dup_huids
        ],
        key=lambda x: -x["sku_count"],
    )
    print("[startup] Caches completely loaded into global application scope.", flush=True)

# Initialize data caches
load_data()

# ─── FastAPI Initialization ───────────────────────────────────────────────────
app = FastAPI(title="NJL Catalogue API", version="1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Securely serve static files (CSS, JS, Images) from the /static folder
static_path = os.path.join(BASE_DIR, "static")
if os.path.exists(static_path):
    app.mount("/static", StaticFiles(directory=static_path), name="static")

# ─── Endpoints ────────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
def serve_html():
    if not os.path.exists(HTML_FILE):
        return HTMLResponse("<h2>Error: Web layout index.html asset not found.</h2>", status_code=404)
    with open(HTML_FILE, "r", encoding="utf-8") as fh:
        return fh.read()

@app.get("/api/stats")
def api_stats():
    return {"total": TOTAL_COUNT}

@app.get("/api/filter-values")
def api_filter_values(
    type:  str = Query("serial", description="'serial', 'huid', or 'warehouse'"),
    q:     str = Query("",       description="optional substring filter"),
    limit: int = Query(MAX_FILTER_RESULTS, ge=1, le=10_000),
):
    if type == "serial":
        cache = SERIAL_CACHE
    elif type == "warehouse":
        cache = WAREHOUSE_CACHE
    else:
        cache = HUID_CACHE

    if q:
        q_lower = q.lower()
        cache = [(v, c) for v, c in cache if q_lower in v.lower()]

    return {"values": cache[:limit]}

@app.get("/api/dropdown-values")
def api_dropdown_values(
    type: str = Query(..., description="category|subcategory|product_group|sub_product_group|sku_status|vendor"),
    q:    str = Query("", description="optional substring filter"),
):
    cache_map = {
        "category":          CATEGORY_CACHE,
        "subcategory":       SUBCATEGORY_CACHE,
        "product_group":     PRODUCT_GROUP_CACHE,
        "sub_product_group": SUB_PROD_GRP_CACHE,
        "sku_status":        SKU_STATUS_CACHE,
        "vendor":            VENDOR_CACHE,
    }
    cache = cache_map.get(type, [])
    if q.strip():
        ql = q.strip().lower()
        cache = [(v, cnt) for v, cnt in cache if ql in v.lower()]
    return {"values": cache}

@app.get("/api/inventory")
def api_inventory(
    page:      int = Query(1,    ge=1),
    page_size: int = Query(20,   ge=1,  le=200),
    sort:      str = Query("default"),
    q:              str = Query(""),
    serials:        str = Query(""),
    huids:          str = Query(""),
    warehouses:     str = Query(""),
    categories:     str = Query(""),
    subcategories:  str = Query(""),
    product_groups: str = Query(""),
    sub_prod_grps:  str = Query(""),
    sku_statuses:   str = Query(""),
    vendors:        str = Query(""),
):
    df = df_global

    if serials.strip():
        serial_set = {s.strip() for s in serials.split(",") if s.strip()}
        df = df[df["SERIALNUMBER"].isin(serial_set)]

    if huids.strip():
        huid_set = {h.strip() for h in huids.split(",") if h.strip()}
        mask = df["HUID"].apply(lambda cell: any(tok.strip() in huid_set for tok in str(cell).split(",") if tok.strip()))
        df = df[mask]

    if warehouses.strip():
        wh_set = {w.strip() for w in warehouses.split(",") if w.strip()}
        df = df[df["WAREHOUSE"].isin(wh_set)]

    def apply_exact(df_in, col, param):
        if not param.strip(): return df_in
        vals = {v.strip() for v in param.split('|') if v.strip()}
        return df_in[df_in[col].isin(vals)] if vals else df_in

    df = apply_exact(df, 'Category',          categories)
    df = apply_exact(df, 'Subcategory',       subcategories)
    df = apply_exact(df, 'Product_Group',     product_groups)
    df = apply_exact(df, 'Sub_Product_Group', sub_prod_grps)
    df = apply_exact(df, 'PWC_SKUSTATUS',      sku_statuses)
    df = apply_exact(df, 'VENDACCOUNT',        vendors)

    if q.strip():
        import re as _re
        terms = [t.lower() for t in _re.split(r"[,\s]+", q.strip()) if t.strip()]
        row_text = ROW_TEXT_CACHE.loc[df.index]
        for term in terms:
            mask = row_text.str.contains(term, na=False, regex=False)
            df = df[mask]
            row_text = row_text.loc[df.index]

    total_filtered = len(df)

    if sort in SORT_MAP:
        sort_col, ascending = SORT_MAP[sort]
        df = df.sort_values(sort_col, ascending=ascending, kind="stable")

    start    = (page - 1) * page_size
    end      = start + page_size
    page_df  = df.iloc[start:end]

    records = page_df.replace({np.nan: None, np.inf: None, -np.inf: None}).to_dict(orient="records")

    for rec in records:
        for key in ("GROSSQTY", "NETWEIGHT"):
            v = rec.get(key)
            rec[key] = float(v) if v is not None else 0.0

    return {
        "total_filtered": total_filtered,
        "total_all":      TOTAL_COUNT,
        "page":           page,
        "page_size":      page_size,
        "total_pages":    math.ceil(total_filtered / page_size) if total_filtered else 0,
        "data":           records,
    }

@app.get("/api/duplicate-huids")
def api_duplicate_huids(
    q:         str = Query("",  description="substring filter on HUID"),
    page:      int = Query(1,   ge=1),
    page_size: int = Query(100, ge=1, le=500),
):
    data = DUP_HUID_CACHE
    if q.strip():
        q_lower = q.strip().lower()
        data = [d for d in data if q_lower in d["huid"].lower()]

    total = len(data)
    start = (page - 1) * page_size
    return {
        "total":     total,
        "page":      page,
        "page_size": page_size,
        "data":      data[start : start + page_size],
    }

# ─── Local Port Execution Guard ───────────────────────────────────────────────
def find_free_port(start_port: int) -> int:
    port = start_port
    while port < start_port + 20:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind((HOST, port))
                return port
            except OSError:
                port += 1
    raise RuntimeError(f"No free port found between {start_port} and {start_port + 20}")

if __name__ == "__main__":
    actual_port = find_free_port(PORT)
    if actual_port != PORT:
        print(f"[startup] ⚠️  Port {PORT} in use — shifting routing target to {actual_port}.", flush=True)
    uvicorn.run("server:app", host=HOST, port=actual_port, reload=False, log_level="info")