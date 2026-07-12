"""
TradeSmart Trading Backend - FastAPI
=====================================
Serves both the frontend HTML and the trading API.
"""

import os
import json
import hashlib
import requests
import openpyxl
from datetime import datetime
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from pathlib import Path

app = FastAPI(title="TradeSmart Trading Bot")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

APP_KEY    = os.getenv("APP_KEY",    "YMUS1087_01")
SECRET_KEY = os.getenv("SECRET_KEY", "mqordqqa5qqlzgu3fm")
CLIENT_ID  = os.getenv("CLIENT_ID",  "YMUS1087")
EXCEL_FILE = os.getenv("EXCEL_FILE", "1. Nifty_Macro_V1.1.xlsm")
SHEET_NAME = "Paper Trading"
WATCH_ROWS = [3, 4, 5, 6, 7]
EXCHANGE   = "NFO"
PRODUCT    = "M"
ORDER_TYPE = "LMT"
BASE_URL   = "https://v2api.tradesmartonline.in/NorenWClientAPIv2"

token_store = {"access_token": None}


class TokenRequest(BaseModel):
    code: str

class OrderRequest(BaseModel):
    row: int
    button: str


@app.get("/", response_class=HTMLResponse)
def serve_frontend():
    html_path = Path(__file__).parent / "index.html"
    if html_path.exists():
        return HTMLResponse(content=html_path.read_text(encoding="utf-8"))
    return HTMLResponse(content="<h1>Place index.html next to main.py</h1>")


def ts_post(endpoint, payload, auth=True):
    url     = f"{BASE_URL}/{endpoint}"
    headers = {}
    if auth and token_store["access_token"]:
        headers["Authorization"] = f"Bearer {token_store['access_token']}"
    body = "jData=" + json.dumps(payload)
    res  = requests.post(url, data=body, headers=headers, timeout=10)
    return res.json()


@app.get("/oauth-url")
def get_oauth_url():
    return {"url": f"https://v2api.tradesmartonline.in/OAuthlogin/authorize/oauth?client_id={APP_KEY}&redirect_uri=https://tradesmart-app.onrender.com/callback"}

@app.get("/callback", response_class=HTMLResponse)
def oauth_callback(code: str = None, error: str = None):
    if error or not code:
        return HTMLResponse(content="""
            <html><body style='font-family:sans-serif;text-align:center;padding:60px;background:#0a0e1a;color:#e2e8f0'>
            <h2 style='color:#ef4444'>Login Failed</h2>
            <p><a href='/' style='color:#3b82f6'>Go back and try again</a></p>
            </body></html>
        """)
    return HTMLResponse(content=f"""
        <html><body style='font-family:sans-serif;text-align:center;padding:60px;background:#0a0e1a;color:#e2e8f0'>
        <h2 style='color:#22c55e'>Login Successful!</h2>
        <p>Completing login...</p>
        <script>
            var code = '{code}';
            if (window.opener) {{
                window.opener.postMessage({{code: code}}, '*');
                setTimeout(() => window.close(), 500);
            }} else {{
                window.location.href = '/?code=' + code;
            }}
        </script>
        </body></html>
    """)


@app.post("/auth/token")
def exchange_token(req: TokenRequest):
    code     = req.code.strip()
    checksum = hashlib.sha256((APP_KEY + SECRET_KEY + code).encode()).hexdigest()
    ret      = ts_post("GenAcsTok", {"code": code, "checksum": checksum}, auth=False)
    if ret and ret.get("stat") == "Ok":
        token_store["access_token"] = ret.get("access_token") or ret.get("susertoken")
        return {"success": True, "uid": ret.get("uid"), "uname": ret.get("uname"), "actid": ret.get("actid")}
    raise HTTPException(status_code=401, detail=ret.get("emsg", "Login failed"))


@app.get("/auth/status")
def auth_status():
    if not token_store["access_token"]:
        return {"logged_in": False}
    ret = ts_post("UserDetails", {"uid": CLIENT_ID})
    if ret and ret.get("stat") == "Ok":
        return {"logged_in": True, "uname": ret.get("uname"), "uid": ret.get("uid")}
    token_store["access_token"] = None
    return {"logged_in": False}


@app.delete("/auth/logout")
def logout():
    token_store["access_token"] = None
    return {"success": True}


@app.get("/rows")
def get_rows():
    rows = []
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True, keep_vba=True)
        ws = wb[SHEET_NAME]
        for row_num in WATCH_ROWS:
            symbol = ws[f"C{row_num}"].value
            qty    = ws[f"B{row_num}"].value
            rate   = ws[f"E{row_num}"].value
            rows.append({
                "row"   : row_num,
                "button": row_num - 2,
                "symbol": str(symbol).strip() if symbol else None,
                "qty"   : int(float(qty)) if qty else None,
                "rate"  : float(rate) if rate else None,
            })
        wb.close()
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"Excel file not found: {EXCEL_FILE}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    return {"rows": rows}


@app.post("/order")
def place_order(req: OrderRequest):
    if not token_store["access_token"]:
        raise HTTPException(status_code=401, detail="Not logged in")

    try:
        wb     = openpyxl.load_workbook(EXCEL_FILE, data_only=True, keep_vba=True)
        ws     = wb[SHEET_NAME]
        symbol = ws[f"C{req.row}"].value
        qty    = ws[f"B{req.row}"].value
        rate   = ws[f"E{req.row}"].value
        wb.close()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Excel read error: {e}")

    if not symbol:
        raise HTTPException(status_code=400, detail=f"Row {req.row}: Symbol (col C) is empty")
    if not qty or float(qty) <= 0:
        raise HTTPException(status_code=400, detail=f"Row {req.row}: Qty is empty or zero")
    if rate is None:
        raise HTTPException(status_code=400, detail=f"Row {req.row}: Rate (col E) is empty")

    trantype = "B" if req.button == "BUY" else "S"
    price    = 0.0 if ORDER_TYPE == "MKT" else float(rate)
    qty_int  = int(float(qty))
    sym      = str(symbol).strip()

    search = ts_post("SearchScrip", {"uid": CLIENT_ID, "exch": EXCHANGE, "stext": sym})
    tsym   = None
    if search and search.get("stat") == "Ok" and search.get("values"):
        for entry in search["values"]:
            if sym.replace(" ", "").upper() in entry["tsym"].replace(" ", "").upper():
                tsym = entry["tsym"]
                break
        if not tsym:
            tsym = search["values"][0]["tsym"]

    if not tsym:
        raise HTTPException(status_code=404, detail=f"Symbol '{sym}' not found on {EXCHANGE}")

    ret = ts_post("PlaceOrder", {
        "uid": CLIENT_ID, "actid": CLIENT_ID, "exch": EXCHANGE,
        "tsym": tsym, "qty": str(qty_int), "prc": str(price),
        "dscqty": "0", "prd": PRODUCT, "trantype": trantype,
        "prctyp": ORDER_TYPE, "ret": "DAY", "ordersource": "API",
        "remarks": f"web_row{req.row}_{datetime.now().strftime('%H%M%S')}",
    })

    if ret and ret.get("stat") == "Ok":
        return {
            "success"  : True,
            "order_no" : ret.get("norenordno"),
            "symbol"   : tsym,
            "trantype" : trantype,
            "qty"      : qty_int,
            "price"    : price,
            "timestamp": datetime.now().strftime("%H:%M:%S"),
        }
    raise HTTPException(status_code=400, detail=ret.get("emsg", "Order failed") if ret else "No response")
