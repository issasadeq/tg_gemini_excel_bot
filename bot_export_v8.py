# -*- coding: utf-8 -*-
import os, io, logging, datetime, json, re
from typing import Dict, Any

from dotenv import load_dotenv
from telegram import Update, InlineKeyboardMarkup, InlineKeyboardButton, InputFile
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, CallbackQueryHandler, filters

import pandas as pd
from pathlib import Path

# ====== PDF (fpdf2) + RTL helpers ======
from fpdf import FPDF
import arabic_reshaper
from bidi.algorithm import get_display

def _rtl(s: str) -> str:
    s = str(s or "")
    if not s.strip():
        return ""
    try:
        return get_display(arabic_reshaper.reshape(s))
    except Exception:
        return s

class StatementPDF(FPDF):
    # نترك الهيدر/الفوتر الافتراضيين فارغين
    pass

def build_statement_pdf_fpdf(out_path: Path, rows_df: pd.DataFrame, header_info: Dict[str, str]):
    """
    يبني PDF كشف حساب بمنطق:
    - رصيد تراكمي من عمود AMT (مدين +، دائن -)
    - أعمدة: التاريخ، نوع المستند، رقم المستند، البيان، رقم المرجع، مدين، دائن، الرصيد
    """
    pdf = StatementPDF(orientation="L", unit="mm", format="A4")  # Landscape مثل الصورة
    pdf.add_page()

    # جرّب خط Arial (موجود غالبًا في ويندوز). لإضافة خط TTF مخصص:
    # pdf.add_font("Noto", "", "fonts/NotoNaskhArabic-Regular.ttf", uni=True); font_name="Noto"
    font_name = "Arial"
    pdf.set_font(font_name, size=14)

    # رأس التقرير
    pdf.set_xy(10, 10)
    pdf.cell(0, 8, _rtl(header_info.get("company", "كشف حساب")), ln=1, align="R")
    pdf.set_font(font_name, size=11)
    pdf.cell(0, 7, _rtl(header_info.get("title", "كشف حساب تحليلي - قبل الترحيل - رصيد")), ln=1, align="R")
    acc = []
    if header_info.get("account"):  acc.append(f"رقم الحساب: {header_info['account']}")
    if header_info.get("analytic"): acc.append(f"الحساب التحليلي: {header_info['analytic']}")
    if acc:
        pdf.cell(0, 7, _rtl(" | ".join(acc)), ln=1, align="R")
    pdf.cell(0, 7, _rtl(f"من تاريخ: {header_info.get('from','')} إلى تاريخ: {header_info.get('to','')}"), ln=1, align="R")
    if header_info.get("currency_label"):
        pdf.cell(0, 7, _rtl(f"العملة: {header_info['currency_label']}"), ln=1, align="R")
    pdf.ln(3)

    # الجدول
    headers = ["التاريخ","نوع المستند","رقم المستند","البيان","رقم المرجع","مدين","دائن","الرصيد"]
    col_w = [35, 40, 28, 100, 30, 30, 30, 35]
    pdf.set_font(font_name, "B", 10)
    pdf.set_fill_color(230,230,230)
    for i, h in enumerate(headers):
        pdf.cell(col_w[i], 8, _rtl(h), border=1, align="C", fill=True)
    pdf.ln(8)

    pdf.set_font(font_name, "", 9)
    bal = 0.0
    # الطباعة سطرًا بسطر
    for _, r in rows_df.iterrows():
        try:
            amt = float(str(r.get("amt","") or "0"))
        except:
            amt = 0.0
        bal += amt

        debit  = str(r.get("debit","") or "").strip()
        credit = str(r.get("credit","") or "").strip()

        row = [
            r.get("date",""),
            _rtl(r.get("document_type","")),
            r.get("voucher_number",""),
            _rtl(r.get("description","")),
            str(r.get("reference_number","") or ""),
            f"{float(debit):,.2f}" if debit else "",
            f"{float(credit):,.2f}" if credit else "",
            f"{bal:,.2f}",
        ]

        # اطبع الأعمدة (يمين للحقول العربية)
        for i, val in enumerate(row):
            align = "R" if i in (1,3) else "C"
            pdf.cell(col_w[i], 7, val, border=1, align=align)
        pdf.ln(7)

        # إذا اقتربنا من نهاية الصفحة — اطبع رأس الجدول في الصفحة التالية
        if pdf.get_y() > 185:
            pdf.add_page()
            pdf.set_font(font_name, "B", 10)
            pdf.set_fill_color(230,230,230)
            for i, h in enumerate(headers):
                pdf.cell(col_w[i], 8, _rtl(h), border=1, align="C", fill=True)
            pdf.ln(8)
            pdf.set_font(font_name, "", 9)

    pdf.output(str(out_path))

# ============= باقي البوت (OCR/Excel/بحث) =============

# Gemini API
import google.generativeai as genai
from num2words import num2words  # للعرض في الرسائل فقط

# ====== Load env ======
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-1.5-flash")

EXCEL_PATH = os.getenv("EXCEL_PATH", "entries.xlsx")
IMAGES_DIR = Path(os.getenv("IMAGES_DIR", "images"))

# ====== Logging ======
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("gemini-structured-bot")

# ====== Gemini setup ======
if not GEMINI_API_KEY:
    raise RuntimeError("الرجاء ضبط GEMINI_API_KEY في .env")
genai.configure(api_key=GEMINI_API_KEY)

# ====== Excel columns (بدون amount_words) ======
COLUMNS = [
    "chat_id","user_id","username","timestamp",
    "bank_or_exchange","document_type","date","voucher_number",
    "sender","beneficiary","description",
    "debit","credit","amt","currency",
    "raw_text","image_path"
]

def _ensure_df(df: pd.DataFrame) -> pd.DataFrame:
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[COLUMNS]

# ====== Sheets by currency ======
INVALID_SHEET_CHARS = "[]:*?/\\"

def sheet_name_for_currency(currency: str) -> str:
    c = (currency or "").strip()
    aliases = {
        "ريال يمني": ["ريال يمني", "ر.ي", "YER", "Yemeni Rial","Yemeni Riyal","يمني"],
        "ريال سعودي": ["ريال سعودي", "ر.س", "SAR", "Saudi Riyal","سعودي"],
        "دولار أمريكي": ["دولار", "دولار أمريكي", "USD", "$"],
        "درهم إماراتي": ["درهم", "درهم إماراتي", "AED", "UAE Dirham"],
    }
    for name, al in aliases.items():
        if any(a in c for a in al):
            return name
    name = c or "غير محدد"
    for ch in INVALID_SHEET_CHARS:
        name = name.replace(ch, "-")
    return name[:31]

def load_entries_all() -> pd.DataFrame:
    xls = Path(EXCEL_PATH)
    if not xls.exists():
        return pd.DataFrame(columns=COLUMNS)
    try:
        sheets = pd.read_excel(EXCEL_PATH, sheet_name=None, dtype=str)
    except ValueError:
        return pd.DataFrame(columns=COLUMNS)
    frames = []
    for sh_name, df in sheets.items():
        df = df.fillna("")
        df = _ensure_df(df)
        df["__sheet__"] = sh_name
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=COLUMNS)
    return pd.concat(frames, ignore_index=True)

def append_to_currency_sheet(row: dict, sheet_name: str):
    xls = Path(EXCEL_PATH)
    if xls.exists():
        try:
            current = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, dtype=str).fillna("")
        except Exception:
            current = pd.DataFrame(columns=COLUMNS)
        mode = "a"
    else:
        current = pd.DataFrame(columns=COLUMNS)
        mode = "w"

    current = _ensure_df(current)
    new_df = pd.concat([current, pd.DataFrame([row])], ignore_index=True)

    if mode == "a":
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            _ensure_df(new_df).to_excel(writer, index=False, sheet_name=sheet_name)
    else:
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="w") as writer:
            _ensure_df(new_df).to_excel(writer, index=False, sheet_name=sheet_name)

def upsert_in_currency_sheet(row: dict, sheet_name: str) -> str:
    xls = Path(EXCEL_PATH)
    if xls.exists():
        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, dtype=str).fillna("")
        except Exception:
            df = pd.DataFrame(columns=COLUMNS)
        mode = "a"
    else:
        df = pd.DataFrame(columns=COLUMNS)
        mode = "w"

    df = _ensure_df(df)

    v = str(row.get("voucher_number","")).strip()
    d = str(row.get("date","")).strip()
    mask = (df["voucher_number"].astype(str) == v) & (df["date"].astype(str) == d)
    if mask.any():
        idx = df[mask].index[0]
        for k in row:
            if k in df.columns:
                df.at[idx, k] = row[k]
        result = "updated"
    else:
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        result = "inserted"

    if mode == "a":
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            _ensure_df(df).to_excel(writer, index=False, sheet_name=sheet_name)
    else:
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="w") as writer:
            _ensure_df(df).to_excel(writer, index=False, sheet_name=sheet_name)
    return result

# ====== Normalization & postprocess ======
ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")

def normalize(s) -> str:
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\u200f","").replace("\u200e","")
    s = s.translate(ARABIC_DIGITS)
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()

def postprocess_fields(d: Dict[str, Any]) -> Dict[str, Any]:
    out = {
        "bank_or_exchange": normalize(d.get("bank_or_exchange") or d.get("company") or d.get("الشركة") or ""),
        "document_type": normalize(d.get("document_type") or d.get("نوع_المستند") or ""),
        "date": normalize(d.get("date") or d.get("التاريخ") or ""),
        "voucher_number": normalize(d.get("voucher_number") or d.get("رقم_السند") or ""),
        "sender": normalize(d.get("sender") or d.get("المرسل") or d.get("العميل") or ""),
        "beneficiary": normalize(d.get("beneficiary") or d.get("المستفيد") or ""),
        "description": normalize(d.get("description") or d.get("البيان") or ""),
        "amount_value": normalize(d.get("amount_value") or d.get("المبلغ") or ""),
        "currency": normalize(d.get("currency") or d.get("العملة") or ""),
    }

    account_holder = normalize(d.get("account_holder") or d.get("صاحب_الحساب") or d.get("صاحب الحساب") or "")
    if not out["beneficiary"] and account_holder:
        out["beneficiary"] = account_holder

    m = re.search(r"(\d+)", out["voucher_number"])
    if m:
        out["voucher_number"] = m.group(1)

    amt = out["amount_value"]
    if amt:
        amt = amt.replace(",", "").replace(" ", "")
        if amt.count(".") > 1:
            amt = amt.replace(".", "")
        out["amount_value"] = re.sub(r"[^0-9\.-]", "", amt)

    m = re.search(r"(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})", out["date"])
    if m:
        y, mo, d2 = m.groups()
        try:
            out["date"] = f"{int(y):04d}-{int(mo):02d}-{int(d2):02d}"
        except:
            pass
    return out

# ====== Amount → Words (للعرض فقط) ======
def number_to_arabic_words(num: str, currency: str = "ريال") -> str:
    try:
        n = int(str(num).strip().replace(",", ""))
    except:
        return ""
    words = num2words(n, lang='ar')
    words = words.replace("واحد ألف", "ألف").replace("واحدة ألف", "ألف")
    return f"{words} {currency} فقط لا غير"

# ====== Business Rules ======
def apply_custom_rules(fields: Dict[str, Any]) -> Dict[str, Any]:
    sender = fields.get("sender", "")
    beneficiary = fields.get("beneficiary", "")
    doc_type = fields.get("document_type", "")

    # 1) المستفيد/صاحب الحساب يحوي "مهدي" => قبض
    if "مهدي" in beneficiary:
        fields["document_type"] = "سند قبض"
    # 2) سند قبض مراجعة + المرسل يحوي "مهدي" => صرف
    elif ("سند قبض مراجعة" in doc_type) and ("مهدي" in sender):
        fields["document_type"] = "سند صرف"
    # 3) المرسل يحوي "مهدي" => صرف
    elif "مهدي" in sender:
        fields["document_type"] = "سند صرف"

    return fields

def split_debit_credit(fields: Dict[str, Any]) -> Dict[str, Any]:
    amt = fields.get("amount_value","")
    doc_type = fields.get("document_type","")
    debit, credit = "", ""
    if doc_type == "سند قبض":
        debit = amt
    elif doc_type == "سند صرف":
        credit = amt
    return {"debit": debit, "credit": credit}

def compute_amt(debit: str, credit: str) -> str:
    d = (debit or "").strip()
    c = (credit or "").strip()
    try:
        if d:
            return str(float(d))   # موجب
        if c:
            return str(-float(c))  # سالب
    except:
        return ""
    return ""

# ====== Search helpers ======
def find_rows(query: str, limit: int = 10) -> pd.DataFrame:
    df = load_entries_all()
    if df.empty:
        return df
    q = normalize(query).lower()
    mask = (
        df["voucher_number"].astype(str).str.contains(q, case=False, na=False) |
        df["sender"].astype(str).str.contains(q, case=False, na=False) |
        df["beneficiary"].astype(str).str.contains(q, case=False, na=False) |
        df["description"].astype(str).str.contains(q, case=False, na=False) |
        df["document_type"].astype(str).str.contains(q, case=False, na=False)
    )
    res = df[mask].copy()
    return res.head(limit)

def find_by_voucher(voucher: str) -> pd.DataFrame:
    df = load_entries_all()
    if df.empty:
        return df
    voucher = re.sub(r"\D", "", voucher or "")
    return df[df["voucher_number"].astype(str) == voucher]

def filter_by_params(df: pd.DataFrame, params: Dict[str,str]) -> pd.DataFrame:
    if df.empty:
        return df
    out = df
    if "date_range" in params and ".." in params["date_range"]:
        start, end = params["date_range"].split("..", 1)
        start = start.strip() or "0001-01-01"
        end = end.strip() or "9999-12-31"
        out = out[(out["date"] >= start) & (out["date"] <= end)]
    for key, col in {
        "type": "document_type",
        "sender": "sender",
        "beneficiary": "beneficiary",
        "currency": "currency",
        "bank": "bank_or_exchange",
    }.items():
        if key in params and params[key]:
            val = normalize(params[key]).lower()
            out = out[out[col].astype(str).str.lower().str.contains(val, na=False)]
    return out

def parse_export_args(text: str) -> Dict[str,str]:
    args = text.split()[1:]
    params = {}
    for token in args:
        if ".." in token and re.match(r"\d{4}-\d{2}-\d{2}..\d{4}-\d{2}-\d{2}", token):
            params["date_range"] = token
        elif "=" in token:
            k, v = token.split("=", 1)
            params[k.strip()] = v.strip()
    return params

# ====== PDF helpers ======
def parse_pdf_args(text: str) -> Dict[str,str]:
    """
    /pdf 2025-01-01..2025-04-30 currency=ريال يمني title=سوبر-ماركت-المسيلة account=1112010001 analytic=1
    """
    args = text.split()[1:]
    params = {"from":"", "to":"", "currency":"", "title":"كشف حساب", "account":"", "analytic":""}
    for token in args:
        if ".." in token and re.match(r"\d{4}-\d{2}-\d{2}..\d{4}-\d{2}-\d{2}", token):
            a,b = token.split("..",1)
            params["from"] = a.strip()
            params["to"] = b.strip()
        elif "=" in token:
            k,v = token.split("=",1)
            params[k.strip()] = v.strip()
    return params

def filter_by_date_and_currency(df: pd.DataFrame, start: str, end: str, currency_label: str) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    if start:
        out = out[out["date"] >= start]
    if end:
        out = out[out["date"] <= end]
    if currency_label:
        # نحول اسم العملة للشييت القياسي (ريال يمني / ريال سعودي ..)
        sheet = sheet_name_for_currency(currency_label)
        if "__sheet__" in out.columns:
            out = out[out["__sheet__"].astype(str) == sheet]
    # تأكد من الأعمدة
    for col in ["amt","debit","credit","document_type","voucher_number","description","date"]:
        if col not in out.columns:
            out[col] = ""
    return out

# ====== Gemini prompts ======
SYSTEM_PROMPT = (
    "You are a precise information extractor for Arabic financial documents (exchange/receipt vouchers). "
    "Return STRICT JSON only with keys: "
    "bank_or_exchange, document_type, date, voucher_number, sender, beneficiary, description, amount_value, currency. "
    "Rules: 1) Convert Eastern Arabic digits to Western. "
    "2) Output date as 'YYYY-MM-DD'. "
    "3) Keep names as-is. "
    "4) amount_value is numeric only. "
    "5) If a field is missing, return empty string. No explanations."
)

def gemini_extract_structured(img_bytes: bytes, mime: str) -> Dict[str, Any]:
    model = genai.GenerativeModel(GEMINI_MODEL)
    parts = [SYSTEM_PROMPT, {"mime_type": mime or "image/jpeg", "data": img_bytes}]
    resp = model.generate_content(parts)
    text = (resp.text or "").strip()
    text = text.strip("`").replace("```json","").replace("```","").replace("json\n","").strip()
    try:
        data = json.loads(text)
    except Exception:
        data = {}
    return postprocess_fields(data)

def gemini_ocr_plain(img_bytes: bytes, mime: str) -> str:
    model = genai.GenerativeModel(GEMINI_MODEL)
    parts = ["Extract ALL text as plain text. Keep line breaks. No commentary.",
             {"mime_type": mime or "image/jpeg", "data": img_bytes}]
    resp = model.generate_content(parts)
    return (resp.text or "").strip()

# ====== Session & UI ======
SESSION: Dict[int, Dict[str, Any]] = {}

SAVE_KB = InlineKeyboardMarkup([[
    InlineKeyboardButton("💾 احفظ في Excel", callback_data="save_excel"),
    InlineKeyboardButton("❌ لا تحفظ", callback_data="discard")
]])

# ====== Handlers ======
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "مرحبا بك في البوت المحاسبي الذكي 👋\n"
        "الأوامر: /find <كلمة-بحث> ، /voucher <رقم-سند> ، /export <تاريخ..تاريخ> [type=][sender=][beneficiary=][currency=][bank=]\n"
        "وأيضًا: /pdf YYYY-MM-DD..YYYY-MM-DD currency=ريال يمني title=اسم-الجهة account=رقم analytic=1"
    )

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "أرسل صورة أو نصًا لاستخراج الحقول ثم الحفظ في Excel.\n"
        "أوامر البحث والتقارير:\n"
        "• /find مهدي\n"
        "• /voucher 9779370654\n"
        "• /export 2025-08-01..2025-08-31 type=سند قبض sender=مهدي\n"
        "• /pdf 2025-01-01..2025-04-30 currency=ريال يمني title=سوبر-ماركت-المسيلة account=1112010001 analytic=1"
    )

async def wherefile_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"ملف الإكسل: {EXCEL_PATH} (كل عملة في ورقة). الصور: {IMAGES_DIR}")

async def cmd_find(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    args = msg.text.split(maxsplit=1)
    if len(args) < 2:
        await msg.reply_text("استخدم: /find <كلمة-بحث>")
        return
    query = args[1]
    df = find_rows(query, limit=10)
    if df.empty:
        await msg.reply_text("لا توجد نتائج مطابقة.")
        return
    lines = ["نتائج البحث (أول 10):"]
    for _, r in df.iterrows():
        lines.append(f"- {r['date']} | {r['voucher_number']} | {r['beneficiary']} | مدين:{r['debit']} دائن:{r['credit']} AMT:{r['amt']} {r['currency']}")
    lines.append("\nاستخدم /voucher <رقم-سند> لعرض التفاصيل.")
    await msg.reply_text("\n".join(lines))

async def cmd_voucher(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    args = msg.text.split(maxsplit=1)
    if len(args) < 2:
        await msg.reply_text("استخدم: /voucher <رقم-سند>")
        return
    df = find_by_voucher(args[1])
    if df.empty:
        await msg.reply_text("لم يتم العثور على سند بهذا الرقم.")
        return
    r = df.iloc[0]
    amount_numeric = r.get('debit') or r.get('credit') or ""
    amount_words = number_to_arabic_words(amount_numeric, r.get('currency') or "ريال")
    await msg.reply_text(
        f"""تفاصيل السند:
- البنك/شركة الصرافة: {r['bank_or_exchange']}
- نوع المستند: {r['document_type']}
- التاريخ: {r['date']}
- رقم السند: {r['voucher_number']}
- المرسل: {r['sender']}
- المستفيد: {r['beneficiary']}
- البيان: {r['description']}
- مدين: {r['debit']} | دائن: {r['credit']} | AMT: {r['amt']} {r['currency']}
- المبلغ كتابة: {amount_words}
- الصورة: {r['image_path'] if r.get('image_path','') else '(لا يوجد)'}"""
    )

async def cmd_export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    params = parse_export_args(msg.text)
    df = load_entries_all()
    if df.empty:
        await msg.reply_text("لا يوجد بيانات للتصدير بعد.")
        return
    fdf = filter_by_params(df, params)
    if fdf.empty:
        await msg.reply_text("لا توجد سجلات مطابقة لمعايير التصدير.")
        return
    ts = datetime.datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    out_path = Path(f"report_{ts}.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        fdf.to_excel(writer, index=False, sheet_name="تقرير")
    await msg.reply_document(document=InputFile(out_path.open("rb"), filename=out_path.name))
    out_path.unlink(missing_ok=True)

async def cmd_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    مثال:
    /pdf 2025-01-01..2025-04-30 currency=ريال يمني title=سوبر-ماركت-المسيلة account=1112010001 analytic=1
    """
    msg = update.message
    params = parse_pdf_args(msg.text)
    if not params.get("from") or not params.get("to"):
        await msg.reply_text("استخدم: /pdf YYYY-MM-DD..YYYY-MM-DD currency=ريال يمني title=الاسم account=رقم analytic=1")
        return

    all_df = load_entries_all()
    if all_df.empty:
        await msg.reply_text("لا توجد بيانات في الإكسل.")
        return

    fdf = filter_by_date_and_currency(all_df, params["from"], params["to"], params.get("currency",""))
    if fdf.empty:
        await msg.reply_text("لا توجد قيود ضمن هذه الفترة/العملة.")
        return

    # رتب حسب التاريخ ثم رقم السند
    try:
        fdf = fdf.sort_values(by=["date","voucher_number"], ascending=[True, True])
    except Exception:
        pass

    header = {
    "company": "كشف حساب - مهدي",
    "title": "كشف حساب إيرادات ومصروفات",
    "account": "حساب: مهدي صالح ناصر الصويدر",
    "analytic": "",
    "from": params["from"],
    "to": params["to"],
    "currency_label": params.get("currency",""),
    }


    ts = datetime.datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    out_path = Path(f"statement_{ts}.pdf")
    try:
        build_statement_pdf_fpdf(out_path, fdf, header)
        await msg.reply_document(document=InputFile(out_path.open("rb"), filename=out_path.name))
    except Exception as e:
        logging.exception("PDF build failed")
        await msg.reply_text(f"تعذر إنشاء PDF: {e}")
    finally:
        try:
            out_path.unlink(missing_ok=True)
        except:
            pass

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message or update.edited_message
    if not msg or not msg.text:
        return
    txt = msg.text.strip()
    fields = postprocess_fields({
        "bank_or_exchange": "",
        "document_type": "",
        "date": txt,
        "voucher_number": txt,
        "sender": "",
        "beneficiary": "",
        "description": txt,
        "amount_value": txt,
        "currency": "",
    })
    fields = apply_custom_rules(fields)
    SESSION[update.effective_user.id] = {"raw_text": txt, "fields": fields, "image_path": ""}

    amount_val = fields.get('amount_value','')
    currency = fields.get('currency','ريال')
    amount_words = number_to_arabic_words(amount_val, currency)
    doc_type = fields.get("document_type","")
    if doc_type == "سند قبض":
        amount_line = f"- المبلغ (مدين): {amount_val} {currency}"
    elif doc_type == "سند صرف":
        amount_line = f"- المبلغ (دائن): {amount_val} {currency}"
    else:
        amount_line = f"- المبلغ: {amount_val} {currency}"

    await msg.reply_text(
        f"""الحقول المستخرجة:
- البنك/شركة الصرافة: {fields['bank_or_exchange']}
- نوع المستند: {doc_type}
- التاريخ: {fields['date']}
- رقم السند: {fields['voucher_number']}
- المرسل: {fields['sender']}
- المستفيد: {fields['beneficiary']}
- البيان: {fields['description']}
{amount_line}
- المبلغ كتابة: {amount_words}

هل تريد حفظها في Excel؟""",
        reply_markup=SAVE_KB
    )

async def handle_photo_or_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message or update.edited_message
    if not msg:
        return
    tg_file = None
    mime = "image/jpeg"
    ext = ".jpg"
    if msg.photo:
        tg_file = await msg.photo[-1].get_file()
        mime = "image/jpeg"; ext = ".jpg"
    elif msg.document and msg.document.mime_type and msg.document.mime_type.startswith("image/"):
        tg_file = await msg.document.get_file()
        mime = msg.document.mime_type
        ext = ".png" if mime == "image/png" else ".webp" if mime == "image/webp" else ".jpg" if mime == "image/jpeg" else ".img"
    if not tg_file:
        await msg.reply_text("أرسل صورة كـ Photo أو Document رجاءً.")
        return

    await msg.chat.send_action("typing")
    bio = io.BytesIO()
    await tg_file.download_to_memory(out=bio)
    bio.seek(0)
    img_bytes = bio.read()

    # حفظ الصورة
    IMAGES_DIR.mkdir(exist_ok=True)
    ts = datetime.datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    image_file = IMAGES_DIR / f"{ts}{ext}"
    with open(image_file, "wb") as f:
        f.write(img_bytes)

    try:
        raw = gemini_ocr_plain(img_bytes, mime)
        fields = gemini_extract_structured(img_bytes, mime)
        fields = apply_custom_rules(fields)
        SESSION[update.effective_user.id] = {"raw_text": raw, "fields": fields, "image_path": str(image_file)}

        amount_val = fields.get('amount_value','')
        currency = fields.get('currency','ريال')
        amount_words = number_to_arabic_words(amount_val, currency)
        doc_type = fields.get("document_type","")
        if doc_type == "سند قبض":
            amount_line = f"- المبلغ (مدين): {amount_val} {currency}"
        elif doc_type == "سند صرف":
            amount_line = f"- المبلغ (دائن): {amount_val} {currency}"
        else:
            amount_line = f"- المبلغ: {amount_val} {currency}"

        await msg.reply_text(
            f"""الحقول المستخرجة:
- البنك/شركة الصرافة: {fields.get('bank_or_exchange','')}
- نوع المستند: {doc_type}
- التاريخ: {fields.get('date','')}
- رقم السند: {fields.get('voucher_number','')}
- المرسل: {fields.get('sender','')}
- المستفيد: {fields.get('beneficiary','')}
- البيان: {fields.get('description','')}
{amount_line}
- المبلغ كتابة: {amount_words}

(تم حفظ نسخة من الصورة في: {image_file})
هل تريد حفظها في Excel؟""",
            reply_markup=SAVE_KB
        )
    except Exception as e:
        log.exception("Gemini structured OCR failed")
        await msg.reply_text(f"حدث خطأ أثناء تحليل الصورة: {e}")

async def on_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    user_id = update.effective_user.id
    data = q.data

    if data == "discard":
        SESSION.get(user_id, {}).pop("_pending_row", None)
        SESSION.get(user_id, {}).pop("_pending_sheet", None)
        await q.edit_message_text("تم الإلغاء. لن يتم حفظ القيم.")
        return

    if data == "save_excel":
        state = SESSION.get(user_id)
        if not state:
            await q.edit_message_text("لا يوجد بيانات محفوظة في هذه الجلسة.")
            return
        f = state["fields"]

        parts = split_debit_credit(f)
        currency = f.get("currency","ريال")
        amt_value = compute_amt(parts["debit"], parts["credit"])

        row = {
            "chat_id": update.effective_chat.id,
            "user_id": user_id,
            "username": (update.effective_user.username or ""),
            "timestamp": datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
            "bank_or_exchange": f.get("bank_or_exchange",""),
            "document_type": f.get("document_type",""),
            "date": f.get("date",""),
            "voucher_number": f.get("voucher_number",""),
            "sender": f.get("sender",""),
            "beneficiary": f.get("beneficiary",""),
            "description": f.get("description",""),
            "debit": parts["debit"],
            "credit": parts["credit"],
            "amt": amt_value,
            "currency": currency,
            "raw_text": state.get("raw_text",""),
            "image_path": state.get("image_path",""),
        }

        sheet = sheet_name_for_currency(currency)

        # فحص التكرار داخل نفس ورقة العملة
        exists_in_sheet = False
        xls = Path(EXCEL_PATH)
        if xls.exists():
            try:
                df_sheet = pd.read_excel(EXCEL_PATH, sheet_name=sheet, dtype=str).fillna("")
                exists_in_sheet = not df_sheet[
                    (df_sheet["voucher_number"].astype(str) == str(row["voucher_number"]).strip()) &
                    (df_sheet["date"].astype(str) == str(row["date"]).strip())
                ].empty
            except Exception:
                exists_in_sheet = False

        if exists_in_sheet:
            SESSION[user_id]["_pending_row"] = row
            SESSION[user_id]["_pending_sheet"] = sheet
            await q.edit_message_text(
                f"⚠️ السند موجود سابقًا في ورقة «{sheet}» (رقم: {row['voucher_number']} | تاريخ: {row['date']}).\n"
                "هل تريد تحديث السند الموجود بهذه البيانات؟",
                reply_markup=InlineKeyboardMarkup([[
                    InlineKeyboardButton("✅ نعم، حدّث", callback_data="confirm_update"),
                    InlineKeyboardButton("❌ لا، تجاهل", callback_data="cancel_update"),
                ]])
            )
            return

        # لا يوجد تكرار → إضافة
        try:
            append_to_currency_sheet(row, sheet)
            await q.edit_message_text("✅ تم الحفظ في Excel ضمن ورقة العملة المناسبة.", reply_markup=None)
        except Exception as e:
            await q.edit_message_text(f"تعذر الحفظ في Excel: {e}", reply_markup=None)
        return

    if data == "cancel_update":
        SESSION.get(user_id, {}).pop("_pending_row", None)
        SESSION.get(user_id, {}).pop("_pending_sheet", None)
        await q.edit_message_text("تم الإلغاء. لم يتم تحديث السند.", reply_markup=None)
        return

    if data == "confirm_update":
        pending = SESSION.get(user_id, {}).pop("_pending_row", None)
        sheet = SESSION.get(user_id, {}).pop("_pending_sheet", None)
        if not pending or not sheet:
            await q.edit_message_text("لا يوجد تحديث معلّق.", reply_markup=None)
            return
        try:
            result = upsert_in_currency_sheet(pending, sheet)
            if result == "updated":
                await q.edit_message_text("✅ تم تحديث السند في Excel.", reply_markup=None)
            else:
                await q.edit_message_text("✅ تم الحفظ كسند جديد في Excel.", reply_markup=None)
        except Exception as e:
            await q.edit_message_text(f"تعذر التحديث: {e}", reply_markup=None)
        return

def main():
    if not BOT_TOKEN:
        raise RuntimeError("الرجاء ضبط BOT_TOKEN في .env")
    if not GEMINI_API_KEY:
        raise RuntimeError("الرجاء ضبط GEMINI_API_KEY في .env")

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("wherefile", wherefile_cmd))
    app.add_handler(CommandHandler("find", cmd_find))
    app.add_handler(CommandHandler("voucher", cmd_voucher))
    app.add_handler(CommandHandler("export", cmd_export))
    app.add_handler(CommandHandler("pdf", cmd_pdf))  # <-- الجديد
    app.add_handler(CallbackQueryHandler(on_button))
    app.add_handler(MessageHandler(filters.PHOTO | (filters.Document.IMAGE), handle_photo_or_doc))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_text))
    app.run_polling(close_loop=False)

if __name__ == "__main__":
    main()
