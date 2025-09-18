# -*- coding: utf-8 -*-
import os, io, logging, datetime, json, re
from typing import Dict, Any, List, Tuple

from dotenv import load_dotenv
from telegram import Update, InlineKeyboardMarkup, InlineKeyboardButton, InputFile
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, CallbackQueryHandler, filters

import pandas as pd
from pathlib import Path

# Gemini API
import google.generativeai as genai

# ====== Load env ======
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-1.5-flash")
EXCEL_PATH = os.getenv("EXCEL_PATH", "entries.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "المدخلات")

# ====== Logging ======
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("gemini-structured-bot")

# ====== Gemini setup ======
if not GEMINI_API_KEY:
    raise RuntimeError("الرجاء ضبط GEMINI_API_KEY في .env")
genai.configure(api_key=GEMINI_API_KEY)

# ====== Excel helpers ======
COLUMNS = [
    "chat_id","user_id","username","timestamp",
    "bank_or_exchange","document_type","date","voucher_number",
    "sender","beneficiary","description","amount_value","currency","raw_text"
]

def _ensure_df(df: pd.DataFrame) -> pd.DataFrame:
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[COLUMNS]

def load_entries() -> pd.DataFrame:
    xls = Path(EXCEL_PATH)
    if not xls.exists():
        return pd.DataFrame(columns=COLUMNS)
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, dtype=str)
    except ValueError:
        return pd.DataFrame(columns=COLUMNS)
    return _ensure_df(df.fillna(""))

def save_entries(df: pd.DataFrame):
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
        _ensure_df(df).to_excel(writer, index=False, sheet_name=SHEET_NAME)

def append_to_excel(row: dict):
    df = load_entries()
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    save_entries(df)

# ====== Normalization helpers ======
ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")

def normalize(s) -> str:
    """تحويل أي نوع إلى نص ثم تنظيفه وتحويل الأرقام العربية إلى الغربية."""
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
    # voucher number digits only
    m = re.search(r"(\d+)", out["voucher_number"])
    if m:
        out["voucher_number"] = m.group(1)
    # amount normalization
    amt = out["amount_value"]
    if amt:
        amt = amt.replace(",", "").replace(" ", "")
        if amt.count(".") > 1:
            amt = amt.replace(".", "")
        out["amount_value"] = re.sub(r"[^0-9\.-]", "", amt)
    # date normalization YYYY-MM-DD
    m = re.search(r"(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})", out["date"])
    if m:
        y, mo, d = m.groups()
        try:
            out["date"] = f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"
        except:
            pass
    return out

# ====== Search helpers ======
def find_rows(query: str, limit: int = 10) -> pd.DataFrame:
    df = load_entries()
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
    df = load_entries()
    if df.empty:
        return df
    voucher = re.sub(r"\D", "", voucher or "")
    return df[df["voucher_number"].astype(str) == voucher]

def filter_by_params(df: pd.DataFrame, params: Dict[str,str]) -> pd.DataFrame:
    if df.empty:
        return df
    out = df
    # date range
    if "date_range" in params and ".." in params["date_range"]:
        start, end = params["date_range"].split("..", 1)
        start = start.strip() or "0001-01-01"
        end = end.strip() or "9999-12-31"
        out = out[(out["date"] >= start) & (out["date"] <= end)]
    # optional filters
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
    """
    يدعم:
      /export 2025-08-01..2025-08-31 type=سند قبض sender=مهدي
    """
    args = text.split()[1:]  # remove command
    params = {}
    for token in args:
        if ".." in token and re.match(r"\d{4}-\d{2}-\d{2}..\d{4}-\d{2}-\d{2}", token):
            params["date_range"] = token
        elif "=" in token:
            k, v = token.split("=", 1)
            params[k.strip()] = v.strip()
    return params

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

# ====== Session ======
SESSION: Dict[int, Dict[str, Any]] = {}

SAVE_KB = InlineKeyboardMarkup([[
    InlineKeyboardButton("💾 احفظ في الإكسل", callback_data="save_excel"),
    InlineKeyboardButton("❌ لا تحفظ", callback_data="discard")
]])

# ====== Business Rules ======
def apply_custom_rules(fields: Dict[str, Any]) -> Dict[str, Any]:
    """تعديل نوع السند حسب اسم المرسل/المستفيد وفق القواعد المخصصة."""
    sender = fields.get("sender", "")
    beneficiary = fields.get("beneficiary", "")

    if "مهدي صويدر" in sender:
        fields["document_type"] = "سند صرف"
    elif ("صالح مهدي" in beneficiary) or ("مهدي صويدر" in beneficiary):
        fields["document_type"] = "سند قبض"
    return fields

# ====== Handlers ======
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "مرحبا بك في البوت المحاسبي الذكي 👋\n"
        "الأوامر: /find <كلمة-بحث> ، /voucher <رقم-سند> ، /export <تاريخ..تاريخ> [type=][sender=][beneficiary=][currency=][bank=]"
    )

async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "أرسل صورة أو نصًا لاستخراج الحقول ثم الحفظ في Excel.\n"
        "أوامر البحث والتقارير:\n"
        "• /find مهدي\n"
        "• /voucher 9779370654\n"
        "• /export 2025-08-01..2025-08-31 type=سند قبض sender=مهدي"
    )

async def wherefile_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"ملف الإكسل: {EXCEL_PATH} (الورقة: {SHEET_NAME})")

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
        lines.append(f"- {r['date']} | {r['voucher_number']} | {r['beneficiary']} | {r['amount_value']} {r['currency']}")
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
    await msg.reply_text(
        f"""تفاصيل السند:
- البنك/شركة الصرافة: {r['bank_or_exchange']}
- نوع المستند: {r['document_type']}
- التاريخ: {r['date']}
- رقم السند: {r['voucher_number']}
- المرسل: {r['sender']}
- المستفيد: {r['beneficiary']}
- البيان: {r['description']}
- المبلغ: {r['amount_value']} {r['currency']}
"""
    )

async def cmd_export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    params = parse_export_args(msg.text)
    df = load_entries()
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
    SESSION[update.effective_user.id] = {"raw_text": txt, "fields": fields}
    await msg.reply_text(
        f"""النص المستخرج (ملخص):
{txt}

---------------------
الحقول المستخرجة:
- البنك/شركة الصرافة: {fields['bank_or_exchange']}
- نوع المستند: {fields['document_type']}
- التاريخ: {fields['date']}
- رقم السند: {fields['voucher_number']}
- المرسل: {fields['sender']}
- المستفيد: {fields['beneficiary']}
- البيان: {fields['description']}
- المبلغ: {fields['amount_value']} {fields['currency']}

هل تريد حفظها في Excel؟""",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("💾 احفظ في الإكسل", callback_data="save_excel"),
            InlineKeyboardButton("❌ لا تحفظ", callback_data="discard")
        ]])
    )

async def handle_photo_or_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message or update.edited_message
    if not msg:
        return
    tg_file = None
    mime = "image/jpeg"
    if msg.photo:
        tg_file = await msg.photo[-1].get_file()
        mime = "image/jpeg"
    elif msg.document and msg.document.mime_type and msg.document.mime_type.startswith("image/"):
        tg_file = await msg.document.get_file()
        mime = msg.document.mime_type
    if not tg_file:
        await msg.reply_text("أرسل صورة كـ Photo أو Document رجاءً.")
        return
    await msg.chat.send_action("typing")
    bio = io.BytesIO()
    await tg_file.download_to_memory(out=bio)
    bio.seek(0)
    img_bytes = bio.read()
    try:
        raw = gemini_ocr_plain(img_bytes, mime)
        fields = gemini_extract_structured(img_bytes, mime)
        fields = apply_custom_rules(fields)
        SESSION[update.effective_user.id] = {"raw_text": raw, "fields": fields}
        preview = raw[:1500] if raw else "(لا يوجد نص)"
        await msg.reply_text(
            f"""النص المستخرج (ملخص):
{preview}

---------------------
الحقول المستخرجة:
- البنك/شركة الصرافة: {fields.get('bank_or_exchange','')}
- نوع المستند: {fields.get('document_type','')}
- التاريخ: {fields.get('date','')}
- رقم السند: {fields.get('voucher_number','')}
- المرسل: {fields.get('sender','')}
- المستفيد: {fields.get('beneficiary','')}
- البيان: {fields.get('description','')}
- المبلغ: {fields.get('amount_value','')} {fields.get('currency','')}

هل تريد حفظها في Excel؟""",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("💾 احفظ في الإكسل", callback_data="save_excel"),
                InlineKeyboardButton("❌ لا تحفظ", callback_data="discard")
            ]])
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
        await q.edit_message_text("تم الإلغاء. لن يتم حفظ القيم.")
        return
    if data == "save_excel":
        state = SESSION.get(user_id)
        if not state:
            await q.edit_message_text("لا يوجد بيانات محفوظة في هذه الجلسة.")
            return
        f = state["fields"]
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
            "amount_value": f.get("amount_value",""),
            "currency": f.get("currency",""),
            "raw_text": state.get("raw_text",""),
        }
        try:
            append_to_excel(row)
            await q.edit_message_text("✅ تم الحفظ في ملف Excel بالأعمدة المنظّمة.")
        except Exception as e:
            await q.edit_message_text(f"تعذر الحفظ في Excel: {e}")

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
    # Arabic aliases (optional)
    app.add_handler(CommandHandler("بحث", cmd_find))
    app.add_handler(CommandHandler("سند", cmd_voucher))
    app.add_handler(CommandHandler("تقرير", cmd_export))

    app.add_handler(CallbackQueryHandler(on_button))
    app.add_handler(MessageHandler(filters.PHOTO | (filters.Document.IMAGE), handle_photo_or_doc))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_text))
    app.run_polling(close_loop=False)

if __name__ == "__main__":
    main()
