"""Excel файлуудыг уншиж Supabase-д хадгалах.

Хоёр төрлийн файл таньдаг:
  - Захиалгын файл (zahialga.xlsx): эхний мөрөнд order_id, date, ... гэсэн гарчиг
  - Дансны хуулга (dansnii_huulga.xlsx): "ДАНСНЫ ХУУЛГА" гэж эхэлдэг

Өмнө орсон мөрийг дахин оруулахгүй. Шинэ гүйлгээг OpenAI-аар ангилна.
"""
import json
import os
import re
from collections import Counter
from datetime import date, datetime, timedelta

import openai
import openpyxl

from db import fetch_all, sb

ORDER_COLUMNS = [
    "order_id", "date", "channel", "customer", "district", "product",
    "quantity", "unit_price", "delivery_fee", "total", "payment_method", "status",
]
ORDER_ID_RE = re.compile(r"ZH-\d+", re.IGNORECASE)
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

CATEGORIES = [
    "Борлуулалт", "Бараа татан авалт", "Хүргэлт", "Сурталчилгаа", "Савлагаа",
    "Шимтгэл", "Буцаалт", "Хувийн", "Хадгаламж",
]

# OpenAI ажиллахгүй үед (түлхүүргүй, интернэтгүй) ашиглах энгийн дүрэм.
# Гүйлгээний утгад тухайн үг байвал энэ ангилалд орно (эхний таарсан нь).
FALLBACK_RULES = [
    ("шимтгэл", "Шимтгэл"),
    ("FACEBK", "Сурталчилгаа"),
    ("НИЙЛҮҮЛЭГЧ", "Бараа татан авалт"),
    ("ХҮРГЭЛТ", "Хүргэлт"),
    ("TAXI", "Хүргэлт"),
    ("SHTS", "Хүргэлт"),
    ("PACKSHOP", "Савлагаа"),
    ("Өөрийн данс", "Хадгаламж"),
]


def _to_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):  # Excel-ийн серийн дугаар
        return date(1899, 12, 30) + timedelta(days=int(value))
    if isinstance(value, str) and DATE_RE.match(value.strip()):
        return date.fromisoformat(value.strip())
    raise ValueError(f"Огноо танигдсангүй: {value!r}")


def _to_int(value):
    if value in (None, ""):
        return 0
    return int(round(float(str(value).replace(",", ""))))


def import_excel(file_storage):
    """Upload хийсэн файлыг таньж зохих хүснэгтэд хадгална. Үр дүнгийн тайлбарыг буцаана."""
    wb = openpyxl.load_workbook(file_storage, read_only=True, data_only=True)
    rows = list(wb.worksheets[0].iter_rows(values_only=True))
    if not rows:
        raise ValueError("Файл хоосон байна.")

    header = [str(c).strip().lower() if c is not None else "" for c in rows[0]]
    if header[: len(ORDER_COLUMNS)] == ORDER_COLUMNS:
        new, skipped = _import_orders(rows[1:])
        linked = link_payments()
        return (f"Захиалга: {new} шинэ, {skipped} өмнө орсон тул алгассан. "
                f"{linked} төлбөрийг захиалгатай нь холболоо.")

    if rows[0][0] and "ДАНСНЫ ХУУЛГА" in str(rows[0][0]).upper():
        new, skipped, how = _import_bank(rows)
        linked = link_payments()
        return (f"Гүйлгээ: {new} шинэ ({how}), {skipped} өмнө орсон тул алгассан. "
                f"{linked} төлбөрийг захиалгатай нь холболоо.")

    raise ValueError("Файл танигдсангүй. Захиалгын эсвэл дансны хуулгын Excel оруулна уу.")


def _import_orders(rows):
    existing = {o["order_id"] for o in fetch_all("orders")}
    records, skipped = [], 0
    for row in rows:
        if not row or not row[0]:
            continue
        r = dict(zip(ORDER_COLUMNS, row))
        order_id = str(r["order_id"]).strip().upper()
        if order_id in existing:
            skipped += 1
            continue
        existing.add(order_id)
        records.append({
            "order_id": order_id,
            "order_date": _to_date(r["date"]).isoformat(),
            "channel": r["channel"],
            "customer": r["customer"],
            "district": r["district"],
            "product": r["product"],
            "quantity": _to_int(r["quantity"]) or 1,
            "unit_price": _to_int(r["unit_price"]),
            "delivery_fee": _to_int(r["delivery_fee"]),
            "total": _to_int(r["total"]),
            "payment_method": r["payment_method"],
            "status": r["status"],
        })
    if records:
        sb().table("orders").insert(records).execute()
    return len(records), skipped


def _parse_bank(rows):
    """Хуулгаас зөвхөн жинхэнэ гүйлгээний мөрүүдийг авна."""
    account_no = None
    current_date = None
    records = []
    for row in rows:
        cells = list(row) + [None] * (5 - len(row))
        first, fc, amount, txn_type, details = cells[:5]
        if first == "Данс:" and fc:
            account_no = str(fc).split()[0]
            continue
        # Огнооны мөр: дараагийн гүйлгээнүүдэд энэ огноог онооно.
        if isinstance(first, (datetime, date)) or (isinstance(first, str) and DATE_RE.match(first.strip())):
            current_date = _to_date(first)
            continue
        txn_type = (txn_type or "").strip().upper()
        # ТӨРӨЛ хоосон мөр = эхний/өдрийн/эцсийн үлдэгдэл, өдрийн орлого/зарлагын нийлбэр → алгасна.
        if txn_type not in ("ОРЛОГО", "ЗАРЛАГА") or current_date is None:
            continue
        description = re.sub(r"\s*\(Ханш:[^)]*\)\s*$", "", str(details or "")).strip()
        match = ORDER_ID_RE.search(description)
        records.append({
            "txn_date": current_date.isoformat(),
            "amount": _to_int(amount),
            "fc_amount": float(fc) if isinstance(fc, (int, float)) else None,
            "type": txn_type,
            "description": description,
            "order_id": match.group(0).upper() if match and txn_type == "ОРЛОГО" else None,
            "account_no": account_no,
        })
    return records


def _txn_key(t):
    return (t["txn_date"], t["type"], int(t["amount"]), t["description"] or "")


def _import_bank(rows):
    records = _parse_bank(rows)
    if not records:
        raise ValueError("Хуулгаас гүйлгээ олдсонгүй.")

    # Нэг өдөр ижил дүн, ижил утгатай гүйлгээ (жишээ нь 100₮ шимтгэл) хэд хэд байж болно.
    # Тиймээс "аль хэдийн хэдэн ширхэг байна" гэж тоолж, илүү гарсныг нь л шинэд тооцно.
    already = Counter(_txn_key(t) for t in fetch_all("bank_transactions"))
    new = []
    for r in records:
        key = _txn_key(r)
        if already[key] > 0:
            already[key] -= 1
        else:
            new.append(r)

    how = "ангилах шаардлагагүй"
    if new:
        how = categorize(new)
        sb().table("bank_transactions").insert(new).execute()
    return len(new), len(records) - len(new), how


def categorize(txns):
    """Гүйлгээ бүрт "category" онооно. OpenAI ажиллахгүй бол энгийн дүрмээр ангилна."""
    customers = sorted({o["customer"] for o in fetch_all("orders") if o["customer"]})
    try:
        categories = _categorize_openai(txns, customers)
        how = "OpenAI ангилсан"
    except Exception as e:
        categories = {}
        how = f"OpenAI ажилласангүй ({_short_error(e)}), энгийн дүрмээр ангилсан"
    for i, t in enumerate(txns):
        t["category"] = categories.get(i) or _fallback_category(t, customers)
    return how


def _categorize_openai(txns, customers):
    if not os.getenv("OPENAI_API_KEY"):
        raise RuntimeError(".env-д OPENAI_API_KEY алга")
    client = openai.OpenAI()
    lines = [f'{i} | {t["type"]} | {t["amount"]} | {t["description"]}' for i, t in enumerate(txns)]
    response = client.chat.completions.create(
        model=os.getenv("OPENAI_MODEL") or "gpt-5-mini",
        messages=[
            {"role": "system", "content": (
                "Чи Cashmere Mongolia гэдэг ноолууран бүтээгдэхүүний онлайн дэлгүүрийн банкны гүйлгээг ангилна. "
                "Ангилал тус бүрийн утга:\n"
                "- Борлуулалт: харилцагчаас орж ирсэн захиалгын төлбөр (ZH-дугаар, QPAY, харилцагчийн нэр, TOLBOR, ZAHIALGA)\n"
                "- Бараа татан авалт: нийлүүлэгчид төлсөн мөнгө\n"
                "- Хүргэлт: хүргэлтийн компани, такси, шатахуун (SHTS)\n"
                "- Сурталчилгаа: Facebook/Instagram зар (FACEBK ADS)\n"
                "- Савлагаа: хайрцаг, уут, савлагааны дэлгүүр (PACKSHOP)\n"
                "- Шимтгэл: банкны гүйлгээний шимтгэл\n"
                "- Буцаалт: харилцагчид буцааж өгсөн мөнгө (ЗАРЛАГА бөгөөд хүлээн авагч нь харилцагч)\n"
                "- Хувийн: кафе, супермаркет, гэр бүл, СӨХ, утасны төлбөр гэх мэт бизнесийн бус зардал\n"
                "- Хадгаламж: өөрийн данс хооронд шилжүүлсэн\n"
                f"Дэлгүүрийн харилцагчдын нэрс: {', '.join(customers) or 'тодорхойгүй'}.\n"
                "Мөр бүрийн дугаар (idx)-аар нь ангиллыг буцаа."
            )},
            {"role": "user", "content": "idx | төрөл | дүн | утга\n" + "\n".join(lines)},
        ],
        response_format={
            "type": "json_schema",
            "json_schema": {
                "name": "transaction_categories",
                "strict": True,
                "schema": {
                    "type": "object",
                    "properties": {
                        "items": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "idx": {"type": "integer"},
                                    "category": {"type": "string", "enum": CATEGORIES},
                                },
                                "required": ["idx", "category"],
                                "additionalProperties": False,
                            },
                        },
                    },
                    "required": ["items"],
                    "additionalProperties": False,
                },
            },
        },
    )
    data = json.loads(response.choices[0].message.content)
    return {it["idx"]: it["category"] for it in data["items"] if it["category"] in CATEGORIES}


def _fallback_category(txn, customers):
    if txn["type"] == "ОРЛОГО":
        return "Борлуулалт"
    text = (txn["description"] or "").upper()
    for keyword, category in FALLBACK_RULES:
        if keyword.upper() in text:
            return category
    if any(c.upper() in text for c in customers):
        return "Буцаалт"
    return "Хувийн"


def _short_error(e):
    if isinstance(e, openai.AuthenticationError):
        return "түлхүүр буруу"
    if isinstance(e, openai.RateLimitError):
        return "хязгаар хэтэрсэн эсвэл төлбөр дууссан"
    if isinstance(e, openai.APIConnectionError):
        return "интернэт холболт"
    if isinstance(e, openai.APIStatusError):
        return f"алдаа {e.status_code}"
    return str(e)[:80]


def link_payments():
    """ZH-дугааргүй орлогыг дүн + харилцагчийн нэрээр нь захиалгатай тааруулна."""
    orders = fetch_all("orders")
    txns = fetch_all("bank_transactions")
    taken = {t["order_id"] for t in txns if t["order_id"]}
    linked = 0
    for t in sorted(txns, key=lambda t: t["txn_date"]):
        if t["order_id"] or t["type"] != "ОРЛОГО":
            continue
        desc = (t["description"] or "").upper()
        candidates = [
            o for o in orders
            if o["order_id"] not in taken
            and o["total"] == t["amount"]
            and o["customer"] and o["customer"].upper() in desc
            and o["order_date"] <= t["txn_date"]
        ]
        if not candidates:
            continue
        order = min(candidates, key=lambda o: o["order_date"])
        sb().table("bank_transactions").update({"order_id": order["order_id"]}).eq("id", t["id"]).execute()
        taken.add(order["order_id"])
        linked += 1
    return linked
