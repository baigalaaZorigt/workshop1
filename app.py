"""Cashmere Mongolia — захиалга, дансны хуулга, тайлан, AI туслах.

Ажиллуулах:  .venv/bin/python app.py   →  http://localhost:5050
"""
import os
from collections import defaultdict
from functools import wraps

import anthropic
import bcrypt
import openai
from flask import Flask, flash, redirect, render_template, request, session, url_for

from db import fetch_all, is_configured, sb
from excel_import import import_excel

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY") or os.urandom(32)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10MB

PAID_METHODS = ("Данс", "QPay")


@app.template_filter("mnt")
def mnt(value):
    return f"{int(value or 0):,}".replace(",", " ") + " ₮"


@app.context_processor
def inject_globals():
    return {"user_email": session.get("email"), "configured": is_configured()}


def login_required(view):
    @wraps(view)
    def wrapper(*args, **kwargs):
        if not session.get("email"):
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapper


# ---------------------------------------------------------------- Нэвтрэх

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        try:
            found = sb().table("users").select("email, password_hash").eq("email", email).execute().data
        except Exception as e:
            flash(f"Өгөгдлийн сантай холбогдож чадсангүй: {e}", "error")
            return render_template("login.html", email=email)
        user = found[0] if found else None
        if user and bcrypt.checkpw(password.encode(), user["password_hash"].encode()):
            session.clear()
            session["email"] = user["email"]
            target = request.args.get("next", "")
            return redirect(target if target.startswith("/") and not target.startswith("//") else url_for("index"))
        flash("И-мэйл эсвэл нууц үг буруу байна.", "error")
        return render_template("login.html", email=email)
    return render_template("login.html")


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------------------------------------------------------------- Нүүр

NON_BUSINESS = ("Хувийн", "Хадгаламж")  # бизнесийн зарлагад тооцохгүй


def short_mnt(value):
    """Графикийн дээрх богино тоо: 2 450 000 → 2.45 сая, 35 000 → 35 мян."""
    if value >= 1_000_000:
        return f"{value / 1_000_000:.2f}".rstrip("0").rstrip(".") + " сая"
    if value >= 1_000:
        return f"{value / 1_000:.0f} мян"
    return f"{value} ₮"


def month_label(month):
    year, m = month.split("-")
    return f"{year} оны {int(m)}-р сар"


def month_summary(txns, month):
    in_month = [t for t in txns if t["txn_date"].startswith(month)]
    income = sum(t["amount"] for t in in_month if t["type"] == "ОРЛОГО")
    business = [t for t in in_month if t["type"] == "ЗАРЛАГА" and t["category"] not in NON_BUSINESS]
    expense = sum(t["amount"] for t in business)

    by_category = defaultdict(int)
    for t in business:
        by_category[t["category"] or "Ангилаагүй"] += t["amount"]
    top = max(by_category.values(), default=0)
    columns = [
        {"name": name, "amount": amount, "short": short_mnt(amount),
         "height": round(amount / top * 100, 1) if top else 0}
        for name, amount in sorted(by_category.items(), key=lambda kv: kv[1], reverse=True)
    ]
    excluded = sum(t["amount"] for t in in_month if t["type"] == "ЗАРЛАГА" and t["category"] in NON_BUSINESS)
    return {"income": income, "expense": expense, "profit": income - expense,
            "columns": columns, "excluded": excluded}


@app.route("/")
@login_required
def index():
    txns = fetch_all("bank_transactions")
    months = sorted({t["txn_date"][:7] for t in txns}, reverse=True)
    month = request.args.get("month")
    if month not in months:
        month = months[0] if months else None
    return render_template(
        "home.html",
        months=[(m, month_label(m)) for m in months],
        month=month,
        s=month_summary(txns, month) if month else None,
    )


# ---------------------------------------------------------------- Excel оруулах

@app.route("/upload", methods=["GET", "POST"])
@login_required
def upload():
    if request.method == "POST":
        file = request.files.get("file")
        if not file or not file.filename.lower().endswith(".xlsx"):
            flash(".xlsx өргөтгөлтэй Excel файл сонгоно уу.", "error")
        else:
            try:
                flash(import_excel(file), "success")
            except Exception as e:
                flash(f"{file.filename}: {e}", "error")
        return redirect(url_for("upload"))
    return render_template("upload.html")


# ---------------------------------------------------------------- Захиалга

@app.route("/orders")
@login_required
def orders():
    rows = fetch_all("orders", order="order_date")
    rows.reverse()  # шинэ нь дээрээ
    paid = {t["order_id"] for t in fetch_all("bank_transactions") if t["order_id"]}

    status = request.args.get("status", "")
    q = request.args.get("q", "").strip().lower()
    statuses = sorted({r["status"] for r in rows if r["status"]})
    if status:
        rows = [r for r in rows if r["status"] == status]
    if q:
        rows = [r for r in rows if q in " ".join(str(r[k] or "") for k in
                ("order_id", "customer", "product", "district", "channel")).lower()]
    for r in rows:
        r["paid_in_bank"] = r["order_id"] in paid
    return render_template("orders.html", rows=rows, statuses=statuses, status=status, q=q,
                           total=sum(r["total"] for r in rows))


# ---------------------------------------------------------------- Тайлан

def build_report():
    orders = fetch_all("orders")
    txns = fetch_all("bank_transactions")
    active = [o for o in orders if o["status"] != "Цуцлагдсан"]
    paid = {t["order_id"] for t in txns if t["order_id"]}

    def group(items, key, value):
        sums = defaultdict(int)
        for it in items:
            sums[it[key] or "Тодорхойгүй"] += it[value]
        return sorted(sums.items(), key=lambda kv: kv[1], reverse=True)

    income = sum(t["amount"] for t in txns if t["type"] == "ОРЛОГО")
    expense = sum(t["amount"] for t in txns if t["type"] == "ЗАРЛАГА")
    return {
        "sales": sum(o["total"] for o in active),
        "order_count": len(active),
        "cancelled": len(orders) - len(active),
        "income": income,
        "expense": expense,
        "net": income - expense,
        "by_product": group(active, "product", "total"),
        "by_channel": group(active, "channel", "total"),
        "by_category": group([t for t in txns if t["type"] == "ЗАРЛАГА"], "category", "amount"),
        "unpaid": sorted(
            (o for o in active if o["payment_method"] in PAID_METHODS and o["order_id"] not in paid),
            key=lambda o: o["order_date"],
        ),
        "has_data": bool(orders or txns),
    }


@app.route("/report")
@login_required
def report():
    return render_template("report.html", r=build_report())


# ---------------------------------------------------------------- AI

AI_SYSTEM = """Чи Cashmere Mongolia гэдэг ноолууран бүтээгдэхүүний жижиг онлайн дэлгүүрийн санхүү, борлуулалтын туслах.
Эзэнд нь захиалга болон дансны хуулгын өгөгдөл дээр тулгуурлан монгол хэлээр, энгийн ойлгомжтой хариулна.
Тоог мөнгөн дүнгээр (₮) зөв тооцоол. Өгөгдөлд байхгүй зүйлийг таамаглахгүй, мэдэхгүй бол шууд хэл.
Хариултаа товч байлга. Markdown тэмдэглэгээ (**, #, хүснэгт) бүү ашигла — энгийн текст, шаардлагатай бол "- " гэсэн жагсаалт хэрэглэ."""


def data_as_text():
    orders = fetch_all("orders", order="order_date")
    txns = fetch_all("bank_transactions", order="txn_date")
    lines = ["ЗАХИАЛГА (order_id | огноо | суваг | харилцагч | дүүрэг | бараа | тоо | нэгж үнэ | хүргэлт | нийт | төлбөр | төлөв)"]
    lines += [" | ".join(str(o[k]) for k in ("order_id", "order_date", "channel", "customer", "district", "product",
                                            "quantity", "unit_price", "delivery_fee", "total", "payment_method", "status"))
              for o in orders]
    lines.append("\nДАНСНЫ ГҮЙЛГЭЭ (огноо | төрөл | дүн | ангилал | холбогдох захиалга | утга)")
    lines += [" | ".join(str(t[k] or "") for k in ("txn_date", "type", "amount", "category", "order_id", "description"))
              for t in txns]
    return "\n".join(lines)


@app.route("/ai", methods=["GET", "POST"])
@login_required
def ai():
    question, answer = "", None
    if request.method == "POST":
        question = request.form.get("question", "").strip()
        # ANTHROPIC_API_KEY байвал Claude, үгүй бол OPENAI_API_KEY-ээр OpenAI ашиглана.
        if os.getenv("ANTHROPIC_API_KEY", "").startswith("sk-"):
            ask, errors = ask_claude, anthropic
        elif os.getenv("OPENAI_API_KEY", "").startswith("sk-"):
            ask, errors = ask_openai, openai
        else:
            ask = None
            flash(".env файлд ANTHROPIC_API_KEY эсвэл OPENAI_API_KEY-г бөглөнө үү.", "error")
        if ask and question:
            try:
                answer = ask(question)
            except errors.AuthenticationError:
                flash("AI түлхүүр буруу байна. .env файлаа шалгана уу.", "error")
            except errors.RateLimitError:
                flash("AI түр завгүй эсвэл төлбөрийн хязгаар дууссан байна. Хэсэг хүлээгээд дахин оролдоно уу.", "error")
            except errors.APIStatusError as e:
                flash(f"AI алдаа ({e.status_code}): {e.message}", "error")
            except errors.APIConnectionError:
                flash("AI руу холбогдож чадсангүй. Интернэтээ шалгана уу.", "error")
    return render_template("ai.html", question=question, answer=answer)


def ask_claude(question):
    client = anthropic.Anthropic()
    response = client.beta.messages.create(
        model="claude-opus-5-5",
        max_tokens=16000,
        output_config={"effort": "medium"},
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
        system=AI_SYSTEM,
        messages=[{
            "role": "user",
            "content": f"<data>\n{data_as_text()}\n</data>\n\nАсуулт: {question}",
        }],
    )
    if response.stop_reason == "refusal":
        return "Уучлаарай, энэ асуултад хариулж чадахгүй байна."
    return "".join(b.text for b in response.content if b.type == "text").strip()


def ask_openai(question):
    client = openai.OpenAI()
    response = client.chat.completions.create(
        model=os.getenv("OPENAI_MODEL") or "gpt-5-mini",
        messages=[
            {"role": "system", "content": AI_SYSTEM},
            {"role": "user", "content": f"<data>\n{data_as_text()}\n</data>\n\nАсуулт: {question}"},
        ],
    )
    message = response.choices[0].message
    if message.refusal:
        return "Уучлаарай, энэ асуултад хариулж чадахгүй байна."
    return (message.content or "").strip()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 5050)), debug=False)
