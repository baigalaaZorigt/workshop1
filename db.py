"""Supabase холболт. Хаяг, түлхүүрийг .env файлаас уншина."""
import os
from functools import lru_cache

from dotenv import load_dotenv
from supabase import create_client

load_dotenv()


def is_configured():
    url = os.getenv("SUPABASE_URL", "")
    key = os.getenv("SUPABASE_SECRET_KEY", "")
    return url.startswith("https://") and bool(key) and "ЭНД" not in key


@lru_cache(maxsize=1)
def sb():
    if not is_configured():
        raise RuntimeError(".env файлд SUPABASE_URL болон SUPABASE_SECRET_KEY-г бөглөнө үү.")
    return create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_SECRET_KEY"])


def fetch_all(table, order=None, page_size=1000):
    """Supabase нэг удаад 1000 мөр л буцаадаг тул хуудаслаж бүгдийг авна."""
    rows, start = [], 0
    while True:
        query = sb().table(table).select("*")
        if order:
            query = query.order(order)
        batch = query.range(start, start + page_size - 1).execute().data
        rows.extend(batch)
        if len(batch) < page_size:
            return rows
        start += page_size
