-- =====================================================================
-- Cashmere Mongolia — Supabase хүснэгтүүд
-- Supabase → SQL Editor дээр бүтнээр нь хуулж ажиллуулна.
-- Өгөгдөл энд оруулахгүй: zahialga.xlsx, dansnii_huulga.xlsx-ийг апп-аар upload хийнэ.
-- =====================================================================

-- bcrypt шифрлэлтэд хэрэгтэй (Supabase дээр extensions схемд байдаг)
create extension if not exists pgcrypto with schema extensions;


-- ---------------------------------------------------------------------
-- 1. orders — zahialga.xlsx
-- Excel багана: order_id, date, channel, customer, district, product,
--               quantity, unit_price, delivery_fee, total, payment_method, status
-- ---------------------------------------------------------------------
create table if not exists public.orders (
  order_id        text primary key,                 -- ZH-001
  order_date      date not null,                    -- Excel-ийн "date"
  channel         text,                             -- Facebook, Instagram, Танил
  customer        text,                             -- Ганаа
  district        text,                             -- Сүхбаатар
  product         text,                             -- Бээлий, Ороолт, Цамц ...
  quantity        integer not null default 1 check (quantity > 0),
  unit_price      bigint  not null default 0 check (unit_price >= 0),   -- ₮
  delivery_fee    bigint  not null default 0 check (delivery_fee >= 0), -- ₮
  total           bigint  not null default 0 check (total >= 0),        -- ₮
  payment_method  text,                             -- QPay, Данс, Бэлэн
  status          text,                             -- Хүргэгдсэн, Цуцлагдсан, Төлбөр хүлээгдэж буй
  created_at      timestamptz not null default now()
);

create index if not exists orders_order_date_idx on public.orders (order_date);


-- ---------------------------------------------------------------------
-- 2. bank_transactions — dansnii_huulga.xlsx
-- Excel багана: ОГНОО (DATE), ВАЛЮТ (IN FC), ДҮН (MNT), ТӨРӨЛ (TYPE),
--               ГҮЙЛГЭЭНИЙ УТГА (DETAILS OF PAYMENT)
-- Огноо нь тусдаа мөрөнд бичигддэг тул апп upload хийхдээ мөр бүрт огноог нь онооно.
-- "ЭХНИЙ ҮЛДЭГДЭЛ", "ӨДРИЙН ОРЛОГО/ЗАРЛАГА/ҮЛДЭГДЭЛ", "НИЙТ ..." мөрүүдийг
-- оруулахгүй — зөвхөн ТӨРӨЛ нь бөглөгдсөн жинхэнэ гүйлгээг хадгална.
-- ---------------------------------------------------------------------
create table if not exists public.bank_transactions (
  id            bigint generated always as identity primary key,
  txn_date      date   not null,                    -- ОГНОО
  amount        bigint not null check (amount >= 0),-- ДҮН (MNT), үргэлж эерэг
  fc_amount     numeric(14,2),                      -- ВАЛЮТ (IN FC), ихэвчлэн хоосон
  type          text   not null check (type in ('ОРЛОГО', 'ЗАРЛАГА')), -- ТӨРӨЛ
  description   text,                               -- ГҮЙЛГЭЭНИЙ УТГА
  category      text,                               -- Борлуулалт, Бараа татан авалт, Хүргэлт, Сурталчилгаа, Савлагаа, Шимтгэл, Буцаалт, Хувийн, Хадгаламж
  order_id      text,                               -- Утгаас олдсон захиалгын дугаар (ZH-001), олдохгүй бол хоосон
  account_no    text,                               -- Хуулгын толгой дахь дансны дугаар
  created_at    timestamptz not null default now()
);

create index if not exists bank_transactions_txn_date_idx on public.bank_transactions (txn_date);
create index if not exists bank_transactions_order_id_idx on public.bank_transactions (order_id);


-- ---------------------------------------------------------------------
-- 3. users — апп руу нэвтрэх хэрэглэгч
-- ---------------------------------------------------------------------
create table if not exists public.users (
  id             bigint generated always as identity primary key,
  email          text not null unique,
  password_hash  text not null,                     -- bcrypt hash, жинхэнэ нууц үг хадгалагдахгүй
  created_at     timestamptz not null default now()
);

-- Анхны хэрэглэгч. 'ЭНД_НУУЦ_ҮГЭЭ_БИЧ'-г өөрийн нууц үгээр солиод ажиллуулна.
-- gen_salt('bf') = bcrypt.
insert into public.users (email, password_hash)
values (
  'ganzorig.aiaa@gmail.com',
  extensions.crypt('ЭНД_НУУЦ_ҮГЭЭ_БИЧ', extensions.gen_salt('bf'))
)
on conflict (email) do nothing;


-- ---------------------------------------------------------------------
-- 4. RLS (Row Level Security)
-- Policy нэмээгүй тул anon / нэвтэрсэн түлхүүрээр юу ч уншиж, бичиж чадахгүй.
-- Апп зөвхөн сервер талаас service_role түлхүүрээр хандана.
-- ---------------------------------------------------------------------
alter table public.orders            enable row level security;
alter table public.bank_transactions enable row level security;
alter table public.users             enable row level security;
