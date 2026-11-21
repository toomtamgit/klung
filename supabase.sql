-- วางทั้งหมดนี้ใน Supabase → SQL Editor → New query → กด Run (รันครั้งเดียวพอ)

create table if not exists public.kv (
  key        text primary key,
  value      text not null,
  updated_at timestamptz not null default now()
);

alter table public.kv enable row level security;

-- เปิดสิทธิ์ให้เว็บแอป (ใช้ publishable key = role "anon") อ่าน/เขียนได้ ไม่ให้ลบ
grant select, insert, update on public.kv to anon;

drop policy if exists "kv read"   on public.kv;
drop policy if exists "kv insert" on public.kv;
drop policy if exists "kv update" on public.kv;

create policy "kv read"   on public.kv for select to anon using (true);
create policy "kv insert" on public.kv for insert to anon with check (true);
create policy "kv update" on public.kv for update to anon using (true) with check (true);
