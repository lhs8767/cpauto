-- 쿠팡 웹앱 Supabase 권한 테이블
-- Supabase SQL Editor에 붙여 넣어 실행한다.

create table if not exists public.user_permissions (
  user_id uuid not null references auth.users(id) on delete cascade,
  permission text not null,
  created_at timestamptz not null default now(),
  primary key (user_id, permission)
);

create table if not exists public.app_files (
  file_key text primary key,
  content_base64 text not null,
  updated_at timestamptz not null default now()
);

alter table public.user_permissions enable row level security;
alter table public.app_files enable row level security;

grant select on public.user_permissions to authenticated;
grant select, insert, update, delete on public.user_permissions to service_role;
grant select, insert, update, delete on public.app_files to service_role;

drop policy if exists "deny direct anon access" on public.user_permissions;
create policy "deny direct anon access"
on public.user_permissions
for all
to anon
using (false)
with check (false);

drop policy if exists "users can read own permissions" on public.user_permissions;
create policy "users can read own permissions"
on public.user_permissions
for select
to authenticated
using (auth.uid() = user_id);

drop policy if exists "service role full access" on public.user_permissions;
create policy "service role full access"
on public.user_permissions
for all
to service_role
using (true)
with check (true);

drop policy if exists "deny direct anon access" on public.app_files;
create policy "deny direct anon access"
on public.app_files
for all
to anon
using (false)
with check (false);

drop policy if exists "service role full access" on public.app_files;
create policy "service role full access"
on public.app_files
for all
to service_role
using (true)
with check (true);

-- 관리자 계정 권한 부여
-- 먼저 Supabase Auth에서 hslee@bonie.co.kr 사용자를 만든 뒤 실행한다.
insert into public.user_permissions (user_id, permission)
select id, permission
from auth.users
cross join (
  values
    ('admin'),
    ('po_convert'),
    ('master'),
    ('sales'),
    ('check'),
    ('pallet')
) as p(permission)
where email = 'hslee@bonie.co.kr'
on conflict (user_id, permission) do nothing;
