-- BALANÇO IMPÉRIO | Schema aplicado no Supabase em 2026-11-05
create extension if not exists pgcrypto;
create schema if not exists app_private;
revoke all on schema app_private from public, anon;
grant usage on schema app_private to authenticated;

create table public.teams(id uuid primary key default gen_random_uuid(), name text not null unique, auth_user_id uuid unique references auth.users(id) on delete cascade, active boolean not null default true, login_email text unique, created_at timestamptz not null default now());
create table public.profiles(id uuid primary key references auth.users(id) on delete cascade, role text not null default 'team' check(role in ('admin','team')), team_id uuid references public.teams(id) on delete set null, created_at timestamptz not null default now());
create table public.materials(id uuid primary key default gen_random_uuid(), code text not null unique, description text not null, unit text not null default '', active boolean not null default true, created_at timestamptz not null default now(), updated_at timestamptz not null default now());
create table public.counts(id uuid primary key default gen_random_uuid(), material_id uuid not null references public.materials(id) on delete cascade, team_id uuid not null references public.teams(id) on delete cascade, quantity numeric(14,3) not null check(quantity>=0), updated_at timestamptz not null default now(), unique(material_id,team_id));
create table public.item_status(material_id uuid primary key references public.materials(id) on delete cascade, status text not null default 'AGUARDANDO' check(status in ('AGUARDANDO','OK','DIVERGENTE')), teams_count integer not null default 0, updated_at timestamptz not null default now());

create or replace function app_private.is_admin() returns boolean language sql stable security definer set search_path='' as $$ select exists(select 1 from public.profiles p where p.id=(select auth.uid()) and p.role='admin') $$;
create or replace function app_private.current_team_id() returns uuid language sql stable security definer set search_path='' as $$ select p.team_id from public.profiles p where p.id=(select auth.uid()) $$;
create or replace function app_private.handle_new_auth_user() returns trigger language plpgsql security definer set search_path='' as $$ begin insert into public.profiles(id,role) values(new.id,'team') on conflict(id) do nothing; return new; end $$;
create trigger on_auth_user_created_profile after insert on auth.users for each row execute function app_private.handle_new_auth_user();

create or replace function app_private.refresh_item_status() returns trigger language plpgsql security definer set search_path='' as $$
declare v_material_id uuid; v_count integer; v_min numeric; v_max numeric; v_status text;
begin
 v_material_id:=coalesce(new.material_id,old.material_id);
 select count(*),min(c.quantity),max(c.quantity) into v_count,v_min,v_max from public.counts c join public.teams t on t.id=c.team_id and t.active=true where c.material_id=v_material_id;
 if v_count<2 then v_status:='AGUARDANDO'; elsif v_min=v_max then v_status:='OK'; else v_status:='DIVERGENTE'; end if;
 insert into public.item_status(material_id,status,teams_count,updated_at) values(v_material_id,v_status,v_count,now())
 on conflict(material_id) do update set status=excluded.status,teams_count=excluded.teams_count,updated_at=now();
 return coalesce(new,old);
end $$;
create trigger counts_refresh_status after insert or update or delete on public.counts for each row execute function app_private.refresh_item_status();
create or replace function app_private.initialize_material_status() returns trigger language plpgsql security definer set search_path='' as $$ begin insert into public.item_status(material_id,status,teams_count) values(new.id,'AGUARDANDO',0) on conflict(material_id) do nothing; return new; end $$;
create trigger materials_initialize_status after insert on public.materials for each row execute function app_private.initialize_material_status();

alter table public.teams enable row level security;
alter table public.profiles enable row level security;
alter table public.materials enable row level security;
alter table public.counts enable row level security;
alter table public.item_status enable row level security;
create policy profiles_read_self_or_admin on public.profiles for select to authenticated using(id=(select auth.uid()) or app_private.is_admin());
create policy teams_admin_all on public.teams for all to authenticated using(app_private.is_admin()) with check(app_private.is_admin());
create policy teams_read_own on public.teams for select to authenticated using(id=app_private.current_team_id());
create policy materials_read_authenticated on public.materials for select to authenticated using(active=true or app_private.is_admin());
create policy materials_admin_write on public.materials for all to authenticated using(app_private.is_admin()) with check(app_private.is_admin());
create policy counts_read_own_or_admin on public.counts for select to authenticated using(team_id=app_private.current_team_id() or app_private.is_admin());
create policy counts_insert_own on public.counts for insert to authenticated with check(team_id=app_private.current_team_id() and not app_private.is_admin());
create policy counts_update_own on public.counts for update to authenticated using(team_id=app_private.current_team_id() and not app_private.is_admin()) with check(team_id=app_private.current_team_id() and not app_private.is_admin());
create policy counts_delete_own on public.counts for delete to authenticated using(team_id=app_private.current_team_id() and not app_private.is_admin());
create policy counts_admin_all on public.counts for all to authenticated using(app_private.is_admin()) with check(app_private.is_admin());
create policy status_read_authenticated on public.item_status for select to authenticated using(true);
create policy status_admin_write on public.item_status for all to authenticated using(app_private.is_admin()) with check(app_private.is_admin());

grant usage on schema public to authenticated;
grant select on public.profiles,public.teams,public.materials,public.item_status to authenticated;
grant insert,update,delete on public.materials to authenticated;
grant select,insert,update,delete on public.counts to authenticated;
grant select,insert,update,delete on public.teams to authenticated;
grant select,update on public.profiles to authenticated;
grant execute on function app_private.is_admin() to authenticated;
grant execute on function app_private.current_team_id() to authenticated;
create index if not exists counts_team_id_idx on public.counts(team_id);
create index if not exists profiles_team_id_idx on public.profiles(team_id);
create index if not exists counts_material_id_idx on public.counts(material_id);
do $$ begin alter publication supabase_realtime add table public.counts; exception when duplicate_object then null; end $$;
do $$ begin alter publication supabase_realtime add table public.item_status; exception when duplicate_object then null; end $$;
