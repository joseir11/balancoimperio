create or replace function app_private.initialize_material_status() returns trigger language plpgsql security definer set search_path='' as $$ begin insert into public.item_status(material_id,status,teams_count) values(new.id,'AGUARDANDO',0) on conflict(material_id) do nothing;return new;end $$;
drop trigger if exists materials_initialize_status on public.materials;
create trigger materials_initialize_status after insert on public.materials for each row execute function app_private.initialize_material_status();
create index if not exists counts_team_id_idx on public.counts(team_id);
create index if not exists profiles_team_id_idx on public.profiles(team_id);
create index if not exists counts_material_id_idx on public.counts(material_id);