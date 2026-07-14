-- Status de voo (drone, projeto "Falha Soca") por talhão, usado para colorir a tag de porte
-- (dias corridos desde o corte) dos talhões "Sem Linhas" conforme a etapa no dronemgmt.

alter table programacao add column if not exists voo_control_status int;
alter table programacao add column if not exists voo_control_status_descricao text;

-- Histórico completo de voos (qualquer projeto), sincronizado do portal dronemgmt da Pedra
-- Agroindustrial — um registro por voo agendado/realizado (não por talhão); um LAYER pode ter
-- várias linhas ao longo do tempo (reagendamentos, novos projetos de voo, etc).
create table if not exists voo_status (
  id                          uuid primary key,
  layer                       bigint,
  harvest                     int,
  section                     text,
  land_plot                   text,
  flight_project_code         text,
  flight_project_descricao    text,
  control_status              int,
  control_status_descricao    text,
  start_date_flight           timestamptz,
  end_date_flight             timestamptz,
  scheduled_date              timestamptz,
  reason_descricao            text,
  company_processing_nome     text,
  image_name                  text,
  image_download              text,
  date_dispatch_processing    timestamptz,
  date_delivery_processing    timestamptz,
  date_disclosure_processing  timestamptz,
  modified_utc                timestamptz,
  synced_at                   timestamptz default now()
);

create index if not exists voo_status_layer_idx on voo_status (layer);

alter table voo_status enable row level security;

create policy "voo_status_select_anon" on voo_status
  for select to anon using (true);

create policy "voo_status_insert_anon" on voo_status
  for insert to anon with check (true);

create policy "voo_status_update_anon" on voo_status
  for update to anon using (true) with check (true);
