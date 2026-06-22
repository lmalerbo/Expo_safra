-- Rastro de reenvio de arquivos de projeto (.dwg/.zip) feito fora do fluxo
-- registrar→consolidar, pela tela "⚙ Gerenciar → Atualizar Projetos". Não
-- altera programacao/log_exportacoes — é só um log de upload (quem, quando,
-- quais fazendas/arquivos), normal ou em bloco (Projeto Personalizado).

create table if not exists reenvios_projeto (
  id            bigserial primary key,
  data          timestamptz default now(),
  usuario       text,
  personalizado boolean default false,
  cod_fazs      text,
  fazendas      text,
  arquivos      text
);

alter table reenvios_projeto enable row level security;

create policy "reenvios_projeto_select_anon" on reenvios_projeto
  for select to anon using (true);

create policy "reenvios_projeto_insert_anon" on reenvios_projeto
  for insert to anon with check (true);
