-- Permite upsert (INSERT ... ON CONFLICT DO UPDATE) de novas linhas em programacao
-- pela publishable key (anon), usado pela atualização de base direto pelo navegador.

create policy "programacao_insert_anon" on programacao
  for insert to anon with check (true);
