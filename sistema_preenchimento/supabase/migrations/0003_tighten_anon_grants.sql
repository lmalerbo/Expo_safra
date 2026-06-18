-- Restringe o que a chave pública (anon) pode escrever, sem quebrar os fluxos
-- legítimos do app (consolidar() e a atualização de base pelo navegador).
--
-- Nota honesta: como o app usa uma chave pública embutida no HTML (sem login
-- real de usuário), não é possível restringir por "quem" está escrevendo —
-- qualquer pessoa com a chave ainda pode chamar os mesmos endpoints que o
-- app legítimo chama. O que esta migration fecha são os write paths que o
-- app NUNCA usa hoje, reduzindo o que dá para fazer de errado com a chave.

-- usuarios: o único PATCH que o app faz é { ha }, dentro de consolidar().
-- "perfis" controla quem tem acesso a preenchimento/dashboard no app — hoje
-- using(true) with check(true) deixa qualquer um se autopromover editando
-- perfis direto pela API. Restringe a coluna gravável a "ha" apenas.
revoke update on usuarios from anon;
grant update (ha) on usuarios to anon;

-- log_exportacoes: nenhuma FK existia entre layer/usuario e as tabelas que
-- eles referenciam. Auditoria confirmou 0 violações nos dados atuais —
-- adiciona a constraint para impedir inserts apontando para layer ou
-- usuário inexistentes (a app nunca precisa fazer isso legitimamente).
alter table log_exportacoes
  add constraint log_exportacoes_layer_fkey
    foreign key (layer) references programacao(layer),
  add constraint log_exportacoes_usuario_fkey
    foreign key (usuario) references usuarios(nome);
