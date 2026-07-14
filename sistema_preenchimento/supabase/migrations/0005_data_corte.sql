-- Adiciona a data de corte (DATA_CORTE da base de talhões) para calcular a tag de porte
-- (dias corridos desde o corte) em talhões 18 Meses / 2° Corte / Sem Linhas.

alter table programacao add column if not exists data_corte date;
