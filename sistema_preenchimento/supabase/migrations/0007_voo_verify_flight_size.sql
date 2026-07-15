-- Campo separado do dronemgmt (flight-consult), distinto de control_status: verifyFlightSize é
-- o valor por trás do badge "Verificar porte" mostrado na tela do dronemgmt. Confirmado que o
-- valor 5 corresponde ao rótulo "Voo liberado" (única condição para considerar o voo agendado).

alter table programacao add column if not exists voo_verify_flight_size int;
alter table voo_status   add column if not exists verify_flight_size int;
