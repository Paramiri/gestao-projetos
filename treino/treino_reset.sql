-- =====================================================================
-- UNIALFA - Reset do ambiente de TREINAMENTO
-- Projeto Supabase: treinamento (ref uuxvdulunrwppbmofyux)
-- =====================================================================
-- Apaga TODOS os dados de projetos/formularios do ambiente de treino
-- (para rodar entre turmas) e reaplica o seed de 7 projetos de exemplo
-- em seguida (treino_seed.sql).
--
-- NAO apaga: as 6 contas de treino (auth.users/perfis), nem
-- public.configuracoes (toggles/papeis). Essas permanecem fixas entre
-- turmas.
--
-- Uso:
--   supabase link --project-ref uuxvdulunrwppbmofyux
--   supabase db query --linked -f treino_reset.sql
--   supabase db query --linked -f treino_seed.sql
--   supabase link --project-ref fiarntunpqteopwjkhjg   (voltar para producao)
-- =====================================================================

truncate table public.registro_historico restart identity;
truncate table public.projeto_historico restart identity;
truncate table public.validador_avaliacoes restart identity;
truncate table public.projeto_equipe restart identity;
delete from public.kv_store;
delete from public.projetos;
