# Scripts do ambiente de treinamento

Scripts SQL usados para popular e resetar o banco do **ambiente de treinamento**
(projeto Supabase `unialfa-treinamento`, ref `uuxvdulunrwppbmofyux`), que é
totalmente separado do banco de produção (`unialfa-projetos`, ref
`fiarntunpqteopwjkhjg`). Não afetam produção de forma alguma.

- **`treino_seed.sql`** — cria 7 projetos fictícios cobrindo todo o ciclo de
  vida do sistema (do Gate 1 pendente até um projeto totalmente encerrado com
  TEP e RLA), com os registros correspondentes em cada um dos 10 formulários
  aplicáveis. Seguro para reaplicar (usa `on conflict do nothing`/`upsert`
  onde faz sentido).
- **`treino_reset.sql`** — apaga todos os dados de projetos/formulários do
  ambiente de treino (para rodar entre turmas), sem afetar as 6 contas fixas
  de treino nem as configurações (`public.configuracoes`).

## Como usar

Requer o [Supabase CLI](https://supabase.com/docs/guides/cli) autenticado.

```bash
supabase link --project-ref uuxvdulunrwppbmofyux

# resetar (opcional, só se já houver dados de uma turma anterior)
supabase db query --linked -f treino/treino_reset.sql

# popular com os 7 projetos de exemplo
supabase db query --linked -f treino/treino_seed.sql

# voltar a apontar o CLI para produção
supabase link --project-ref fiarntunpqteopwjkhjg
```

## Contas de treinamento

6 contas fixas (`@demo.sgp.local`, senha `Treino@2026`), uma por papel —
criadas separadamente via Auth Admin API, não fazem parte destes scripts.
Peça a quem administra o ambiente caso precise recriá-las. Na tela de login do
ambiente de treino, a pessoa não digita e-mail/senha: só seleciona o papel
desejado num menu suspenso e clica em "Entrar" — a conta e a senha correspondentes
são preenchidas automaticamente pelo próprio front-end.

## Sobre as Edge Functions no ambiente de treino

As mesmas 3 Edge Functions de produção (`send-notification`,
`analisar-transcricao-ata`, `transcrever-audio-ata`) estão publicadas no
projeto de treinamento, mas **deliberadamente sem nenhuma chave de API real
configurada** (`RESEND_API_KEY`, `VAPID_*`, `ANTHROPIC_API_KEY`,
`OPENAI_API_KEY`) — isso garante que nenhum e-mail real, notificação push
real ou chamada de IA paga saia do ambiente de treino, independente de
toggles ligados durante uma demonstração.
