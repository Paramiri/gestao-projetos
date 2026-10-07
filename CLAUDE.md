# UNIALFA — Sistema de Gestão de Projetos

Site estático (HTML/CSS/JS) publicado no GitHub Pages, com Supabase como backend (auth, tabelas via REST). Repositório: `https://github.com/Paramiri/unialfa-gestao-projetos`. Ao contrário do que uma versão antiga desta nota dizia, a maioria das tabelas **tem RLS habilitado** (`perfis`, `perfis_pendentes`, `projeto_equipe`, `projeto_historico`, `projetos`, `registro_historico`, `validador_avaliacoes`, `configuracoes`, `kv_store`, `push_subscriptions`) — confirmado direto no banco em 10/08/2026, não presuma "sem RLS" sem checar (`select relname, relrowsecurity from pg_class ...`). A maior parte das políticas usa a função `is_admin()` (`SECURITY DEFINER`, evita o bug de recursão de subconsultar a própria tabela — ver "Lições técnicas" abaixo) para liberar ações restritas a Admin; tabelas com dado por usuário (ex.: `push_subscriptions`, `perfis_pendentes`) também restringem por `auth.email() = <coluna>`. A maioria das regras de negócio mais finas continua só no JS de cada formulário, **exceto** a regra de quem pode editar/excluir um registro do `kv_store` (criador/Gerente de Projetos/Admin edita, Gerente de Projetos/Admin exclui, equipe do projeto continua exigida nos 7 formulários com essa restrição) — desde 15/09/2026 essa regra específica também é aplicada via RLS (`supabase/migrations/20260915143738_kv_store_row_level_access_control.sql`, funções `meu_papel()`/`sou_membro_projeto()`/`safe_uuid()` e colunas geradas `tipo`/`created_by`/`projeto_id`), não só no navegador.

**Regra permanente (RLS do kv_store em sincronia com o JS):** a política de edição/exclusão acima foi desenhada para espelhar exatamente `possoEditar()`/`possoExcluir()` (funções JS duplicadas nos 9 formulários de registro — ver `01- solicitacao-demanda.html` como referência) e o `souMembro()` já existente. Se essa regra mudar no JS no futuro (quem pode editar, quem pode excluir, ou quais formulários exigem equipe), a política SQL em `supabase/migrations/20260915143738_kv_store_row_level_access_control.sql` também precisa ser atualizada e reaplicada (`supabase db query --linked --file supabase/migrations/<arquivo>.sql`, em treino primeiro, depois produção) — senão as duas camadas ficam divergentes e um comportamento (ex.: alguém autorizado no JS mas bloqueado no banco, ou vice-versa) passa a ser inconsistente. Ficam deliberadamente fora dessa política (continuam só no JS): a decisão específica de quem aprova Gate 1/Gate 2 (mais restrita que a regra geral de edição) e os 3 relatórios de chave única (`plancom_data`, `relsit_data`, `relent_data`).

**Status atual (desde 18/09/2026):** o site de produção está **desligado** — `CNAME` (`gestaoprojetos.alfa.br`) e o workflow `.github/workflows/deploy-pages.yml` foram removidos (commits `3e7e011` e `5a8e882`) e o GitHub Pages foi despublicado manualmente. Um push no `main` **não** republica mais nada; o polling do "Padrão de deploy" abaixo só volta a fazer sentido se o site for religado (reverter esses dois commits + Settings → Pages → Source: GitHub Actions). O trabalho atual é só no código.

## Documentação oficial de regras de acesso

`Regras de Acesso e Permissoes - Sistema UNIALFA.docx` (na raiz do repositório) é o **documento oficial** e versionado de todas as regras de acesso e permissão do sistema: login obrigatório, acesso sem login (convidado), papéis de usuário, gates de aprovação (Gate 1/Gate 2) e restrição por equipe de projeto.

**Regra permanente:** sempre que uma regra de acesso for incluída, alterada ou removida no código — por exemplo, um novo formulário passa a exigir login/equipe, uma opção de "sem login" é adicionada, um novo papel é criado, uma trava de gate muda — este documento deve ser atualizado no mesmo commit (ou logo em seguida) para continuar refletindo o estado real do código. Isso vale tanto para mudanças feitas por mim quanto pelo usuário.

Como atualizar:
1. Editar o conteúdo em `gerar-regras-acesso.ps1` (raiz do repo) — cada seção é montada com as funções auxiliares `H1`/`H2`/`P`/`Bul`/tabela.
2. Rodar o script via PowerShell: `powershell -File "gerar-regras-acesso.ps1"` — ele usa automação COM do Microsoft Word (`New-Object -ComObject Word.Application`) para gerar o `.docx` diretamente na raiz do repo, sobrescrevendo o anterior.
3. Conferir o resultado (ex.: exportar para PDF via `$doc.SaveAs2($pdfPath, 17)` e ler o PDF) antes de commitar.
4. Commitar o `.docx` e o `.ps1` junto com a mudança de código que motivou a atualização.

Motivo de usar Word COM em vez do pacote `docx` (Node.js) da skill padrão: este ambiente não tem Node.js, pandoc, nem LibreOffice instalados — apenas o Microsoft Word está disponível localmente.

## Manual de Uso da ferramenta

`Manual de Uso - Ferramenta de Gestao de Projetos.docx` (na raiz do repositório) é o **manual oficial** e versionado de uso do sistema, com prints de tela e passo a passo de cada formulário, papéis, gates e relatórios.

**Regra permanente:** sempre que uma funcionalidade for incluída, alterada ou removida no código — novo formulário, novo campo relevante, nova regra de negócio (ex.: registro de riscos, sinalização automática, prioridade da EAP), mudança de fluxo — este manual deve ser atualizado no mesmo commit (ou logo em seguida) para continuar refletindo o estado real do sistema. Isso vale tanto para mudanças feitas por mim quanto pelo usuário.

Como atualizar:
1. Editar o conteúdo em `gerar-manual-uso.ps1` (raiz do repo) — inclui capa (com changelog de versão), TOC automático, seções por formulário com tabelas de campos e prints.
2. Rodar o script via PowerShell: `powershell -File "gerar-manual-uso.ps1" -ImgDir "<pasta com os prints>"` — **atenção**: `-ImgDir` precisa apontar para a pasta que já contém os prints existentes (não uma pasta vazia/nova), senão o regenerado perde todas as imagens antigas silenciosamente. Se novos prints forem necessários, tirar antes e colocar nessa mesma pasta.
3. Conferir o resultado (exportar para PDF via `$doc.SaveAs2($pdfPath, 17)` e ler o PDF) antes de commitar — comparar a contagem de páginas com a versão anterior como sinal rápido de que nenhuma imagem foi perdida.
4. Commitar o `.docx` e o `.ps1` junto com a mudança de código que motivou a atualização.

Se essa atualização for delegada a um agente, a verificação do passo 3 (conferir o PDF de fato, não só o texto) deve ser feita antes de reportar a tarefa como concluída — uma checagem só de texto (ex. `pdftotext`) não detecta perda de imagens.

## Central de Ajuda (in-app)

`15 - central-ajuda.html` é uma **terceira superfície de documentação**, independente dos dois `.docx` acima — é a versão navegável (com busca e menu lateral) do Manual de Uso, publicada como página do próprio site, e é a que os usuários realmente abrem a partir do ícone `?` em cada formulário. Uma mudança de funcionalidade documentada no Manual de Uso ou nas Regras de Acesso e esquecida aqui deixa o `?` do sistema desatualizado mesmo com os `.docx` corretos — já aconteceu uma vez (notificações push).

**Regra permanente:** sempre que o Manual de Uso ou as Regras de Acesso forem atualizados por uma mudança de funcionalidade, checar se `15 - central-ajuda.html` também precisa de uma seção nova/ajustada (adicionar `<a>` no menu lateral `#sbNav` + `<section class="sec" id="...">` correspondente) — no mesmo commit da mudança de código, junto com os `.docx`.

Diferente de `app.css`/`app.js`, essa página não é referenciada por `<link>`/`<script>` num único lugar, e sim por um link `?` (`class="doc-help"`/`"head-help"`) em cada um dos 12 formulários + o item de menu em `app.js` — todos com cache-busting `?v=N` (ex.: `href="15 - central-ajuda.html?v=1#passo-1"`). Sempre que o conteúdo de `15 - central-ajuda.html` mudar, incrementar o `N` em **todos** esses links (não só no arquivo em si), senão o GitHub Pages/CDN pode continuar servindo a versão em cache por até 10 minutos.

## Ambiente de Treinamento (sincronização)

`Paramiri/unialfa-gestao-projetos-treino` é um repositório-espelho, publicado separadamente no GitHub Pages (`https://paramiri.github.io/unialfa-gestao-projetos-treino/`), usado como ambiente de treinamento/demonstração: mesmo código, backend Supabase totalmente separado (projeto `unialfa-treinamento`, ref `uuxvdulunrwppbmofyux`), selecionado em runtime pelo hostname (`IS_TREINO = location.href.indexOf('treino')>-1`, já presente em todos os arquivos HTML). Scripts de seed/reset dos dados de exemplo desse ambiente ficam versionados em `treino/` neste repositório; o botão que roda esse reset a partir da Administração (produção) chama a Edge Function `reset-treino`.

**Regra permanente:** sempre que um arquivo do site de produção for alterado — HTML, `app.js`/`app.css`, ou uma Edge Function em `supabase/functions/` — replicar a mesma mudança no repositório-espelho, no mesmo commit (ou logo em seguida) da mudança em produção, para o ambiente de treinamento continuar refletindo a mesma versão do sistema. Isso vale tanto para mudanças feitas por mim quanto pelo usuário, exatamente como já vale para o Manual de Uso e a Central de Ajuda.

Como sincronizar:
1. Copiar os arquivos alterados de `unialfa-gestao-projetos` (produção) para a working copy local do repositório-espelho, preservando as poucas divergências propositais dele: sem arquivo `CNAME` (usa o endereço padrão do GitHub Pages) e sem a senha das contas de treino em texto puro no `README.md`.
2. Commitar e dar push no repositório-espelho (mesmo padrão de commit em português + `Co-Authored-By`).
3. Confirmar a publicação fazendo polling no GitHub Pages do espelho (`https://paramiri.github.io/unialfa-gestao-projetos-treino/...`) até o conteúdo novo aparecer — mesmo procedimento do "Padrão de deploy" abaixo, aplicado ao repositório-espelho.
4. Se a mudança envolver uma Edge Function nova ou alterada, reimplantá-la também no projeto de treinamento (`supabase functions deploy <nome> --project-ref uuxvdulunrwppbmofyux`) — sem copiar nenhum secret real de e-mail/IA (`RESEND_API_KEY`, `VAPID_*`, `ANTHROPIC_API_KEY`, `OPENAI_API_KEY`) para lá, propositalmente, para o ambiente de treino nunca poder disparar e-mail real, push real nem chamada de IA paga.

**Regra permanente (dados de exemplo):** sempre que uma mudança alterar o formato dos dados salvos por algum formulário (novo campo no objeto do `kv_store`, nova coluna em `projetos`/`projeto_equipe`/`projeto_historico`/`registro_historico`) ou introduzir uma funcionalidade que dependa de dado que os 7 projetos de exemplo do treino ainda não têm (ex.: severidade de risco no TAP, datas reais no cronograma, nomes de entrega/categoria casando entre formulários) — atualizar `treino/treino_seed.sql` para os exemplos passarem a **demonstrar** a funcionalidade nova, não só aceitar o dado sem erro. Isso vale tanto para mudanças feitas por mim quanto pelo usuário — mesmo padrão de obrigatoriedade já usado para o Manual de Uso e a Central de Ajuda, não uma avaliação opcional. Já aconteceu de o seed ficar desatualizado por vários ciclos de feature sem ninguém perceber (riscos sem severidade, cronograma com campos antigos) — o `?` do sistema não mostra isso, só fica visível quando alguém entra no ambiente de treino.

Como sincronizar os dados de exemplo (além do código, no item acima):
1. Editar `treino/treino_seed.sql` com o ajuste necessário, validando cada bloco JSON alterado antes de aplicar (ex.: `ConvertFrom-Json` via PowerShell, já que este ambiente não tem Node/Python).
2. Replicar a MESMA mudança dentro da constante `SEED_SQL` embutida em `supabase/functions/reset-treino/index.ts` — os dois precisam ficar com o mesmo conteúdo (ver o comentário no topo do arquivo); reconstruir o arquivo colando o novo `treino_seed.sql` entre o prefixo/sufixo existentes é mais seguro que editar o `.ts` linha a linha.
3. Reimplantar a Edge Function atualizada em produção: `supabase functions deploy reset-treino --project-ref fiarntunpqteopwjkhjg`.
4. Rodar o reset diretamente no banco de treinamento via Supabase CLI, para o ambiente já refletir a mudança sem exigir que alguém clique no botão manualmente:
   `supabase link --project-ref uuxvdulunrwppbmofyux` → `supabase db query --linked --file treino/treino_reset.sql` → `supabase db query --linked --file treino/treino_seed.sql` → `supabase link --project-ref fiarntunpqteopwjkhjg` (voltar para produção).
5. Conferir o resultado com uma consulta rápida (`supabase db query --linked "select ..."`) antes de considerar concluído.
6. Commitar `treino/treino_seed.sql` e `supabase/functions/reset-treino/index.ts` juntos — esses dois arquivos só existem no repositório de produção, não no espelho, então não entram no passo de cópia para o repositório-espelho acima.

## Padrão de deploy

`git add` dos arquivos específicos → commit em português com `Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>` → `git push` → confirmar publicação fazendo polling no GitHub Pages (`https://paramiri.github.io/unialfa-gestao-projetos/...`) até o conteúdo novo aparecer.

`app.css`/`app.js` são referenciados em `index.html` com cache-busting `?v=N` — incrementar o número sempre que esses dois arquivos mudarem.

## Lições técnicas (migradas da memória local do Claude Code em 07/10/2026)

Antes ficavam só na memória do Claude Code da máquina antiga (`memoria-claude-backup.zip`); trazidas para cá para não se perderem numa próxima troca de máquina.

- **RLS — nunca checar papel com subconsulta na tabela de perfis.** Uma política como `exists (select 1 from perfis where id = auth.uid() and papel = 'admin')` em `perfis` (ou em qualquer outra tabela que subconsulte `perfis`, já que isso dispara a política SELECT de `perfis`) entra em recursão infinita e aparece só como `500` opaco do PostgREST. Sintoma já visto: `window.currentProfile` ficava `null` para todos, e o fallback `profile?profile.papel:'solicitante'` escondia o erro — todo usuário aparecia como "Solicitante" e gates/admin/equipe paravam de funcionar sem erro visível. Sempre encapsular a checagem numa função `SECURITY DEFINER` (`is_admin()`, `meu_papel()`, `sou_membro_projeto()`). Diagnóstico: papel sempre no valor padrão → olhar a aba Network atrás de `500` na consulta de perfil antes de suspeitar do JS.
- **Login por magic link — redirect via repositório-relé.** O Supabase truncava o redirect para `https://paramiri.github.io/` (sem o sub-caminho), em qualquer configuração de Redirect URLs. Solução em uso: repositório separado `Paramiri/paramiri.github.io` com um `index.html` que só repassa `location.hash` para `/unialfa-gestao-projetos/`; os formulários chamam `/auth/v1/otp` **sem** `email_redirect_to`. Se o login quebrar com redirect para a raiz/localhost, a correção é nesse relé, não na URL Configuration do Supabase. (Com o site desligado e sem domínio próprio, rever isso se o site for religado em outro endereço.)
- **E-mail (Auth e notificações) via Resend.** SMTP customizado no Supabase (Authentication → Emails → SMTP Settings, `smtp.resend.com:465`, usuário `resend`) com domínio verificado `sistemas.alfa.br` (remetente `sistemas@sistemas.alfa.br`; notificações saem de `notificacoes@sistemas.alfa.br`); templates Magic Link/Confirm signup traduzidos para português. Sem SMTP próprio, o limite `over_email_send_rate_limit` (429) do Supabase é por projeto inteiro, não por destinatário. A Edge Function `send-notification` é genérica (`{to, subject, html}`), chamada por `window.unialfaNotificar.enviar()` nos formulários; um gatilho novo de e-mail só precisa de nova chamada no JS, não de nova função.
- **Windows (máquina local do usuário).** O `winget` **não** tem pacote do Supabase CLI — baixar o `.zip` da release no GitHub, descobrindo o nome real do asset pela API (`Invoke-RestMethod https://api.github.com/repos/supabase/cli/releases/latest`), pois ele inclui a versão. Conferir se o usuário está no PowerShell (`PS C:\>`) e não no `cmd` (`C:\>`); `curl` no PowerShell é alias de `Invoke-WebRequest` (usar `-Headers @{...}`); `$env:PATH += ...` só vale para a janela atual; rodar `supabase link` de dentro da pasta do repositório (em pasta protegida dá erro `FileSystem.makeDirectory`); para trocar de unidade no `cmd`, `cd /d`. `python`/`python3` nessa máquina são stubs da Microsoft Store.
- **Word COM — exportação para PDF pode travar.** `SaveAs2($pdf, 17)`/`ExportAsFixedFormat` já travaram indefinidamente até com documento em branco. Se passar de ~60–90 s, não insistir: matar o `winword.exe` (`Get-Process winword | Stop-Process -Force`) e verificar o `.docx` direto pelo XML — descompactar, extrair texto de `<w:t>` em `word/document.xml`, e comparar `word/media/` com o número de chamadas `Img "` em `gerar-manual-uso.ps1` para garantir que nenhuma imagem sumiu. Avisar o usuário que a verificação visual/de páginas não foi feita.

## Backups fora do Git

Itens que **não** estão no repositório e precisam ser recuperados manualmente numa troca de máquina (backups feitos em 18/09/2026, guardados pelo usuário junto com o backup do OneDrive):

- `prints-manual-de-uso-backup.zip` — pasta de prints do Manual de Uso; extrair e apontar o `-ImgDir` de `gerar-manual-uso.ps1` para ela (sem isso o Manual é regerado sem imagens, silenciosamente).
- `memoria-claude-backup.zip` — memória antiga do Claude Code; o conteúdo útil já foi incorporado na seção acima, guardar só como histórico.
