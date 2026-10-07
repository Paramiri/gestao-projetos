param([string]$OutPath = (Join-Path $PSScriptRoot "Guia de Replicacao Tecnica - SGP.docx"))
$ErrorActionPreference = "Stop"

# Gera "Guia de Replicacao Tecnica - SGP.docx" via automacao COM do Microsoft Word
# (Node.js/pandoc/LibreOffice nao estao disponiveis neste ambiente).
# Documento de analise tecnica: o que foi usado neste projeto e o passo a passo cronologico
# para construir do zero uma aplicacao semelhante.

function RGB($r,$g,$b) { return [int]($r + ($g*256) + ($b*65536)) }
$colRed   = RGB 0xB9 0x1D 0x2E
$colInk   = RGB 0x1A 0x1A 0x1A
$colMuted = RGB 0x6A 0x6A 0x70
$colWhite = RGB 0xFF 0xFF 0xFF
$colGood  = RGB 0x05 0x96 0x69

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()
$sel = $word.Selection

$sec = $doc.Sections.Item(1)
$sec.PageSetup.PageWidth = $word.CentimetersToPoints(21.59)
$sec.PageSetup.PageHeight = $word.CentimetersToPoints(27.94)

$footer = $sec.Footers.Item(1)
$footer.Range.Font.Size = 8.5
$footer.Range.Font.Color = $colMuted
$footer.Range.Text = "SGP - Guia de Replicacao Tecnica | Pagina "
$footer.Range.Collapse(0) | Out-Null
$footer.Range.Fields.Add($footer.Range, 33) | Out-Null  # wdFieldPage = 33

function P($text, $size=11, $bold=$false, $italic=$false, $color=$colInk, $align="left", $spaceAfter=8) {
  $sel.Style = $doc.Styles.Item(-1)
  $sel.Font.Name = "Montserrat"
  $sel.Font.Size = $size
  $sel.Font.Bold = $bold
  $sel.Font.Italic = $italic
  $sel.Font.Color = $color
  $sel.ParagraphFormat.Alignment = if($align -eq "center"){1}else{0}
  $sel.ParagraphFormat.SpaceAfter = $spaceAfter
  $sel.ParagraphFormat.LineSpacing = 14
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Font.Bold = $false
  $sel.Font.Italic = $false
}

function H1($text) {
  $sel.Style = $doc.Styles.Item(-2)
  $sel.Font.Name = "Montserrat"
  $sel.Font.Size = 17
  $sel.Font.Bold = $true
  $sel.Font.Color = $colInk
  $sel.Font.Italic = $false
  $sel.ParagraphFormat.Alignment = 0
  $sel.ParagraphFormat.SpaceBefore = 4
  $sel.ParagraphFormat.SpaceAfter = 10
  $sel.ParagraphFormat.Borders.Item(3).LineStyle = 1
  $sel.ParagraphFormat.Borders.Item(3).Color = RGB 0xE4 0xE4 0xE7
  $sel.ParagraphFormat.PageBreakBefore = $true
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Borders.Item(3).LineStyle = 0
  $sel.ParagraphFormat.PageBreakBefore = $false
  $sel.Style = $doc.Styles.Item(-1)
}

function H2($text) {
  $sel.Style = $doc.Styles.Item(-3)
  $sel.Font.Name = "Montserrat"
  $sel.Font.Size = 13.5
  $sel.Font.Bold = $true
  $sel.Font.Color = $colRed
  $sel.Font.Italic = $false
  $sel.ParagraphFormat.Alignment = 0
  $sel.ParagraphFormat.SpaceBefore = 16
  $sel.ParagraphFormat.SpaceAfter = 6
  $sel.ParagraphFormat.PageBreakBefore = $false
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item(-1)
  $sel.Font.Color = $colInk
}

function H3($text) {
  $sel.Style = $doc.Styles.Item(-1)
  $sel.Font.Name = "Montserrat"
  $sel.Font.Size = 12
  $sel.Font.Bold = $true
  $sel.Font.Color = $colInk
  $sel.ParagraphFormat.SpaceBefore = 12
  $sel.ParagraphFormat.SpaceAfter = 5
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Font.Bold = $false
}

function Bul($text) {
  $sel.Style = $doc.Styles.Item(-1)
  $sel.Font.Name = "Montserrat"
  $sel.Font.Size = 11
  $sel.Font.Bold = $false
  $sel.Font.Color = $colInk
  $sel.ParagraphFormat.SpaceAfter = 4
  $sel.ParagraphFormat.LineSpacing = 13
  $sel.ParagraphFormat.LeftIndent = $word.CentimetersToPoints(0.6)
  $sel.Range.ListFormat.ApplyBulletDefault()
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Range.ListFormat.RemoveNumbers()
  $sel.ParagraphFormat.LeftIndent = 0
}

function Nota($text){
  $sel.Style = $doc.Styles.Item(-1)
  $sel.ParagraphFormat.LeftIndent = $word.CentimetersToPoints(0.5)
  $sel.ParagraphFormat.Borders.Item(1).LineStyle = 1
  $sel.ParagraphFormat.Borders.Item(1).Color = $colRed
  $sel.ParagraphFormat.Borders.Item(1).LineWidth = 6
  $sel.Font.Name = "Montserrat"
  $sel.Font.Size = 10.5
  $sel.Font.Italic = $true
  $sel.Font.Color = $colMuted
  $sel.ParagraphFormat.SpaceBefore = 8
  $sel.ParagraphFormat.SpaceAfter = 10
  $sel.Font.Bold = $true
  $sel.Font.Color = $colRed
  $sel.TypeText("ATENCAO - ")
  $sel.Font.Bold = $false
  $sel.Font.Color = $colMuted
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Font.Italic = $false
  $sel.ParagraphFormat.Borders.Item(1).LineStyle = 0
  $sel.ParagraphFormat.LeftIndent = 0
}

function Exemplo($text){
  $sel.Style = $doc.Styles.Item(-1)
  $sel.ParagraphFormat.LeftIndent = $word.CentimetersToPoints(0.5)
  $sel.ParagraphFormat.Borders.Item(1).LineStyle = 1
  $sel.ParagraphFormat.Borders.Item(1).Color = $colGood
  $sel.ParagraphFormat.Borders.Item(1).LineWidth = 6
  $sel.Font.Name = "Consolas"
  $sel.Font.Size = 10
  $sel.Font.Color = $colInk
  $sel.ParagraphFormat.SpaceBefore = 8
  $sel.ParagraphFormat.SpaceAfter = 10
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Borders.Item(1).LineStyle = 0
  $sel.ParagraphFormat.LeftIndent = 0
}

function TableSimple($rows, $colWidthsCm){
  $nRows = $rows.Count; $nCols = $rows[0].Count
  $range = $sel.Range
  $table = $doc.Tables.Add($range, $nRows, $nCols)
  $table.Borders.Enable = $true
  $table.Borders.InsideLineStyle = 1
  $table.Borders.OutsideLineStyle = 1
  $table.Borders.InsideColor = RGB 0xD1 0xD5 0xDB
  $table.Borders.OutsideColor = RGB 0xD1 0xD5 0xDB
  for ($r=0; $r -lt $nRows; $r++) {
    for ($c=0; $c -lt $nCols; $c++) {
      $cell = $table.Cell($r+1, $c+1)
      $cell.Range.Text = $rows[$r][$c]
      $cell.Range.Font.Size = 9.5
      $cell.Range.Font.Name = "Montserrat"
      if ($r -eq 0) {
        $cell.Range.Font.Bold = $true
        $cell.Range.Font.Color = $colWhite
        $cell.Shading.BackgroundPatternColor = $colInk
      } else {
        $cell.Range.Font.Bold = $false
        $cell.Range.Font.Color = $colInk
      }
    }
  }
  for ($c=0; $c -lt $nCols; $c++) { $table.Columns.Item($c+1).Width = $word.CentimetersToPoints($colWidthsCm[$c]) }
  $sel.EndKey(6) | Out-Null
  $sel.Style = $doc.Styles.Item(-1)
  $sel.ParagraphFormat.SpaceAfter = 10
  $sel.TypeParagraph()
}

# ============================================================
# CAPA
# ============================================================
P "SGP - SISTEMA DE GESTAO DE PROJETOS" 10 $true $false $colRed "left" 4
$sel.Style = $doc.Styles.Item(-1)
$sel.Font.Name="Montserrat"; $sel.Font.Size = 26; $sel.Font.Bold = $true; $sel.Font.Color = $colInk
$sel.ParagraphFormat.SpaceAfter = 4
$sel.TypeText("GUIA DE REPLICACAO TECNICA")
$sel.TypeParagraph()
$sel.Font.Bold = $false
P "Sistema de Gestao de Projetos (SGP)" 15 $false $false $colInk "left" 4
P "Analise da arquitetura atual e passo a passo cronologico para construir do zero uma aplicacao semelhante (mesmo modelo: site estatico + Supabase + GitHub Pages)." 12 $false $true $colMuted "left" 30
P "Andre Torres - Gerente de Projetos" 11 $false $false $colMuted "left" 2
P "Documento de analise tecnica - versao 1.0 - 09 de setembro de 2026" 11 $false $false $colMuted "left" 2

$sel.InsertBreak(7) | Out-Null

# ============================================================
# SUMARIO (TOC automatico)
# ============================================================
$sel.Style = $doc.Styles.Item(-1)
$sel.Font.Name="Montserrat"; $sel.Font.Size = 16; $sel.Font.Bold = $true; $sel.Font.Color = $colInk
$sel.ParagraphFormat.SpaceAfter = 10
$sel.TypeText("SUMARIO")
$sel.TypeParagraph()
$sel.Font.Bold = $false
$tocRange = $sel.Range
$toc = $doc.TablesOfContents.Add($tocRange, $true, 1, 2)
$sel.EndKey(6) | Out-Null
$sel.InsertBreak(7) | Out-Null

# ============================================================
# 1. VISAO GERAL
# ============================================================
H1 "1. Visao geral - o que e este sistema, por baixo do capo"
P "Este documento descreve, de forma tecnica, tudo o que foi utilizado para construir e publicar o Sistema de Gestao de Projetos (SGP), e apresenta o passo a passo, em ordem cronologica, de como construir uma aplicacao semelhante do zero. E um documento de analise e planejamento, nao um manual de uso do sistema (esse ja existe: `Manual de Uso - Ferramenta de Gestao de Projetos.docx`)."

H2 "1.1 Resumo da arquitetura"
P "O sistema e um site estatico (sem framework, sem etapa de build/compilacao) publicado no GitHub Pages, com todo o backend fornecido pelo Supabase (banco de dados Postgres, autenticacao, armazenamento de arquivos e funcoes de servidor). Nao existe servidor proprio: tudo o que precisa rodar `no servidor` (chamadas a IA, envio de e-mail, notificacao push) roda em Funcoes de Borda (Edge Functions) do proprio Supabase."
$rArq = @(
  @("Camada","Tecnologia usada neste projeto"),
  @("Frontend","HTML + CSS + JavaScript puro, um arquivo por formulario/pagina, sem framework e sem build step"),
  @("Hospedagem do site","GitHub Pages, publicado automaticamente por GitHub Actions a cada push na branch main"),
  @("Dominio proprio","gestaoprojetos.suaempresa.com.br, apontado via DNS (CNAME) para o GitHub Pages"),
  @("Banco de dados","Supabase (Postgres gerenciado), acessado direto via API REST (PostgREST), sem SDK"),
  @("Autenticacao","Supabase Auth - link magico por e-mail (OTP) e login com Microsoft (Azure AD/SSO)"),
  @("E-mail transacional","Resend, usado como SMTP customizado do Supabase Auth e chamado pelas Edge Functions"),
  @("Armazenamento de arquivos","Supabase Storage (buckets privados, um por tipo de anexo)"),
  @("Funcoes de servidor","Supabase Edge Functions (Deno/TypeScript), 6 funcoes no total"),
  @("Inteligencia artificial","API da Anthropic (Claude) para analise/sugestao de texto; API da OpenAI (Whisper) para transcricao de audio"),
  @("Notificacao push","Web Push API + chaves VAPID, com Service Worker (PWA)"),
  @("Documentacao","Scripts PowerShell com automacao COM do Microsoft Word, gerando .docx; uma versao navegavel dentro do proprio site")
)
TableSimple $rArq @(4.5,11.5)

H2 "1.2 Por que essas escolhas"
Bul "Site estatico sem framework: qualquer editor de texto e um `git push` bastam para publicar - nao ha etapa de build, nao ha dependencia de Node.js/npm no ambiente de manutencao (que de fato nao tem Node.js instalado)."
Bul "Supabase no lugar de um backend proprio: da banco de dados, autenticacao, storage e funcoes de servidor prontos, com plano gratuito suficiente para comecar, e todas as regras de seguranca ficam centralizadas no banco (Row Level Security), nao espalhadas pelo codigo do frontend."
Bul "GitHub Pages: hospedagem gratuita, HTTPS automatico, e o mesmo repositorio que guarda o codigo ja publica o site - sem servidor para administrar."
Nota "Este e um modelo deliberadamente simples e barato de operar. Ele troca flexibilidade de backend (nao ha logica de servidor customizada fora das Edge Functions) por velocidade de manutencao e custo quase zero de infraestrutura. Para um sistema com regras de negocio muito mais complexas ou alto volume, um backend dedicado pode valer mais a pena."

# ============================================================
# 2. PRE-REQUISITOS
# ============================================================
H1 "2. Pre-requisitos - contas e acessos necessarios"
P "Antes de iniciar a construcao, e preciso providenciar as contas e acessos abaixo. Alguns dependem de outras areas (TI/infraestrutura de dominio, administrador do Microsoft 365/Azure AD da organizacao), entao vale levantar isso com antecedencia."
$rPre = @(
  @("Conta / acesso","Para que serve","Custo aproximado"),
  @("Conta no GitHub","Guardar o codigo-fonte e publicar o site (GitHub Pages)","Gratuito (repo publico); pago se precisar de repo privado com Pages"),
  @("Conta no Supabase","Banco de dados, autenticacao, storage e funcoes de servidor","Gratuito para comecar; planos pagos por uso conforme o projeto cresce"),
  @("Dominio proprio + acesso ao DNS","Endereco proprio do site (ex.: nomedosistema.suaempresa.com.br) e verificacao do remetente de e-mail","Custo do registro do dominio (se ainda nao existir); alteracao de DNS costuma ser gratuita"),
  @("Conta no Resend","Envio de e-mail transacional (login por link e notificacoes)","Gratuito ate um certo volume; pago acima disso"),
  @("Acesso ao Azure AD / Microsoft Entra ID da organizacao","Cadastrar o aplicativo que permite Entrar com Microsoft","Geralmente sem custo adicional se a empresa ja usa Microsoft 365/Azure AD"),
  @("Conta na Anthropic (console.anthropic.com)","Chave de API para os recursos de IA (sugestao de preenchimento, analise de documentos)","Pago por uso (tokens processados)"),
  @("Conta na OpenAI (platform.openai.com)","Chave de API para transcricao de audio (Whisper)","Pago por uso (minutos de audio)"),
  @("Microsoft Word instalado","Gerar a documentacao oficial (manual, regras de acesso) via automacao COM","Ja incluso se a organizacao usa Microsoft Office")
)
TableSimple $rPre @(4.2,7.5,4.3)
Nota "As contas de IA (Anthropic/OpenAI) e de e-mail (Resend) sao opcionais no sentido de que o sistema funciona sem elas - mas os recursos de IA e as notificacoes automaticas ficam indisponiveis ate que as respectivas chaves sejam configuradas."

# ============================================================
# 3. FASE 1 - REPOSITORIO E HOSPEDAGEM
# ============================================================
H1 "3. Fase 1 - Repositorio no GitHub e publicacao do site"

H2 "3.1 Criar o repositorio"
P "Criar um repositorio novo no GitHub (publico, para poder usar o GitHub Pages gratuito). Definir a branch principal como `main`."

H2 "3.2 Habilitar o GitHub Pages via GitHub Actions"
P "Em Settings > Pages, escolher como fonte de publicacao (`Source`) a opcao `GitHub Actions` (em vez de publicar direto de uma branch). Isso permite controlar exatamente o que e publicado atraves de um workflow."

H2 "3.3 Criar o workflow de deploy automatico"
P "Criar o arquivo `.github/workflows/deploy-pages.yml` com um job que faz checkout do codigo, prepara o Pages e publica o conteudo da raiz do repositorio a cada push na branch main:"
Exemplo "on: push (branch main) -> actions/checkout -> actions/configure-pages -> actions/upload-pages-artifact (path: '.') -> actions/deploy-pages"
P "A partir desse momento, todo `git push` na branch main publica o site automaticamente, sem nenhum passo manual - o unico `deploy` que existe neste projeto e um commit."

H2 "3.4 Resultado"
P "O site fica acessivel em `https://<usuario-ou-organizacao>.github.io/<nome-do-repositorio>/`. O dominio proprio (se houver) e configurado depois, na Fase 6 - o site ja funciona neste endereco padrao antes disso."

# ============================================================
# 4. FASE 2 - BACKEND SUPABASE
# ============================================================
H1 "4. Fase 2 - Criar o projeto Supabase (banco de dados e backend)"

H2 "4.1 Criar o projeto"
P "Em supabase.com, criar um novo projeto: escolher organizacao, nome, senha do banco de dados (guardar em local seguro) e a regiao mais proxima dos usuarios (ex.: `sa-east-1` para o Brasil)."

H2 "4.2 Anotar as credenciais publicas"
P "Em Project Settings > API, anotar dois valores que vao para o codigo do frontend:"
Bul "Project URL (ex.: https://xxxxxxxxxxxx.supabase.co)"
Bul "Chave publica (`anon` / `publishable key`)"
Nota "Essas duas informacoes NAO sao segredo - elas ficam visiveis no codigo JavaScript do site. A seguranca dos dados nao depende de esconde-las, e sim das politicas de Row Level Security (RLS) configuradas no banco (Fase 3). Ja a chave `service_role` (tambem em Project Settings > API) e um segredo real - ela ignora todas as politicas de RLS e nunca deve aparecer em codigo de frontend, apenas dentro das Edge Functions (Fase 9)."

# ============================================================
# 5. FASE 3 - MODELAGEM DO BANCO
# ============================================================
H1 "5. Fase 3 - Modelar as tabelas e a seguranca (RLS)"

H2 "5.1 Tabelas usadas neste projeto"
P "Este sistema usa 10 tabelas no schema publico. A maioria dos formularios nao tem uma tabela dedicada: eles gravam um registro no formato JSON numa tabela generica de chave-valor (`kv_store`), o que reduz a modelagem SQL as custas de exigir mais disciplina no codigo JavaScript."
$rTab = @(
  @("Tabela","Papel"),
  @("perfis","Um registro por usuario logado: e-mail, nome, papel (Solicitante, Gerente de Projetos, Gestor Responsavel, Dono do Negocio, Alta Gestao, PMO/Admin)"),
  @("perfis_pendentes","Cadastros de usuario aguardando aprovacao/definicao de papel pelo Admin"),
  @("projetos","Cadastro compartilhado de projetos (nome, status, datas) usado pelos 9 formularios do ciclo de vida"),
  @("projeto_equipe","Vinculo de quais usuarios pertencem a equipe de cada projeto (usado na restricao por equipe)"),
  @("projeto_historico","Historico de eventos importantes de cada projeto (aprovacoes, decisoes de SMP, encerramento)"),
  @("kv_store","Armazenamento generico chave-valor: guarda o JSON de cada registro de cada formulario (Demanda, Canvas, TAP, etc.)"),
  @("registro_historico","Historico generico de alteracoes de qualquer registro (usado pelo botao `Ver historico`)"),
  @("validador_avaliacoes","Avaliacoes feitas na ferramenta Validador de Projetos"),
  @("configuracoes","Tabela chave/valor de liga-desliga de funcionalidades (ex.: permitir solicitacao sem login)"),
  @("push_subscriptions","Inscricoes de notificacao push de cada usuario/dispositivo")
)
TableSimple $rTab @(4,12)

H2 "5.2 Habilitar Row Level Security (RLS) em todas as tabelas"
P "Logo apos criar cada tabela, habilitar RLS:"
Exemplo "alter table nome_da_tabela enable row level security;"
Nota "Uma tabela sem RLS habilitado fica acessivel por qualquer pessoa que tenha a chave publica (ou seja, qualquer visitante do site), mesmo sem login. RLS e o que faz a chave publica ser segura de expor no frontend."

H2 "5.3 Criar a funcao is_admin() ANTES de qualquer politica de papel"
P "Erro comum a evitar: uma politica de RLS que verifica `e admin?` fazendo uma subconsulta direta na propria tabela de perfis causa recursao infinita no Postgres - o sintoma e um erro 500 sem mensagem clara, e a interface simplesmente nunca mostra o papel real do usuario (fica sempre no papel padrao/`Solicitante`), sem erro visivel."
Exemplo "create or replace function public.is_admin() returns boolean language sql security definer set search_path = public stable as `$`$ select exists(select 1 from perfis where id = auth.uid() and papel = 'admin'); `$`$;"
P "Toda politica que precisar checar `este usuario e admin?` deve chamar `is_admin()` em vez de repetir a subconsulta - inclusive politicas em outras tabelas (ex.: `projeto_equipe`) que tambem consultem a tabela de perfis."

H2 "5.4 Escrever as politicas de acesso"
P "Para cada tabela, escrever as politicas de SELECT/INSERT/UPDATE/DELETE de acordo com a regra de negocio: normalmente `auth.uid() = dono_do_registro or is_admin()` para dados pessoais, e `is_admin()` sozinho para acoes restritas a administrador. Tabelas com dado por usuario tambem podem restringir por `auth.email() = coluna_de_email`."
Nota "Regras de negocio mais finas (por exemplo, quem pode editar um registro especifico de um formulario) neste projeto ficam no JavaScript de cada pagina, nao no banco - o RLS garante o minimo (isolamento entre usuarios e papeis), e a interface refina o resto. Para dados realmente sensiveis, prefira aplicar a regra tambem no banco."

# ============================================================
# 6. FASE 4 - LOGIN POR E-MAIL
# ============================================================
H1 "6. Fase 4 - Login por link magico e validacao do dominio de e-mail"

H2 "6.1 Habilitar o provedor de e-mail"
P "No Supabase, em Authentication > Providers, o provedor `Email` ja vem habilitado por padrao, permitindo login sem senha via link enviado por e-mail (OTP)."

H2 "6.2 Configurar um SMTP proprio (essencial em producao)"
P "Sem um SMTP customizado, o Supabase usa um servico de e-mail proprio com limite de envio baixo (compartilhado entre todos os projetos gratuitos) e o remetente/textos padrao vem em ingles. Em Authentication > Emails > SMTP Settings, configurar um provedor de e-mail transacional proprio."

H2 "6.3 Criar conta no Resend e gerar uma API key"
P "Criar conta em resend.com e gerar uma API key em API Keys."

H2 "6.4 Validar (verificar) o dominio de envio - o passo de validar o e-mail"
P "Este e o passo que efetivamente libera o envio de e-mail para qualquer destinatario. Sem ele, o Resend fica em modo sandbox e so permite enviar e-mail de teste para o proprio dono da conta."
Bul "Em Resend > Domains, adicionar o dominio (ou subdominio dedicado, ex.: sistemas.suaempresa.com.br) que vai ser usado como remetente."
Bul "O Resend gera registros DNS (SPF, DKIM e, geralmente, DMARC) que precisam ser cadastrados no DNS do dominio - normalmente por quem administra o dominio (TI/infraestrutura)."
Bul "Aguardar a propagacao e confirmar a verificacao dentro do proprio Resend (o status muda para `Verified`); ferramentas como dnschecker.org ajudam a confirmar que os registros ja propagaram."
Nota "Neste projeto, o dominio principal da organizacao ja estava com o limite de registros do provedor de DNS esgotado, entao foi usado um dominio alternativo da mesma empresa com um subdominio dedicado (`sistemas.<dominio-alternativo>`). O plano gratuito do Resend permite verificar apenas 1 dominio por conta - planeje isso com antecedencia se a organizacao tiver mais de um sistema."

H2 "6.5 Apontar o Supabase para o SMTP do Resend"
$rSmtp = @(
  @("Campo","Valor"),
  @("Host","smtp.resend.com"),
  @("Porta","465"),
  @("Usuario","resend"),
  @("Senha","a API key gerada no Resend"),
  @("Remetente (From)","um endereco no dominio ja verificado (ex.: sistemas@sistemas.suaempresa.com.br)")
)
TableSimple $rSmtp @(4,12)

H2 "6.6 Traduzir os templates de e-mail"
P "Em Authentication > Email Templates, editar os templates `Confirm signup` e `Magic Link` (entre outros usados) para o idioma da organizacao."

H2 "6.7 Testar de ponta a ponta com um destinatario externo"
P "Enviar um link magico para um e-mail que NAO seja o dono da conta Resend, para confirmar que o dominio realmente saiu do modo sandbox e que o e-mail chega (nao so na caixa de entrada do proprio remetente)."

# ============================================================
# 7. FASE 5 - LOGIN COM MICROSOFT
# ============================================================
H1 "7. Fase 5 - Login com Microsoft (Azure AD / SSO)"
P "Esta e uma segunda forma de login, independente da anterior - o usuario pode entrar com a conta institucional Microsoft em vez de receber um link por e-mail."

H2 "7.1 Registrar o aplicativo no Azure"
P "No portal.azure.com, ir em Azure Active Directory (ou Microsoft Entra ID) > App registrations > New registration. Isso exige um usuario com permissao administrativa no tenant Azure AD/Microsoft 365 da organizacao."
Bul "Nome do aplicativo (ex.: `Sistema de Gestao de Projetos - Login`)."
Bul "Tipo de conta suportada: normalmente `Accounts in this organizational directory only` (mono-tenant), para restringir o login aos usuarios da propria empresa."
Bul "Redirect URI (tipo Web):"
Exemplo "https://<ref-do-projeto>.supabase.co/auth/v1/callback"

H2 "7.2 Gerar as credenciais do aplicativo"
Bul "Anotar o `Application (client) ID`, exibido na pagina de visao geral do app registrado."
Bul "Em Certificates & secrets > New client secret, gerar um segredo e copiar o VALOR imediatamente (ele so aparece uma vez)."

H2 "7.3 Conceder as permissoes de API"
P "Em API permissions, garantir as permissoes delegadas `openid`, `email`, `profile` e `User.Read` (Microsoft Graph), e conceder consentimento em nome da organizacao (`Grant admin consent`) se exigido pela politica do tenant."

H2 "7.4 Configurar o provedor no Supabase"
P "Em Authentication > Providers > Azure, habilitar o provedor e colar o Client ID e o Client Secret gerados no passo anterior."

H2 "7.5 Chamar o login a partir do frontend"
Exemplo "location.href = SUPABASE_URL + '/auth/v1/authorize?provider=azure&scopes=openid%20email%20profile%20User.Read';"
P "O Supabase cuida de todo o fluxo OAuth e devolve uma sessao valida do mesmo jeito que o login por link magico - o restante do codigo do frontend nao precisa distinguir como o usuario entrou."

# ============================================================
# 8. FASE 6 - DOMINIO PROPRIO
# ============================================================
H1 "8. Fase 6 - Dominio proprio e HTTPS"

H2 "8.1 Criar o arquivo CNAME no repositorio"
P "Na raiz do repositorio, criar um arquivo chamado exatamente `CNAME` (sem extensao), contendo uma unica linha com o dominio desejado:"
Exemplo "gestaoprojetos.suaempresa.com.br"

H2 "8.2 Apontar o DNS do dominio"
P "No provedor de DNS do dominio, criar um registro CNAME apontando o subdominio escolhido para o dominio padrao do GitHub Pages:"
Exemplo "gestaoprojetos.suaempresa.com.br  CNAME  <usuario-ou-organizacao>.github.io."

H2 "8.3 Confirmar no GitHub e habilitar HTTPS"
P "Em Settings > Pages, o GitHub detecta o CNAME e emite automaticamente um certificado HTTPS para o dominio (pode levar algumas horas). Assim que disponivel, marcar `Enforce HTTPS` para forcar conexoes seguras."

H2 "8.4 Atualizar a configuracao de URL no Supabase"
P "Em Authentication > URL Configuration, atualizar a `Site URL` e a lista de `Redirect URLs` para o novo dominio proprio, senao os links de login continuam apontando para o endereco antigo do GitHub Pages."

# ============================================================
# 9. FASE 7 - REDIRECIONAMENTO DO LINK MAGICO
# ============================================================
H1 "9. Fase 7 - Um problema real encontrado: redirecionamento do link magico"
P "Esta secao documenta um problema especifico deste projeto e a solucao aplicada - vale conhecer antes de descartar a possibilidade de que ele se repita numa nova aplicacao."

H2 "9.1 O problema"
P "Quando o site mora num subcaminho do GitHub Pages (`usuario.github.io/nome-do-repositorio/`, antes de um dominio proprio estar configurado, ou mesmo depois em alguns casos), o link enviado por e-mail as vezes e redirecionado apenas para a raiz do dominio de Pages (`usuario.github.io/`), perdendo o subcaminho do aplicativo - isso aconteceu de forma consistente neste projeto, mesmo com a Redirect URL cadastrada corretamente (exata ou com curinga) na configuracao do Supabase."

H2 "9.2 A solucao aplicada"
P "Foi criado um SEGUNDO repositorio no GitHub, com o nome exato `usuario.github.io` (a convencao especial do GitHub para `site pessoal/da organizacao`), contendo apenas um `index.html` que le o hash da URL recebida e redireciona (`location.replace`) para o endereco correto dentro do repositorio do aplicativo, preservando o hash (onde o Supabase embute o token de acesso)."
P "Todas as chamadas de login por link magico no frontend passaram a ser feitas SEM o parametro `email_redirect_to`, deixando o Supabase usar a Site URL configurada (que aponta para esse repositorio-relay) por padrao."
Nota "Se a nova aplicacao ja for publicada direto num dominio proprio configurado desde o inicio (Fase 6) e a Site URL do Supabase apontar exatamente para esse dominio, esse problema pode nao se repetir. Mas se aparecer (o sintoma e o link cair numa URL truncada, tipo raiz do dominio, sem o caminho esperado), este repositorio-relay e uma solucao ja validada."

# ============================================================
# 10. FASE 8 - ARMAZENAMENTO DE ARQUIVOS
# ============================================================
H1 "10. Fase 8 - Upload e download de arquivos (Storage)"

H2 "10.1 Criar um bucket por tipo de anexo"
P "Em Storage, criar um bucket para cada categoria de arquivo anexado (por exemplo, um bucket para anexos de um formulario e outro para anexos de outro), marcado como privado (nao publico), com limite de tamanho e tipos de arquivo permitidos definidos na criacao do bucket."

H2 "10.2 Criar as politicas de acesso do bucket"
P "Cada bucket precisa de politicas de RLS na tabela interna `storage.objects`, tipicamente restringindo SELECT, INSERT e DELETE a usuarios autenticados (`to authenticated`)."

H2 "10.3 Usar a API REST de Storage direto do frontend"
$rStorage = @(
  @("Acao","Chamada"),
  @("Enviar arquivo","POST /storage/v1/object/<bucket>/<caminho>"),
  @("Baixar arquivo","GET /storage/v1/object/authenticated/<bucket>/<caminho>"),
  @("Remover arquivo","DELETE /storage/v1/object/<bucket>/<caminho>")
)
TableSimple $rStorage @(4,12)
P "Neste projeto, o caminho de cada arquivo e composto por `id-do-registro/timestamp_nome-aleatorio_nome-original`, evitando colisao entre uploads e mantendo os arquivos de um mesmo registro agrupados."

# ============================================================
# 11. FASE 9 - FUNCOES DE BORDA
# ============================================================
H1 "11. Fase 9 - Funcoes de servidor (Edge Functions) e chaves de IA"
P "Como o site e estatico, qualquer operacao que exija um segredo de verdade (chave de API paga, senha de envio de e-mail) nao pode rodar no navegador do usuario - ela roda numa Edge Function do Supabase, que e codigo de servidor (Deno/TypeScript) hospedado pelo proprio Supabase."

H2 "11.1 Instalar a ferramenta de linha de comando do Supabase"
Exemplo "npm install -g supabase   (ou o instalador correspondente ao sistema operacional)"
Exemplo "supabase login"
Exemplo "supabase link --project-ref <ref-do-projeto>"

H2 "11.2 Criar cada funcao"
P "Cada funcao vive em `supabase/functions/<nome-da-funcao>/index.ts`. Este projeto usa 6 funcoes:"
$rFuncoes = @(
  @("Funcao","O que faz","Segredos que usa"),
  @("analisar-documento-demanda","Le um documento anexado e sugere o preenchimento de campos via IA","ANTHROPIC_API_KEY"),
  @("analisar-transcricao-ata","Le a transcricao de uma reuniao e sugere o preenchimento da ata via IA","ANTHROPIC_API_KEY"),
  @("sugerir-preenchimento","Le os documentos ja registrados de um projeto e sugere o preenchimento de outro formulario via IA","ANTHROPIC_API_KEY"),
  @("transcrever-audio-ata","Transcreve um audio de reuniao em texto","OPENAI_API_KEY"),
  @("send-notification","Envia e-mail e/ou notificacao push num unico ponto central, chamado nas mudancas de status","RESEND_API_KEY, VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY"),
  @("reset-treino","Restaura os dados de exemplo do ambiente de treinamento (uso interno, ver Fase 14)","TREINO_DB_URL")
)
TableSimple $rFuncoes @(3.8,7.2,5.2)
P "Todas as funcoes tambem recebem automaticamente `SUPABASE_URL` e `SUPABASE_SERVICE_ROLE_KEY` (a Supabase injeta essas duas em toda funcao, sem precisar configurar), usadas para validar o token do usuario que fez a chamada e consultar o papel dele antes de executar qualquer acao."

H2 "11.3 Definir os segredos (nunca no codigo)"
Exemplo "supabase secrets set ANTHROPIC_API_KEY=xxxx"
Exemplo "supabase secrets set OPENAI_API_KEY=xxxx"
Exemplo "supabase secrets set RESEND_API_KEY=xxxx"
Exemplo "supabase secrets set VAPID_PUBLIC_KEY=xxxx VAPID_PRIVATE_KEY=xxxx"

H2 "11.4 Publicar (deploy) cada funcao"
Exemplo "supabase functions deploy <nome-da-funcao>"

H2 "11.5 Chamar a funcao a partir do frontend"
P "O frontend chama `https://<ref-do-projeto>.supabase.co/functions/v1/<nome-da-funcao>` via fetch, sempre enviando o token de acesso do usuario logado no cabecalho `Authorization: Bearer <token>` - a funcao usa esse token para confirmar quem esta chamando e se essa pessoa tem permissao, antes de gastar a chave de IA ou enviar um e-mail."

# ============================================================
# 12. FASE 10 - NOTIFICACOES POR E-MAIL E PUSH
# ============================================================
H1 "12. Fase 10 - Notificacoes automaticas (e-mail e push)"

H2 "12.1 E-mail nas mudancas de status"
P "A funcao `send-notification` centraliza todo envio de e-mail do sistema, recebendo um payload generico `{ to, subject, html }` e usando a mesma `RESEND_API_KEY` ja configurada na Fase 4. Cada formulario chama essa funcao no exato momento em que o status relevante muda (por exemplo, uma aprovacao ou uma decisao registrada)."
Nota "Neste projeto os disparos de e-mail nao tem uma chave de liga-desliga separada - eles disparam sempre que o codigo chega naquele ponto. A unica forma de impedir o envio real e nao configurar o segredo `RESEND_API_KEY` naquele ambiente (e exatamente assim que o ambiente de treinamento, Fase 14, evita mandar e-mail de verdade)."

H2 "12.2 Notificacao push (PWA)"
P "Alem do e-mail, o sistema pode enviar notificacoes push para o celular/computador do usuario, mesmo com o site fechado, usando a Web Push API do proprio navegador."
Bul "Gerar um par de chaves VAPID (chave publica e chave privada) - existem geradores prontos (bibliotecas `web-push` em varias linguagens)."
Bul "A chave publica vai para o codigo do frontend (nao e segredo); a chave privada vira secret da Edge Function `send-notification` (`VAPID_PRIVATE_KEY`)."
Bul "Criar um `manifest.json` (deixa o site instalavel como aplicativo/PWA) e um `service-worker.js` (recebe o evento de push em segundo plano e mostra a notificacao mesmo com o site fechado)."
Bul "No frontend, pedir permissao de notificacao ao usuario e gravar a inscricao (`subscription`) retornada pelo navegador na tabela `push_subscriptions`, associada ao e-mail do usuario logado."
Nota "No iPhone/iPad (iOS), a API de notificacao push so fica disponivel depois que o usuario adiciona o site a Tela de Inicio (`Adicionar a Tela de Inicio`) - e uma limitacao do proprio iOS, nao do sistema."

# ============================================================
# 13. FASE 11 - CONSTRUCAO DO FRONTEND
# ============================================================
H1 "13. Fase 11 - Construcao do frontend"

H2 "13.1 Um arquivo HTML autossuficiente por formulario"
P "Cada formulario/pagina e um unico arquivo `.html` contendo seu proprio `<style>` e `<script>` - sem framework (React/Vue/etc.), sem etapa de build. Essa escolha prioriza simplicidade de manutencao (abrir um arquivo, editar, publicar) em troca de alguma duplicacao de codigo entre os arquivos."

H2 "13.2 Codigo compartilhado"
P "Apenas o essencial e compartilhado entre todas as paginas: `app.css` (estilo do menu/cabecalho comuns) e `app.js` (navegacao lateral). Ambos sao referenciados a partir de `index.html` com um parametro de versao na URL (`?v=N`), incrementado a cada alteracao para evitar que o navegador do usuario continue usando uma versao antiga em cache."

H2 "13.3 Autenticacao no frontend"
P "Cada arquivo HTML que exige login tem um bloco de interface `AUTH GATE` (a tela de login em si) e um pequeno modulo JavaScript (`unialfaAuth` neste projeto) que fala diretamente com a API REST do Supabase Auth via `fetch` - sem usar a biblioteca oficial `supabase-js`, para nao depender de um empacotador/build."

H2 "13.4 Um unico codigo para producao e treinamento"
P "Um teste simples de hostname (`location.href` contem `treino`?) troca, em tempo de execucao, a URL/chave do Supabase e libera a opcao extra de login por senha fixa usada so no ambiente de treinamento - o mesmo arquivo HTML roda nos dois ambientes sem duplicacao de codigo (ver Fase 14)."

# ============================================================
# 14. FASE 12 - DEPLOY CONTINUO
# ============================================================
H1 "14. Fase 12 - Fluxo de publicacao continua"
P "Nao ha um passo de `publicar` separado: publicar o sistema e simplesmente enviar o codigo para o GitHub."
Bul "Editar o(s) arquivo(s) necessario(s)."
Bul "`git add` dos arquivos alterados."
Bul "`git commit` com uma mensagem descrevendo a mudanca."
Bul "`git push` para a branch main."
Bul "O workflow do GitHub Actions (Fase 1) publica automaticamente em 1 a 2 minutos."
Nota "Sempre que `app.css` ou `app.js` mudam, e preciso incrementar o numero de versao (`?v=N`) no lugar onde sao referenciados, senao alguns navegadores continuam servindo a versao antiga em cache por ate alguns minutos."

# ============================================================
# 15. FASE 13 - AMBIENTE DE TREINAMENTO
# ============================================================
H1 "15. Fase 13 - Ambiente de treinamento/demonstracao (recomendado, opcional)"
P "Um ambiente separado, com dados ficticios, permite demonstrar e treinar o uso do sistema sem qualquer risco de alterar dados reais de producao."

H2 "15.1 Repositorio espelho"
P "Criar um segundo repositorio publico no GitHub (GitHub Pages gratuito nao publica repositorio privado), sem arquivo `CNAME` (usa o endereco padrao do GitHub Pages em vez de um dominio proprio), com o mesmo codigo do repositorio de producao."

H2 "15.2 Projeto Supabase separado"
P "Criar um segundo projeto no Supabase, replicando exatamente o mesmo schema de tabelas, as mesmas politicas de RLS e a mesma funcao `is_admin()` (Fase 3)."

H2 "15.3 Publicar as funcoes SEM os segredos reais"
P "Publicar as mesmas Edge Functions necessarias no projeto de treinamento, mas deliberadamente sem configurar `RESEND_API_KEY`, `VAPID_*`, `ANTHROPIC_API_KEY` ou `OPENAI_API_KEY` nesse projeto - essa e a real trava de seguranca que impede o ambiente de treinamento de mandar e-mail de verdade, notificacao push de verdade ou gastar uma chamada de IA paga."

H2 "15.4 Contas de teste fixas"
P "Criar algumas contas de login fixo (usuario/senha), uma por papel de usuario, habilitadas apenas quando o codigo detecta o hostname de treinamento (Fase 11.4) - alem do login por link magico e do login com Microsoft, que tambem continuam funcionando no ambiente de treino se configurados."

H2 "15.5 Dados de exemplo"
P "Preparar um script SQL de carga (seed) com alguns registros ficticios cobrindo o ciclo de vida completo do sistema, e um script de reset que apaga apenas os dados de exemplo (preservando as contas fixas e as configuracoes) para poder recomecar uma demonstracao do zero."

# ============================================================
# 16. FASE 14 - DOCUMENTACAO VIVA
# ============================================================
H1 "16. Fase 14 - Documentacao viva do sistema"
P "Um sistema sem codigo comentado para usuario final precisa de documentacao mantida em dia. Este projeto usa tres camadas, todas atualizadas junto com qualquer mudanca de funcionalidade:"
Bul "Um manual de uso (.docx) com o passo a passo de cada tela, incluindo capturas de tela."
Bul "Um documento de regras de acesso e permissoes (.docx), separado do manual, focado exclusivamente em quem pode fazer o que."
Bul "Uma versao navegavel (com busca) dentro do proprio site, aberta pelo icone de ajuda de cada formulario."

H2 "16.1 Como os dois .docx sao gerados"
P "Como o ambiente de manutencao nao tem Node.js, pandoc nem LibreOffice instalados - apenas o Microsoft Word - os dois documentos oficiais sao gerados por um script PowerShell que usa automacao COM do Word (`New-Object -ComObject Word.Application`) para montar o documento inteiro programaticamente (capa, sumario automatico, secoes, tabelas e imagens) e salvar em `.docx`."
Nota "Ao reexecutar esses scripts para incluir capturas de tela novas, sempre apontar para a MESMA pasta de imagens ja usada anteriormente (nunca uma pasta vazia/nova) - senao o documento gerado perde todas as imagens antigas silenciosamente."

H2 "16.2 Disciplina de manutencao"
P "A regra adotada neste projeto: toda vez que uma funcionalidade e incluida, alterada ou removida no codigo, os tres materiais (manual, regras de acesso e a versao navegavel no site) sao revisados no mesmo commit da mudanca, ou logo em seguida - essa regra fica escrita no proprio repositorio (arquivo `CLAUDE.md`) para orientar qualquer pessoa (ou assistente de IA) que for alterar o codigo depois."

# ============================================================
# 17. RESUMO FINAL
# ============================================================
H1 "17. Resumo final - checklist de construcao"
P "Ordem cronologica resumida, do zero ate um sistema publicado e funcional:"
Bul "1. Criar o repositorio no GitHub e habilitar o GitHub Pages via Actions (Fase 1)."
Bul "2. Criar o projeto no Supabase e anotar URL + chave publica (Fase 2)."
Bul "3. Criar as tabelas, habilitar RLS e escrever as politicas de acesso, comecando pela funcao is_admin() (Fase 3)."
Bul "4. Configurar o login por e-mail: SMTP proprio via Resend e VALIDAR o dominio de envio (Fase 4)."
Bul "5. Configurar o login com Microsoft: registrar o app no Azure AD e ligar ao Supabase (Fase 5)."
Bul "6. Apontar um dominio proprio (opcional) e atualizar a URL Configuration do Supabase (Fase 6)."
Bul "7. Testar o link magico de ponta a ponta e resolver qualquer problema de redirecionamento (Fase 7)."
Bul "8. Criar os buckets de armazenamento de arquivos e suas politicas (Fase 8)."
Bul "9. Publicar as Edge Functions e configurar os segredos de IA/e-mail/push (Fase 9)."
Bul "10. Ligar os disparos de e-mail e (opcionalmente) notificacao push nos pontos certos do fluxo (Fase 10)."
Bul "11. Construir as paginas do frontend, com o padrao de autenticacao compartilhado (Fase 11)."
Bul "12. Deploy passa a ser so `git push` a partir daqui (Fase 12)."
Bul "13. (Opcional) Montar um ambiente de treinamento espelho, com Edge Functions sem segredos reais (Fase 13)."
Bul "14. Criar e manter a documentacao viva do sistema (Fase 14)."

H2 "17.1 O que precisa de decisao humana (nao da para automatizar)"
Bul "Aprovacao/registro do dominio da empresa e acesso ao DNS (depende de TI/infraestrutura)."
Bul "Permissao de administrador no Azure AD/Microsoft 365 para registrar o app de login com Microsoft."
Bul "Definicao de quem paga e mantem as chaves de API da Anthropic, OpenAI e Resend (cartao/contrato)."
Bul "Modelagem das regras de negocio especificas da nova aplicacao (papeis de usuario, fluxo de aprovacao, campos de cada formulario) - tudo isso e especifico do dominio do novo sistema, nao e reaproveitavel diretamente deste."

$word.ActiveDocument.TablesOfContents.Item(1).Update() | Out-Null
$doc.SaveAs2($OutPath, 16)
$doc.Close()
$word.Quit()
Write-Output "SAVED: $OutPath"
