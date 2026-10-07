param([string]$OutPath = (Join-Path $PSScriptRoot "Guia de Migracao para Servidor Proprio - SGP.docx"))
$ErrorActionPreference = "Stop"

# Gera "Guia de Migracao para Servidor Proprio - SGP.docx" via automacao COM do Microsoft Word
# (Node.js/pandoc/LibreOffice nao estao disponiveis neste ambiente de geracao do documento).
# Documento de planejamento: como levar o sistema, como ele esta hoje (site + banco de dados +
# autenticacao + storage + funcoes de IA), da nuvem (GitHub Pages + Supabase Cloud) para um
# servidor da propria empresa, mantendo o mesmo comportamento.

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
$footer.Range.Text = "SGP - Guia de Migracao para Servidor Proprio | Pagina "
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
  $sel.Font.Size = 9.5
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
$sel.Font.Name="Montserrat"; $sel.Font.Size = 25; $sel.Font.Bold = $true; $sel.Font.Color = $colInk
$sel.ParagraphFormat.SpaceAfter = 4
$sel.TypeText("GUIA DE MIGRACAO PARA SERVIDOR PROPRIO")
$sel.TypeParagraph()
$sel.Font.Bold = $false
P "Sistema de Gestao de Projetos (SGP)" 15 $false $false $colInk "left" 4
P "Como levar o sistema, exatamente como esta hoje - site, banco de dados, autenticacao, arquivos anexados e funcoes de IA - da nuvem (GitHub Pages + Supabase Cloud) para um servidor da propria empresa, preservando o mesmo comportamento." 12 $false $true $colMuted "left" 30
P "Andre Torres - Gerente de Projetos" 11 $false $false $colMuted "left" 2
P "Documento de planejamento tecnico - versao 1.0 - 09 de setembro de 2026" 11 $false $false $colMuted "left" 2

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
H1 "1. Visao geral desta migracao"

H2 "1.1 O que este documento cobre"
P "Hoje o sistema roda inteiramente em servicos de terceiros: o site (as paginas HTML) e publicado pelo GitHub Pages, e todo o backend - banco de dados, login, armazenamento de arquivos e as funcoes que chamam IA e enviam e-mail - roda no Supabase Cloud (projeto de producao com identificador fiarntunpqteopwjkhjg). Este documento descreve como migrar TUDO isso para dentro de um servidor da propria empresa: o site e o Supabase inteiro (banco Postgres, autenticacao, storage e as 6 funcoes de servidor), auto-hospedados via Docker."
P "Esta e a opcao de maior escopo entre as duas formas possiveis de migrar o sistema. A alternativa mais simples - mover so a hospedagem do site e manter o banco de dados na nuvem - nao e o que este documento cobre, pois foi decidido levar o backend completo tambem para o servidor da empresa."

H2 "1.2 O que continua igual, do ponto de vista de quem usa o sistema"
Bul "O mesmo endereco (gestaoprojetos.suaempresa.com.br), as mesmas telas, os mesmos formularios, exatamente como sao hoje - nenhuma tela precisa ser redesenhada ou reescrita."
Bul "O mesmo login (link magico por e-mail e Entrar com Microsoft), os mesmos usuarios e papeis ja cadastrados."
Bul "Os mesmos projetos, registros e historico ja existentes hoje - nada e perdido, tudo e migrado junto."
Bul "Os mesmos anexos ja enviados em Solicitacoes de Demanda e SMPs."
Bul "As mesmas notificacoes por e-mail e os mesmos recursos de IA (sugestao de preenchimento, analise de documento, transcricao de audio)."

H2 "1.3 O que muda de fato (mesmo com tudo funcionando igual para quem usa)"
P "Migrar de um servico gerenciado (Supabase Cloud) para uma instalacao propria (self-hosted) preserva o COMPORTAMENTO do sistema, mas transfere para a equipe de TI da empresa varias responsabilidades que hoje sao automaticas:"
Bul "Backup do banco de dados: hoje o Supabase Cloud faz backup automatico; apos a migracao, isso passa a ser uma rotina que a propria empresa precisa configurar e monitorar (Fase 12)."
Bul "Atualizacao de seguranca: hoje o Supabase atualiza a infraestrutura por conta propria; apos a migracao, atualizar as imagens Docker do banco, da autenticacao, do storage etc. passa a ser tarefa da equipe."
Bul "Disponibilidade (uptime): hoje o Supabase e o GitHub garantem o site no ar; apos a migracao, isso depende do servidor, da rede e da energia da propria empresa."
Bul "Certificado HTTPS: hoje e automatico (GitHub Pages); apos a migracao, precisa ser configurado e renovado (Fase 9)."
Bul "Escala: o Supabase Cloud escala sozinho conforme o uso cresce; num servidor proprio, crescer exige planejar mais capacidade de hardware."
Nota "Este documento assume que a empresa esta ciente dessas mudancas de responsabilidade e decidiu seguir mesmo assim (razoes tipicas: exigencia de manter os dados dentro da infraestrutura da propria empresa, ou politica interna de nao depender de servicos externos). Se o unico objetivo fosse reduzir dependencia de terceiros no site em si, migrar so o frontend e manter o Supabase Cloud seria uma opcao de risco e esforco muito menores."

H2 "1.4 Recomendacao de metodo de trabalho"
P "Migrar em paralelo, nunca `no escuro`: o ambiente novo (servidor da empresa) deve ser construido e testado de ponta a ponta enquanto o ambiente atual (GitHub Pages + Supabase Cloud) continua no ar normalmente. So depois que TODOS os testes da Fase 10 passarem no ambiente novo e que se faz a troca final (Fase 11) - que, alem disso, deve ser reversivel por um tempo (plano de rollback)."

# ============================================================
# 2. INVENTARIO
# ============================================================
H1 "2. Inventario - tudo que existe hoje em producao e precisa ser levado"
P "Antes de comecar, confirmar que esta lista bate com a realidade atual do projeto Supabase de producao (nao presumir - conferir cada item no proprio painel do Supabase e no repositorio GitHub)."
$rInv = @(
  @("O que e","Onde esta hoje","O que precisa acontecer na migracao"),
  @("Codigo do site (18 paginas HTML + app.js/app.css)","Repositorio GitHub Paramiri/unialfa-gestao-projetos","Publicar no servidor da empresa; trocar a URL/chave do Supabase em cada arquivo (Fase 8)"),
  @("Banco de dados (10 tabelas, politicas de RLS, funcao is_admin)","Postgres do projeto Supabase fiarntunpqteopwjkhjg","Exportar e importar no Postgres do servidor novo (Fase 3)"),
  @("Contas de usuario ja cadastradas (perfis, autenticacao)","Supabase Auth do projeto de producao","Migram junto com o banco; cada pessoa precisa entrar de novo uma vez (Fase 3 e 6)"),
  @("Configuracao de login por e-mail (SMTP Resend)","Supabase Auth > Emails > SMTP Settings","Reconfigurar as mesmas credenciais no ambiente novo (Fase 6)"),
  @("Configuracao de login com Microsoft (Azure AD)","Supabase Auth > Providers > Azure","Reconfigurar com um novo Redirect URI apontando pro servidor novo (Fase 6)"),
  @("Arquivos anexados (buckets anexos-demanda e anexos-smp)","Supabase Storage do projeto de producao","Baixar todos os arquivos e reenviar no Storage novo (Fase 4)"),
  @("6 funcoes de servidor (Edge Functions) e seus segredos","Supabase Functions do projeto de producao","Copiar o codigo e reconfigurar os segredos no ambiente novo (Fase 5)"),
  @("Dominio proprio (gestaoprojetos.suaempresa.com.br)","DNS apontando para o GitHub Pages","Reapontar o DNS para o servidor da empresa (Fase 9)"),
  @("Ambiente de treinamento (repositorio + projeto Supabase separados)","GitHub + Supabase Cloud (projeto uuxvdulunrwppbmofyux)","Decisao em aberto - ver nota abaixo")
)
TableSimple $rInv @(4.8,5,6.2)
Nota "Ambiente de treinamento: como ele usa dados ficticios e nao tem nenhum segredo real configurado (e assim de proposito - ver o Guia de Replicacao Tecnica, Fase 13), recomenda-se MANTE-LO na nuvem mesmo depois da migracao do sistema principal - baixo custo, baixo risco, e nao ha motivo de conformidade/dado sensivel para justificar move-lo tambem. Esta e uma recomendacao, nao uma obrigacao: o mesmo processo deste documento pode ser repetido para o ambiente de treinamento se a empresa preferir migrar os dois."

# ============================================================
# 3. REQUISITOS DO SERVIDOR
# ============================================================
H1 "3. Requisitos do servidor da empresa"

H2 "3.1 Hardware minimo recomendado"
$rHw = @(
  @("Recurso","Minimo recomendado pela Supabase para o kit self-hosted"),
  @("Memoria (RAM)","4 GB (recomendado 8 GB ou mais, considerando IA e uso real, nao so o minimo de instalacao)"),
  @("Processador","2 nucleos (recomendado 4 ou mais)"),
  @("Armazenamento em disco","40 GB livres (mais espaco proporcional ao volume de anexos e ao tamanho do banco de dados ao longo do tempo)"),
  @("Sistema operacional","Linux (recomendado) com Docker Engine e Docker Compose; ou Windows Server com Docker Desktop/WSL2 (ver nota abaixo)")
)
TableSimple $rHw @(4.5,11.5)
Nota "O kit oficial de auto-hospedagem da Supabase e desenhado e documentado para Linux. Rodar em Windows Server exige Docker Desktop com o back-end WSL2 (Subsistema Windows para Linux) habilitado - funciona, mas e um caminho menos testado pela propria Supabase para producao. Se a empresa tiver a opcao de escolher, um servidor Linux (Ubuntu Server LTS, por exemplo) e o caminho mais direto e mais documentado."

H2 "3.2 Software necessario"
Bul "Docker Engine + Docker Compose (ou Docker Desktop, no caso de Windows)"
Bul "Git"
Bul "Supabase CLI (usada para exportar o banco de dados atual - Fase 3)"
Bul "Um servidor web para publicar o site: IIS (se o servidor for Windows) ou Nginx/Apache (se for Linux) - ver Fase 7"
Bul "Um proxy reverso com certificado HTTPS na frente de tudo (pode ser o proprio IIS/Nginx, ou uma ferramenta dedicada como Caddy) - ver Fase 9"

H2 "3.3 Rede"
Bul "Um dominio (o mesmo gestaoprojetos.suaempresa.com.br ja usado hoje, ou um novo) apontando para o IP do servidor."
Bul "Portas liberadas no firewall da empresa para acesso externo: 443 (HTTPS) obrigatoria; 80 (HTTP) recomendada so para redirecionar automaticamente para 443."
Bul "Se o Supabase auto-hospedado ficar num endereco/porta separada do site (comum no kit oficial, que expoe tudo por padrao na porta 8000), decidir se ele fica acessivel so pela rede interna ou tambem pela internet (o site publico PRECISA conseguir alcanca-lo para funcionar, entao normalmente fica exposto atras do mesmo dominio, num subcaminho ou subdominio dedicado - ex.: api.gestaoprojetos.suaempresa.com.br)."

# ============================================================
# 4. FASE 1
# ============================================================
H1 "4. Fase 1 - Provisionar o servidor e instalar o Docker"
Bul "Provisionar a maquina (fisica ou virtual) com o sistema operacional escolhido (Fase 3.1)."
Bul "Instalar o Docker Engine e o Docker Compose (em Linux) ou o Docker Desktop com WSL2 (em Windows)."
Bul "Instalar o Git."
Bul "Instalar a Supabase CLI (usada na Fase 3 para exportar o banco de dados de producao)."
Bul "Confirmar que o servidor consegue acessar a internet (para baixar as imagens Docker) e que a rede da empresa consegue alcancar o servidor nas portas definidas na Fase 3.3."

# ============================================================
# 5. FASE 2
# ============================================================
H1 "5. Fase 2 - Instalar o Supabase auto-hospedado"
P "O proprio Supabase distribui um kit oficial de auto-hospedagem baseado em Docker Compose, com todos os servicos necessarios: banco de dados Postgres, Auth (login), PostgREST (API REST do banco), Storage (arquivos), Realtime e o runtime das Edge Functions, tudo atras de um gateway (Envoy) numa unica porta."

H2 "5.1 Baixar o kit oficial"
Exemplo "git clone --depth 1 --branch self-hosted/v0.8.1 https://github.com/supabase/supabase"
Exemplo "mkdir supabase-project"
Exemplo "cp -rf supabase/docker/. supabase-project"
Exemplo "cd supabase-project && cp .env.example .env"
Nota "Conferir no repositorio oficial (github.com/supabase/supabase) qual e a branch/tag self-hosted mais recente no momento da instalacao - o numero de versao acima (v0.8.1) e o vigente na data deste documento e pode ter mudado."

H2 "5.2 Gerar as chaves e preencher o .env"
P "O kit inclui scripts prontos para gerar os segredos necessarios (chave JWT, chaves publica e secreta da API):"
Exemplo "sh utils/generate-keys.sh"
Exemplo "sh utils/add-new-auth-keys.sh"
P "Depois, editar o arquivo .env e preencher, no minimo, os campos abaixo:"
$rEnv = @(
  @("Variavel","O que colocar"),
  @("POSTGRES_PASSWORD","Uma senha nova e forte, so letras e numeros (evita problema de codificacao)"),
  @("SUPABASE_PUBLISHABLE_KEY / SUPABASE_SECRET_KEY","Geradas pelos scripts acima - substituem a chave publica e a chave secreta usadas hoje no projeto Supabase Cloud"),
  @("SUPABASE_PUBLIC_URL","O endereco publico do backend novo (ex.: https://api.gestaoprojetos.suaempresa.com.br)"),
  @("API_EXTERNAL_URL","O mesmo endereco publico, seguido de /auth/v1 (usado no fluxo de login)"),
  @("SITE_URL","O endereco do site (ex.: https://gestaoprojetos.suaempresa.com.br) - para onde o login redireciona por padrao"),
  @("DASHBOARD_USERNAME / DASHBOARD_PASSWORD","Login de acesso ao Studio (painel administrativo do banco, equivalente ao painel do Supabase Cloud)"),
  @("SMTP_HOST / SMTP_PORT / SMTP_USER / SMTP_PASS","As MESMAS credenciais do Resend ja usadas hoje (Fase 6)")
)
TableSimple $rEnv @(5.3,10.7)

H2 "5.3 Subir os containers e validar"
Exemplo "sh run.sh start"
P "Aguardar cerca de um minuto e conferir se todos os servicos aparecem como saudaveis:"
Exemplo "docker compose ps"
P "Depois, acessar o Studio (painel administrativo) pelo navegador, no endereco do servidor na porta configurada (por padrao 8000), e entrar com o DASHBOARD_USERNAME/DASHBOARD_PASSWORD definidos no .env - se a tela carregar e mostrar um banco de dados vazio, o ambiente esta pronto para receber os dados migrados."

# ============================================================
# 6. FASE 3
# ============================================================
H1 "6. Fase 3 - Migrar o banco de dados"
P "Aqui e onde todo o conteudo ja existente - projetos cadastrados, usuarios, historico, configuracoes, tudo o que esta hoje na producao - e efetivamente transferido para o servidor novo. Usar sempre a Supabase CLI (nao o pg_dump direto), pois ela aplica filtros proprios do Supabase (remove schemas internos, ajusta papeis reservados)."

H2 "6.1 Exportar do projeto de producao atual"
P "A partir de uma maquina com a Supabase CLI instalada e a string de conexao do banco de producao (disponivel no painel do projeto fiarntunpqteopwjkhjg em Project Settings > Database):"
Exemplo "supabase db dump --db-url `"[STRING_DE_CONEXAO]`" -f roles.sql --role-only"
Exemplo "supabase db dump --db-url `"[STRING_DE_CONEXAO]`" -f schema.sql"
Exemplo "supabase db dump --db-url `"[STRING_DE_CONEXAO]`" -f data.sql --use-copy --data-only"
P "Isso gera 3 arquivos: os papeis/permissoes do banco, a estrutura (tabelas, politicas de RLS, a funcao is_admin, indices) e os dados propriamente ditos."

H2 "6.2 Importar no banco novo (self-hosted)"
Exemplo "psql --single-transaction --variable ON_ERROR_STOP=1 --file roles.sql --file schema.sql --command 'SET session_replication_role = replica' --file data.sql --dbname `"postgres://postgres:[SENHA_DO_ENV]@[ENDERECO_DO_SERVIDOR]:5432/postgres`""
P "O comando SET session_replication_role = replica desativa temporariamente os gatilhos (triggers) durante a importacao dos dados, evitando efeitos colaterais (como duplicar operacoes de criptografia de senha) - depois da importacao, o comportamento normal do banco volta sozinho."

H2 "6.3 Conferir depois de importar"
Bul "Abrir o Studio do ambiente novo e confirmar que as 10 tabelas existem, com o mesmo numero de linhas do ambiente de producao (comparar um `select count(*)` de cada tabela nos dois ambientes)."
Bul "Confirmar que a funcao is_admin() existe e que RLS esta habilitado nas 10 tabelas."
Bul "Confirmar, no Authentication do Studio novo, que os usuarios ja cadastrados aparecem na lista."
Nota "O ambiente novo gera um segredo de assinatura (JWT) proprio, diferente do usado na nuvem. Isso significa que toda sessao/token de login ja emitido deixa de valer no ambiente novo - as CONTAS continuam existindo (o cadastro veio junto no banco), mas cada pessoa precisa fazer login de novo uma vez apos a migracao. Isso e esperado e nao e perda de dado."
Nota "Se o projeto de producao na nuvem estiver rodando uma versao do Postgres mais nova do que a versao padrao do kit self-hosted, considerar subir o ambiente novo ja na mesma versao do Postgres da nuvem (ajustavel no .env/docker-compose antes do primeiro `sh run.sh start`), para evitar incompatibilidade de schema."

# ============================================================
# 7. FASE 4
# ============================================================
H1 "7. Fase 4 - Migrar os arquivos anexados (Storage)"
P "Os arquivos enviados pelos usuarios (anexos da Solicitacao de Demanda e da SMP) ficam fora do banco de dados, no servico de Storage - por isso precisam de uma transferencia separada da Fase 3."

H2 "7.1 Metodo recomendado: copiar via API"
P "Como o codigo deste sistema ja fala com o Storage direto pela API REST (sem nenhuma biblioteca/SDK), o caminho mais simples e escrever um pequeno script que, para cada um dos 2 buckets (anexos-demanda e anexos-smp):"
Bul "Lista todos os arquivos do bucket no ambiente de producao (GET /storage/v1/object/list/<bucket>, com a chave de servico do projeto de producao)."
Bul "Baixa cada arquivo (GET /storage/v1/object/authenticated/<bucket>/<caminho>)."
Bul "Recria o mesmo bucket no ambiente novo (mesmo nome, mesma configuracao de privado/publico e limite de tamanho) e reenvia cada arquivo no mesmo caminho (POST /storage/v1/object/<bucket>/<caminho>, com a chave de servico do ambiente novo)."
Nota "Preservar exatamente o mesmo caminho de cada arquivo (o formato usado hoje e id-do-registro/timestamp_nome-aleatorio_nome-original) e essencial: e esse caminho que fica gravado dentro de cada registro do banco de dados (kv_store) - se o caminho mudar na copia, o link do anexo dentro do sistema quebra mesmo com o arquivo tecnicamente migrado."

H2 "7.2 Alternativa para volumes grandes"
P "Se o volume de arquivos for grande, em vez do backend de arquivos local (padrao do kit self-hosted, que guarda tudo em disco dentro da pasta de volumes), o Storage tambem pode ser configurado para usar um armazenamento compativel com S3 (por exemplo, um MinIO instalado no proprio servidor da empresa) - nesse caso, ferramentas como o rclone permitem sincronizar os arquivos direto entre o bucket da nuvem e o novo destino S3-compativel, sem passar arquivo por arquivo pela API. Para o volume atual do sistema (dois buckets de anexos de formularios), o metodo via API da Fase 7.1 e suficiente e mais simples de auditar."

H2 "7.3 Recriar as politicas de acesso"
P "Cada bucket precisa das mesmas politicas de RLS de storage.objects ja usadas hoje (SELECT/INSERT/DELETE restritos a authenticated) - reescrever essas politicas no SQL Editor do Studio novo, ja que elas nao vem automaticamente so pela copia dos arquivos."

# ============================================================
# 8. FASE 5
# ============================================================
H1 "8. Fase 5 - Migrar as funcoes de servidor (Edge Functions)"

H2 "8.1 Copiar o codigo de cada funcao"
P "No kit self-hosted, cada funcao fica numa pasta propria dentro de volumes/functions/<nome-da-funcao>/index.ts. Copiar o codigo das 6 funcoes ja existentes no repositorio de producao (pasta supabase/functions/ do repositorio GitHub) para dentro dessa mesma estrutura no servidor novo."
$rFuncoes = @(
  @("Funcao","Segredos que precisa"),
  @("analisar-documento-demanda","ANTHROPIC_API_KEY"),
  @("analisar-transcricao-ata","ANTHROPIC_API_KEY"),
  @("sugerir-preenchimento","ANTHROPIC_API_KEY"),
  @("transcrever-audio-ata","OPENAI_API_KEY"),
  @("send-notification","RESEND_API_KEY, VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY"),
  @("reset-treino","TREINO_DB_URL (so relevante se o ambiente de treinamento tambem for migrado)")
)
TableSimple $rFuncoes @(5,11)

H2 "8.2 Configurar os segredos no ambiente novo"
P "No kit self-hosted, os segredos customizados das funcoes ficam num arquivo de ambiente proprio dentro de volumes/functions/ (o mesmo padrao usado no desenvolvimento local do Supabase), carregado automaticamente pelo servico de functions do docker-compose. Adicionar la as MESMAS chaves ja usadas hoje (ANTHROPIC_API_KEY, OPENAI_API_KEY, RESEND_API_KEY, VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY) - podem ser exatamente as mesmas chaves de API ja pagas e em uso, nao ha necessidade de gerar chaves novas nos provedores (Anthropic/OpenAI/Resend), so de configura-las no novo local."
Nota "Confirmar o nome exato do arquivo de variaveis de ambiente das funcoes na versao do kit self-hosted baixada na Fase 2 (o comentario dentro do proprio docker-compose.yml do kit indica o caminho) - esse detalhe pode variar entre versoes do kit."

H2 "8.3 Aplicar e reiniciar"
Exemplo "sh run.sh restart functions"
P "Sempre que o codigo de uma funcao mudar, o comando acima recarrega. Se o que mudou foi uma variavel de ambiente/segredo, usar em vez disso:"
Exemplo "sh run.sh recreate functions"

H2 "8.4 Testar cada funcao isoladamente"
P "Antes de seguir, chamar cada uma das 6 funcoes manualmente (por exemplo, via um cliente HTTP como o Postman/Insomnia, ou um teste simples dentro de uma das telas do sistema apontada para o ambiente novo) e confirmar que cada uma responde como esperado, sem erro de segredo ausente."

# ============================================================
# 9. FASE 6
# ============================================================
H1 "9. Fase 6 - Reconfigurar a autenticacao"

H2 "9.1 Login por e-mail (SMTP)"
P "Como o SMTP ja foi preenchido no .env na Fase 2.2 com as mesmas credenciais do Resend, o login por link magico deve funcionar sem nenhuma mudanca adicional de configuracao - o dominio de envio (ex.: sistemas.suaempresa.com.br) ja esta verificado no Resend e nao precisa ser verificado de novo, pois a verificacao e do lado do Resend, nao do Supabase."
Bul "Conferir, no Studio novo, em Authentication > Templates, se os templates em portugues (Confirm signup, Magic Link) precisam ser recriados manualmente - eles nao vem automaticamente do banco de dados migrado, pois ficam configurados a parte."

H2 "9.2 Login com Microsoft (Azure AD)"
P "O login com Microsoft precisa de um Redirect URI novo, ja que o endereco de callback do backend mudou de <ref>.supabase.co para o dominio do servidor da empresa."
Bul "No Azure Portal, no mesmo app registration ja usado hoje (ou um novo, se a organizacao preferir separar), adicionar um Redirect URI adicional apontando para o endereco novo:"
Exemplo "https://api.gestaoprojetos.suaempresa.com.br/auth/v1/callback"
Bul "No .env do ambiente novo, preencher AZURE_ENABLED=true, AZURE_CLIENT_ID e AZURE_SECRET com os mesmos valores ja usados hoje (o Client ID nao muda; o Client Secret pode ser reaproveitado se ainda estiver valido, ou gerado um novo em Certificates and secrets)."
Bul "Reiniciar o servico de autenticacao para aplicar (sh run.sh recreate auth, ou o comando equivalente indicado pelo kit baixado)."
Nota "Manter o Redirect URI antigo (apontando para o Supabase Cloud) cadastrado no Azure ate a migracao estar totalmente concluida e validada - assim o login com Microsoft continua funcionando nos dois ambientes durante o periodo de testes em paralelo (secao 1.4)."

# ============================================================
# 10. FASE 7
# ============================================================
H1 "10. Fase 7 - Publicar o site no servidor da empresa"
P "O site continua sendo um conjunto de arquivos estaticos (HTML/CSS/JS) - so muda ONDE esses arquivos sao servidos. Duas opcoes, conforme o sistema operacional do servidor escolhido (secao 3.1)."

H2 "10.1 Opcao Windows Server (IIS)"
Bul "Instalar o papel Web Server (IIS) pelo Gerenciador do Servidor, se ainda nao estiver instalado."
Bul "Criar um novo site no IIS Manager, apontando a pasta fisica para uma copia local do repositorio (todos os arquivos .html, .css, .js, a pasta icons/ etc.)."
Bul "Definir index.html como documento padrao do site."
Bul "Configurar a vinculacao (binding) HTTPS na porta 443 com o certificado do dominio (Fase 9)."

H2 "10.2 Opcao Linux (Nginx)"
P "Instalar o Nginx e criar um arquivo de configuracao de site apontando para a pasta com os arquivos do repositorio, por exemplo:"
Exemplo "server { listen 443 ssl; server_name gestaoprojetos.suaempresa.com.br; root /var/www/sgp; index index.html; }"
Bul "Copiar todos os arquivos do repositorio para o caminho definido em root (ex.: /var/www/sgp)."
Bul "Apontar o certificado HTTPS (Fase 9) nas diretivas ssl_certificate / ssl_certificate_key."
Bul "Recarregar o Nginx apos qualquer mudanca de configuracao (nginx -s reload)."

H2 "10.3 Como os arquivos chegam no servidor"
P "Manter o mesmo fluxo de hoje o quanto for possivel: o codigo continua vivendo no repositorio Git (GitHub ou um Git interno da empresa, se preferirem trazer isso tambem para dentro), e uma rotina simples (manual, ou um script agendado) sincroniza a ultima versao da branch principal para a pasta que o IIS/Nginx esta servindo - por exemplo, um git pull agendado, ou um pequeno pipeline que copia os arquivos a cada atualizacao."

# ============================================================
# 11. FASE 8
# ============================================================
H1 "11. Fase 8 - Trocar as credenciais do Supabase no codigo do site"
P "Cada uma das 18 paginas HTML do sistema, mais o arquivo push-notifications.js, tem a URL e a chave publica do Supabase escritas diretamente no codigo (nao ha um arquivo de configuracao central). Apos concluir as Fases 2 a 6, e preciso substituir, em TODOS esses arquivos, o endereco antigo pelo endereco novo:"
$rTroca = @(
  @("Onde estava (producao na nuvem)","Onde passa a apontar (servidor da empresa)"),
  @("https://fiarntunpqteopwjkhjg.supabase.co","https://api.gestaoprojetos.suaempresa.com.br  (ou o endereco escolhido para o backend novo)"),
  @("sb_publishable_RO-UPexCYhZ0rVZiIYWunA_a7YqodnQ","A nova SUPABASE_PUBLISHABLE_KEY gerada na Fase 2.2")
)
TableSimple $rTroca @(8,8)
P "Como as duas strings aparecem repetidas em cada um dos ~19 arquivos, a forma mais segura de trocar e uma busca-e-substituicao em lote (find-and-replace) em todos os arquivos de uma vez, seguida de uma conferencia rapida (buscar o endereco antigo de novo depois da troca, para confirmar que nao sobrou nenhuma ocorrencia)."
Nota "O endereco e a chave do ambiente de TREINAMENTO (a outra metade do if/else presente em cada arquivo, usado quando o hostname contem `treino`) NAO precisam mudar, caso a decisao (secao 2, nota) seja manter o ambiente de treinamento na nuvem."

# ============================================================
# 12. FASE 9
# ============================================================
H1 "12. Fase 9 - Dominio e HTTPS"

H2 "12.1 Redirecionar o dominio"
P "No provedor de DNS do dominio gestaoprojetos.suaempresa.com.br, trocar o registro que hoje aponta para o GitHub Pages por um registro apontando para o IP publico do servidor da empresa (registro A) ou para o nome do servidor (registro CNAME, se aplicavel)."
Bul "Se o backend novo ficar num subdominio proprio (ex.: api.gestaoprojetos.suaempresa.com.br), criar tambem esse registro apontando para o mesmo servidor (ou outro, se o backend rodar numa maquina separada do site)."

H2 "12.2 Certificado HTTPS"
P "Diferente do GitHub Pages (onde o certificado e emitido e renovado automaticamente), no servidor proprio isso precisa ser configurado:"
Bul "Opcao mais simples: usar o Let's Encrypt (certificado gratuito, renovacao automatica via certbot no Linux, ou uma ferramenta equivalente no Windows/IIS)."
Bul "Alternativa: usar um certificado ja emitido pela propria empresa (se a organizacao tiver uma autoridade certificadora interna)."
Nota "O guia oficial da Supabase para producao recomenda colocar um proxy reverso (Nginx ou Caddy, por exemplo) na frente do gateway do backend self-hosted para lidar com o HTTPS - especialmente importante aqui, ja que o login com Microsoft (OAuth) exige HTTPS valido para funcionar."

H2 "12.3 Atualizar a configuracao de URL no Supabase novo"
P "Em Authentication > URL Configuration do Studio novo, confirmar que Site URL e a lista de Redirect URLs apontam para o dominio definitivo do site (https://gestaoprojetos.suaempresa.com.br), do mesmo jeito que ja e feito hoje no projeto da nuvem."

# ============================================================
# 13. FASE 10
# ============================================================
H1 "13. Fase 10 - Testes de validacao antes do corte"
P "So avancar para a Fase 11 depois que TODOS os itens abaixo forem confirmados funcionando no ambiente novo, acessando pelo endereco definitivo (ou um endereco de teste equivalente, se o dominio definitivo ainda nao puder ser trocado)."
Bul "Login por link magico chega por e-mail e autentica corretamente."
Bul "Login com Microsoft autentica corretamente e devolve o papel certo do usuario."
Bul "Um projeto ja existente (migrado da producao) aparece certo em todas as telas: Painel Executivo, Ficha do Projeto, TAP, Planejamento etc."
Bul "O historico de alteracoes de um registro antigo aparece completo (Ver historico)."
Bul "Um anexo ja existente (migrado do Storage antigo) abre e baixa corretamente a partir do sistema."
Bul "Um upload de anexo NOVO funciona (envia, aparece na lista, pode ser baixado e removido)."
Bul "As 3 funcoes de sugestao/analise por IA respondem sem erro de segredo ausente."
Bul "A transcricao de audio funciona."
Bul "Um evento que dispara notificacao (ex.: aprovacao de um TAP) chega por e-mail normalmente."
Bul "A notificacao push (se em uso) e recebida no celular/computador inscrito."
Bul "Todas as regras de acesso (quem ve o que, restricao por equipe, os dois gates de aprovacao) continuam se comportando exatamente como no ambiente antigo - comparar com o documento Regras de Acesso e Permissoes."
Bul "O Central de Ajuda e os links de documentacao dentro do site continuam abrindo normalmente."
Nota "Vale testar com pelo menos uma conta de cada papel de usuario (Solicitante, Gerente de Projetos, Gestor Responsavel, Dono do Negocio, Alta Gestao, PMO/Admin), nao so com uma conta de administrador."

# ============================================================
# 14. FASE 11
# ============================================================
H1 "14. Fase 11 - Corte de producao e plano de rollback"

H2 "14.1 Congelar o ambiente antigo por um instante"
P "Pouco antes da troca final, evitar novos registros/edicoes no ambiente antigo (nuvem) por uma janela curta, para nao perder dados criados DEPOIS do ultimo dump da Fase 3 e ANTES do dominio apontar para o servidor novo. Se a janela de corte precisar ser maior que alguns minutos, repetir a exportacao/importacao do banco (Fase 3) e do storage (Fase 4) uma ultima vez, `a quente`, imediatamente antes da troca de DNS."

H2 "14.2 Trocar o DNS"
P "Executar a mudanca de DNS descrita na Fase 9.1. A propagacao pode levar de alguns minutos a algumas horas, dependendo do TTL configurado no registro antigo."

H2 "14.3 Acompanhar de perto"
P "Nas primeiras horas/dias apos o corte, acompanhar ativamente: erros no console do navegador reportados pelos usuarios, uso de CPU/memoria/disco do servidor novo, e volume de chamadas as funcoes de IA (para garantir que nao ha loop ou uso indevido gerando custo inesperado nas chaves da Anthropic/OpenAI)."

H2 "14.4 Plano de rollback"
P "Enquanto o projeto Supabase Cloud de producao nao for desligado, ele continua sendo uma copia valida do estado ANTES da migracao. Se algo critico for encontrado logo apos o corte, reverter o DNS para o endereco antigo (GitHub Pages) traz o sistema de volta ao estado anterior - com a ressalva de que qualquer dado criado no ambiente novo depois do corte fica para tras nesse cenario de rollback (por isso a importancia de resolver problemas rapido, ou repetir a migracao de dados antes de tentar de novo)."
Nota "So desligar/cancelar o projeto Supabase Cloud de producao e o repositorio antigo do GitHub Pages depois de um periodo de operacao estavel no ambiente novo (recomenda-se pelo menos 1 a 2 semanas de uso normal sem incidentes)."

# ============================================================
# 15. FASE 12
# ============================================================
H1 "15. Fase 12 - Nova rotina operacional (responsabilidades da empresa)"
P "A partir da migracao, os itens abaixo deixam de ser automaticos e passam a exigir uma rotina definida por quem for operar o servidor."

H2 "15.1 Backup"
Bul "Backup diario do banco de dados Postgres (pg_dump agendado, ou snapshot do volume Docker do banco), guardado em local separado do proprio servidor."
Bul "Backup periodico dos arquivos de Storage (a pasta de volumes onde os anexos ficam guardados)."
Bul "Testar a restauracao do backup periodicamente - um backup nunca testado nao e garantia de nada."

H2 "15.2 Atualizacoes"
Bul "Acompanhar novas versoes do kit self-hosted da Supabase (correcoes de seguranca, principalmente) e atualizar as imagens Docker periodicamente (docker compose pull seguido de sh run.sh start de novo)."
Bul "Manter o sistema operacional do servidor e o proprio Docker atualizados."

H2 "15.3 Certificado e dominio"
Bul "Se usar Let's Encrypt, confirmar que a renovacao automatica esta configurada e realmente rodando (ela expira a cada ~90 dias)."

H2 "15.4 Monitoramento"
Bul "Definir quem e avisado se o site sair do ar, se o banco de dados parar, ou se o disco do servidor ficar cheio - hoje isso e responsabilidade do Supabase/GitHub; apos a migracao, precisa de um responsavel e, idealmente, um alerta automatizado (mesmo que simples)."

H2 "15.5 Segredos e senhas"
Bul "Guardar as chaves de API (Anthropic, OpenAI, Resend, VAPID) e as senhas do banco/Studio num local seguro (cofre de senhas da empresa), documentando quem tem acesso - o proprio guia oficial da Supabase recomenda, para producao, mover os segredos do arquivo .env para uma ferramenta dedicada de gestao de segredos, se a empresa ja tiver uma em uso."

# ============================================================
# 16. RISCOS
# ============================================================
H1 "16. Riscos e recomendacao final"
$rRisco = @(
  @("Risco","Como reduzir"),
  @("Perda de dados na migracao do banco/storage","Migrar em paralelo (secao 1.4), conferir contagens de linhas e testar restauracao antes do corte final (Fase 6.3, Fase 10)"),
  @("Login com Microsoft nao funcionar por causa do HTTPS/Redirect URI","Configurar o certificado HTTPS antes de testar o login com Microsoft (Fase 9.2 depende da Fase 9.1/9.2 - Fase 6.2), e manter o Redirect URI antigo cadastrado durante a transicao"),
  @("Ficar sem backup automatico e nao perceber","Definir a rotina de backup (Fase 12.1) ANTES do corte de producao, nao depois"),
  @("Uma das 19 paginas ficar apontando para o endereco antigo por engano","Fazer a troca de credenciais (Fase 8) com busca-e-substituicao em lote, e conferir com uma nova busca pelo endereco antigo depois"),
  @("Servidor da empresa ficar indisponivel (energia, rede, hardware)","Definir redundancia/plano de contingencia proporcional a criticidade real do sistema para a operacao da empresa"),
  @("Versao do kit self-hosted desatualizar e ficar sem suporte","Acompanhar atualizacoes do kit e manter uma rotina de atualizacao (Fase 12.2)")
)
TableSimple $rRisco @(6.5,9.5)
P "Recomendacao final: esta migracao e tecnicamente viavel e o sistema pode, sim, continuar funcionando exatamente como hoje - mas ela troca uma operacao praticamente sem manutencao (Supabase Cloud + GitHub Pages) por uma infraestrutura que a propria empresa passa a manter, com backup, atualizacao, certificado e disponibilidade sob responsabilidade interna. Vale confirmar, antes de iniciar, que essa troca de responsabilidade e realmente o objetivo (por exemplo, exigencia de manter os dados dentro da empresa) - e nao apenas um objetivo de reduzir custo ou dependencia de terceiros, caso em que migrar so o site (mantendo o Supabase Cloud) atingiria o mesmo resultado com uma fracao do esforco e do risco deste documento."

# ============================================================
# 17. CHECKLIST
# ============================================================
H1 "17. Checklist resumido"
Bul "1. Conferir o inventario do que existe hoje em producao (Fase 2)."
Bul "2. Provisionar o servidor e instalar Docker/Git/Supabase CLI (Fase 1)."
Bul "3. Instalar o kit self-hosted da Supabase e configurar o .env (Fase 2)."
Bul "4. Exportar o banco de producao e importar no ambiente novo (Fase 3)."
Bul "5. Migrar os arquivos do Storage (Fase 4)."
Bul "6. Copiar as Edge Functions e configurar os segredos (Fase 5)."
Bul "7. Reconfigurar SMTP e login com Microsoft (Fase 6)."
Bul "8. Publicar o site no servidor (IIS ou Nginx) (Fase 7)."
Bul "9. Trocar a URL/chave do Supabase em todos os arquivos do site (Fase 8)."
Bul "10. Apontar o dominio e configurar HTTPS (Fase 9)."
Bul "11. Rodar TODOS os testes da Fase 10 antes de prosseguir."
Bul "12. Fazer o corte de producao com plano de rollback pronto (Fase 11)."
Bul "13. Colocar a nova rotina de backup/atualizacao/monitoramento em pratica (Fase 12)."
Bul "14. So depois de um periodo estavel, desligar o ambiente antigo (Fase 11.4)."

$word.ActiveDocument.TablesOfContents.Item(1).Update() | Out-Null
$doc.SaveAs2($OutPath, 16)
$doc.Close()
$word.Quit()
Write-Output "SAVED: $OutPath"
