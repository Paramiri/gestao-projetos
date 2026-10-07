param(
  [string]$OutPath = (Join-Path $PSScriptRoot "Manual de Uso - Ferramenta de Gestao de Projetos.docx"),
  [Parameter(Mandatory=$true)][string]$ImgDir
)
$ErrorActionPreference = "Stop"

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
$footer.Range.Text = "SGP - Manual de Uso da Ferramenta de Gestao de Projetos | Pagina "
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
  $sel.ParagraphFormat.LeftIndent = 0
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
  $sel.Font.Name = "Montserrat"
  $sel.Font.Size = 10.5
  $sel.Font.Color = $colInk
  $sel.ParagraphFormat.SpaceBefore = 8
  $sel.ParagraphFormat.SpaceAfter = 10
  $sel.Font.Bold = $true
  $sel.Font.Color = $colGood
  $sel.TypeText("EXEMPLO - ")
  $sel.Font.Bold = $false
  $sel.Font.Color = $colInk
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Borders.Item(1).LineStyle = 0
  $sel.ParagraphFormat.LeftIndent = 0
}

function Img($filename, $caption, $widthIn=5.6){
  $path = Join-Path $ImgDir $filename
  if (-not (Test-Path $path)) { Write-Warning "Imagem nao encontrada: $path"; return }
  $sel.Style = $doc.Styles.Item(-1)
  $sel.ParagraphFormat.Alignment = 1
  $sel.ParagraphFormat.SpaceBefore = 6
  $sel.ParagraphFormat.SpaceAfter = 2
  $shape = $sel.InlineShapes.AddPicture($path)
  $ratio = $shape.Height / $shape.Width
  $shape.Width = $word.InchesToPoints($widthIn)
  $shape.Height = $shape.Width * $ratio
  $shape.Line.Visible = $true
  $shape.Line.ForeColor.RGB = RGB 0xD1 0xD5 0xDB
  $shape.Line.Weight = 0.75
  $sel.TypeParagraph()
  if ($caption) {
    $sel.Font.Name = "Montserrat"
    $sel.Font.Size = 9.5
    $sel.Font.Italic = $true
    $sel.Font.Color = $colMuted
    $sel.ParagraphFormat.SpaceAfter = 14
    $sel.TypeText($caption)
    $sel.TypeParagraph()
    $sel.Font.Italic = $false
  }
  $sel.ParagraphFormat.Alignment = 0
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
$sel.Font.Name="Montserrat"; $sel.Font.Size = 28; $sel.Font.Bold = $true; $sel.Font.Color = $colInk
$sel.ParagraphFormat.SpaceAfter = 4
$sel.TypeText("MANUAL DE USO")
$sel.TypeParagraph()
$sel.Font.Bold = $false
P "Ferramenta de Gestao de Projetos (SGP)" 15 $false $false $colInk "left" 20
P "Guia passo a passo: do mapa de diretrizes e da Solicitacao de Demanda a geracao dos Relatorios de Situacao e de Entregas e Beneficios - incluindo login, papeis de usuario, gates de aprovacao, restricao por equipe e o Validador de Projetos." 12 $false $true $colMuted "left" 30
P "Andre Torres - Gerente de Projetos" 11 $false $false $colMuted "left" 2
P "Versao 2.92 - 7 de outubro de 2026 (substitui a versao 2.91 de 17/09/2026)" 11 $false $false $colMuted "left" 2

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
# 1. INTRODUCAO
# ============================================================
H1 "1. Introducao"

H2 "1.1 Sobre este manual"
P "Este manual explica, passo a passo, como usar as ferramentas eletronicas que dao suporte a gestao de projetos da instituicao. Ele cobre a sequencia completa de uso: do login e do mapa de diretrizes, passando pelo registro de uma nova demanda, ate a geracao dos dois relatorios de fechamento - o Relatorio de Situacao de Projetos (SGP-11) e o Relatorio de Entregas e Beneficios (SGP-12)."
P "Esta e a versao 2.0 do manual. Em relacao a versao 1.0 (15/07/2026), foram adicionadas as secoes sobre login obrigatorio, acesso sem login, papeis de usuario, os dois gates de aprovacao, restricao por equipe do projeto, o Validador de Projetos e a pagina de Administracao - alem de capturas de tela reais de cada ferramenta."
P "Esta e a versao 2.1 do manual. Em relacao a versao 2.0 (28/07/2026), foram adicionadas as secoes sobre o registro vivo de riscos e as sugestoes de riscos recorrentes no TAP (Passo 3), a sinalizacao automatica de risco de atraso no Relatorio de Situacao (secao 4.1) e a prioridade leve (P0/P1/P2) das atividades da EAP (Passo 5)."
P "Esta e a versao 2.2 do manual. Em relacao a versao 2.1 (29/07/2026), foi adicionada a implementacao tecnica do Gate 2 - Pactuacao no Relatorio de Entregas e Beneficios (secao 4.2): status controlado pelo Admin, data e aprovador preenchidos automaticamente, e trava de edicao do relatorio enquanto pactuado."
P "Esta e a versao 2.3 do manual. Em relacao a versao 2.2 (29/07/2026), foi adicionada a Gestao de Entraves e Encaminhamentos no Relatorio de Situacao (secao 4.1): dois cartoes de registro (Descricao, Responsavel, Prazo, Status Aberto/Resolvido/Cancelado), exibidos lado a lado no Painel e incluidos no CSV exportado."
P "Esta e a versao 2.4 do manual. Em relacao a versao 2.3 (29/07/2026), foi adicionada a propagacao automatica da decisao da SMP (Passo 7) para o historico compartilhado do projeto vinculado: toda vez que uma SMP e aprovada ou nao aprovada, o evento fica visivel em `Ver historico` em qualquer formulario ligado aquele projeto, sem alterar o Status de Gate 1."
P "Esta e a versao 2.5 do manual. Em relacao a versao 2.4 (29/07/2026), foi adicionado o campo opcional `Projeto vinculado (Gate 1)` a cada projeto listado no Relatorio de Situacao (secao 4.1) e no Relatorio de Entregas e Beneficios (secao 4.2): a selecao preenche o nome automaticamente e libera o `Ver historico` compartilhado, sem aplicar a restricao por equipe usada nos outros 7 formularios vinculados."
P "Esta e a versao 2.6 do manual. Em relacao a versao 2.5 (29/07/2026), foi adicionado o botao `Importar do TAP` na secao Indicadores do projeto do Relatorio de Entregas e Beneficios (secao 4.2): reaproveita os indicadores de resultado ja preenchidos no TAP do mesmo Projeto vinculado, evitando redigitacao, com protecao contra indicadores duplicados."
P "Esta e a versao 2.7 do manual. Em relacao a versao 2.6 (30/07/2026), foi adicionado o versionamento historico ao Relatorio de Situacao (secao 4.1) e ao Relatorio de Entregas e Beneficios (secao 4.2): cada `Salvar e ver painel` arquiva uma copia completa dos dados, consultavel pelo botao `Historico de versoes` no rodape do Painel, com visualizacao somente leitura de versoes antigas e trava de edicao ate voltar a versao atual."
P "Esta e a versao 2.8 do manual. Em relacao a versao 2.7 (01/08/2026), foi adicionado o cartao `Resultados alcancados` ao Relatorio de Situacao (secao 4.1): registra indicadores, entregas e beneficios de fato alcancados frente ao planejado, atendendo o Reporte de Resultados (D06.5) sem exigir um 13o formulario separado. A tabela de artefatos (secao 8) e o texto sobre cobertura das Diretrizes (secao 1.4) foram atualizados de acordo."
P "Esta e a versao 2.9 do manual. Em relacao a versao 2.8 (01/08/2026), foi adicionada a nova secao 2.8 - Notificacoes por e-mail em transicoes de estado: seis transicoes (Gate 1, Canvas aprovado, TAP aprovado, decisao da SMP, registro de TEP e Gate 2 pactuado) agora disparam automaticamente um e-mail via Resend para quem criou o registro, a equipe do projeto vinculado e os Admins."
P "Esta e a versao 2.10 do manual. Em relacao a versao 2.9 (01/08/2026), foi adicionada a nova secao 6.4 - Painel Executivo: uma 15a pagina, restrita a Admin, que agrega em tempo real os 10 formularios de registro e os 2 relatorios numa unica tela (gates de aprovacao, entraves/encaminhamentos/resultados, contagem por formulario e portfolio de projetos), sem exigir nenhuma mudanca no banco de dados."
P "Esta e a versao 2.11 do manual. Em relacao a versao 2.10 (03/08/2026), foi adicionada a nova secao 2.9 - Lembretes por notificacao push no celular: um segundo canal de aviso, alem do e-mail, ativado por cada usuario individualmente na pagina inicial. Dispara nos mesmos seis momentos de decisao ja descritos na secao 2.8."
P "Esta e a versao 2.12 do manual. Em relacao a versao 2.11 (04/08/2026), foi adicionado um setimo momento de notificacao (secao 2.8): ao registrar uma nova Solicitacao de Demanda, os Admins agora recebem um aviso imediato (e-mail e push) de que ha um Gate 1 - Triagem pendente, em vez de depender de alguem checar a lista manualmente."
P "Esta e a versao 2.13 do manual. Em relacao a versao 2.12 (06/08/2026), foi adicionado o campo obrigatorio Prioridade (Critica/Alta/Media/Baixa) a Solicitacao de Demanda (Passo 1): exibido como selo colorido na lista de demandas cadastradas e no detalhe do registro, ajudando o Admin a priorizar a triagem do Gate 1."
P "Esta e a versao 2.14 do manual. Em relacao a versao 2.13 (06/08/2026), foi adicionado o pre-cadastro de usuario e o campo Telefone a aba Usuarios da Administracao (secao 6.1 e 3.2 das Regras de Acesso): o Admin pode informar nome, telefone, e-mail e papel de uma pessoa antes do primeiro login dela, e os dados sao aplicados automaticamente ao perfil assim que ela loga pela primeira vez."
P "Esta e a versao 2.15 do manual. Em relacao a versao 2.14 (10/08/2026), foi adicionado o botao Remover a aba Usuarios da Administracao (secao 6.1 e 3.3 das Regras de Acesso): o Admin so consegue excluir um usuario que nao tenha nenhum vinculo com projetos ou registros no sistema, checado automaticamente antes de confirmar a remocao."
P "Esta e a versao 2.16 do manual. Em relacao a versao 2.15 (10/08/2026), foi adicionada a restricao de acesso por papel ao Painel Executivo, ao Relatorio de Situacao e ao Relatorio de Entregas (secao 6.3 e 3.4 das Regras de Acesso): o Admin marca, por caixas de selecao, quais papeis podem ver cada uma dessas tres paginas, podendo marcar mais de um papel por pagina."
P "Esta e a versao 2.17 do manual. Em relacao a versao 2.16 (11/08/2026), foi adicionada a importacao de transcricao por IA na Ata de Reuniao (secao 2.10, Passo 6 e 6.3 - substitui o antigo painel `Preencher com Read AI`, que nunca funcionou em producao): o usuario cola a transcricao de uma reuniao e a IA sugere pauta, participantes, resumo, encaminhamentos e entraves para revisao, com interruptor geral e lista de papeis permitidos configuraveis pelo Admin."
P "Esta e a versao 2.18 do manual. Em relacao a versao 2.17 (12/08/2026), a importacao de transcricao por IA (secao 2.10) passou a aceitar tambem arquivo `.txt`, `.docx` ou `.pdf` anexado, alem de colar o texto diretamente - o sistema extrai o texto do arquivo automaticamente para revisao antes de analisar."
P "Esta e a versao 2.19 do manual. Em relacao a versao 2.18 (12/08/2026), foi adicionada a importacao de audio por IA na Ata de Reuniao (secao 2.11): alem de colar/anexar a transcricao em texto, o usuario pode anexar a propria gravacao da reuniao (.mp3, .m4a, .aac, .wav, .ogg...), dividida e transcrita automaticamente em pedacos pelo sistema - com interruptor geral e lista de papeis permitidos proprios, independentes dos da importacao por texto, configuraveis pelo Admin (secao 6.3)."
P "Esta e a versao 2.20 do manual. Em relacao a versao 2.19 (12/08/2026), os antigos botoes `Exportar CSV` e `Imprimir` da Ata de Reuniao (Passo 6), que agiam sobre a lista inteira e nao geravam um documento util, foram substituidos por `Imprimir` e `Exportar Word` no painel de uma ata especifica: ambos geram o documento no formato oficial do SGP-07 (cabecalho, unidade, pauta, participantes, descricao, saidas e entraves)."
P "Esta e a versao 2.21 do manual. Em relacao a versao 2.20 (13/08/2026), o mesmo ajuste foi estendido aos demais formularios que ainda tinham os antigos botoes `Exportar CSV` e `Imprimir` sem funcionalidade real: Solicitacao de Demanda (Passo 1), Canvas de Projeto (Passo 2), TAP (Passo 3), Planejamento e Desenvolvimento (Passo 4), EAP (Passo 5), SMP (Passo 7), TEP (Passo 8), RLA (Passo 9), Relatorio de Situacao (secao 4.1) e Relatorio de Entregas e Beneficios (secao 4.2). Em todos, `Imprimir` e `Exportar Word` agora geram o documento no formato padrao do respectivo codigo SGP (SGP-01 a SGP-12)."
P "Esta e a versao 2.22 do manual. Em relacao a versao 2.21 (13/08/2026), o Plano de Comunicacao de Projeto (Passo 10) ganhou as 4 colunas que faltavam em relacao ao documento oficial SGP-08 - Quando Comunicar, Com Quem se Comunicar, Como Comunicar e Quem Comunica, alem das ja existentes Tipo de Comunicacao e O que Comunicar - e passou a ter a aba `Editar dados` restrita por papel (padrao PMO/Admin, configuravel em Administracao > Configuracoes, secao 6.3), mantendo a visualizacao e impressao livres para qualquer usuario autenticado."
P "Esta e a versao 2.23 do manual. Em relacao a versao 2.22 (13/08/2026), foi adicionada a nova secao 6.5 - Ambiente de Treino: um ambiente de treinamento/demonstracao totalmente separado da producao, com 7 projetos ficticios cobrindo o ciclo completo, e uma quarta aba na Administracao (visivel so em producao) com um botao que reseta e repopula esse ambiente com um clique, restrito a Admin (secao 3.8 das Regras de Acesso)."
P "Esta e a versao 2.24 do manual. Em relacao a versao 2.23 (14/08/2026), foi adicionada a nova secao 2.12 - Assistente de preenchimento por IA (Canvas e TAP): um botao `Sugerir com IA` que le a Solicitacao de Demanda, o Canvas (no caso do TAP) e as Atas de Reuniao do projeto vinculado para sugerir o preenchimento dos campos ainda vazios, com a mesma logica de interruptor geral e lista de papeis liberados ja usada na importacao de transcricao (secao 2.10)."
P "Esta e a versao 2.25 do manual. Em relacao a versao 2.24 (19/08/2026), foi adicionada a nova secao 2.13 - Importacao de documento por IA (Solicitacao de Demanda): um painel `Importar documento preenchido` que le o formulario oficial SGP-01 preenchido a mao em Word ou PDF (colado ou anexado) e sugere o preenchimento dos campos do formulario online, com a mesma logica de interruptor geral e lista de papeis liberados das demais importacoes por IA - sem restricao por equipe de projeto, ja que e o primeiro formulario da esteira."
P "Esta e a versao 2.26 do manual. Em relacao a versao 2.25 (19/08/2026), o assistente de preenchimento por IA (secao 2.12) foi estendido do Canvas/TAP para tambem cobrir o Planejamento e Desenvolvimento de Projeto (Passo 4: pre-projeto, beneficios e marcos iniciais de cronograma, a partir do TAP e do Canvas) e a EAP (Passo 5: pacotes de trabalho e entregas, a partir do Planejamento e do TAP, so quando a arvore estiver totalmente vazia) - Fase 2 do assistente, usando o mesmo interruptor geral ja existente."
P "Esta e a versao 2.27 do manual. Em relacao a versao 2.26 (19/08/2026), o assistente de preenchimento por IA (secao 2.12) foi estendido para tambem cobrir a SMP (Passo 7: descricao da mudanca e analise de impactos, contrastando o combinado no TAP/Planejamento com o que as Atas mais recentes discutem), o TEP (Passo 8: justificativa, atividades encerradas e consideracoes finais, comparando planejado e realizado) e o RLA (Passo 9: as 4 secoes de texto livre - Visao geral, Destaques, Desafios e Tarefas pos-projeto - nunca os blocos de avaliacao estruturada Sim/Nao/Parcial com placar, que continuam sendo autoavaliacao manual) - Fase 3, concluindo o assistente nos 7 formularios da esteira principal, usando o mesmo interruptor geral ja existente."
P "Esta e a versao 2.28 do manual. Em relacao a versao 2.27 (19/08/2026), o assistente de preenchimento por IA no Planejamento (Passo 4, secao 2.12) passou a ler tambem a Solicitacao de Demanda e a sugerir mais 5 campos que ainda ficavam de fora: o Solicitante (capa, copiado da Demanda), as Macro fases do projeto (aba Pre-projeto, resumidas do cronograma do TAP) e a Introducao, Situacao atual e Proposta de mudanca (aba Viabilidade, redigidas com a mesma base factual do Contexto e do Objetivo geral do Pre-projeto, mas com redacao propria para cada campo). O campo Negocio (capa) permanece de fora do assistente, por nao ter fonte confiavel nos documentos lidos."
P "Esta e a versao 2.29 do manual. Em relacao a versao 2.28 (20/08/2026), a Solicitacao de Demanda (Passo 1) passou a aceitar upload real de arquivo nos Anexos - antes era so um campo de texto livre com nomes de documentos ou links. Agora e possivel anexar diretamente ate 15 MB por arquivo (.txt, .docx ou .pdf), guardado num repositorio de arquivos proprio do sistema (Supabase Storage) e disponivel para download por quem acessar o registro; o campo de texto livre continua existindo, renomeado para Links de documentos, para casos como Drive ou SharePoint que nao sao um arquivo enviado."
P "Esta e a versao 2.30 do manual. Em relacao a versao 2.29 (20/08/2026), o vinculo `Projeto vinculado (Gate 1)` do Relatorio de Situacao (secao 4.1) e do Relatorio de Entregas e Beneficios (secao 4.2) passou a puxar mais campos automaticamente do projeto selecionado, alem do Nome do projeto que ja preenchia sozinho: no Relatorio de Situacao, Responsavel e Inicio/Termino previsto (a partir do cronograma do TAP); no Relatorio de Entregas, Responsavel pelo projeto, Objetivo, Publico-alvo e um resumo dos Riscos em aberto (a partir do TAP). E preenchimento direto, sem IA - mesma logica do botao `Importar do TAP` que ja existia para os Indicadores. Campos como % Execucao, Status e os codigos internos (Codigo/ID, Programa vinculado, Codigo da acao) continuam manuais, por nao terem fonte confiavel em nenhum documento do sistema."
P "Esta e a versao 2.31 do manual. Em relacao a versao 2.30 (20/08/2026), o assistente de preenchimento por IA (secao 2.12) foi estendido do preenchimento direto (Fase A) para tambem sugerir por IA, por linha de projeto: no Relatorio de Situacao (secao 4.1), o que Merece atencao e o que Merece destaque, com base no TAP, nas Atas e na situacao ja digitada na linha (status, % execucao, datas, sinalizacao automatica); no Relatorio de Entregas (secao 4.2), o Impacto/Resultado esperado do projeto, com base no TAP e nas Atas. Diferente dos outros 7 formularios do assistente, aqui o botao `Sugerir com IA` aparece por linha, nao uma vez so para o formulario inteiro - Fase B, concluindo a proposta tecnica de preenchimento automatico dos relatorios de portfolio."
P "Esta e a versao 2.32 do manual. Em relacao a versao 2.31 (20/08/2026), o Relatorio de Situacao (secao 4.1) ganhou o botao `Importar marcos do TAP` na secao Marcos de cada linha de projeto: complementa a copia direta ja existente (Responsavel, Inicio/Termino previsto) trazendo tambem os marcos do cronograma de entregas macro do TAP como linhas de Marcos, sem duplicar os que ja estiverem na lista - mesma logica do `Importar do TAP` ja usado nos Indicadores do Relatorio de Entregas (secao 4.2)."
P "Esta e a versao 2.33 do manual. Em relacao a versao 2.32 (20/08/2026), duas melhorias na Ata de Reuniao (Passo 6): os campos e linhas preenchidos pela importacao de transcricao por IA (secao 2.10) agora ficam com o mesmo selo `IA` e destaque visual ja usado nos demais assistentes do sistema, o que antes nao acontecia nesse formulario; e a coluna SEQ das tabelas de Encaminhamentos e Entraves passou a ser preenchida automaticamente (numeracao sequencial, so contando linhas com conteudo), liberando mais espaco de tela para o texto do Encaminhamento/Entrave em si."
P "Esta e a versao 2.34 do manual. Em relacao a versao 2.33 (21/08/2026), o Planejamento e Desenvolvimento do Projeto (Passo 4) passou a agrupar visualmente Objetivo Geral e Objetivo Especifico sob um unico item `Objetivo do projeto`, e Escopo Incluido e Escopo Excluido sob `Escopo` - reorganizacao apenas de apresentacao (o mesmo destaque de campo aplicado a um grupo inteiro), sem mudanca nos dados salvos nem nos campos em si."
P "Esta e a versao 2.35 do manual. Em relacao a versao 2.34 (21/08/2026), foi adicionada a nova secao 6.6 - Backup automatizado: um bloco na aba Configuracoes da Administracao onde o Admin/PMO define destino, frequencia, retencao, e-mail de alerta e papeis com acesso a essas proprias configuracoes de backup - complemento operacional a proposta tecnica `Backup Automatizado` entregue separadamente, que descreve a rotina agendada (GitHub Actions) que efetivamente le esses parametros e executa a copia."
P "Esta e a versao 2.36 do manual. Em relacao a versao 2.35 (26/08/2026), foi adicionada a nova secao 6.7 - Relatorio de Resultados: uma 16a pagina em `Gestao`, separada do Painel Executivo e com acesso configuravel por papel (padrao Admin), reunindo os indicadores estrategicos do portfolio (marcos atrasados, projetos dentro do orcamento e entraves em aberto) e um Plano de Acao com todo entrave/encaminhamento em aberto do Relatorio de Situacao. De apoio a essa nova pagina, o Relatorio de Situacao (secao 4.1) ganhou a secao `Custos realizados` por categoria em cada projeto (importavel do TAP, mesma logica do `Importar marcos do TAP`) e os campos opcionais `Por que?` e `Resultado esperado` em cada entrave/encaminhamento."
P "Esta e a versao 2.37 do manual. Em relacao a versao 2.36 (27/08/2026), os 3 indicadores do Relatorio de Resultados (secao 6.7) passaram de contagens simples para o padrao meta + barra de progresso + selo de status (Normal/Atencao/Critico) + tendencia: `Marcos atrasados` virou `Marcos no prazo` (percentual) e `Entraves em aberto` virou `Entraves resolvidos` (comparando com o ultimo salvamento do relatorio, usando um novo identificador estavel em cada entrave/encaminhamento). A tendencia e calculada a partir do historico de versoes que o Relatorio de Situacao ja mantinha (secao 4.1). As metas de cada indicador sao configuraveis pelo Admin em Administracao > Configuracoes (secao 6.3)."
P "Esta e a versao 2.38 do manual. Em relacao a versao 2.37 (27/08/2026), o Relatorio de Resultados (secao 6.7) ganhou dois botoes de exportacao, no mesmo modelo visual de um arquivo de referencia fornecido para a reuniao de resultados: `Exportar PDF` (reaproveita o mecanismo de impressao ja usado nos demais relatorios) e `Gerar Apresentacao (PPTX)`, que monta no proprio navegador uma apresentacao de 3 slides (capa, indicadores estrategicos com meta/barra/selo/tendencia, e Plano de Acao) com os mesmos numeros exibidos na tela."
P "Esta e a versao 2.39 do manual. Em relacao a versao 2.38 (27/08/2026), o Backup automatizado (secao 6.6) teve seu destino externo (SharePoint) confirmado funcionando de ponta a ponta em 28/08/2026, apos o TI concluir a concessao pontual de acesso (`Sites.Selected`, restrita a um unico site) que faltava desde a revogacao das permissoes mais amplas concedidas inicialmente - fecha o ultimo item em aberto da proposta tecnica `Backup Automatizado`, sem mudanca nos campos da tela."
P "Esta e a versao 2.40 do manual. Em relacao a versao 2.39 (28/08/2026), o registro de um TEP (Passo 8) passou a encerrar o projeto no cadastro compartilhado (secao 2.5): o campo `Projeto vinculado` dos 9 formularios que o usam (Canvas, TAP, Planejamento, EAP, SMP, TEP, RLA, Relatorio de Situacao e Relatorio de Entregas) so oferece projetos ainda ativos para novos vinculos - antes, um projeto encerrado continuava aparecendo para sempre. Documentos ja vinculados antes do encerramento continuam mostrando o vinculo normalmente. O Painel Executivo (secao 6.4) ganhou o card `Encerrados` no Portfolio de projetos."
P "Esta e a versao 2.41 do manual. Em relacao a versao 2.40 (02/09/2026), o Painel Executivo (secao 6.4) ganhou a nova secao `Visao atual por projeto` - Fase 1 de uma proposta em fases para dar visibilidade por projeto individual, nao so agregada. Uma linha por projeto do cadastro compartilhado (secao 2.5), cruzando o Status do cadastro, a Situacao mais recente do Relatorio de Situacao (secao 4.1, com o percentual concluido quando reportado) e o Progresso do ciclo de vida (quantos dos 9 formularios ja tem registro vinculado aquele projeto, com a lista do que falta)."
P "Esta e a versao 2.42 do manual. Em relacao a versao 2.41 (03/09/2026), foi adicionada a nova pagina Ficha do Projeto (secao 6.4.1) - Fase 2 da proposta de visao por projeto. Aberta clicando no nome de um projeto na tabela `Visao atual por projeto` do Painel Executivo, reune numa unica tela os dados do cadastro, a situacao mais recente do Relatorio de Situacao (status, %, marcos, custos realizados), o cronograma e custos planejados do TAP, o progresso do projeto em cada um dos 9 formularios do ciclo de vida (com link direto para cada um), a equipe do projeto e o historico de alteracoes - sem restricao por papel, aberta a qualquer usuario autenticado."
P "Esta e a versao 2.43 do manual. Em relacao a versao 2.42 (03/09/2026), foi adicionada a Saude do projeto (secao 6.4.1) - Fase 3 da proposta de visao por projeto. Um selo geral (Normal/Atencao/Critico/Encerrado/Sem dados) calculado a partir de tres sinais - situacao atrasada, marco ativo com data-alvo vencida e orcamento estourado (custo realizado acima do estimado no TAP) - alem de tres numeros de apoio: % de marcos concluidos, atraso em dias do marco mais atrasado e % do orcamento realizado. Aparece na Ficha do Projeto e como nova coluna `Saude` na tabela `Visao atual por projeto` do Painel Executivo (secao 6.4). Nenhum dado novo no banco - tudo calculado em cima do que ja era lido pela Ficha (Fase 2)."
P "Esta e a versao 2.44 do manual. Em relacao a versao 2.43 (03/09/2026), a SMP - Solicitacao de Mudanca de Projeto (Passo 7) ganhou a mesma secao Anexos ja existente na Solicitacao de Demanda (Passo 1, versao 2.29): upload real de arquivo (.txt, .docx ou .pdf, ate 15 MB) guardado num repositorio proprio do sistema (Supabase Storage, bucket dedicado `anexos-smp`) e disponivel para download por quem acessar o registro, alem do campo de texto livre Links de documentos para referencias que nao sao um arquivo enviado."
P "Esta e a versao 2.45 do manual. Em relacao a versao 2.44 (03/09/2026), o Registro de riscos do TAP (Passo 3, secao 9) ganhou os campos Probabilidade (Baixa/Media/Alta), Impacto (Baixo/Medio/Alto) e Mitigacao (texto livre), complementares aos ja existentes Status, Responsavel e Ultima revisao - cada risco agora descreve tambem a chance de ocorrer, a gravidade do efeito e a resposta planejada, com Probabilidade e Impacto exibidos como selo colorido por severidade na tela de detalhes."
P "Esta e a versao 2.46 do manual. Em relacao a versao 2.45 (03/09/2026), o documento gerado pelos botoes `Imprimir` e `Exportar Word` da Solicitacao de Demanda (Passo 1) passou a incluir a secao 10 - Aprovacao ao final, com linhas para assinatura do Gestor Responsavel e do Responsavel, e a data - alinhando o documento gerado pelo sistema ao formulario oficial SGP-01 em Word. O logo usado em todos os documentos gerados por Imprimir/Exportar Word (Solicitacao de Demanda, Canvas, TAP, Planejamento, EAP, SMP, Ata, TEP, RLA, Relatorio de Situacao e Relatorio de Entregas) tambem foi atualizado para a identidade visual do sistema, conforme o template padrao."
P "Esta e a versao 2.47 do manual. Em relacao a versao 2.46 (09/09/2026), o cronograma de entregas macro do TAP (Passo 3, secao 10) passou a ser a fonte unica do cronograma do projeto: ganhou o campo `Conclusao real` em cada marco, e o Planejamento (Passo 4, aba Cronograma) e o Relatorio de Situacao (secao 4.1, secao Marcos) passaram a mostrar esses mesmos marcos sempre ao vivo, so leitura, em vez de manter copias proprias - o botao `Importar marcos do TAP` do Relatorio de Situacao foi removido por nao ser mais necessario. O calculo de Saude do projeto (secao 6.4.1) e o indicador Marcos no prazo (secao 6.7) passaram a considerar a Conclusao real: um marco concluido no TAP nunca conta como atrasado, mesmo que o Termino previsto ja tenha passado."
P "Esta e a versao 2.48 do manual. Em relacao a versao 2.47 (10/09/2026), o cronograma de entregas macro do TAP (Passo 3, secao 10) ganhou o campo opcional `Depende de` em cada marco - uma dependencia simples, por nome, de outro marco do mesmo cronograma. Um marco cuja dependencia ainda nao tenha Conclusao real e ja tenha passado do Termino previsto ganha um aviso `bloqueado`, visivel em todo lugar que mostra o cronograma: o proprio TAP, o espelho do Planejamento (Passo 4) e os Marcos do Relatorio de Situacao (secao 4.1)."
P "Esta e a versao 2.49 do manual. Em relacao a versao 2.48 (10/09/2026), a edicao do cronograma de entregas macro passou do TAP (Passo 3) para o Planejamento e Desenvolvimento de Projeto (Passo 4, aba 3): agora e la que os marcos sao cadastrados e alterados, e o TAP passou a mostra-los so leitura, ao vivo - o inverso do que valia desde a versao 2.47. Cada marco tambem ganhou uma numeracao automatica (`Nº`, coluna renumerada ao vivo conforme a ordem das linhas), e o campo de dependencia passou a referenciar outro marco pelo numero (`Depende do nº`) em vez do nome - mais simples de usar e mais resistente a renomear um marco depois. O calculo de Saude do projeto (secao 6.4.1), o indicador Marcos no prazo (secao 6.7) e os Marcos do Relatorio de Situacao (secao 4.1) passaram a ler o cronograma do Planejamento em vez do TAP; o TAP continua sendo a fonte do custo total estimado, usado no calculo de orcamento."
P "Esta e a versao 2.50 do manual. Em relacao a versao 2.49 (11/09/2026), foi corrigido um problema de sobreposicao nas colunas Custo e Depende do nº da tabela de Cronograma do Planejamento (Passo 4, aba 3), que apareciam cortadas/sobrepostas em telas mais estreitas por causa da largura minima interna dos campos. O rotulo da coluna de dependencia tambem foi encurtado de `Depende do nº` para `Depende nº`, evitando quebra de linha no cabecalho. Sem mudanca de comportamento - apenas ajuste de layout."
P "Esta e a versao 2.51 do manual. Em relacao a versao 2.50 (11/09/2026), a tabela de Cronograma do Planejamento (Passo 4, aba 3) ganhou um botao `Linha do tempo`, que alterna para uma visualizacao em barras (estilo Gantt) dos marcos - uma barra por marco, do Inicio ao Termino previstos, colorida por status (no prazo, concluido, atrasado, bloqueado por dependencia), com uma linha vertical marcando o dia de hoje. A mesma visualizacao, com os mesmos dados ao vivo, tambem foi adicionada ao espelho do cronograma no TAP (Passo 3, secao 10) e a secao Marcos do Relatorio de Situacao (secao 4.1). E somente para visualizacao - a edicao continua sempre pela tabela, e os documentos gerados (Imprimir/Exportar Word) continuam mostrando a tabela, sem a linha do tempo."
P "Esta e a versao 2.52 do manual. Em relacao a versao 2.51 (11/09/2026), a linha do tempo (Gantt) do cronograma passou a destacar um caminho critico leve: quando pelo menos dois marcos estao encadeados por `Depende nº`, o caminho que determina a data final do cronograma ganha contorno tracejado, icone de alvo e setas ligando um marco ao proximo. E calculado andando para tras, a partir do marco com o Termino previsto mais tardio, pela cadeia de dependencias - nao e um calculo completo de folga (CPM) com multiplos caminhos. Disponivel nos tres lugares que ja mostram a linha do tempo: Planejamento (Passo 4), espelho do TAP (Passo 3) e Relatorio de Situacao (secao 4.1)."
P "Esta e a versao 2.53 do manual. Em relacao a versao 2.52 (11/09/2026), a tabela de Cronograma do Planejamento (Passo 4, aba 3) ganhou um terceiro botao, `Carga por responsavel`, que agrupa os marcos pelos nomes digitados no campo Responsavel (aceita mais de um nome/area por marco, separados por barra, virgula ou ponto-e-virgula) e desenha uma barra por marco na linha de cada responsavel, na mesma escala de tempo da linha do tempo. Barras sobrepostas ficam mais escuras, e um responsavel com marcos de datas sobrepostas ganha o aviso `possivel sobrecarga`. E uma leitura do texto ja digitado no campo Responsavel - nao soma horas nem calcula capacidade de trabalho. Disponivel nos mesmos tres lugares da linha do tempo: Planejamento (Passo 4), espelho do TAP (Passo 3) e Relatorio de Situacao (secao 4.1) - fecha as 5 fases da proposta de melhoria de controle de cronograma."
P "Esta e a versao 2.54 do manual. Em relacao a versao 2.53 (11/09/2026), o Registro de riscos do TAP (Passo 3, secao 9) ganhou a Fase 1 de uma nova proposta de melhoria, desta vez em gestao de riscos: uma coluna `Severidade`, calculada automaticamente (Probabilidade x Impacto, escala Baixo/Medio/Alto/Critico) e exibida ao vivo como selo colorido direto na tabela - sem novo campo para preencher. Um botao `Matriz de risco`, acima da tabela, alterna para um heatmap 3x3 (Probabilidade x Impacto) com a contagem de riscos em aberto (Status diferente de Encerrado) em cada combinacao. Ambos sao calculos sobre os campos que ja existiam - nenhum dado novo precisa ser digitado."
P "Esta e a versao 2.55 do manual. Em relacao a versao 2.54 (11/09/2026), a Fase 2 da gestao de riscos leva a severidade calculada no TAP (Passo 3) para o portfolio: o calculo de Saude do projeto (secao 6.4.1 e coluna `Saude` do Painel Executivo, secao 6.4) passou a considerar um quarto sinal - pelo menos um risco em aberto com severidade Alto ou Critico -, junto com atraso de marco, orcamento estourado e situacao atrasada. A Ficha do Projeto ganhou um quinto cartao de KPI, `Riscos altos/criticos`, na secao Saude. O Painel Executivo tambem ganhou a secao `Riscos do portfolio (TAP)`, com a contagem total de riscos criticos e riscos altos em aberto somando todos os TAPs cadastrados - a mesma matriz de severidade da Fase 1, agora em nivel de portfolio."
P "Esta e a versao 2.56 do manual. Em relacao a versao 2.55 (11/09/2026), a Fase 3 fecha a proposta de gestao de riscos: o Relatorio de Resultados (secao 6.7) ganhou um quarto indicador estrategico, `Projetos sem risco alto/critico`, com meta configuravel (Administracao > Configuracoes, secao 6.3), selo de status e tendencia dos ultimos salvamentos - no mesmo padrao ja usado por Marcos no prazo e Projetos dentro do orcamento. Como o risco vive inteiramente no TAP e nao tem nenhum campo proprio no Relatorio de Situacao que mude entre salvamentos, o Relatorio de Situacao passou a gravar, a cada salvamento, uma contagem de quantos projetos com TAP estao sem risco alto/critico em aberto - e esse numero, congelado em cada versao salva, que alimenta a tendencia. A apresentacao PPTX exportada tambem foi ajustada para caber os 4 indicadores lado a lado no slide."
P "Esta e a versao 2.57 do manual. Em relacao a versao 2.56 (11/09/2026), a secao Custos realizados do Relatorio de Situacao (secao 4.1) ganhou a Fase 1 de uma nova proposta de melhoria, desta vez em controle de custos: o botao `Comparativo por categoria`, ao lado da lista editavel, mostra uma tabela cruzando, categoria por categoria, o Estimado no TAP com o Realizado aqui digitado - casando pelo nome da categoria, o mesmo nome ja usado pelo botao Importar categorias do TAP. Cada categoria ganha um selo `Dentro do estimado`, `Acima do estimado` ou `Sem estimativa no TAP`. E somente para visualizacao, sem nenhum campo novo para preencher - a edicao continua sempre pela lista."
P "Esta e a versao 2.58 do manual. Em relacao a versao 2.57 (11/09/2026), a Fase 2 da proposta de custos por categoria leva a comparacao Estimado x Realizado do Relatorio de Situacao (Fase 1) para o portfolio: o Painel Executivo (secao 6.4) ganhou a secao `Custos por categoria (portfolio)`, com uma tabela listando cada categoria de custo, em quantos projetos ela tem estimativa no TAP e em quantos desses o Realizado ultrapassou o Estimado - mesmo casamento por nome ja usado no Comparativo por categoria e no botao Importar categorias do TAP, agora em nivel de portfolio."
P "Esta e a versao 2.59 do manual. Em relacao a versao 2.58 (11/09/2026), a Fase 3 fecha a proposta de custos por categoria: a Ficha do Projeto (secao 6.4.1) ganhou a secao `Custos: Estimado (TAP) x Realizado`, um drill-down do cartao de Orcamento da Saude do projeto - a mesma tabela do Comparativo por categoria do Relatorio de Situacao, agora consolidada na visao de projeto unico, sem precisar abrir os dois formularios separadamente para comparar."
P "Esta e a versao 2.60 do manual. Em relacao a versao 2.59 (11/09/2026), a EAP (Passo 5) ganhou a Fase 1 de uma nova proposta de melhoria, desta vez em acompanhamento de execucao: cada Entrega (nivel 2) passou a exibir um status automatico - Concluida, Atrasada, No prazo ou Sem cronograma vinculado - casando pelo nome com o marco correspondente no cronograma do Planejamento (Passo 4), sem nenhum campo novo para preencher na propria EAP. Um resumo no topo da arvore mostra o percentual de entregas concluidas do projeto."
P "Esta e a versao 2.61 do manual. Em relacao a versao 2.60 (11/09/2026), a Fase 2 fecha a proposta de acompanhamento de execucao da EAP: o Painel Executivo (secao 6.4) ganhou a secao `Entregas da EAP (portfolio)`, com o percentual de entregas concluidas de todas as EAPs e quantos projetos tem entrega atrasada, e a Ficha do Projeto (secao 6.4.1) ganhou a secao `Entregas da EAP: status automatico`, listando cada entrega do projeto com seu selo de status - ambas usando o mesmo casamento automatico pelo nome com o cronograma do Planejamento introduzido na Fase 1, sem nenhum campo novo para preencher em nenhum dos tres lugares."
P "Esta e a versao 2.62 do manual. Em relacao a versao 2.61 (11/09/2026), a aba Usuarios da Administracao (secao 6.1) ganhou a Fase 1 de uma nova proposta de melhoria, desta vez em controle de usuarios e papeis: um link `Ver historico` por usuario, reaproveitando a mesma tabela e o mesmo padrao `Ver historico` ja usado nos 9 formularios do ciclo de vida do projeto. Toda troca de papel ou de telefone de um usuario ja cadastrado passa a ficar registrada - valor anterior, novo valor, quando e quem alterou -, fechando uma lacuna em que a troca de papel (que controla o acesso a praticamente toda a ferramenta) nao deixava nenhum rastro."
P "Esta e a versao 2.63 do manual. Em relacao a versao 2.62 (11/09/2026), a Fase 2 da proposta de controle de usuarios e papeis acrescenta uma trava de seguranca ao seletor de papel da aba Usuarios (secao 6.1): nao e mais possivel tirar o papel Admin de alguem (a propria pessoa ou outra) quando essa pessoa e a unica com esse papel no sistema - o sistema recusa a troca com um aviso, em vez de deixar a ferramenta sem ninguem que consiga acessar a propria Administracao para reverter."
P "Esta e a versao 2.64 do manual. Em relacao a versao 2.63 (11/09/2026), a Fase 3 fecha a proposta de controle de usuarios e papeis, deixando a remocao de usuario mais segura (secao 6.1): a checagem automatica antes de excluir um perfil passou a incluir tambem inscricoes de notificacao push registradas naquele e-mail (antes nao verificadas, o que podia deixar uma inscricao orfa), e a propria remocao passou a ficar registrada no historico de alteracoes - quem removeu quem, com qual papel e quando -, fechando a ultima lacuna de rastreabilidade da proposta."
P "Esta e a versao 2.65 do manual. Em relacao a versao 2.64 (11/09/2026), a tabela de Cronograma do Planejamento (Passo 4, aba 3) ganhou a Fase 1 de uma nova proposta de melhoria: um marco pode, opcionalmente, ganhar uma lista de tarefas internas (nome e concluida), aberta pelo botao `+ Tarefas` na ultima coluna, sem data ou responsavel proprios - quem tem isso e o marco. O botao mostra quantas ja foram concluidas (ex.: `Tarefas (2/3)`) e fica com contorno laranja enquanto houver pendencia. Nao aparece no espelho do TAP nem no Relatorio de Situacao ainda - fica so na tabela de edicao do Planejamento por enquanto."
P "Esta e a versao 2.66 do manual. Em relacao a versao 2.65 (14/09/2026), a Fase 2 fecha a proposta de detalhamento do cronograma: a Atividade (nivel 3) da EAP (Passo 5) ganhou o mesmo status automatico que a Entrega ja tinha, um nivel mais fundo - casando pelo nome da atividade com uma tarefa de qualquer marco do cronograma (Passo 4). Como tarefa nao tem data, os selos possiveis sao so `Concluida`, `Pendente` ou `Sem tarefa vinculada` - sem o estado `Atrasada` que a Entrega tem. Fecha a amarracao pedida: agora da para ver na EAP o que ficou pendente no detalhamento do cronograma."
P "Esta e a versao 2.67 do manual. Em relacao a versao 2.66 (14/09/2026), a Fase 3 (opcional) fecha de vez a proposta de detalhamento do cronograma: a tabela de Cronograma do Planejamento (Passo 4, aba 3) ganhou um resumo do percentual de tarefas concluidas somando todos os marcos, ao lado do botao `+ Adicionar linha`, e o Painel Executivo (secao 6.4) ganhou a secao `Tarefas do cronograma (portfolio)`, com o mesmo percentual agregado de todos os projetos e em quantos deles existe pelo menos uma tarefa pendente."
P "Esta e a versao 2.68 do manual. Em relacao a versao 2.67 (14/09/2026), foi corrigido um aperto de layout na tabela de Cronograma do Planejamento (Passo 4, aba 3): a largura minima da tabela aumentou (para acomodar as 9 colunas de forma legivel, apos a coluna Tarefas ter sido adicionada na Fase 1 do detalhamento do cronograma) e as colunas Marco/Entrega, Responsavel e as tres datas ganharam mais espaco relativo, as custas de Custo e Depende nº, que normalmente tem conteudo mais curto. Sem mudanca de comportamento - apenas ajuste de layout, corrigindo texto cortado/muito espremido em campos com nomes ou responsaveis mais longos."
P "Esta e a versao 2.69 do manual. Em relacao a versao 2.68 (14/09/2026), o mesmo ajuste de layout foi estendido as tabelas de Verificacao e Analise Critica do Planejamento (Passo 4, abas 6 e 8), que tinham o mesmo problema: a largura minima de ambas aumentou e suas colunas de texto mais longas ganharam mais espaco relativo. As tabelas de Saidas e Validacao (abas 5 e 7) tambem receberam um ajuste menor, dando largura propria as colunas de selecao (Status/Resultado/Aprovacao), que antes ficavam mais estreitas que o restante sem necessidade. Sem mudanca de comportamento - apenas ajuste de layout."
P "Esta e a versao 2.70 do manual. Em relacao a versao 2.69 (14/09/2026), foi adicionada a nova secao 2.14 - Edicao simultanea: os 12 formularios de registro (e os 3 relatorios de chave unica - Plano de Comunicacao, Relatorio de Situacao e Relatorio de Entregas) passam a detectar quando outra pessoa salvou o mesmo registro enquanto alguem editava, bloqueando o salvamento com um aviso claro em vez de sobrescrever silenciosamente as mudancas da outra pessoa. Nada e perdido automaticamente - o formulario continua aberto com o texto digitado, e a pessoa so precisa atualizar a pagina e refazer a alteracao."
P "Esta e a versao 2.71 do manual. Em relacao a versao 2.70 (14/09/2026), a secao 2.14 - Edicao simultanea ganhou um segundo mecanismo, preventivo: assim que alguem abre um registro para editar, uma faixa amarela avisa quem mais estiver editando aquele mesmo registro (ou relatorio, nos tres de tela unica) e desde quando, renovada a cada minuto e expirando sozinha apos cerca de 10 minutos sem atividade. E so um aviso - nao bloqueia a edicao - e complementa a trava na hora de salvar, ja existente desde a versao 2.70."
P "Esta e a versao 2.72 do manual. Em relacao a versao 2.71 (14/09/2026), a secao 2.14 - Edicao simultanea ganhou um terceiro mecanismo, ao vivo: se outra pessoa efetivamente salvar uma alteracao no mesmo registro enquanto alguem esta editando, uma faixa vermelha aparece na hora no topo da pagina avisando quem salvou, sem precisar esperar ate a proxima tentativa de salvar para descobrir. Usa atualizacao em tempo real do banco de dados (Supabase Realtime) - nao muda o formato de nenhum dado salvo nem exige nenhuma acao nova do usuario, so torna o aviso mais imediato."
P "Esta e a versao 2.73 do manual. Em relacao a versao 2.72 (14/09/2026), a trava de conflito da secao 2.14 - Edicao simultanea foi estendida as acoes rapidas de aprovar/reprovar direto da listagem (Gate 1/Gate 2), que antes nao passavam por ela - so o formulario completo de edicao estava protegido. Sem mudanca visivel de fluxo: a acao continua sendo um clique so, mas agora busca o estado mais recente do registro antes de gravar, evitando sobrescrever silenciosamente uma mudanca feita por outra pessoa entre a abertura da ficha e a decisao do gate."
P "Esta e a versao 2.74 do manual. Em relacao a versao 2.73 (14/09/2026), o Relatorio de Situacao (secao 4.1) e o Relatorio de Entregas (secao 4.2) pararam de bloquear o salvamento inteiro quando duas pessoas editam projetos diferentes do mesmo relatorio ao mesmo tempo - agora as edicoes sao mescladas automaticamente por projeto/item, e so bloqueiam quando ha colisao real (mesma linha, ou os campos gerais do relatorio). O Plano de Comunicacao nao muda, por ser um documento unico sem lista de projetos."
P "Esta e a versao 2.76 do manual. Em relacao a versao 2.75 (15/09/2026), foi adicionada a nova secao 2.16 - Quem pode imprimir ou exportar um registro: nos 7 formularios com restricao por equipe, os botoes Imprimir e Exportar Word de um registro ja salvo passam a exigir fazer parte da equipe do projeto (ou ser Admin), sem excecao para o Gerente de Projetos. Os botoes ficam ocultos para quem nao tem permissao. O documento Regras de Acesso e Permissoes tambem foi atualizado (secao 5.4)."
P "Esta e a versao 2.77 do manual. Em relacao a versao 2.76 (15/09/2026), foi adicionada a nova pagina Painel de Prazos (secao 6.4.2) - uma 17a pagina em `Gestao`, pensada para ficar aberta numa TV/monitor, com atualizacao automatica a cada 2 minutos, mostrando o cronograma de todos os projetos ativos com marcos atrasados e vencendo em breve em destaque. Junto com ela, um mecanismo diario e automatico de aviso por e-mail/push (secao 6.4.3) passou a notificar quem e responsavel por um marco (5 dias antes, 2 dias antes, no dia e, repetido semanalmente, quando ja atrasado) - o cronograma do Planejamento (Passo 4, aba 3) ganhou o campo opcional `Responsavel (conta)` para vincular cada marco a uma conta real da equipe do projeto, alem do texto livre Responsavel ja existente. O documento Regras de Acesso e Permissoes tambem foi atualizado (secao 3.4)."
P "Esta e a versao 2.78 do manual. Em relacao a versao 2.77 (15/09/2026), `+ Adicionar usuario` (secao 6.1) passou a criar a conta de verdade na hora (via Supabase Auth Admin API), em vez de so um pre-cadastro que virava conta no primeiro login - agora a pessoa ja pode ser adicionada a equipe de um projeto (secao 6.2) sem esperar ela logar. Pre-cadastros antigos ainda funcionam e ganharam um botao `Criar conta agora` para converte-los imediatamente."
P "Esta e a versao 2.79 do manual. Em relacao a versao 2.78 (15/09/2026), o e-mail de aviso automatico de prazo (secao 6.4.3) passou a incluir um link direto para a Ficha do Projeto do projeto correspondente."
P "Esta e a versao 2.80 do manual. Em relacao a versao 2.79 (15/09/2026), o link da Ficha do Projeto no e-mail de aviso automatico de prazo (secao 6.4.3) foi substituido por uma indicacao da Gerencia de Projetos como alternativa de contato, com o nome do Gerente puxado do proprio campo `Gerente do projeto` do Planejamento."
P "Esta e a versao 2.81 do manual. Em relacao a versao 2.80 (15/09/2026), foi corrigido um bug critico de mesclagem no Relatorio de Situacao (secao 4.1) e no Relatorio de Entregas (secao 4.2): edicao concorrente em um item sem identidade estavel podia colapsar ou duplicar a lista salva. Foi adicionada a nova pagina Meu Painel (secao 6.4.4) - uma 18a pagina em `Gestao`, sem restricao de papel, reunindo marcos, tarefas, riscos, entradas e entraves/encaminhamentos atribuidos a conta logada, com um botao para concluir cada item direto da lista. De apoio a essa nova pagina, o campo `Responsavel (conta)` (ja existente no cronograma desde a versao 2.77) foi estendido ao Registro de riscos do TAP (Passo 3), a aba Entradas do Planejamento (Passo 4) e aos Entraves/Encaminhamentos do Relatorio de Situacao (secao 4.1)."
P "Esta e a versao 2.82 do manual. Em relacao a versao 2.81 (15/09/2026), foi adicionada a nova pagina Gestao de Atividades (secao 6.4.5) - uma 19a pagina em `Gestao`, com acesso configuravel por papel (padrao Admin), que agrega os mesmos itens do Meu Painel por conta responsavel, dando ao PMO uma visao da carga de trabalho de toda a equipe num cartao por pessoa, alem de um indicador dos itens em aberto ainda sem Responsavel (conta) atribuido. De passagem, foi corrigida uma lacuna encontrada na configuracao de acesso do Painel de Prazos (secao 6.4.2): a linha correspondente nunca tinha sido criada na tabela de configuracoes do sistema, entao o Admin nao conseguia de fato salvar pela tela quem mais podia ver aquela pagina, apesar dos campos aparecerem normalmente."
P "Esta e a versao 2.83 do manual. Em relacao a versao 2.82 (16/09/2026), o campo `Responsavel (conta)` (cronograma do Planejamento, Registro de riscos do TAP, Entradas do Planejamento e Entraves/Encaminhamentos do Relatorio de Situacao) passou a aceitar mais de uma pessoa por item - um painel com uma caixa de selecao por conta, em vez de um menu de escolha unica - e a mostrar o nome cadastrado de cada conta em vez do e-mail (caindo para o e-mail so quando a conta nao tem nome cadastrado). Foi tambem corrigido um bug que deixava o texto de cada opcao invisivel dentro desse painel (so o quadrado de selecao aparecia) em navegadores reais, causado por uma regra de estilo generica do formulario que forcava a largura do quadrado de selecao."
P "Esta e a versao 2.84 do manual. Em relacao a versao 2.83 (16/09/2026), o e-mail de aviso automatico de prazo (secao 6.4.3) passou a trazer um link `Abrir o Meu Painel`, para quem recebe o aviso ja ir direto tratar o item, em vez de precisar abrir o sistema e navegar ate a tela por conta propria."
P "Esta e a versao 2.85 do manual. Em relacao a versao 2.84 (16/09/2026), cada marco do cronograma (Passo 4) ganhou uma coluna Execucao - autorrelato de A iniciar/Em andamento/Atrasado/Concluido, com a data da ultima mudanca gravada automaticamente. Diferente do calculo automatico de atraso (que so olha as datas), e a propria pessoa sinalizando se esta de fato trabalhando no marco - visivel/editavel tambem no Meu Painel (secao 6.4.4) e exibido ao lado de cada marco atrasado ou vencendo no Painel de Prazos (secao 6.4.2), que passa a destacar os marcos sem nenhuma atualizacao de execucao."
P "Esta e a versao 2.86 do manual. Em relacao a versao 2.85 (17/09/2026), o aviso automatico de prazo (secao 6.4.3) ganhou um segundo mecanismo, de silencio: quando um marco ja comecou e fica 5 dias ou mais sem nenhuma atualizacao no campo Execucao, um e-mail avisa os mesmos destinatarios do aviso de prazo, repetindo a cada 5 dias enquanto continuar parado. Diferente do aviso de prazo (que so olha a data final), este dispara por causa da inatividade - fecha a lacuna de um marco com vencimento ainda distante que podia ficar semanas parado sem ninguem ser avisado."
P "Esta e a versao 2.87 do manual. Em relacao a versao 2.86 (17/09/2026), os dias das duas regras do aviso automatico (secao 6.4.3) - janelas do aviso de prazo, repeticao do atrasado, limite e repeticao do aviso de silencio - deixaram de ser fixos no codigo e passaram a ser configuraveis em Administracao > Configuracoes (secao 6.3), num novo bloco `Regras de aviso automatico de prazo e execucao`, no mesmo padrao ja usado pelas metas do Relatorio de Resultados."
P "Esta e a versao 2.88 do manual. Em relacao a versao 2.87 (17/09/2026), a Gestao de Atividades (secao 6.4.5) passou a agrupar Marcos, Tarefas, Riscos (TAP) e Entradas (Planejamento) por projeto dentro do cartao de cada pessoa, em vez de listar os itens soltos: cada grupo de projeto mostra a contagem de itens e, em Marcos, quantos estao atrasados, com os grupos de Marcos ordenados pelo mais critico primeiro; cada marco tambem passou a indicar explicitamente o campo da data (`Termino previsto`). Pessoas com pelo menos um item atrasado agora abrem automaticamente ao carregar a pagina."
P "Esta e a versao 2.89 do manual. Em relacao a versao 2.88 (17/09/2026), o campo Anexos da aba Viabilidade no Planejamento e Desenvolvimento de Projeto (Passo 4) deixou de ser so um texto livre e ganhou upload real de arquivo (.txt/.docx/.pdf, ate 15 MB), no mesmo padrao ja usado pela Solicitacao de Demanda (Passo 1) e pela SMP (Passo 7) - com lista de arquivos enviados, download e remocao. O texto livre antigo continua disponivel, renomeado para `Links de documentos`."
P "Esta e a versao 2.90 do manual. Em relacao a versao 2.89 (17/09/2026), a Ficha do Projeto (secao 6.4.1) ganhou a secao `Anexos do projeto`, reunindo num so lugar todo arquivo enviado nos 3 pontos do sistema que aceitam upload real - Solicitacao de Demanda, SMP e a aba Viabilidade do Planejamento - com a origem de cada arquivo e um botao para baixar direto, sem precisar abrir cada formulario separadamente."
P "Esta e a versao 2.91 do manual. Em relacao a versao 2.90 (17/09/2026), a EAP (Passo 5) trocou o assistente `Sugerir com IA` pelo botao `Importar do Cronograma`: em vez de a IA redigir pacotes e entregas com texto proprio, o novo botao copia os nomes dos marcos do Cronograma do Planejamento sem nenhuma reformulacao, garantindo que o casamento automatico Entrega x Marco (usado para o status automatico da entrega e na Ficha do Projeto) sempre funcione."
P "Esta e a versao 2.92 do manual. Em relacao a versao 2.91 (17/09/2026), a ferramenta passou a se chamar SGP - Sistema de Gestao de Projetos, sem vinculo com nenhuma instituicao: os codigos dos 12 formularios passaram a ser SGP-01 a SGP-12 (iguais ao numero de cada formulario, com a EAP ganhando o SGP-05), o login passou a ser so por link magico por e-mail (o botao Entrar com Microsoft foi desativado), as notificacoes passaram a sair do e-mail do proprio sistema e todas as capturas de tela foram refeitas com a nova identidade visual e os dados ficticios de demonstracao."
P "O manual nao substitui as Diretrizes para a Gestao de Projetos (documento que define o framework D01 a D07) nem o documento Regras de Acesso e Permissoes (que detalha cada regra de controle de acesso); ele e o guia operacional de como usar cada ferramenta na pratica."

H2 "1.2 Visao geral da ferramenta"
P "A ferramenta e composta por 16 paginas eletronicas independentes:"
Bul "12 formularios de registro (codigo SGP-01 a SGP-12), cada um correspondendo a um artefato institucional - da Solicitacao de Demanda ao Relatorio de Entregas e Beneficios."
Bul "1 ferramenta de apoio a decisao, o Validador de Projetos, que nao gera registros de artefato mas ajuda a avaliar projetos."
Bul "1 pagina de Administracao, restrita a Admin, para gerenciar usuarios, equipes e configuracoes do sistema."
Bul "1 Painel Executivo, tambem restrito a Admin, com a visao agregada de todo o sistema (secao 6.4)."
Bul "1 Ficha do Projeto, a visao consolidada de um unico projeto, aberta a partir do Painel Executivo e sem restricao por papel (secao 6.4.1)."
P "Cada formulario de registro tem duas partes: uma tela de preenchimento (`+Novo...`) e uma tela de consulta (`...cadastrados`), onde ficam listados todos os registros ja salvos, com opcao de abrir, editar o status ou excluir cada um."
P "O Mapa de Diretrizes (pagina inicial) apresenta visualmente as 7 diretrizes (D01 a D07) e mostra, para cada etapa do processo, quais ferramentas usar. E o ponto de partida recomendado para quem quer entender o processo antes de preencher os formularios - e a unica pagina, junto com o Validador de Projetos, que nao exige login."
Img "00_home.png" "Mapa de diretrizes - pagina inicial. Nao exige login." 5.8

H2 "1.3 Como acessar e fazer login"
P "Todos os 12 formularios de registro e a pagina de Administracao exigem login antes de carregar ou salvar qualquer dado. Ao abrir qualquer um deles sem sessao ativa, aparece a tela de entrada:"
Img "01_login_gate.png" "Tela de login, exibida ao abrir qualquer formulario sem sessao ativa." 4.6
P "O login e feito sem necessidade de senha:"
Bul "Link magico por e-mail - digite seu e-mail e clique em `Enviar link de acesso`. Um link chega no seu e-mail; clique nele e a pagina de origem atualiza sozinha."
Nota "No primeiro login, cada pessoa recebe automaticamente o papel Solicitante. Para obter outro papel (ex.: Gerente de Projetos, Admin), peca a um Admin para altera-lo em Administracao > Usuarios (secao 6)."
P "Depois de logado, uma barra preta no topo da pagina mostra `Conectado como [seu e-mail]` e o seu papel atual, com um botao `Sair`."

H2 "1.4 Como a ferramenta se relaciona com as Diretrizes (D01-D07)"
P "Os 12 formularios cobrem os 13 artefatos previstos nas Diretrizes. O Reporte de Resultados (D06.5) nao tem um formulario proprio - e atendido pela secao `Resultados alcancados` do Relatorio de Situacao (SGP-11, secao 4.1)."
$rows = @(
  @("Passo","Formulario","Diretriz / Estrategia"),
  @("1","Solicitacao de Demanda (SGP-01)","D01.1 - Recebimento e Registro de Demandas"),
  @("-","Gate 1 - Triagem","D01.4/D01.5 - aprovacao do Gestor Responsavel (Admin)"),
  @("2","Canvas de Projeto (SGP-02)","D01.6 - Elaboracao do Canvas de Projeto"),
  @("3","TAP - Termo de Abertura (SGP-03)","D02.1 - Desenvolvimento do TAP"),
  @("4","Planejamento e Desenvolvimento (SGP-04)","D02.2 a D02.10"),
  @("5","EAP - Estrutura Analitica do Projeto","D02.9 - Desenvolvimento da EAP"),
  @("-","Gate 2 - Pactuacao","D06.1/D06.2 - autorizacao da Alta Gestao"),
  @("6","Ata de Reuniao (SGP-07)","D01 a D07 - transversal, qualquer interacao formal"),
  @("7","SMP - Solicitacao de Mudanca (SGP-06)","D04.4 - Controle Integrado de Mudancas"),
  @("8","TEP - Termo de Encerramento (SGP-09)","D05.1.9 - Emissao do Termo de Encerramento"),
  @("9","RLA - Registro de Licoes Aprendidas (SGP-10)","D05.1.6 - Conducao de Licoes Aprendidas"),
  @("10","Plano de Comunicacao de Projeto (SGP-08)","D03.5 - Planejamento das Comunicacoes"),
  @("11","Relatorio de Situacao de Projetos (SGP-11)","D04.3 - Elaboracao de Relatorios"),
  @("12","Relatorio de Entregas e Beneficios (SGP-12)","D06.1/D06.2 - Governanca e Tomada de Decisao")
)
TableSimple $rows @(1.5,7.5,7.0)

# ============================================================
# 2. ANTES DE COMECAR
# ============================================================
H1 "2. Antes de comecar"

H2 "2.1 Onde encontrar as ferramentas"
P "O Mapa de Diretrizes (pagina inicial) e o ponto central: o menu lateral lista os 12 formularios, o Validador de Projetos e a Administracao. A tabela no Anexo (secao 8) traz o nome de cada artefato e seu codigo, para referencia rapida."

H2 "2.2 Convencoes usadas neste manual"
Bul "Campos marcados com asterisco (*) nos formularios sao obrigatorios. O sistema nao deixa registrar enquanto algum campo obrigatorio estiver vazio - ele destaca o campo em vermelho e rola a tela ate ele."
Bul "Ao salvar, cada formulario gera um numero de protocolo automatico no formato CODIGO-ANO-SEQUENCIAL (ex.: SGP-01-2026-001). Guarde esse numero para localizar o registro depois."
Bul "Todo registro tem um status (ex.: Pendente de aprovacao, Aprovado, Reprovado), alteravel a partir da tela de consulta - em alguns formularios, so o Admin pode altera-lo (ver secao 2.4)."
Bul "A tela de consulta de cada formulario tem um campo de busca. Ao abrir um registro especifico, o rodape do painel lateral traz os botoes `Imprimir` e `Exportar Word`, que geram o documento no formato padrao do respectivo codigo SGP, pronto para impressao ou para abrir no Word."

H2 "2.3 Papeis de usuario e o que cada um pode fazer"
P "Cada pessoa que faz login recebe um papel, usado para liberar ou restringir acoes especificas no sistema. Papeis sao atribuidos e alterados por um Admin, em Administracao > Usuarios."
$rows2 = @(
  @("Papel","Observacao"),
  @("Solicitante","Papel padrao, atribuido automaticamente a todo novo usuario no primeiro login."),
  @("Gerente de Projetos","Sem restricoes adicionais alem das descritas neste manual."),
  @("Gestor Responsavel","Sem restricoes adicionais alem das descritas neste manual."),
  @("Dono do Negocio","Sem restricoes adicionais alem das descritas neste manual."),
  @("Alta Gestao","Sem restricoes adicionais alem das descritas neste manual."),
  @("Admin (PMO/Admin)","Papel de administracao do sistema - unico que pode executar as acoes da lista abaixo.")
)
TableSimple $rows2 @(5.0,11.0)
P "Acoes exclusivas de Admin:"
Bul "Acessar a pagina de Administracao (qualquer outro papel ve `Acesso restrito`)."
Bul "Alterar o papel de outros usuarios."
Bul "Adicionar ou remover membros da equipe de um projeto."
Bul "Ativar ou desativar as opcoes de `sem login` da Solicitacao de Demanda e da Ata de Reuniao."
Bul "Aprovar ou reprovar o Gate 1 na Solicitacao de Demanda."
Nota "Um Admin e sempre considerado membro de qualquer equipe de projeto automaticamente - nao precisa ser adicionado manualmente para editar registros vinculados a um projeto (ver secao 2.5)."

H2 "2.4 Os dois gates de aprovacao"
P "O sistema tem dois pontos de decisao formal no ciclo de vida de um projeto:"
Bul "Gate 1 - Triagem: ocorre logo apos o registro da demanda. So o Admin consegue alterar o status da Solicitacao de Demanda para Aprovada ou Reprovada; qualquer outro papel ve o campo travado."
Bul "Gate 2 - Pactuacao: ocorre ao final do planejamento, quando o Relatorio de Entregas e Beneficios e apresentado e pactuado com o Dono do Negocio perante a Alta Gestao (ver secao 4.2)."
Img "14_f01_gate1.png" "Detalhe de uma Solicitacao de Demanda - o campo Status, no rodape, so pode ser alterado pelo Admin (Gate 1)." 5.4

H2 "2.5 Restricao por equipe do projeto"
P "Sete dos doze formularios exigem vincular o registro a um `Projeto vinculado` (um projeto ja Aprovado no Gate 1): Canvas, TAP, Planejamento e Desenvolvimento, EAP, SMP, TEP e RLA. Nesses formularios, apenas quem faz parte da equipe daquele projeto - ou um Admin - pode criar ou editar um registro vinculado a ele."
Img "17_f03_novo.png" "TAP com um projeto vinculado selecionado: unidade, gerente e status sao preenchidos automaticamente a partir do projeto." 5.6
Bul "Se a pessoa nao for da equipe daquele projeto (e nao for Admin), aparece um aviso na tela e o botao de enviar/registrar fica desabilitado."
Bul "A equipe de cada projeto e definida por um Admin, em Administracao > Equipes (secao 6)."
Bul "A mesma regra vale para Imprimir e Exportar Word (secao 2.16): nesses 7 formularios, so quem e da equipe do projeto (ou Admin) ve esses dois botoes num registro ja salvo."
Nota "Os outros 5 formularios (Solicitacao de Demanda, Ata de Reuniao, Plano de Comunicacao, Relatorio de Situacao e Relatorio de Entregas) nao usam esse conceito de forma obrigatoria - qualquer usuario autenticado pode criar/editar registros neles, sem restricao por equipe. O Relatorio de Situacao e o Relatorio de Entregas ganharam, no entanto, um campo opcional `Projeto vinculado (Gate 1)` em cada projeto listado (secoes 4.1 e 4.2) - ele so serve para ligar aquele projeto ao historico compartilhado, sem aplicar a trava de equipe."
Nota "Lista mostra so projetos ativos: o campo `Projeto vinculado` so oferece projetos ainda em andamento para NOVOS vinculos - assim que um TEP e registrado para um projeto (Passo 8), ele some da lista de opcoes em todos os formularios, inclusive Relatorio de Situacao e Relatorio de Entregas. Um documento que ja estava vinculado a esse projeto antes dele encerrar continua mostrando o vinculo normalmente, so que com o status ao lado (ex.: `Nome do Projeto - Concluido`)."

H2 "2.6 Acesso sem login - Solicitacao de Demanda e Ata de Reuniao"
P "Dois formularios podem, opcionalmente, ser abertos por um Admin para aceitar registros sem login: Solicitacao de Demanda e Ata de Reuniao. Cada um tem um interruptor independente em Administracao > Configuracoes (secao 6) - podem estar ligados, desligados ou so um dos dois."
Img "02_sem_login_gate.png" "Quando a opcao esta ativa, a tela de login ganha um terceiro caminho: Continuar sem login." 4.6
Img "03_sem_login_dados.png" "A pessoa sem login informa apenas nome completo e e-mail para registrar." 4.6
Bul "Quem usa esse caminho so consegue criar um registro novo - nunca editar, excluir ou alterar status, mesmo em um registro que ela mesma criou."
Bul "O registro criado sem login recebe um selo `Sem login` na listagem, junto com o nome e e-mail informados."

H2 "2.7 Onde os dados ficam armazenados"
P "Todos os 12 formularios de registro gravam os dados em um banco de dados real (PostgreSQL, hospedado no Supabase), associados a conta autenticada de quem salvou o registro (ou ao nome/e-mail informado, no caso de registro sem login). Nao ha mais uma chave unica compartilhada por todos: cada sessao usa o token da propria conta, e o acesso as ferramentas segue as regras de login, papel e equipe descritas neste capitulo."
P "Cada formulario ainda grava seus registros de forma independente - nao existe herenca automatica de informacoes entre eles. Ao preencher o TAP logo depois do Canvas, por exemplo, e preciso informar novamente o nome do projeto, a unidade e o gerente (ou selecionar o mesmo `Projeto vinculado`, quando o formulario tiver esse campo). Preencha os dados de identificacao da mesma forma em todos os formularios de um mesmo projeto, para manter a rastreabilidade entre eles."

H2 "2.8 Notificacoes por e-mail em transicoes de estado"
P "Sete momentos do sistema disparam automaticamente um e-mail de notificacao, enviado por uma Supabase Edge Function pelo servidor de e-mail (SMTP) do sistema (as credenciais ficam guardadas no servidor, nunca no navegador):"
Bul "Registro de nova Solicitacao de Demanda: assim que a demanda e salva, antes de qualquer decisao - avisa os Admins que ha um Gate 1 - Triagem pendente de analise."
Bul "Gate 1 - Solicitacao de Demanda: quando o Admin muda o status para `Aprovada` ou `Reprovada`."
Bul "Canvas de Projeto: quando o status muda para `Aprovado`."
Bul "TAP: quando o status muda para `Aprovado`."
Bul "SMP: quando a decisao e finalizada como `Aprovada` ou `Nao aprovada`."
Bul "TEP: ao registrar um novo Termo de Encerramento."
Bul "Gate 2 - Relatorio de Entregas: quando o status muda para `Pactuado`."
P "Nos seis momentos de decisao (todos exceto o registro de nova demanda), os destinatarios sao: quem criou o registro (e-mail capturado no login, ou informado no registro sem login), todos os e-mails cadastrados na equipe do Projeto vinculado (Administracao > Equipes) e todos os usuarios com papel Admin. Ja o aviso de registro de nova demanda vai somente para os Admins, ja que ainda nao ha Gate 1 aprovado nem equipe designada nesse momento. Se o envio falhar por qualquer motivo (ex.: instabilidade do provedor de e-mail), o registro e salvo normalmente mesmo assim - a notificacao nunca bloqueia o fluxo de trabalho."

H2 "2.9 Lembretes por notificacao push no celular"
P "Alem do e-mail (secao 2.8), o sistema oferece um segundo canal de aviso: uma notificacao push, que aparece no celular ou no computador mesmo com o navegador fechado, igual a uma notificacao de aplicativo. E opcional e ativada individualmente por cada usuario - nao existe um interruptor geral do Admin para isso."
Bul "Onde ativar: na pagina inicial (mapa de diretrizes), logado, aparece o botao `Ativar lembretes no celular` na barra preta do topo."
Bul "Ao tocar no botao, o navegador pede permissao para mostrar notificacoes - aceite para concluir a ativacao. O botao muda para `Lembretes ativos`."
Bul "Para desativar, basta tocar novamente no botao ja ativo."
Bul "Dispara nos mesmos sete momentos da secao 2.8 (registro de nova demanda, Gate 1, Canvas aprovado, TAP aprovado, decisao da SMP, registro de TEP e Gate 2 pactuado), para os mesmos destinatarios - mas so quem tiver ativado o push naquele navegador/celular recebe a notificacao."
Nota "No iPhone/iPad, a Apple so libera notificacao push para sites adicionados a Tela de Inicio. Antes de tocar em `Ativar lembretes no celular`, va em Compartilhar > Adicionar a Tela de Inicio no Safari, abra o sistema pelo icone criado, e so entao ative - num navegador comum (aba do Safari) o botao explica essa limitacao em vez de ativar direto. Em Android e computador nao ha essa exigencia."

H2 "2.10 Importacao de transcricao por IA (Ata de Reuniao)"
P "A Ata de Reuniao pode, opcionalmente, preencher pauta, data/horario, participantes, resumo, encaminhamentos e entraves automaticamente a partir da transcricao de uma reuniao - o texto e analisado por IA (Claude, via uma Supabase Edge Function; a chave da API fica guardada no servidor, nunca no navegador) e os campos sugeridos ficam disponiveis para revisao antes de salvar. O usuario nunca perde o controle: nada e enviado ao Supabase automaticamente, so o preenchimento do formulario."
P "O texto pode ser fornecido de duas formas: colando diretamente na caixa de texto, ou anexando um arquivo `.txt`, `.docx` (Word) ou `.pdf` pelo botao `Anexar arquivo` - o sistema extrai o texto do arquivo e preenche a caixa automaticamente para revisao antes de clicar em `Analisar e preencher`. Limite de 15 MB por arquivo."
Nota "PDFs escaneados (so imagem, sem texto real por tras) nao sao suportados - a extracao depende do PDF ter texto selecionavel. Nesse caso, copie e cole o texto manualmente."
P "Diferente do envio de e-mail e do push (secoes 2.8 e 2.9), essa funcionalidade tem um interruptor geral do Admin, desligado por padrao - cada analise tem um custo real de API. Em Administracao > Configuracoes (secao 6.3), o Admin:"
Bul "Liga ou desliga a importacao para o sistema inteiro - desligada, ninguem tem acesso ao painel, nem o proprio Admin."
Bul "Com a importacao ligada, marca quais papeis podem usa-la (pode marcar mais de um) - PMO/Admin sempre tem acesso quando a importacao esta ligada, mesmo sem estar marcado na lista."
Nota "Quem nao tem permissao simplesmente nao ve o painel `Importar transcricao da reuniao` na Ata - nao ha mensagem de erro, o painel so aparece para quem pode usa-lo."
Bul "Cada campo preenchido pela IA (pauta, data/horario, resumo) e cada linha das tabelas de participantes/encaminhamentos/entraves ficam com um selo `IA` e destaque visual ate serem editados - o mesmo padrao usado no assistente de preenchimento por IA dos demais formularios (secao 2.12)."

H2 "2.11 Importacao de audio por IA (Ata de Reuniao)"
P "Alem de colar/anexar a transcricao em texto, a Ata de Reuniao pode transcrever a propria gravacao da reuniao: no painel `Importar transcricao da reuniao`, o botao `Anexar audio` aceita `.mp3`, `.m4a`, `.aac`, `.wav`, `.ogg` e outros formatos comuns. O sistema divide o audio automaticamente em pedacos de cerca de 8 minutos direto no navegador (por causa de um limite tecnico da API de transcricao) e transcreve cada pedaco por IA (OpenAI, via uma segunda Supabase Edge Function; a chave fica guardada no servidor, igual a da Claude). O texto de cada pedaco vai sendo concatenado na caixa de transcricao, junto com uma lista de progresso (`Parte 1 de N`, `Parte 2 de N`...) - reunioes longas podem levar alguns minutos para processar."
Bul "Se algum pedaco falhar (ex.: instabilidade de rede), aparece um botao `Tentar novamente` especifico para aquele pedaco - os pedacos ja transcritos com sucesso nao precisam ser refeitos."
Bul "Depois de transcrito, o texto fica disponivel na mesma caixa da importacao por texto, para revisar e so entao clicar em `Analisar e preencher` normalmente."
P "A importacao de audio tem uma regra de habilitacao SEPARADA da importacao de texto (secao 2.10) - dois interruptores independentes em Administracao > Configuracoes (secao 6.3), cada um desligado por padrao. Na pratica, o botao `Anexar audio` so aparece para quem tem as duas permissoes ativas ao mesmo tempo (importar transcricao E importar audio), ja que o audio so serve para gerar o texto que depois passa pela mesma analise."
Nota "PDFs e audio muito longos consomem mais tempo de processamento; nao ha limite de duracao da reuniao em si, mas o navegador precisa decodificar o arquivo inteiro antes de dividir em pedacos - arquivos muito grandes (varias horas) podem demorar mais para começar a transcrever."

H2 "2.12 Assistente de preenchimento por IA (Canvas, TAP, Planejamento, SMP, TEP, RLA, Relatorio de Situacao e Relatorio de Entregas)"
P "O Canvas de Projeto (Passo 2), o TAP (Passo 3), o Planejamento e Desenvolvimento de Projeto (Passo 4), a SMP (Passo 7), o TEP (Passo 8) e o RLA (Passo 9) tem um botao `Sugerir com IA`, que aparece assim que um Projeto vinculado e selecionado. Diferente da importacao de transcricao (secao 2.10), que le um texto colado pelo usuario, este assistente le os proprios documentos ja registrados do projeto - seguindo a esteira (Demanda alimenta o Canvas; Canvas e Demanda alimentam o TAP; TAP e Canvas alimentam o Planejamento; TAP e Planejamento alimentam a SMP; TAP, Planejamento e EAP alimentam o TEP; TEP e SMPs alimentam o RLA), sempre incluindo as Atas de Reuniao do projeto - e usa IA (Claude, pela mesma Supabase Edge Function em espirito da secao 2.10) para sugerir o preenchimento dos campos ainda vazios. A EAP (Passo 5) NAO tem mais este assistente - ver secao 5 do Passo 5, `Importar do Cronograma`."
Bul "So preenche o que estiver vazio: um campo ja digitado pelo usuario nunca e sobrescrito pela sugestao."
Bul "Cada campo preenchido pela IA fica com um selo `IA` e destaque visual ate ser editado - assim fica claro, campo a campo, o que veio de sugestao e o que foi escrito pela pessoa."
Bul "Uma linha no rodape do bloco lista os documentos usados como base (ex.: `Sugestao baseada em: Solicitacao de Demanda SGP-01-2026-001, Canvas de Projeto SGP-02-2026-002`)."
Bul "No TAP, alem dos campos de texto, a sugestao tambem propoe linhas iniciais para as tabelas de riscos, cronograma de entregas macro, custos e partes interessadas - sempre que a tabela ainda estiver vazia. No Planejamento, propoe tambem marcos iniciais na aba Cronograma."
Bul "No Planejamento, o Solicitante (capa) e copiado direto da Solicitacao de Demanda, e as Macro fases do projeto (aba Pre-projeto) sao um resumo em texto do cronograma de entregas macro do TAP. Na aba Viabilidade, a Introducao, a Situacao atual e a Proposta de mudanca sao redigidas com a mesma base factual do Contexto e do Objetivo geral do Pre-projeto, mas com redacao propria para o proposito de cada campo - nunca uma copia literal de um campo para o outro. O campo Negocio (capa) fica de fora do assistente, por nao ter fonte confiavel nos documentos lidos."
Bul "Na SMP, a logica tambem e diferente: como o TAP e o Planejamento so mostram o que foi combinado originalmente (nao a mudanca em si), a IA procura nas Atas mais recentes alguma mudanca de fato sendo discutida e contrasta com esse combinado - nunca copia os campos do TAP/Planejamento diretamente. Se as Atas nao discutirem nenhuma mudanca concreta, a sugestao devolve todos os campos vazios em vez de inventar uma mudanca hipotetica."
Bul "No TEP, a IA compara o planejado (TAP, Planejamento, EAP) com o que as Atas mais recentes confirmam como realizado, para redigir a justificativa de encerramento, o resumo de atividades encerradas e as consideracoes finais."
Bul "No RLA, a IA le o TEP, as Atas e todas as SMPs do projeto para preencher as 4 secoes de texto livre - Visao geral, Destaques, Desafios (baseado nos entraves recorrentes das Atas e nos motivos de mudanca das SMPs) e Tarefas pos-projeto. Os blocos de avaliacao estruturada Sim/Nao/Parcial com placar percentual (Fase de planejamento, Execucao, Fatores humanos, Geral) nao sao preenchidos pela IA - continuam sendo uma autoavaliacao manual da equipe."
Bul "No Relatorio de Situacao (secao 4.1) e no Relatorio de Entregas (secao 4.2), o botao `Sugerir com IA` aparece POR LINHA de projeto, nao uma vez so para o formulario - ja que um unico relatorio cobre varios projetos ao mesmo tempo, cada um vinculado (ou nao) a um projeto diferente. No Relatorio de Situacao, a IA le o TAP e as Atas do projeto daquela linha (mais a situacao ja digitada nela - status, % execucao, datas, sinalizacao automatica de atraso) para sugerir o que Merece atencao e o que Merece destaque. No Relatorio de Entregas, le o TAP e as Atas do projeto vinculado para sugerir o Impacto/Resultado esperado do projeto - sempre redigido como uma expectativa, ja que o Gate 2 ocorre antes da execucao comecar, nunca como um resultado ja medido."
P "Mesma logica de custo e permissao da importacao de transcricao (secao 2.10): interruptor geral do Admin, desligado por padrao, e lista de papeis liberados em Administracao > Configuracoes (secao 6.3) - um unico interruptor cobre os 8 formularios acima, e PMO/Admin sempre tem acesso quando esta ligado. Alem disso, so quem faz parte da equipe do projeto selecionado (secao 2.5) pode usar o botao - no caso dos dois relatorios, a equipe conferida e a do projeto vinculado aquela linha especifica, nao do relatorio em si (que continua sem restricao por equipe)."
Nota "Como qualquer sugestao de IA, o resultado pode estar incompleto ou impreciso quando os documentos anteriores tambem estiverem - revise sempre antes de clicar em `Registrar`."

H2 "2.13 Importacao de documento por IA (Solicitacao de Demanda)"
P "A Solicitacao de Demanda (Passo 1) tem um painel `Importar documento preenchido`, no topo do formulario, para quem ja preencheu o formulario oficial SGP-01 em Word ou PDF fora do sistema e agora precisa transcrever esses dados para a tela. O texto pode ser colado diretamente, ou o arquivo (`.txt`, `.docx` ou `.pdf`) anexado pelo botao `Anexar arquivo` - o sistema extrai o texto automaticamente para revisao antes de clicar em `Analisar e preencher`. Limite de 15 MB por arquivo."
Bul "A IA identifica nome do projeto, solicitante, departamento, justificativa, objetivo, escopo, prazo, orcamento, partes interessadas e anexos, e preenche so os campos ainda vazios do formulario - exatamente como as demais importacoes por IA (secoes 2.10 e 2.12)."
Bul "Departamento e Prioridade so sao preenchidos se o valor identificado pela IA corresponder exatamente a uma das opcoes da lista suspensa - caso contrario, o campo fica em branco para preenchimento manual."
Nota "PDFs escaneados (so imagem, sem texto real por tras) nao sao suportados - a extracao depende do PDF ter texto selecionavel. Nesse caso, copie e cole o texto manualmente."
P "Mesma logica de custo e permissao das demais importacoes por IA: interruptor geral do Admin, desligado por padrao, e lista de papeis liberados em Administracao > Configuracoes (secao 6.3). Diferente das secoes 2.10 e 2.12, nao ha restricao por equipe de projeto aqui - a Solicitacao de Demanda e o primeiro formulario, preenchida antes de existir um projeto aprovado - mas o acesso sem login (secao 2.6) nunca tem acesso a este painel, mesmo com a importacao habilitada."

H2 "2.14 Edicao simultanea"
P "Como mais de uma pessoa pode editar o mesmo registro (a equipe de um projeto - secao 2.5 - ou qualquer usuario nos formularios sem restricao de equipe), o sistema avisa quando isso realmente colide, em vez de uma pessoa salvar por cima da outra sem perceber."
P "Ao abrir um registro para editar, o sistema guarda o horario da ultima alteracao conhecida daquele registro. Se, entre esse momento e o clique em salvar, outra pessoa salvar o mesmo registro primeiro, o sistema bloqueia o salvamento e mostra o aviso: `Este registro foi alterado por outra pessoa enquanto voce editava`."
Bul "Nada e perdido automaticamente: o formulario continua aberto com o que a pessoa digitou, exatamente como estava."
Bul "Para continuar, e preciso atualizar a pagina (para carregar a versao mais recente, ja com a alteracao da outra pessoa) e refazer manualmente as mudancas."
Bul "Vale para os 12 formularios de registro. No Plano de Comunicacao (Passo 10, documento unico, sem lista de projetos), o aviso aparece se qualquer outra pessoa salvar entre o momento em que a tela de edicao foi aberta e o `Salvar`."
Nota "Isso nao e uma trava de `quem chega primeiro bloqueia o registro` - duas pessoas podem abrir o mesmo registro ao mesmo tempo sem problema. O aviso so aparece na hora de salvar, e so para quem salvar por ultimo."
P "No Relatorio de Situacao (secao 4.1) e no Relatorio de Entregas (secao 4.2) - que reunem varios projetos numa unica tela - o sistema e mais esperto: se uma pessoa editar o projeto A e outra editar o projeto B (ou outra parte independente do relatorio) ao mesmo tempo, as duas alteracoes sao salvas juntas automaticamente, sem bloquear ninguem. O aviso de conflito so aparece quando as edicoes realmente colidem: a mesma linha alterada pelas duas pessoas, ou os campos gerais do relatorio (referencia, responsavel, situacao financeira, pactuacao do Gate 2) alterados por outra pessoa durante a edicao."
P "Alem dessa trava na hora de salvar, o sistema tambem avisa preventivamente, assim que um registro e aberto para edicao: se outra pessoa ja estiver editando esse mesmo registro (ou o mesmo relatorio, nos tres de tela unica), aparece uma faixa amarela no topo da pagina - `Fulano esta editando este registro desde HH:MM. Combine com essa pessoa antes de salvar, para evitar conflito`."
Bul "E so um aviso - nao impede a pessoa de continuar editando e tentando salvar."
Bul "Enquanto o registro ficar aberto para edicao, o sistema renova esse aviso a cada minuto, para que outras pessoas continuem vendo que alguem esta editando."
Bul "O aviso some sozinho depois de cerca de 10 minutos sem renovacao (por exemplo, se a outra pessoa fechar a aba sem clicar em nada) - nao e preciso esperar mais que isso para editar em paz."
P "Por fim, ha um terceiro aviso, ao vivo: se, enquanto alguem esta editando, outra pessoa efetivamente salvar uma alteracao nesse mesmo registro, uma faixa vermelha aparece na hora no topo da pagina - `Fulano salvou uma alteracao neste registro agora, enquanto voce editava. Atualize a pagina antes de continuar, para nao perder nada.` - sem precisar esperar ate a proxima tentativa de salvar para descobrir. Basta clicar em `Ok` para dispensar o aviso."

H2 "2.15 Quem pode editar ou excluir um registro ja criado"
P "Alem da restricao por equipe (secao 2.5 - quem pode CRIAR um registro vinculado a um projeto), o sistema tambem restringe quem pode EDITAR ou EXCLUIR um registro depois de criado, nos 9 formularios que tem essa acao: Solicitacao de Demanda, Canvas, TAP, Planejamento e Desenvolvimento, EAP, SMP, Ata de Reuniao, TEP e RLA."
Bul "Editar: apenas quem criou o registro, quem tem o papel Gerente de Projetos, ou o Admin."
Bul "Excluir: apenas quem tem o papel Gerente de Projetos, ou o Admin."
Nota "O papel Gerente de Projetos e um atributo global da conta (Administracao > Usuarios, secao 6.1), nao um vinculo especifico a um projeto - qualquer conta com esse papel pode editar/excluir registros de qualquer projeto, nao so dos que participa."
Bul "Os botoes `Editar dados` e `Excluir` ficam ocultos para quem nao tem permissao, em vez de aparecerem e barrarem a acao so no clique."
Nota "Nao se aplica ao Plano de Comunicacao de Projeto nem aos Relatorios de Situacao/Entregas - sao documentos unicos por projeto, sem o conceito de `registro individual criado por alguem`."

H2 "2.16 Quem pode imprimir ou exportar um registro"
P "Nos 7 formularios com restricao por equipe (secao 2.5 - Canvas, TAP, Planejamento e Desenvolvimento, EAP, SMP, TEP e RLA), os botoes Imprimir e Exportar Word de um registro ja salvo tambem passam a exigir fazer parte da equipe daquele projeto, ou ser Admin."
Bul "Diferente da regra de editar/excluir (secao 2.15), aqui nao ha excecao para o Gerente de Projetos - precisa realmente estar na equipe do projeto (ou ser Admin) para imprimir ou exportar."
Bul "Os botoes `Imprimir` e `Exportar Word` ficam ocultos para quem nao e da equipe, da mesma forma que os botoes de editar/excluir."
Nota "Nao se aplica aos outros 5 formularios (Solicitacao de Demanda, Ata de Reuniao, Plano de Comunicacao, Relatorio de Situacao e Relatorio de Entregas) nem aos relatorios agregados - continuam sem restricao de impressao/exportacao por equipe."

# ============================================================
H1 "3. Passo a passo do fluxo de um projeto"
P "Esta secao percorre os 12 formularios de registro na ordem em que normalmente sao usados ao longo da vida de um projeto. Os Passos 1 a 5 e 8 a 9 seguem a esteira sequencial do projeto; os Passos 6, 7 e 10 sao de uso recorrente ou condicional; os relatorios finais (Passos 11 e 12) sao detalhados na secao 4."

H2 "Passo 1 - Solicitacao de Demanda (SGP-01)"
P "Quando usar: no inicio de tudo, para registrar formalmente uma ideia, necessidade ou problema institucional antes de qualquer outra acao. Pode ser preenchida com login normal ou, se a opcao estiver ativa, sem login (ver secao 2.6)."
$r1 = @(
  @("Campo","Obrig.","Descricao"),
  @("Nome do projeto","Sim","Nome que vai identificar a iniciativa em todos os registros futuros"),
  @("Solicitante","Sim","Nome de quem esta pedindo o projeto"),
  @("Departamento / unidade","Sim","Lista de areas da instituicao (Reitoria, Diretorias, Coordenacoes, RH, Marketing, TI etc.)"),
  @("Justificativa / necessidade","Sim","Por que o projeto e necessario"),
  @("Objetivo / resultado esperado","Sim","O que deve ser entregue ao final"),
  @("Escopo","Sim","O que esta incluido e o que nao esta"),
  @("Prazo desejado","Sim","Data limite desejada pelo solicitante"),
  @("Prioridade","Sim","Critica, Alta, Media ou Baixa - urgencia da demanda, ajuda o Admin a priorizar a triagem do Gate 1"),
  @("Orcamento estimado","Nao","Valor aproximado, se ja houver"),
  @("Partes interessadas","Sim","Areas, equipes ou pessoas impactadas"),
  @("Arquivos anexados","Nao","Upload real de arquivo (.txt, .docx ou .pdf, ate 15 MB cada)"),
  @("Links de documentos","Nao","Links de apoio (Drive, SharePoint etc.) que nao sao um arquivo enviado")
)
TableSimple $r1 @(5.0,1.8,9.2)
Img "12_f01_novo.png" "Tela de novo registro da Solicitacao de Demanda." 5.6
Img "13_f01_lista.png" "Demandas cadastradas, com protocolo, projeto, prazo e status." 5.6
Exemplo "`Automatizacao de servicos` - solicitante Hudson Lucas Aleixo, unidade Relacionamento, justificativa: reduzir o tempo de espera dos alunos e desafogar o atendimento presencial/manual da secretaria, oferecendo disponibilidade 24/7 para solicitacoes basicas."
P "Como salvar: clique em `Registrar solicitacao`. O sistema gera o protocolo e leva voce para `Demandas cadastradas`. Clique em qualquer linha da tabela para abrir o registro e revisar os dados. Ao salvar, todos os Admins recebem um aviso automatico (e-mail e, se ativado, push) de que ha um Gate 1 pendente de analise (secao 2.8)."
Nota "E neste ponto que ocorre o Gate 1 - Triagem: so um Admin pode mudar o status para Aprovada ou Reprovada (secao 2.4). So avance para o Passo 2 depois que o status estiver Aprovada. Essa mudanca de status dispara um e-mail de notificacao (secao 2.8)."
Bul "No rodape do painel de detalhes de uma demanda, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do SGP-01, pronto para impressao ou para abrir no Word - incluindo, ao final, a secao 10 - Aprovacao, com linhas para assinatura do Gestor Responsavel e do Responsavel, e a data."
Bul "Assistente de preenchimento por IA: o painel `Importar documento preenchido`, no topo do formulario, extrai os dados de uma copia do formulario oficial ja preenchida em Word ou PDF fora do sistema (secao 2.13)."
Bul "Anexos: o botao `Anexar arquivo`, na secao Anexos, envia o arquivo direto para um repositorio proprio do sistema (ate 15 MB, .txt/.docx/.pdf) - cada arquivo enviado aparece numa lista com opcoes `Baixar` e `Remover`, tanto no formulario quanto no painel de detalhes do registro. O campo de texto Links de documentos continua disponivel a parte, para links que nao sao um arquivo enviado."

H2 "Passo 2 - Canvas de Projeto (SGP-02)"
P "Quando usar: assim que a demanda for aprovada. O Canvas e o primeiro documento estruturado do projeto - reune, em uma unica tela, a motivacao, o produto, os parceiros, as entregas, os riscos e os custos. Exige selecionar um `Projeto vinculado` ja Aprovado no Gate 1, e so membros da equipe daquele projeto (ou Admin) podem registrar (secao 2.5)."
$r2 = @(
  @("Campo","Obrig.","Descricao"),
  @("Nome / Unidade / Gerente / Data","Sim","Identificacao do projeto"),
  @("Justificativas","Sim","Por que estamos aqui - dores e problemas nao atendidos"),
  @("Objetivo","Sim","Objetivo SMART do projeto"),
  @("Beneficios","Nao","O que a organizacao ganha depois da implantacao"),
  @("Produto","Sim","O que sera construido ou entregue"),
  @("Parceiros","Nao","Pessoas ou entidades que apoiam, sem se subordinar ao gerente"),
  @("Grupo de entregas","Sim","Componentes do produto, que depois viram a EAP"),
  @("Restricoes","Nao","Limitacoes que reduzem a liberdade de decisao da equipe"),
  @("Riscos","Sim","Eventos incertos que podem afetar os objetivos"),
  @("Custos","Sim","Estimativa de gastos por grupo de entrega")
)
TableSimple $r2 @(5.0,1.8,9.2)
Img "15_f02_novo.png" "Tela de novo Canvas, com o seletor de Projeto vinculado no topo." 5.6
Img "16_f02_lista.png" "Canvas cadastrados." 5.6
P "Como salvar: clique em `Registrar canvas`. O status inicial e Pendente de aprovacao; altere para Aprovado ao validar o Canvas com o Dono do Negocio e o Gerente de Projetos. Ao marcar Aprovado, um e-mail de notificacao e disparado automaticamente (secao 2.8)."
Bul "No rodape do painel de detalhes de um canvas, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do SGP-02, pronto para impressao ou para abrir no Word."
Bul "Assistente de preenchimento por IA: apos selecionar o Projeto vinculado, o botao `Sugerir com IA` le a Solicitacao de Demanda e as Atas de Reuniao do projeto e sugere o preenchimento dos campos ainda vazios (secao 2.12)."

H2 "Passo 3 - TAP, Termo de Abertura de Projeto (SGP-03)"
P "Quando usar: depois do Canvas aprovado, para autorizar formalmente a existencia do projeto. Tambem exige `Projeto vinculado` e segue a restricao por equipe (secao 2.5) - a figura da secao 2.5 mostra este formulario com um projeto ja selecionado."
$r3 = @(
  @("Campo","Obrig.","Descricao"),
  @("Identificacao (nome, unidade, gerente, data)","Sim","Cabecalho do TAP"),
  @("Alinhamento estrategico / Programa vinculado","Nao","Relacao com outras acoes da Holding, se houver"),
  @("Justificativa","Sim","Problema ou oportunidade que motiva o projeto"),
  @("Objetivos","Sim","O que o projeto entrega"),
  @("Publico-alvo / Beneficios / Exclusoes","Nao","Detalhamento do escopo - inclui o que fica de fora"),
  @("Premissas / Restricoes / Criterios","Nao","Condicionantes do projeto"),
  @("Registro de riscos (tabela)","Nao","Registro vivo: risco, probabilidade, impacto, severidade calculada, mitigacao, status, responsavel e data da ultima revisao"),
  @("Cronograma de entregas macro","Nao","So leitura - editado no Planejamento (Passo 4); mostrado aqui ao vivo, marco por marco, com numero, datas, conclusao real e dependencia"),
  @("Custo total estimado (tabela)","Nao","Soma automatica a medida que as linhas sao preenchidas"),
  @("Partes interessadas e Equipe (tabelas)","Nao","Nome e unidade de cada pessoa"),
  @("Indicadores de resultado (tabela)","Nao","Valor inicial e valor final esperado, com datas")
)
TableSimple $r3 @(5.6,1.6,8.8)
P "Registro de riscos (secao 9 do formulario): diferente dos demais campos do TAP, o registro de riscos e vivo - pode ser reaberto e atualizado a qualquer momento do projeto pelo botao `Editar dados`, sem se limitar ao preenchimento inicial. Cada linha da tabela tem Risco, Probabilidade (Baixa, Media ou Alta), Impacto (Baixo, Medio ou Alto), Mitigacao (texto livre com a resposta planejada para reduzir a chance ou o efeito do risco), Status (Aberto, Monitorando, Mitigado, Materializado ou Encerrado, exibido como selo colorido), Responsavel e Ultima revisao (data). Probabilidade e Impacto tambem sao exibidos como selo colorido (verde/laranja/vermelho por severidade) na tela de detalhes."
P "Severidade e matriz de risco: a partir de Probabilidade e Impacto, o sistema calcula automaticamente a Severidade do risco (Baixo/Medio/Alto/Critico) e mostra como selo colorido, ao vivo, direto na tabela - nao e um campo preenchido manualmente. O botao `Matriz de risco`, acima da tabela, alterna para um heatmap 3x3 (Probabilidade x Impacto) com a contagem de riscos em cada combinacao, considerando so os riscos em aberto (Status diferente de Encerrado) - a mesma leitura de exposicao atual do projeto que orienta onde focar a mitigacao."
P "Responsavel (conta) em cada risco: alem do campo de texto livre Responsavel, cada risco tem um segundo campo opcional, `Responsavel (conta)` - um botao que abre um painel com uma caixa de selecao para cada conta da equipe do projeto selecionado, permitindo marcar uma ou mais pessoas ao mesmo tempo (mesmo padrao ja usado no cronograma do Planejamento, secao 6.4.4, Meu Painel). A lista mostra o nome cadastrado de cada conta, caindo para o e-mail so quando a conta nao tem nome cadastrado, e o botao resume a escolha (nome de quem foi marcado ou, com mais de uma pessoa, o primeiro nome seguido de `+N`). Quando preenchido, o risco passa a aparecer no Meu Painel (secao 6.4.4) de cada pessoa marcada, com um botao para encerra-lo (muda o Status para Encerrado) direto de la."
P "Sugestoes de riscos recorrentes: ao selecionar um Projeto vinculado, o formulario analisa automaticamente o campo Entraves de todas as Atas de Reuniao ja registradas para aquele projeto e destaca, logo abaixo da tabela de riscos, qualquer entrave que se repita em 2 ou mais atas diferentes - com a contagem de atas e um botao `+ Adicionar ao registro` que insere a sugestao direto na tabela (status Aberto, revisao com a data de hoje). E uma regra automatica de comparacao de texto, sem chamar nenhuma IA - ela so aproveita entraves ja registrados nas atas, inclusive os que tiverem sido preenchidos via `Importar transcricao da reuniao` (Passo 6, secao 2.10)."
Img "37_f03_riscos.png" "Registro de riscos preenchido, com os campos Probabilidade, Impacto e Mitigacao, alem de Status, Responsavel e Ultima revisao." 5.6
Img "18_f03_lista.png" "TAPs cadastrados." 5.6
P "Cronograma de entregas macro (secao 10 do formulario): e editado no Planejamento (Passo 4) - aqui o TAP mostra, so leitura, o mesmo cronograma ao vivo: numero do marco (Nº), responsavel, Inicio/Termino previstos, Conclusao real, custo e dependencia, com um botao `Linha do tempo` para ver os mesmos marcos em barras (ver detalhe no Passo 4). Para alterar um marco, use o Planejamento. Alem de Inicio e Termino previstos, cada marco tem um campo `Conclusao real`, preenchido quando o marco de fato termina - e o que permite ao sistema distinguir um marco `atrasado agora` de um marco que `terminou atrasado`: um marco com Conclusao real preenchida nunca conta como atrasado no calculo de Saude do projeto (secao 6.4.1) nem no indicador Marcos no prazo (secao 6.7), mesmo que o Termino previsto ja tenha passado."
P "Dependencia entre marcos: cada marco tambem tem um campo opcional `Depende nº`, onde se digita o numero de outro marco do mesmo cronograma (por exemplo, o marco 2 pode depender do marco 1). Se o marco indicado ainda nao tiver Conclusao real e o Termino previsto ja tiver passado, o marco dependente ganha um aviso `bloqueado` - visivel no proprio Planejamento (edicao, detalhes e impressao/exportacao), no espelho do TAP e nos Marcos do Relatorio de Situacao. E uma verificacao simples, de um nivel so e por numero - nao calcula um caminho critico completo nem cadeias de varios marcos."
P "Como salvar: clique em `Registrar TAP`. Um projeto sem TAP aprovado nao deve avancar para a execucao. Ao marcar Aprovado, um e-mail de notificacao e disparado automaticamente (secao 2.8)."
Bul "No rodape do painel de detalhes de um TAP, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do SGP-03 (incluindo o registro de riscos, o cronograma vindo ao vivo do Planejamento e as tabelas de custos, partes interessadas, equipe e indicadores), pronto para impressao ou para abrir no Word."
Bul "Assistente de preenchimento por IA: apos selecionar o Projeto vinculado, o botao `Sugerir com IA` le a Solicitacao de Demanda, o Canvas e as Atas de Reuniao do projeto e sugere o preenchimento dos campos de texto ainda vazios, alem de rascunhos iniciais para as tabelas de riscos, custos e partes interessadas (secao 2.12) - diferente da sugestao de riscos recorrentes acima, que e uma regra automatica sem IA."

H2 "Passo 4 - Planejamento e Desenvolvimento de Projeto (SGP-04)"
P "Quando usar: logo apos o TAP, para detalhar o projeto em profundidade. E o dossie mais extenso da ferramenta, organizado em 8 abas internas, navegaveis pela barra de estagios no topo."
$r4 = @(
  @("Aba","Conteudo"),
  @("1. Pre-projeto","Produtos impactados, escopo, contexto, objetivos, entregaveis, premissas, restricoes, partes interessadas, macro fases"),
  @("2. Viabilidade","Situacao atual, proposta de mudanca, beneficios, alternativas, impactos, plano de implementacao, conclusao"),
  @("3. Cronograma","Marcos e prazos do projeto - editado aqui, com numeracao automatica (Nº) e dependencia entre marcos; aparece so leitura no TAP e no Relatorio de Situacao"),
  @("4. Entradas","Tabela de requisitos de entrada, com 7 categorias pre-preenchidas, para completar com descricao, fonte, responsavel e status"),
  @("5. Saidas","Tabela de entregaveis ligados a um requisito de entrada, com criterios de aceitacao"),
  @("6. Verificacao","Tabela de verificacao de conformidade de cada saida"),
  @("7. Validacao","Tabela de validacao do uso pretendido, com metodo e resultado"),
  @("8. Analise critica","Tabela de pontos de controle por etapa do projeto")
)
TableSimple $r4 @(3.0,13.0)
Img "19_f04_novo.png" "Capa de identificacao do dossie, com a barra das 8 abas no rodape visivel." 5.6
Img "20_f04_lista.png" "Projetos cadastrados (estado vazio, antes do primeiro dossie ser salvo)." 5.6
P "Como salvar: preencha a capa (nome do projeto, unidade e gestor sao obrigatorios) e navegue pelas 8 abas. Clique em `Registrar dossie`. E este documento que alimenta, junto com o Canvas, a elaboracao do Relatorio de Entregas e Beneficios (Passo 12)."
Bul "No rodape do painel de detalhes de um dossie, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do SGP-04, com as 8 abas em sequencia, pronto para impressao ou para abrir no Word."
P "Cronograma de entregas macro (aba 3): cada marco recebe automaticamente um numero sequencial (Nº) assim que o campo Marco/Entrega e preenchido - renumerado ao vivo ao reordenar, adicionar ou remover linhas. Alem de Inicio e Termino previstos, cada marco tem um campo `Conclusao real`. E este cronograma que aparece, so leitura, no TAP (Passo 3) e no Relatorio de Situacao (secao 4.1) do mesmo projeto, e que alimenta o calculo de Saude do projeto (secao 6.4.1) e o indicador Marcos no prazo (secao 6.7)."
P "Dependencia entre marcos: cada marco tambem tem um campo opcional `Depende nº`, onde se digita o numero de outro marco do mesmo cronograma. Se o marco indicado ainda nao tiver Conclusao real e o Termino previsto ja tiver passado, o marco dependente ganha um aviso `bloqueado` - visivel no proprio Planejamento, no espelho do TAP e nos Marcos do Relatorio de Situacao."
P "Linha do tempo (Gantt): acima da tabela de marcos, o botao `Linha do tempo` alterna para uma visualizacao em barras, uma por marco, com o mesmo periodo de Inicio/Termino previstos - azul para marcos no prazo, verde para concluidos, vermelho para atrasados, e um contorno vermelho extra para marcos bloqueados por dependencia. Uma linha vertical marca o dia de hoje. E somente para visualizacao - a edicao continua sempre pela aba Tabela. Marcos sem nenhuma data preenchida nao aparecem na linha do tempo, mas continuam na tabela normalmente. A mesma visualizacao (com os mesmos dados, ao vivo) tambem esta disponivel no espelho do TAP (secao 10) e na secao Marcos do Relatorio de Situacao (secao 4.1)."
P "Caminho critico: quando pelo menos dois marcos estao encadeados por `Depende nº`, a linha do tempo destaca esse caminho com um contorno tracejado, o icone de alvo e setas ligando um marco ao proximo - a sequencia de marcos que, se atrasar, atrasa a data final do cronograma. E calculado a partir do marco com o Termino previsto mais tardio, andando para tras pela cadeia de dependencias; nao e um calculo completo de folga (CPM) com multiplos caminhos, so a cadeia que hoje determina o fim do projeto."
P "Carga por responsavel: o botao `Carga por responsavel`, ao lado da linha do tempo, agrupa os marcos pelos nomes digitados no campo Responsavel - se o campo tiver mais de um nome ou area separados por barra, virgula ou ponto-e-virgula (ex.: `Direcao/ Gerencia`), cada um vira uma linha propria. Cada marco atribuido vira uma barra semi-transparente no periodo do Inicio ao Termino previstos; barras que se sobrepoem ficam mais escuras, e um responsavel com marcos de datas sobrepostas ganha o aviso `possivel sobrecarga`. E uma leitura simples do texto ja digitado - nao soma horas nem capacidade de trabalho."
P "Execucao (autorrelato): cada marco tem uma coluna `Execucao` - um seletor com `A iniciar / Em andamento / Atrasado / Concluido`, com a data da ultima mudanca gravada automaticamente (nao editada a mao). E diferente do que o sistema ja calculava sozinho a partir das datas (atrasado ou nao): e a propria pessoa sinalizando se esta de fato trabalhando no marco - o unico jeito de saber isso entre o Inicio previsto e o vencimento, ja que antes so dava pra descobrir que nada andou quando o prazo ja tinha passado. Aparece tambem no Meu Painel (secao 6.4.4, onde o responsavel normalmente atualiza, sem precisar abrir o Planejamento inteiro) e no Painel de Prazos (secao 6.4.2), ao lado de cada marco atrasado ou vencendo. Escolher `Concluido` tambem preenche a Conclusao real com a data de hoje, para os dois campos nao ficarem discordando."
P "Responsavel (conta): alem do campo de texto livre Responsavel, cada marco tem um segundo campo opcional, `Responsavel (conta)` - um botao que abre um painel com uma caixa de selecao para cada conta da equipe do projeto selecionado (secao 2.5), permitindo marcar uma ou mais pessoas para o mesmo marco. A lista mostra o nome cadastrado de cada conta (caindo para o e-mail quando nao houver nome), e o botao resume a escolha (nome, ou `nome +N` com mais de uma pessoa marcada). Diferente do texto livre, que e so uma anotacao, esse campo vincula o marco a contas de verdade: e a partir dele (mais o papel Gerente de Projetos da equipe do projeto) que o sistema sabe para quem mandar o aviso automatico de prazo (secao 6.4.3), para popular o Painel de Prazos (secao 6.4.2) e para popular o Meu Painel (secao 6.4.4) de cada conta marcada. Preencher esse campo e opcional - um marco sem `Responsavel (conta)` continua funcionando normalmente, so nao gera aviso individual nem aparece no Meu Painel de ninguem."
P "Responsavel (conta) nas Entradas (aba 4): no mesmo padrao do cronograma acima, cada linha da aba Entradas tem um botao `Responsavel (conta)` opcional que abre um painel com uma caixa de selecao por conta da equipe do projeto, permitindo marcar mais de uma pessoa. Quando preenchido, a entrada passa a aparecer no Meu Painel (secao 6.4.4) de cada pessoa marcada, com um botao para conclui-la (muda o Status para OK) direto de la."
P "Anexos na Viabilidade (aba 2): o botao `Anexar arquivo` envia o arquivo direto para um repositorio proprio do sistema (ate 15 MB, .txt/.docx/.pdf) - mesmo recurso ja usado na Solicitacao de Demanda (Passo 1) e na SMP (Passo 7). Cada arquivo enviado aparece numa lista com opcoes `Baixar` e `Remover`, tanto no formulario quanto no painel de detalhes do dossie; na exportacao para Word, os arquivos aparecem listados por nome (o conteudo nao e embutido no documento gerado). O campo `Links de documentos`, logo abaixo, continua disponivel a parte, para links (Drive, SharePoint etc.) que nao sao um arquivo enviado."
P "Tarefas dentro de um marco: quando um marco precisar de mais detalhe, o botao `+ Tarefas` (ultima coluna da tabela) abre uma lista simples logo abaixo da linha - nome da tarefa e uma caixa de concluida, sem data ou responsavel proprios (quem tem isso e o marco). O botao muda para `Tarefas (2/3)`, por exemplo, mostrando quantas ja foram concluidas, e fica com contorno laranja enquanto houver alguma pendente. E totalmente opcional - um marco sem nenhuma tarefa continua funcionando normalmente. Ao lado do botao `+ Adicionar linha`, um resumo mostra o percentual de tarefas concluidas somando todos os marcos do cronograma."
Bul "Assistente de preenchimento por IA: apos selecionar o Projeto vinculado, o botao `Sugerir com IA` le a Solicitacao de Demanda, o TAP, o Canvas e as Atas de Reuniao do projeto e sugere o Solicitante (capa), o preenchimento da aba Pre-projeto (incluindo Macro fases, resumidas dos entregaveis/escopo do TAP), um rascunho inicial do cronograma (aba 3, quando ainda vazio) e a Introducao/Situacao atual/Proposta de mudanca/Beneficios na aba Viabilidade. O campo Negocio (capa) fica de fora, por nao ter fonte confiavel nos documentos (secao 2.12)."

H2 "Passo 5 - EAP, Estrutura Analitica do Projeto"
P "Quando usar: em paralelo ao planejamento, para decompor visualmente o escopo em pacotes de trabalho."
Img "21_f05_novo.png" "Construtor hierarquico de 3 niveis: Pacote de trabalho, Entrega e Atividade." 5.6
Img "22_f05_lista.png" "EAPs cadastradas (estado vazio)." 5.6
Bul "Informe o nome do projeto (obrigatorio), a unidade e o gerente."
Bul "Use o construtor hierarquico de 3 niveis: `+ Adicionar pacote de trabalho`, depois `+ Adicionar entrega` dentro do pacote, depois `+ Atividade` dentro da entrega."
Bul "Clique em `Visualizar arvore` a qualquer momento para conferir o diagrama antes de salvar."
Bul "Cada Atividade (nivel 3) pode receber, de forma opcional, uma prioridade: `P0 - Critica`, `P1 - Alta` ou `P2 - Normal` (ou nenhuma), exibida como selo colorido no construtor, na `Visualizar arvore` e no painel de detalhes. Serve para repriorizar o trabalho no dia a dia sem precisar abrir uma SMP - Solicitacao de Mudanca de Projeto (Passo 7). EAPs que ja existiam sem prioridade definida continuam funcionando normalmente."
Img "39_f05_prioridade.png" "Atividades de nivel 3 com os tres niveis de prioridade (P0, P1 e P2), exibidos como selo colorido ao lado de cada atividade." 5.6
Bul "Status automatico da entrega: ao clicar em `Visualizar arvore` (ou ao abrir uma EAP ja registrada), cada Entrega (nivel 2) ganha um selo - `Concluida`, `Atrasada`, `No prazo` ou `Sem cronograma vinculado` - calculado automaticamente a partir do cronograma do Planejamento (Passo 4), casando pelo nome da entrega com o nome do marco. Nao ha nenhum campo novo para preencher na EAP: uma entrega fica `Concluida` se o marco correspondente tiver Conclusao real preenchida no cronograma, `Atrasada` se nao tiver conclusao e o Termino previsto ja tiver passado, e `Sem cronograma vinculado` se nao houver marco com o mesmo nome. Um resumo no topo da arvore mostra o percentual de entregas concluidas do projeto."
Bul "Status automatico da atividade: o mesmo casamento por nome desce mais um nivel - cada Atividade (nivel 3) ganha um selo `Concluida`, `Pendente` ou `Sem tarefa vinculada`, casando pelo nome da atividade com uma tarefa cadastrada dentro de qualquer marco do cronograma (botao `+ Tarefas` de cada marco, Passo 4, secao 3). Como tarefa nao tem data - so nome e concluida - nao existe estado `Atrasada` aqui, so `Concluida`/`Pendente`. Tambem sem nenhum campo novo para preencher na EAP."
P "Clique em `Registrar EAP` para gerar o protocolo."
Bul "No rodape do painel de detalhes de uma EAP, os botoes `Imprimir` e `Exportar Word` geram um documento com a arvore hierarquica completa (pacotes, entregas e atividades, com a prioridade de cada uma), pronto para impressao ou para abrir no Word. O documento sai com o codigo SGP-05 - o layout segue o padrao visual dos demais documentos gerados pelo sistema."
Bul "Importar do Cronograma: apos selecionar o Projeto vinculado, o botao `Importar do Cronograma` cria (ou reaproveita, se ja existir) um pacote de trabalho fixo `Marcos do Cronograma` e adiciona uma Entrega (nivel 2) para cada marco do Cronograma do Planejamento (Passo 4), com o nome copiado de forma identica, sem nenhuma reformulacao - e esse nome identico que permite o casamento automatico de status explicado acima. Clicar de novo so acrescenta marcos que ainda nao viraram entrega, sem duplicar os ja importados nem mexer em pacotes/entregas criados manualmente. Este botao substituiu um antigo assistente por IA (`Sugerir com IA`) que sugeria pacotes e entregas com texto proprio - como o casamento com o Cronograma exige nome exatamente igual, a sugestao de IA (que reformulava o texto) nem sempre garantia esse casamento; a importacao mecanica garante sempre."

H2 "Passo 6 - Ata de Reuniao (SGP-07) - uso recorrente"
P "Quando usar: a qualquer momento do projeto, para registrar formalmente qualquer reuniao - nao faz parte da esteira sequencial, fica sempre disponivel. Pode ser preenchida com login normal ou, se a opcao estiver ativa, sem login (ver secao 2.6)."
Img "25_f07_novo.png" "Tela de nova ata, com o painel `Importar transcricao da reuniao` para preenchimento automatico por IA." 5.6
Img "26_f07_lista.png" "Atas cadastradas (estado vazio)." 5.6
Bul "Selecione a(s) unidade(s) envolvidas, preencha Pauta e Projeto (obrigatorios), participantes, resumo, encaminhamentos e entraves."
Bul "Nas tabelas de Encaminhamentos e Entraves, a coluna SEQ e preenchida automaticamente - numera em sequencia (01, 02...) assim que o Encaminhamento/Entrave da linha e digitado, sem precisar informar o numero manualmente. Uma linha ainda vazia nao entra na contagem; reordenar (`▲`/`▼`) ou remover uma linha renumera as demais automaticamente."
Bul "Opcional, quando habilitado pelo Admin: em `Importar transcricao da reuniao`, cole o texto, anexe um arquivo `.txt`/`.docx`/`.pdf`, ou anexe o audio da gravacao (`.mp3`/`.m4a`/`.aac`/`.wav`/`.ogg`) para transcricao automatica (secao 2.11), depois clique em `Analisar e preencher` - a IA sugere pauta, data/horario (se mencionados), participantes, resumo, encaminhamentos e entraves. Revise sempre os dados sugeridos antes de registrar (secao 2.10)."
P "Clique em `Registrar ata` para gerar o protocolo."
Bul "Na lista `Atas cadastradas`, clique numa ata para abri-la e usar os botoes `Imprimir` e `Exportar Word`, no rodape do painel lateral: ambos geram o documento no formato oficial do SGP-07 (cabecalho, unidade, pauta, participantes, descricao, saidas e entraves), pronto para impressao ou para abrir no Word."

H2 "Passo 7 - SMP, Solicitacao de Mudanca de Projeto (SGP-06) - uso condicional"
P "Quando usar: sempre que for necessario alterar escopo, cronograma, custo ou qualidade de um projeto ja em andamento. Nenhuma mudanca deve ser feita sem passar por este formulario."
Img "23_f06_novo.png" "Tela de nova SMP." 5.6
Img "24_f06_lista.png" "SMPs cadastradas (estado vazio)." 5.6
Bul "Identifique o projeto e a mudanca (titulo e solicitante sao obrigatorios); descreva a mudanca, os beneficios e o impacto de nao implementa-la."
Bul "Preencha a analise de impactos nas 8 dimensoes: objetivo, cronograma, escopo, custo, alinhamento estrategico, qualidade, riscos e outros impactos."
Bul "Marque a decisao: Aprovada, Nao aprovada ou Pendente de avaliacao, com justificativa."
Bul "Assistente de preenchimento por IA: apos selecionar o Projeto vinculado, o botao `Sugerir com IA` procura nas Atas de Reuniao mais recentes alguma mudanca sendo discutida em relacao ao combinado no TAP/Planejamento, e sugere a descricao da mudanca e a analise de impactos - se nao houver mudanca clara sendo discutida, nenhum campo e preenchido (secao 2.12)."
P "Clique em `Registrar SMP` - o status do registro acompanha automaticamente a decisao marcada."
Bul "Anexos: mesmo recurso ja usado na Solicitacao de Demanda (Passo 1) - o botao `Anexar arquivo`, na secao Anexos, envia o arquivo direto para um repositorio proprio do sistema (ate 15 MB, .txt/.docx/.pdf), com opcoes `Baixar` e `Remover` tanto no formulario quanto no painel de detalhes do registro. O campo de texto Links de documentos continua disponivel a parte, para links que nao sao um arquivo enviado."
Img "58_smp_anexos.png" "Secao Anexos da SMP: arquivo enviado, com opcoes de baixar/remover, e o campo Links de documentos." 5.6
Bul "No rodape do painel de detalhes de uma SMP, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do SGP-06 (incluindo a analise das 8 dimensoes de impacto), pronto para impressao ou para abrir no Word."
Nota "Propagacao para o projeto vinculado: sempre que a decisao da SMP e finalizada como `Aprovada` ou `Nao aprovada` (seja ao registrar/editar a SMP ou ao trocar o Status na tela de detalhes), o sistema grava automaticamente um evento no historico compartilhado do projeto vinculado (o mesmo historico acessivel pelo botao `Ver historico` em qualquer formulario com `Projeto vinculado` - secao 2.5). Isso torna as decisoes de mudanca visiveis para quem estiver no Canvas, TAP, Planejamento, EAP, TEP ou RLA daquele projeto, sem precisar abrir a SMP. Essa propagacao apenas registra o evento no historico - ela nao altera o Status do projeto usado no Gate 1 (Aprovado/Reprovado). A mesma decisao (Aprovada ou Nao aprovada) tambem dispara um e-mail de notificacao (secao 2.8)."

H2 "Passo 8 - TEP, Termo de Encerramento de Projeto (SGP-09)"
P "Quando usar: ao encerrar o projeto, seja por conclusao, paralisacao ou cancelamento."
Img "28_f09_novo.png" "Tela de novo TEP." 5.6
Img "29_f09_lista.png" "TEPs cadastrados (estado vazio)." 5.6
Bul "Preencha a identificacao e o programa vinculado, se houver; selecione o tipo de encerramento (Concluido, Paralisado ou Cancelado)."
Bul "Se Paralisado ou Cancelado, o campo Justificativa aparece automaticamente."
Bul "Registre entregas de resultados, atividades encerradas, o link da pasta do projeto e a analise de efetividade."
Bul "Assistente de preenchimento por IA: apos selecionar o Projeto vinculado, o botao `Sugerir com IA` compara o TAP, o Planejamento e a EAP com o que as Atas mais recentes confirmam realizado, sugerindo a justificativa de encerramento, as atividades encerradas e as consideracoes finais (secao 2.12)."
P "Clique em `Registrar TEP` para gerar o protocolo. O registro de um novo TEP dispara um e-mail de notificacao (secao 2.8)."
Nota "Encerra o projeto no cadastro compartilhado: ao registrar o TEP, o projeto vinculado passa a refletir o tipo de encerramento escolhido (Concluido/Paralisado/Cancelado) e some da lista de `Projeto vinculado` para novos registros nos demais formularios (secao 2.5)."
Bul "No rodape do painel de detalhes de um TEP, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do SGP-09, pronto para impressao ou para abrir no Word."

H2 "Passo 9 - RLA, Registro de Licoes Aprendidas (SGP-10)"
P "Quando usar: junto com o TEP, ao final do projeto (ou tambem em pontos intermediarios), para capturar o aprendizado organizacional."
Img "30_f10_novo.png" "Tela de novo RLA." 5.6
Img "31_f10_lista.png" "RLAs cadastrados (estado vazio)." 5.6
Bul "Responda as perguntas abertas dos blocos Visao geral, Destaques, Desafios e Pos-projeto."
Bul "Nos blocos de avaliacao estruturada, marque Sim/Nao/Parcial/N-A para cada afirmacao - cada bloco calcula automaticamente um placar percentual."
Bul "Assistente de preenchimento por IA: apos selecionar o Projeto vinculado, o botao `Sugerir com IA` le o TEP, as Atas e todas as SMPs do projeto para sugerir as respostas das 4 secoes de texto livre - Visao geral, Destaques, Desafios e Tarefas pos-projeto. Os blocos de avaliacao estruturada Sim/Nao/Parcial (secoes 6 a 9, com placar percentual) nao sao preenchidos pela IA - continuam sendo uma autoavaliacao manual da equipe (secao 2.12)."
P "Clique em `Registrar RLA`. O score global fica salvo junto com o registro."
Bul "No rodape do painel de detalhes de um RLA, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do SGP-10 (incluindo os quatro blocos de avaliacao estruturada com o placar de cada um), pronto para impressao ou para abrir no Word."

H2 "Passo 10 - Plano de Comunicacao de Projeto (SGP-08)"
P "Quando usar: para planejar o que sera comunicado, a quem e em qual momento do projeto. Este artefato tem uma ferramenta eletronica propria, organizada como uma tabela de referencia com abas Painel (visualizacao) e Editar dados."
Img "27_f08_painel.png" "Painel do Plano de Comunicacao: cada linha da tabela e um tipo de comunicacao do projeto." 5.6
$r10 = @(
  @("Coluna","Descricao"),
  @("Tipo de Comunicacao","Nome do momento/evento de comunicacao (ex.: Status semanal, Go/No-Go)"),
  @("O que Comunicar","Conteudo tratado nessa comunicacao"),
  @("Quando Comunicar","Periodicidade ou gatilho (ex.: Semanal, Em ate 2h do evento)"),
  @("Com Quem se Comunicar","Publico-alvo da comunicacao"),
  @("Como Comunicar","Canal e formato (ex.: Reuniao + ata, E-mail resumo + dashboard)"),
  @("Quem Comunica","Responsavel por comunicar (ex.: GP, GP / PMO)")
)
TableSimple $r10 @(4.5,11.5)
P "As 15 linhas padrao do plano (Comeco do projeto, Diario de equipe, Status semanal, Report executivo, Comite de Mudancas, Gestao de Riscos, Comunicacao de incidentes, Entregas e marcos, Integracao com stakeholders externos, Comunicacao de mudancas organizacionais, Treinamentos, Homologacao/UAT, Go/No-Go, Pos-Lancamento e Licoes aprendidas & Encerramento) vem preenchidas com o conteudo do documento oficial SGP-08 nas 6 colunas - use `Editar dados` para ajustar aos seus projetos."
Nota "Edicao restrita por papel: diferente dos demais formularios, so PMO/Admin pode usar `Editar dados` e salvar alteracoes no Plano de Comunicacao - controlavel pelo Admin em Administracao > Configuracoes (secao 6.3), mesmo esquema de lista de papeis usado no Painel Executivo e nos dois relatorios. Qualquer usuario autenticado continua podendo abrir a aba Painel, visualizar todas as colunas e usar o botao Imprimir; quem nao tem permissao de edicao ve a aba `Editar dados` escondida e um aviso no lugar dela."

# ============================================================
# 4. RELATORIOS FINAIS
# ============================================================
H1 "4. Gerando os relatorios finais"
P "Esta secao detalha os dois relatorios que fecham o ciclo de uso da ferramenta: o Relatorio de Situacao de Projetos (visao de controle do portfolio) e o Relatorio de Entregas e Beneficios (ficha executiva de programa e projetos, base do Gate 2)."

H2 "4.1 Relatorio de Situacao de Projetos (SGP-11)"
P "Visao consolidada de todos os projetos do portfolio: status, percentual de execucao, marcos e pontos de atencao. Tem tres abas: Painel, Editar dados e Importar Project."
Img "32_f11_painel.png" "Painel do Relatorio de Situacao (estado vazio, antes do primeiro preenchimento)." 5.6
Img "33_f11_editar.png" "Aba Editar dados: cabecalho do relatorio e botao para adicionar projetos/marcos." 5.6
Img "34_f11_importar.png" "Aba Importar Project: arraste uma exportacao do MS Project (Excel ou CSV) para preencher em lote." 5.6
Bul "Preencha o cabecalho (mes/ano de referencia, responsavel, previsao financeira, entregas planejadas)."
Bul "Adicione cada projeto do portfolio, com status, % execucao, datas e os campos `Merece atencao` e `Merece destaque`."
Nota "Projeto vinculado (Gate 1) - opcional: cada linha de projeto pode, opcionalmente, ser associada a um Projeto ja Aprovado no Gate 1 (o mesmo cadastro compartilhado usado pelo Canvas, TAP, Planejamento, EAP, SMP, TEP e RLA - secao 2.5). Ao selecionar um projeto na lista, os campos Nome do projeto, Responsavel (a partir do gestor do projeto) e Inicio previsto/Termino previsto (a partir da menor e maior data do cronograma de entregas macro do Planejamento vinculado, se houver) sao preenchidos automaticamente - sempre so nos campos ainda vazios, sem sobrescrever o que ja foi digitado. Tambem aparece o botao `Ver historico`, que mostra as alteracoes desse projeto vindas de qualquer outro formulario vinculado a ele. Deixar em `Nenhum` mantem o comportamento anterior - o projeto do Relatorio de Situacao continua sendo texto livre, sem vinculo. Se o Planejamento nao tiver nenhuma data de cronograma preenchida, Inicio previsto e Termino previsto ficam em branco para preenchimento manual - o sistema nao tenta adivinhar a partir de outro documento."
Bul "Assistente de preenchimento por IA: com um Projeto vinculado selecionado, o botao `Sugerir com IA` le o TAP e as Atas daquele projeto, mais a situacao ja digitada na linha, para sugerir o que Merece atencao e o que Merece destaque (secao 2.12)."
Bul "Marcos vem do Planejamento: a secao Marcos de cada linha de projeto mostra, automaticamente (sem botao de importar), os marcos do cronograma de entregas macro do Planejamento e Desenvolvimento de Projeto daquele projeto - nome, Inicio/Termino previsto e Conclusao real vem sempre ao vivo do Planejamento (para mudar um marco, edite o Planejamento). O unico campo proprio aqui e o Status de cada marco (A iniciar, No prazo, Atrasado, Paralisado, Concluido ou Cancelado), que continua sendo ajustado manualmente. Um marco com dependencia bloqueada (`bloqueado`) mostra o mesmo aviso aqui. O botao `Linha do tempo`, acima da lista, mostra os mesmos marcos em barras (ver detalhe no Passo 4)."
Img "42_f11_projeto_vinculado.png" "Projeto vinculado selecionado, com o nome preenchido automaticamente e o historico compartilhado exibido." 5.6
Bul "Alternativa mais rapida: use `Importar Project` - o sistema detecta automaticamente as colunas de nome, % concluido, inicio, termino e responsavel."
Bul "Clique em `Salvar e ver painel`. No rodape, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do SGP-11 com o retrato completo do portfolio (visao geral dos projetos, entraves e encaminhamentos, resultados alcancados e o resumo de atencao/destaque por programa), pronto para impressao ou para abrir no Word. Se uma versao historica estiver sendo visualizada (Nota abaixo), os dois botoes exportam essa versao, nao a atual."
Nota "Sinalizacao automatica: a tabela de projetos do Painel ganhou a coluna `Sinalizacao` e o quadro de indicadores ganhou o card `Em risco de atraso`. E um calculo automatico, independente do campo Status manual: compara o % Execucao informado com o % que seria esperado pelo tempo ja decorrido entre o Inicio previsto e o Termino previsto do projeto. Se a defasagem for de 15 pontos percentuais ou mais, aparece `Risco de atraso`; se o Termino previsto ja passou e o projeto nao esta em 100%, aparece `Prazo vencido`. Projetos Concluido, Cancelado ou Paralisado ficam fora desse calculo. O resultado tambem sai no documento exportado."
Img "38_f11_sinalizacao.png" "Painel com o card `Em risco de atraso` e a coluna Sinalizacao mostrando os selos Risco de atraso e Prazo vencido." 5.6
Bul "Custos realizados: logo abaixo de Marcos, cada projeto tem sua propria lista de categorias de custo com o valor efetivamente gasto ate o momento, base do indicador `Projetos dentro do orcamento` do novo Relatorio de Resultados (secao 6.7)."
Nota "Importar categorias do TAP: o botao `Importar categorias do TAP` (visivel com um Projeto vinculado selecionado) traz as mesmas categorias de custo estimado cadastradas no TAP daquele projeto - so falta preencher o valor efetivamente gasto em cada categoria, e categorias repetidas nao sao duplicadas. O total gasto (soma de todas as categorias) e recalculado automaticamente a cada valor digitado. Diferente dos Marcos acima (que agora vem sempre ao vivo do TAP, sem botao), os Custos realizados continuam sendo uma copia unica - o botao pode ser clicado de novo se o TAP ganhar novas categorias depois."
Nota "Comparativo por categoria: o botao `Comparativo por categoria`, ao lado da lista de Custos realizados, mostra uma tabela cruzando, categoria por categoria, o valor Estimado no TAP com o Realizado aqui - casando pelo nome da categoria (o mesmo nome que o botao Importar categorias do TAP ja usa). Cada linha ganha um selo: `Dentro do estimado`, `Acima do estimado`, ou `Sem estimativa no TAP` quando uma categoria realizada nao tem contrapartida no TAP. E somente para visualizacao - a edicao continua sempre pela lista."
Img "50_f11_custos_realizados.png" "Secao Custos realizados de um projeto, apos importar as categorias do TAP e preencher os valores gastos." 5.6
Bul "Na aba Editar dados, alem dos projetos, ha dois cartoes independentes: `Entraves` e `Encaminhamentos`. Cada item tem Descricao, Responsavel, Prazo, Status (Aberto / Resolvido / Cancelado) e mais dois campos opcionais, `Por que?` e `Resultado esperado`."
Nota "Responsavel (conta): diferente do mesmo campo no cronograma do Planejamento e no Registro de riscos do TAP (que listam so a equipe do projeto), aqui o botao `Responsavel (conta)` abre um painel com uma caixa de selecao para cada conta do sistema - Entraves e Encaminhamentos nao tem um projeto vinculado unico, entao nao ha uma equipe para restringir a lista - permitindo marcar uma ou mais pessoas. A lista mostra o nome cadastrado de cada conta (caindo para o e-mail quando nao houver nome), e o botao resume a escolha (nome, ou `nome +N` com mais de uma pessoa marcada). Quando preenchido, o item passa a aparecer no Meu Painel (secao 6.4.4) de cada conta marcada, com um botao `Resolver` que muda o Status para Resolvido direto de la. E opcional."
Nota "Gestao de Entraves e Encaminhamentos: entraves sao bloqueios que impedem o avanco do portfolio (ex.: dependencia externa, aprovacao pendente); encaminhamentos sao os proximos passos e decisoes acordadas para desbloquear ou avancar o trabalho. Os dois sao listas simples, independentes dos projetos cadastrados - continuam visiveis no Painel mesmo que nenhum projeto tenha sido informado ainda. No Painel, aparecem lado a lado na secao `Entraves e encaminhamentos`, com o Status de cada item destacado por selo colorido, e tambem saem no documento exportado. Os campos `Por que?` (causa do entrave/encaminhamento) e `Resultado esperado` sao opcionais e nao aparecem no documento exportado deste relatorio - servem de insumo para o Plano de Acao do Relatorio de Resultados (secao 6.7), preenchidos so quando ha algo relevante a registrar."
Img "41_f11_entraves_encaminhamentos.png" "Secao Entraves e encaminhamentos no Painel, com itens Aberto e Resolvido lado a lado." 5.6
Bul "Ainda em Editar dados, o cartao `Resultados alcancados` registra o desempenho real do portfolio: Descricao, Tipo (Indicador / Entrega / Beneficio), Planejado e Alcancado."
Nota "Reporte de Resultados (D06.5): este cartao e a base do momento em que o Gerente de Projetos e o Dono do Negocio apresentam as instancias de decisao no que o projeto foi alcancado - nao so o que foi planejado, mas indicadores, entregas e beneficios de fato realizados, lado a lado com a meta original. E uma lista simples, independente dos projetos cadastrados, com o mesmo comportamento de Entraves e Encaminhamentos: continua visivel no Painel mesmo sem nenhum projeto informado, e sai no documento exportado. Junto com os Entraves e Encaminhamentos (D06.6/D06.7), completa os tres artefatos de governanca da Diretriz D06 sem exigir um 13o formulario separado - o Reporte de Resultados e atendido dentro do proprio Relatorio de Situacao."
Img "47_f11_resultados_alcancados.png" "Secao Resultados alcancados no Painel, com selos coloridos por Indicador, Entrega e Beneficio." 5.6
Bul "No rodape do Painel, o botao `Historico de versoes` mostra todas as vezes que este relatorio foi salvo, com data/hora e quem salvou."
Nota "Versionamento: toda vez que o relatorio e salvo (`Salvar e ver painel`), o sistema arquiva uma copia completa dos dados daquele momento, alem de atualizar a versao atual. Isso preserva as referencias mensais anteriores, que antes eram sobrescritas a cada novo salvamento. Ao clicar em `Visualizar` numa versao antiga, o Painel passa a mostrar aquele instantaneo, com um aviso amarelo no topo (`Visualizando versao... - somente leitura`) e um botao `Voltar a versao atual`. Enquanto uma versao antiga esta sendo visualizada, as abas `Editar dados` e `Importar Project` ficam bloqueadas - e preciso voltar a versao atual antes de editar."
Img "46_f11_versao_historica_banner.png" "Aviso de versao historica (somente leitura), com o botao para voltar a versao atual." 5.6
Nota "Este relatorio nao tem aprovacao/status formal - e uma ferramenta viva de acompanhamento, atualizada sempre que a situacao dos projetos mudar."

H2 "4.2 Relatorio de Entregas e Beneficios (SGP-12)"
P "O relatorio mais importante do ponto de vista de governanca: e sobre ele que ocorre o Gate 2 - Pactuacao, quando o Dono do Negocio assume formalmente, perante a Alta Gestao, o compromisso de que todo o trabalho planejado sera executado."
Img "35_f12_painel.png" "Painel do Relatorio de Entregas e Beneficios (estado vazio)." 5.6
Img "36_f12_editar.png" "Aba Editar dados: Ficha do Programa, o primeiro bloco a preencher." 5.6
Bul "Preencha a Ficha do Programa (codigo, nome, unidade, responsavel, justificativa, objetivo, alinhamento estrategico)."
Bul "Cadastre indicadores e valores estimados do programa, depois cada projeto vinculado (com suas proprias entregas, indicadores e valores)."
Nota "Projeto vinculado (Gate 1) - opcional: assim como no Relatorio de Situacao (secao 4.1), cada projeto cadastrado aqui pode ser associado a um Projeto ja Aprovado no Gate 1. A selecao preenche automaticamente o Nome do projeto, o Responsavel pelo projeto (a partir do gestor do projeto) e, a partir do TAP vinculado, o Objetivo, o Publico-alvo e um resumo dos Riscos em aberto (status Aberto ou Monitorando) - sempre so nos campos ainda vazios. Tambem libera o botao `Ver historico`, mostrando as alteracoes feitas nesse projeto em qualquer outro formulario vinculado a ele - incluindo decisoes de SMP (Passo 7)."
Bul "Assistente de preenchimento por IA: com um Projeto vinculado selecionado, o botao `Sugerir com IA` le o TAP e as Atas daquele projeto para sugerir o Impacto/Resultado do projeto - sempre redigido como uma expectativa, ja que o Gate 2 ocorre antes da execucao comecar (secao 2.12)."
Img "43_f12_projeto_vinculado.png" "Projeto vinculado selecionado no Relatorio de Entregas, com nome e historico preenchidos." 5.6
Nota "Importar indicadores do TAP: quando um Projeto vinculado e selecionado, a secao `Indicadores do projeto` ganha o botao `Importar do TAP`. Ele busca o(s) TAP(s) registrados para aquele mesmo projeto e traz os indicadores de resultado ja preenchidos la (secao 9 do TAP - Passo 3), evitando digitar tudo de novo: o Valor inicial do TAP vira o Valor atual aqui, e o Valor final vira a Meta. Indicadores com o mesmo nome que ja estiverem na lista nao sao duplicados - o botao pode ser clicado varias vezes com seguranca. Se nenhum TAP for encontrado para o projeto selecionado, aparece um aviso."
Img "44_f12_importar_tap.png" "Indicadores do projeto apos importar do TAP - Valor atual e Meta preenchidos automaticamente." 5.6
Bul "Clique em `Salvar e ver painel`. No rodape, os botoes `Imprimir` e `Exportar Word` geram o documento no formato oficial do SGP-12 (ficha do programa, indicadores, valores estimados e a ficha de cada projeto com suas entregas), pronto para levar a reuniao de pactuacao ou abrir no Word."
Nota "Gate 2 - Pactuacao: no topo da aba Editar dados ha um cartao `Gate 2 - Pactuacao` com Status (Pendente de pactuacao / Pactuado), Data de pactuacao e Aprovador. Assim como o Gate 1 da Solicitacao de Demanda, so o Admin consegue alterar o Status - qualquer outro papel ve o controle travado com um aviso. Ao marcar `Pactuado`, a Data e o Aprovador sao preenchidos automaticamente (editaveis). Enquanto o status estiver `Pactuado`, todos os demais campos do relatorio ficam bloqueados para edicao, inclusive para o Admin, ate que o Gate 2 seja reaberto (status volte para `Pendente de pactuacao`)."
Img "40_f12_gate2.png" "Cartao Gate 2 - Pactuacao preenchido, com o relatorio travado apos a pactuacao." 5.6
Nota "So depois do Gate 2 pactuado o projeto deve avancar para a execucao. Um projeto que comeca a ser executado sem essa pactuacao corre o risco de ser questionado ou desautorizado em momentos criticos. Ao marcar Pactuado, um e-mail de notificacao e disparado automaticamente (secao 2.8)."
Bul "No rodape do Painel, o botao `Historico de versoes` lista todas as vezes que este relatorio foi salvo."
Nota "Versionamento: assim como no Relatorio de Situacao (secao 4.1), toda vez que este relatorio e salvo, uma copia completa fica arquivada com data/hora e quem salvou - util para reconstituir o programa como ele estava em uma pactuacao (Gate 2) anterior. Clicar em `Visualizar` numa versao antiga mostra o mesmo aviso de somente leitura e bloqueia a edicao ate voltar a versao atual."
Img "45_f12_historico_versoes.png" "Lista de versoes salvas do Relatorio de Entregas, com data, programa e quem salvou." 5.6

# ============================================================
# 5. VALIDADOR DE PROJETOS
# ============================================================
H1 "5. Validador de Projetos"
P "O Validador de Projetos e uma ferramenta de apoio a decisao sobre Fatores Criticos de Sucesso (FCS) e Beneficios em instituicoes de ensino. Diferente dos 12 formularios de registro, ele nao gera um artefato com codigo SGP - e um painel de reflexao e diagnostico."
P "E a unica ferramenta, alem do Mapa de Diretrizes, que nao exige login. Todo o conteudo funciona livremente; login so e pedido para vincular uma avaliacao a um projeto e salvar o veredito no historico daquele projeto (secao 2.6 do documento Regras de Acesso detalha esse caso)."
H2 "5.1 Quadro de conexoes"
P "Mostra a matriz analitica entre 7 fatores criticos de sucesso (F1-F7) e 5 dimensoes de beneficio (B1-B5). Clique em um fator ou dimensao para ver as conexoes."
Img "05_validador_quadro.png" "Quadro de conexoes com o fator F1 selecionado - o painel a direita mostra os beneficios que ele habilita." 5.8
H2 "5.2 Simulador - e se...?"
P "Ajuste o quanto cada fator esta presente no seu projeto (0-100%) e veja o potencial estimado de cada dimensao de beneficio recalcular em tempo real."
Img "06_validador_simulador.png" "Simulador de fatores x beneficios. Abaixo dele, o aviso de login para quem quiser vincular a um projeto." 5.8
H2 "5.3 Assistente de decisao em 8 perguntas"
P "Responda a oito perguntas do semaforo de decisao, uma de cada vez. Ao final, o painel devolve um veredito com a regra aplicada."
Img "07_validador_assistente.png" "Primeira pergunta do assistente de decisao." 5.8
Img "08_validador_veredito.png" "Veredito final, com o placar de respostas verdes/amarelas/vermelhas." 5.8
Nota "Regra aplicada: qualquer resposta vermelha leva a encerrar/repactuar; tres ou mais amarelas levam a replanejar; caso contrario, o veredito e concluir."

# ============================================================
# 6. ADMINISTRACAO DO SISTEMA
# ============================================================
H1 "6. Administracao e Painel Executivo"
P "Quatro paginas independentes, listadas no menu lateral em `Gestao`: a Administracao (gerencia usuarios, equipes e configuracoes, sempre restrita a Admin - secao 2.3), o Painel Executivo (visao consolidada e agregada de todo o sistema, com acesso configuravel por papel - secao 6.3), o Painel de Prazos (cronograma dos projetos em andamento com avisos automaticos de prazo, tambem com acesso configuravel por papel - secoes 6.3 e 6.4.2) e o Relatorio de Resultados (indicadores estrategicos e plano de acao para a reuniao mensal de resultados, tambem com acesso configuravel por papel - secoes 6.3 e 6.7)."
H2 "6.1 Usuarios"
P "Lista todos os usuarios que ja fizeram login pelo menos uma vez, com nome, telefone (editavel diretamente na lista), um seletor para alterar o papel, e um link `Ver historico` por usuario."
Img "09_admin_usuarios.png" "Aba Usuarios: nome, telefone, papel atual e seletor de alteracao de papel por usuario." 5.8
P "No topo da lista, o Admin pode cadastrar uma pessoa que ainda nao fez login, preenchendo nome, telefone (opcional), e-mail e papel e clicando em `+ Adicionar usuario`. A conta e criada de verdade na hora (via Supabase Auth) - a pessoa ja aparece na lista de usuarios normalmente, com o papel e o nome definidos, e ja pode ser adicionada a equipe de qualquer projeto (secao 6.2) sem precisar esperar ela logar."
Nota "Quando a pessoa eventualmente fizer o primeiro login (link magico), o Supabase Auth reconhece o e-mail e autentica direto na conta ja criada - nao gera uma conta duplicada, e o perfil (nome, telefone, papel) ja definido pelo Admin permanece do jeito que estava."
P "Antes de 15/09/2026, cadastrar alguem aqui criava so um pre-cadastro pendente, que so virava conta de verdade no primeiro login da pessoa - por isso ela nao podia ser adicionada a uma equipe antes disso. Um pre-cadastro feito daquele jeito (selo `Pendente - 1o login` na lista) ainda funciona: o Admin pode ajustar telefone e papel, remove-lo pelo botao `Remover`, ou usar o novo botao `Criar conta agora` para converte-lo numa conta de verdade imediatamente, sem esperar o login."
P "Cada usuario ja cadastrado tem um link `Ver historico`, que expande um painel logo abaixo da linha listando toda troca de papel e de telefone ja feita nesse perfil - valor anterior, novo valor, quando aconteceu e quem alterou. Clicar de novo (`Fechar historico`) recolhe o painel."
Nota "Como o papel de um usuario controla o acesso a praticamente toda a ferramenta (paginas restritas a Admin, listas configuraveis de quem ve cada relatorio, contagem de equipe), esse historico existe para que uma promocao ou rebaixamento de papel nunca fique sem rastro de quem fez e quando - o mesmo principio de rastreabilidade ja usado no `Ver historico` dos 9 formularios do ciclo de vida do projeto."
Nota "Trava contra ficar sem nenhum Admin: o seletor de papel recusa tirar o papel Admin de alguem (a propria pessoa ou outra) quando essa pessoa e a unica com esse papel no sistema, com um aviso explicando o motivo. Sem essa trava, um Admin poderia acidentalmente se rebaixar (ou rebaixar o unico outro Admin) e ninguem mais conseguiria acessar esta tela para reverter, ja que ela e sempre exclusiva a quem tem papel Admin."
P "Cada usuario ja cadastrado (menos o proprio Admin logado) tem um botao `Remover` na ultima coluna. Ao clicar, o sistema checa na hora se a pessoa faz parte da equipe de algum projeto, criou algum projeto, tem registro criado em seu nome em algum dos 9 formularios que gravam essa informacao, ou tem alguma inscricao de notificacao push registrada nesse e-mail. Se algum desses vinculos existir, a remocao e recusada e o Admin ve exatamente qual foi o motivo; caso contrario, pede confirmacao final antes de excluir o perfil."
Nota "Remover aqui apaga so o perfil (papel, nome, telefone) do sistema - nao impede a pessoa de logar de novo depois; se logar, um perfil novo e criado do zero, com o papel padrao Solicitante. Os tres formularios que sao documentos unicos e compartilhados (Plano de Comunicacao, Relatorio de Situacao e Relatorio de Entregas) nao entram na checagem automatica, por nao terem o conceito de `quem criou o registro`."
Nota "A propria remocao tambem fica registrada: antes de excluir o perfil, o sistema grava no historico de alteracoes (a mesma tabela usada pelo `Ver historico`) quem removeu quem, com qual papel a pessoa estava e quando - mesmo que o perfil nao exista mais para abrir o historico depois pela tela."
H2 "6.2 Equipes"
P "Define quem faz parte da equipe de cada projeto - a base da restricao por equipe descrita na secao 2.5. Selecione um projeto, escolha uma pessoa na lista e clique em `+ Adicionar a equipe`."
Img "10_admin_equipes.png" "Aba Equipes, com um projeto selecionado e um membro ja cadastrado." 5.8
Nota "A lista de `+ Adicionar a equipe` so mostra contas de verdade (secao 6.1), sem contar quem ja esta na equipe daquele projeto - se alguem que deveria aparecer nao estiver na lista, confira se essa pessoa ja tem conta criada (nao um pre-cadastro antigo ainda pendente) e se ela ja nao faz parte da equipe."
H2 "6.3 Configuracoes"
P "Controla os dois interruptores de acesso sem login (Solicitacao de Demanda e Ata de Reuniao), descritos na secao 2.6."
Img "11_admin_config.png" "Aba Configuracoes: os dois interruptores de acesso sem login, hoje desativados." 5.8
P "Logo abaixo, cinco grupos de caixas de selecao controlam quem pode ver o Painel Executivo, o Painel de Prazos, o Relatorio de Resultados, o Relatorio de Situacao de Projetos e o Relatorio de Entregas e Beneficios: o Admin marca um ou mais papeis por pagina, e quem tiver um papel fora da lista marcada ve `Acesso restrito` ao tentar abrir aquela pagina. Por padrao, o Painel Executivo, o Painel de Prazos e o Relatorio de Resultados comecam marcados so para Admin (mesmo comportamento de antes), e os dois relatorios finais comecam com todos os papeis marcados (acesso livre, tambem o comportamento de antes) - a restricao so entra em vigor se o Admin desmarcar algum papel."
P "Um sexto grupo, `Quem pode editar o Plano de Comunicacao de Projeto`, controla so a aba `Editar dados` desse formulario (Passo 10) - diferente dos cinco grupos acima, que controlam a pagina inteira. Comeca marcado apenas para PMO/Admin; qualquer papel fora da lista continua vendo e imprimindo o Painel normalmente, so nao consegue editar."
Nota "O papel Admin sempre tem acesso a essas seis configuracoes, mesmo que fique desmarcado por engano - isso evita que o proprio Admin perca o acesso e fique sem como reverter a configuracao."
P "Por ultimo, dois interruptores independentes controlam a importacao por IA na Ata de Reuniao - cada um desligado por padrao, ja que cada uso tem custo real de API, e cada um vale para todo mundo sem excecao, inclusive Admin: um para a importacao de transcricao em texto (secao 2.10) e outro, separado, para a importacao de audio (secao 2.11). Cada interruptor tem seu proprio grupo de papeis logo abaixo, com Admin sempre incluido quando ligado. Como a importacao de audio so serve para gerar o texto que passa pela mesma analise, o botao `Anexar audio` na Ata so aparece para quem tem as duas permissoes ativas ao mesmo tempo."
H2 "6.4 Painel Executivo"
P "Agrega em tempo real dados de todos os 10 formularios de registro e dos 2 relatorios - sem exigir nenhuma mudanca no banco de dados, apenas consultando o que ja esta salvo. Pensada para dar ao PMO uma visao geral do sistema inteiro em uma unica tela, sem precisar abrir cada formulario individualmente. Por padrao so o Admin ve esta pagina, mas isso pode ser ampliado para outros papeis em Administracao > Configuracoes (secao 6.3)."
Bul "Gates de aprovacao: contagem de Solicitacoes de Demanda por status (Gate 1) e o status atual do Relatorio de Entregas (Gate 2 - Pactuacao)."
Bul "Governanca (dados do Relatorio de Situacao - secao 4.1): quantidade de Entraves e Encaminhamentos ainda `Aberto`, e quantos Resultados alcancados ja foram registrados."
Bul "Riscos do portfolio (dados do TAP - Passo 3, secao 9): contagem total de riscos em aberto (Status diferente de Encerrado) com severidade Critico e com severidade Alto, somando o Registro de riscos de todos os TAPs cadastrados - mesma matriz de severidade (Probabilidade x Impacto) usada na Matriz de risco de cada TAP."
Bul "Custos por categoria (portfolio): tabela listando cada categoria de custo (casada pelo nome, o mesmo casamento do Comparativo por categoria do Relatorio de Situacao, secao 4.1) com em quantos projetos do portfolio ela tem estimativa no TAP e em quantos desses o Realizado ultrapassou o Estimado - para achar categorias que estouram o orcamento de forma sistematica, nao so num projeto isolado. So entram projetos com Projeto vinculado, custo estimado no TAP e custo realizado registrado."
Bul "Entregas da EAP (portfolio): % de entregas (nivel 2) concluidas somando todas as EAPs cadastradas, e em quantos projetos existe pelo menos uma entrega atrasada - mesmo casamento pelo nome com o cronograma do Planejamento e mesmo calculo automatico ja usado na propria EAP (Passo 5, secao 6.4.1)."
Bul "Tarefas do cronograma (portfolio): soma as tarefas cadastradas dentro dos marcos do cronograma do Planejamento (Passo 4, aba 3) de todos os projetos, mostrando o percentual concluido e em quantos projetos existe pelo menos uma tarefa pendente."
Bul "Registros por formulario: uma linha por artefato (Canvas, TAP, Planejamento, EAP, SMP, Ata de Reuniao, TEP, RLA, alem da propria Solicitacao de Demanda), com o total e a contagem por status de cada um."
Bul "Portfolio de projetos: contagem de projetos Aprovados, Encerrados (Concluido/Paralisado/Cancelado, via TEP), Reprovados e em avaliacao no Gate 1, a partir do cadastro compartilhado (secao 2.5)."
Bul "Visao atual por projeto: uma linha por projeto do cadastro compartilhado, com Status do cadastro, Situacao mais recente do Relatorio de Situacao (secao 4.1, com o percentual concluido quando ja reportado), Saude do projeto (secao 6.4.1) e Progresso do ciclo de vida - quantos dos 9 formularios (Demanda a RLA) ja tem registro vinculado aquele projeto, com a lista do que ainda falta."
Img "48_painel_executivo.png" "Painel Executivo: Gates de aprovacao e indicadores de governanca do Relatorio de Situacao." 5.8
Img "49_painel_executivo_tabela.png" "Registros por formulario e portfolio de projetos, com contagem por status." 5.8
Img "56_painel_executivo_visao_projeto.png" "Visao atual por projeto: status do cadastro, situacao do Relatorio de Situacao, saude do projeto e progresso do ciclo de vida, projeto a projeto." 5.8
Nota "Por reunir dados de varios formularios numa unica tela, o carregamento faz uma consulta para cada um dos 10 formularios de registro - pode levar alguns segundos a mais que os outros formularios da ferramenta. Clique em `Atualizar` para recarregar os numeros mais recentes a qualquer momento."
Nota "Na coluna Situacao, o selo `Nao reportado` significa que o projeto ainda nao tem uma linha propria no Relatorio de Situacao - nao e um erro, apenas indica que ninguem preencheu o acompanhamento daquele projeto ainda."
H2 "6.4.1 Ficha do Projeto"
P "Visao consolidada de um unico projeto, reunindo em uma tela so o que hoje esta espalhado pelos 9 formularios do ciclo de vida. Aberta clicando no nome de um projeto na tabela `Visao atual por projeto` do Painel Executivo (secao 6.4); abrir a pagina direto, sem escolher um projeto antes, mostra um seletor com todos os projetos cadastrados. Diferente do Painel Executivo, nao tem restricao por papel - qualquer usuario autenticado pode ver a ficha de qualquer projeto, a mesma logica ja usada pelo botao `Ver historico` em cada formulario."
Bul "Dados do projeto: nome, unidade, gerente, status do cadastro compartilhado e quando foi criado/atualizado pela ultima vez."
Bul "Saude do projeto: selo geral (Normal/Atencao/Critico/Encerrado/Sem dados), % de marcos concluidos, atraso do marco mais atrasado (em dias), % do orcamento realizado frente ao estimado no TAP e quantidade de riscos altos/criticos em aberto - ver detalhe do calculo logo abaixo."
Bul "Custos: Estimado (TAP) x Realizado: drill-down do cartao de Orcamento acima - tabela cruzando, categoria por categoria, o custo Estimado no TAP com o Realizado no Relatorio de Situacao, mesmo casamento por nome ja usado pelo Comparativo por categoria do Relatorio de Situacao (secao 4.1). Cada categoria ganha um selo `Dentro do estimado`, `Acima do estimado` ou `Sem estimativa no TAP`."
Bul "Entregas da EAP: status automatico: lista cada Entrega (nivel 2) da EAP (Passo 5) deste projeto com um selo de status - `Concluida`, `Atrasada`, `No prazo` ou `Sem cronograma vinculado` - calculado automaticamente casando pelo nome com o cronograma do Planejamento, sem nenhum campo novo para preencher. Um resumo no topo mostra o percentual de entregas concluidas do projeto."
Bul "Situacao atual: a linha desse projeto no Relatorio de Situacao - status, % concluido, datas previstas, marcos, custos realizados e os campos `Merece destaque`/`Merece atencao`, quando preenchidos."
Bul "Progresso do ciclo de vida: quantos registros existem em cada um dos 9 formularios (Demanda a RLA) para esse projeto, com o status de cada um e um link para abrir o formulario correspondente."
Bul "Anexos do projeto: reune, num so lugar, todo arquivo enviado nos 3 pontos do sistema que aceitam upload real - Solicitacao de Demanda (Passo 1), SMP (Passo 7) e a aba Viabilidade do Planejamento (Passo 4) - cada um com a origem (formulario, e o protocolo quando ha mais de uma SMP) e um botao `Baixar`. Nao e preciso abrir cada formulario separadamente para achar um documento anexado."
Bul "Cronograma (Planejamento) e custos planejados (TAP): o cronograma de entregas macro do Planejamento e Desenvolvimento de Projeto e o custo total estimado do TAP desse projeto, para comparar com os custos realizados do Relatorio de Situacao."
Bul "Equipe do projeto: mesma lista de Administracao > Equipes (secao 6.2), com o papel de cada membro vindo do cadastro global de usuarios."
Bul "Historico de alteracoes: o mesmo `Ver historico` ja disponivel em cada formulario vinculado a esse projeto."
Img "57_ficha_projeto.png" "Ficha do Projeto: dados do cadastro, saude do projeto, situacao atual, progresso do ciclo de vida e cronograma do Planejamento." 5.8
Nota "Como o selo de saude e calculado: compara quatro sinais - a situacao do projeto esta `Atrasado` no Relatorio de Situacao; existe pelo menos um marco do cronograma do Planejamento, sem Conclusao real preenchida, cuja data-alvo ja passou (mesma definicao usada no indicador `Marcos no prazo` do Relatorio de Resultados, secao 6.7 - o que conta e a data-alvo e a conclusao real, nao o status manual do marco); o custo realizado ultrapassa o estimado no TAP; e existe pelo menos um risco em aberto (Status diferente de Encerrado) com severidade Alto ou Critico no Registro de riscos do TAP. Dois ou mais sinais ruins = Critico; um sinal = Atencao; nenhum = Normal. Projetos Concluido/Cancelado mostram `Encerrado`, e projetos sem nenhuma linha no Relatorio de Situacao mostram `Sem dados` - nunca inventa um status a partir de dado ausente. A mesma coluna `Saude` aparece na tabela `Visao atual por projeto` do Painel Executivo (secao 6.4), com o mesmo calculo."
Nota "Riscos do portfolio: o Painel Executivo tambem mostra, em uma secao propria, a contagem total de riscos criticos e altos em aberto somando o Registro de riscos de todos os TAPs do portfolio - a mesma leitura de severidade usada na matriz de risco do TAP (Passo 3), agregada em nivel de portfolio."
Nota "Cada secao mostra um estado vazio proprio quando nao ha dado ainda (ex.: `Nenhum TAP encontrado para este projeto ainda`) - nunca um erro, apenas indica que aquela etapa ainda nao foi preenchida."
H2 "6.4.2 Painel de Prazos"
P "Pagina pensada para ficar aberta numa TV ou monitor da sala de projetos, atualizando sozinha a cada 2 minutos, com o cronograma (marcos do Planejamento e Desenvolvimento de Projeto, Passo 4) de todos os projetos ainda ativos - projetos ja encerrados via TEP nao entram. Por padrao so o Admin ve esta pagina, mas isso pode ser ampliado para outros papeis em Administracao > Configuracoes (secao 6.3), da mesma forma que o Painel Executivo."
Bul "Tres indicadores no topo: total de marcos atrasados no portfolio (sem Conclusao real, com o Termino previsto ja vencido), total de marcos vencendo em ate 5 dias, e % de marcos concluidos somando todos os projetos ativos."
Bul "Um cartao por projeto ativo com Planejamento registrado, ordenado com os projetos com mais marcos atrasados primeiro: nome do projeto, gerente, barra de % concluido, e a lista dos marcos atrasados (destacados em vermelho) e vencendo em breve (destacados em amarelo), cada um com o responsavel do cronograma e ha quantos dias esta atrasado ou em quantos dias vence. Um projeto sem nenhum marco atrasado ou vencendo mostra `Tudo em dia`."
Nota "Execucao ao lado de cada marco: alem do prazo, cada marco atrasado ou vencendo mostra o que a pessoa sinalizou no campo Execucao (secao 6.2, Passo 4) e ha quanto tempo - por exemplo `Em andamento - atualizado ha 2 dias`. Um marco sem nenhuma atualizacao mostra `sem atualizacao de execucao` em destaque - e o sinal de que ninguem sinalizou estar trabalhando naquilo, util para separar quem esta atrasado mas ativo de quem esta atrasado em silencio."
Nota "Este painel so mostra os dados - quem efetivamente recebe um aviso por e-mail/push quando um prazo se aproxima e o mecanismo descrito a seguir, que roda sozinho todo dia sem precisar desta tela aberta."
H2 "6.4.3 Aviso automatico de prazo (e-mail e push)"
P "Complementando o Painel de Prazos, um mecanismo roda sozinho uma vez por dia (sem depender de ninguem abrir o sistema) e avisa por e-mail - e por push, para quem tiver ativado (secao 2.9) - quando um marco do cronograma esta chegando ou ja passou do prazo, sem Conclusao real preenchida:"
Bul "5 dias antes do Termino previsto."
Bul "2 dias antes."
Bul "No dia do vencimento."
Bul "Atrasado - repete a cada 7 dias enquanto o marco continuar sem Conclusao real, para nao lembrar todo dia da mesma pendencia."
Nota "Os numeros acima (5, 2 e 7 dias) sao o padrao de fabrica - o Admin pode altera-los em Administracao > Configuracoes > Regras de aviso automatico de prazo e execucao (secao 6.3), sem mexer em codigo. Um valor em branco ou invalido faz a regra correspondente voltar ao padrao, sem deixar de funcionar."
P "Quem recebe cada aviso: toda conta vinculada no campo `Responsavel (conta)` do cronograma (Passo 4, aba 3), quando preenchido - pode ser uma ou mais de uma pessoa -, mais toda conta com papel Gerente de Projetos que estiver na equipe daquele projeto (secao 2.5) - nao o campo de texto livre `Gerente` do cadastro do projeto. Se nenhuma conta existir para aquele marco (nem `Responsavel (conta)`, nem Gerente de Projetos na equipe), ninguem recebe o aviso automatico, mas o indicador visual no Painel de Prazos continua aparecendo normalmente."
Nota "O campo `Responsavel (conta)` e nao a pessoa preenche o texto livre `Responsavel` ja existente: e um segundo campo, um botao que abre um painel com uma caixa de selecao por conta da equipe do projeto selecionado, podendo marcar mais de uma pessoa para o mesmo marco. A lista mostra o nome cadastrado de cada conta (caindo para o e-mail quando nao houver nome), e o botao resume a escolha (nome, ou `nome +N` com mais de uma pessoa marcada). O texto livre `Responsavel` continua existindo do lado, sem nenhuma mudanca - o campo novo serve so para o sistema saber para quais contas mandar o aviso automatico."
Nota "Cada aviso e enviado uma unica vez por marco e por janela (5 dias, 2 dias, no dia) - nao repete todo dia dentro da mesma janela. Ja o aviso de atraso repete a cada 7 dias, para continuar chegando enquanto a pendencia nao for resolvida, sem virar um lembrete diario."
Bul "O e-mail traz um link `Abrir o Meu Painel`, para quem recebeu o aviso ja ir direto tratar o item (secao 6.4.4) - e um link generico, sem nada especifico da pessoa; quem clicar precisa estar logado com a propria conta para ver o que esta atribuido a ela."
Bul "O e-mail termina indicando a Gerencia de Projetos como alternativa de contato, caso a pessoa prefira ajustar o prazo diretamente em vez de seguir so pelo aviso automatico - o nome do Gerente vem do campo `Gerente do projeto` do proprio Planejamento (aba 1, capa), nao e fixo; se esse campo estiver vazio no projeto, a frase aparece sem nome."
H2 "Aviso automatico de silencio (e-mail e push)"
P "Complementa o aviso de prazo acima, mas olha para outra coisa: em vez do vencimento, olha se a pessoa esta de fato sinalizando execucao. O mesmo mecanismo diario dispara um segundo aviso quando um marco ja comecou (Inicio previsto ja chegou) e esta 5 dias ou mais sem nenhuma mudanca no campo Execucao (secao 6.2, Passo 4) - contando a partir da ultima atualizacao, ou do proprio Inicio previsto se o campo nunca foi preenchido. Repete a cada 5 dias enquanto continuar parado, com o mesmo destinatario e o mesmo link para o Meu Painel do aviso de prazo."
Nota "Por que existe: o aviso de prazo acima so olha a data final - um marco com o Termino previsto ainda longe (semanas ou meses) podia ficar parado o tempo todo sem ninguem ser avisado, mesmo com o campo Execucao existindo desde a criacao dele. O aviso de silencio fecha essa lacuna, avisando por causa da inatividade, nao do prazo. No mesmo dia em que o aviso de prazo ja foi enviado para um marco, o aviso de silencio nao dispara tambem, para nao mandar dois e-mails de uma vez. Os dois numeros (5 e 5 dias) tambem sao configuraveis em Administracao > Configuracoes (secao 6.3), junto com os do aviso de prazo."
H2 "6.4.4 Meu Painel"
P "Visao pessoal, sem restricao de papel - qualquer usuario autenticado abre esta pagina e ve, reunido numa unica tela, tudo que esta atribuido a sua propria conta em qualquer projeto, com um botao para concluir cada item direto da lista, sem precisar abrir o formulario de origem."
$r644 = @(
  @("Secao","De onde vem","Botao"),
  @("Meus projetos","Projetos em que a conta faz parte da equipe (secao 2.5)","-"),
  @("Marcos atribuidos a mim","Cronograma do Planejamento (Passo 4, aba 3), campo Responsavel (conta) do marco","Concluir - grava a Conclusao real com a data de hoje"),
  @("Tarefas atribuidas a mim","Tarefas dentro de um marco (Passo 4, aba 3) cujo marco tenha Responsavel (conta)","Concluir - marca a tarefa como concluida"),
  @("Riscos atribuidos a mim (TAP)","Registro de riscos do TAP (Passo 3), campo Responsavel (conta) do risco","Encerrar - muda o Status do risco para Encerrado"),
  @("Entradas atribuidas a mim (Planejamento)","Aba Entradas do Planejamento (Passo 4, aba 4), campo Responsavel (conta) da linha","Concluir - muda o Status da entrada para OK"),
  @("Entraves e encaminhamentos atribuidos a mim","Relatorio de Situacao (secao 4.1), campo Responsavel (conta) - aqui com todas as contas do sistema, nao so a equipe do projeto","Resolver - muda o Status para Resolvido")
)
TableSimple $r644 @(4.2,7.6,4.2)
Img "59_meu_painel.png" "Meu Painel: marcos, tarefas, riscos, entradas e entraves/encaminhamentos atribuidos a conta logada, cada um com seu botao de conclusao." 5.6
Nota "Responsavel (conta) aceita mais de uma pessoa: nas quatro linhas de Marcos/Tarefas, Riscos, Entradas e Entraves/encaminhamentos acima, o campo Responsavel (conta) pode ter uma ou varias contas marcadas ao mesmo tempo - o item aparece no Meu Painel de cada pessoa marcada, nao so de uma, e concluir pelo botao de qualquer uma delas resolve o item para todas."
Nota "Execucao, ao lado de cada marco: um seletor com A iniciar/Em andamento/Atrasado/Concluido, salvando na hora ao ser alterado (mesmo padrao do botao Concluir) - e aqui, no Meu Painel, que normalmente faz mais sentido atualizar, sem precisar abrir o Planejamento inteiro. Grava tambem a data/hora da ultima mudanca, usada no Painel de Prazos (secao 6.4.2) para mostrar ha quanto tempo aquele marco foi atualizado pela ultima vez."
Nota "A aba Verificacao do Planejamento (aba 6) nao entra no Meu Painel nem tem campo Responsavel (conta) - ficou de fora deliberadamente."
Nota "Concluir um item aqui grava direto no mesmo registro compartilhado do formulario de origem (Planejamento, TAP ou Relatorio de Situacao) - nao e uma copia separada. Se duas pessoas mexerem no mesmo registro ao mesmo tempo (por exemplo, uma concluindo um item pelo Meu Painel enquanto outra edita o formulario de origem), o sistema busca a versao mais recente do registro e tenta gravar de novo automaticamente antes de desistir - mesma logica de edicao simultanea ja usada nos demais formularios (secao 2.14)."
P "O botao `Atualizar`, no topo da pagina, recarrega a lista sem precisar sair e voltar."

H2 "6.4.5 Gestao de Atividades"
P "Visao agregada, para Admin/PMO, da carga de trabalho de toda a equipe: agrupa os mesmos itens do Meu Painel (secao 6.4.4) - marcos e tarefas do cronograma, riscos do TAP, entradas do Planejamento, entraves/encaminhamentos do Relatorio de Situacao - mas por conta responsavel, nao por login individual. Por padrao so o Admin ve esta pagina, mas isso pode ser ampliado para outros papeis em Administracao > Configuracoes (secao 6.3), da mesma forma que o Painel Executivo e o Painel de Prazos."
Bul "Quatro indicadores no topo: quantas pessoas tem pelo menos um item pendente, total de itens em aberto no portfolio, quantos estao atrasados (marcos vencidos sem Conclusao real e entraves/encaminhamentos com prazo vencido) e quantos itens em aberto ainda nao tem nenhuma conta responsavel atribuida."
Bul "Um cartao por pessoa, com nome, papel e a contagem de itens em cada uma das 5 categorias, mais quantos estao atrasados - clicar no cartao expande a lista detalhada de cada item. Pessoas com pelo menos um item atrasado ja aparecem expandidas ao carregar a pagina, sem precisar clicar. E somente visualizacao: nao ha botao para concluir um item em nome de outra pessoa."
Bul "Marcos, Tarefas, Riscos (TAP) e Entradas (Planejamento) agrupados por projeto: dentro de cada categoria, os itens da pessoa ficam organizados sob o nome do projeto a que pertencem, em vez de soltos numa lista so. O cabecalho de cada grupo mostra quantos itens tem e, em Marcos, quantos deles estao atrasados. Os grupos de Marcos aparecem ordenados pelo projeto mais critico primeiro (mais itens atrasados, depois o prazo mais proximo); nas demais categorias a ordem e alfabetica pelo nome do projeto. Cada marco mostra explicitamente de qual campo a data e (`Termino previsto`), o mesmo nome usado na coluna correspondente do cronograma do Planejamento (Passo 4)."
Bul "Itens sem responsavel (conta) atribuido: contagem por categoria dos itens em aberto que so tem o campo de texto livre Responsavel preenchido, sem o seletor de conta - ajuda o PMO a encontrar pendencias sem dono definido antes que virem atraso."
Img "60_gestao_atividades.png" "Gestao de Atividades: indicadores gerais e um cartao expandido mostrando o detalhamento por categoria de uma pessoa." 5.6
Nota "Controle de acesso (Row Level Security no Supabase): diferente das demais listas de papeis da Administracao, a permissao desta pagina tambem libera enxergar nome e papel de todos os usuarios do sistema, nao so abrir a tela - e o dado que a pagina usa para agrupar as pendencias por pessoa. A funcao `listar_perfis_gestao_atividades` (SECURITY DEFINER) confere essa permissao direto no banco antes de devolver qualquer linha - mesmo alguem tentando contornar a tela nao consegue essa lista sem estar autorizado."
H2 "6.5 Ambiente de Treino"
P "O sistema tem um ambiente de treinamento/demonstracao totalmente separado da producao - mesma ferramenta, mesmo schema de dados, mas um banco proprio, populado com 7 projetos ficticios cobrindo o ciclo completo (de uma Solicitacao de Demanda ainda pendente ate um projeto totalmente encerrado com TEP e RLA). Serve para apresentacoes, estudo e capacitacao sem tocar em nenhum dado real. Veja o roteiro `Como rodar o Ambiente de Treinamento` para o endereco, as contas de acesso e o passo a passo completo."
P "Quando a Administracao e acessada em producao, aparece uma quarta aba, `Ambiente de Treino`, com um unico botao: `Resetar ambiente de treinamento`. Ao clicar (e confirmar o aviso), o sistema apaga todos os projetos e registros de formulario atualmente no ambiente de treinamento e recria os 7 projetos de exemplo do zero - util para deixar o ambiente limpo antes de uma nova turma, sem precisar rodar nenhum comando manualmente."
Bul "A aba so aparece quando a Administracao e acessada pelo endereco de producao - quem acessa a Administracao pelo proprio ambiente de treinamento nao ve essa opcao."
Bul "Nao afeta nenhum dado de producao em nenhuma hipotese - a acao roda numa Edge Function que so tem permissao de escrita no banco do ambiente de treinamento."
Bul "As 6 contas fixas de treinamento e as configuracoes do ambiente de treino nao sao apagadas pelo reset - so os projetos e os registros de formulario."
Nota "Assim como as outras acoes exclusivas de Admin, a Edge Function confere o papel do usuario direto no banco de producao antes de fazer qualquer alteracao - nao basta ver o botao na tela."
H2 "6.6 Backup automatizado"
P "No fim da aba Configuracoes (secao 6.3), um bloco separado deixa o Admin/PMO definir os parametros de uma rotina de backup periodico do sistema inteiro (banco de dados, arquivos anexados e codigo-fonte) para um destino fora do GitHub - complemento operacional a proposta `Backup Automatizado` entregue separadamente."
Bul "Habilitar backup automatizado: interruptor geral da rotina."
Bul "Destino: OneDrive/SharePoint institucional, outro armazenamento de nuvem ou servidor da instituicao, mais o caminho/pasta exato onde gravar."
Bul "Frequencia: diaria ou semanal."
Bul "Retencao (dias, opcional): depois de quantos dias os backups antigos podem ser descartados; vazio mantem tudo indefinidamente."
Bul "E-mail de alerta: para onde avisar automaticamente se uma execucao falhar."
Bul "E-mails com acesso ao destino externo: campo de registro/referencia de quem deveria ter permissao na pasta de destino - conceder esse acesso de fato continua sendo feito manualmente onde o destino estiver (OneDrive, repositorio etc.)."
Bul "Quem pode gerenciar essas configuracoes: papeis liberados para alterar esta tela, alem do PMO/Admin, que sempre tem acesso."
Bul "Ultima execucao: data/hora e status (sucesso ou falha) da ultima vez que a rotina automatizada rodou, preenchido por ela mesma."
Nota "Esta tela define os parametros - quem efetivamente le essas configuracoes e executa a copia dos dados e uma rotina agendada fora do sistema (GitHub Actions). Como o agendamento desse tipo de rotina roda em horarios fixos (nao pode ser alterado em tempo real por este formulario), ela e programada para rodar com frequencia maior do que o necessario e, a cada execucao, checar aqui se ja esta na hora de rodar de acordo com a frequencia configurada - na pratica, o efeito para quem usa a tela e o mesmo de definir a frequencia diretamente."
Nota "Destino no SharePoint confirmado: desde 28/08/2026, o envio da copia externa para o SharePoint institucional esta confirmado funcionando de ponta a ponta, com o aplicativo usando apenas a permissao `Sites.Selected` (acesso restrito a um unico site, o modelo de menor privilegio recomendado desde o inicio) - sem depender de nenhuma permissao ampla sobre todos os sites da instituicao."
P "Por ultimo, um bloco `Metas dos indicadores estrategicos` define a meta em % de cada um dos 4 indicadores do Relatorio de Resultados (secao 6.7) - Marcos no prazo, Projetos dentro do orcamento, Entraves resolvidos e Projetos sem risco alto/critico. Enquanto uma meta ficar vazia, o cartao correspondente mostra `Meta nao definida` em vez de um selo de status."
Img "53_admin_metas_resultados.png" "Bloco Metas dos indicadores estrategicos, com os campos preenchidos." 5.8
P "Um ultimo bloco, `Regras de aviso automatico de prazo e execucao`, define em dias as duas regras do mecanismo diario descrito na secao 6.4.3: quantos dias antes do Termino previsto cada janela do aviso de prazo comeca (padrao 5 e 2), a cada quantos dias o aviso de atrasado se repete (padrao 7), apos quantos dias sem atualizar Execucao o aviso de silencio dispara (padrao 5) e a cada quantos dias ele se repete (padrao 5). Um campo em branco ou com valor invalido faz a regra correspondente usar o padrao de fabrica, sem deixar de funcionar."
H2 "6.7 Relatorio de Resultados"
P "Pagina separada do Painel Executivo, pensada para a reuniao mensal de resultados: reune os indicadores estrategicos do portfolio e o plano de acao dos pontos em aberto, no mesmo estilo visual das demais telas de gestao. Por padrao so o Admin ve esta pagina, mas isso pode ser ampliado para outros papeis em Administracao > Configuracoes (secao 6.3)."
Bul "Marcos no prazo: % dos marcos de projetos ainda ativos no Relatorio de Situacao (secao 4.1) que NAO estao atrasados - a data-alvo (do cronograma do Planejamento) ainda nao passou, o marco ja tem Conclusao real preenchida no Planejamento, ou o marco esta marcado como Concluido/Cancelado."
Bul "Projetos dentro do orcamento: % dos projetos avaliados cujo total de Custos realizados (secao 4.1) fica dentro do total estimado no TAP daquele projeto. So entram na conta os projetos com projeto vinculado e ao menos uma categoria de custo realizado preenchida."
Bul "Entraves resolvidos: dos entraves que estavam Aberto no ultimo salvamento do Relatorio de Situacao, quantos ja aparecem como Resolvido agora."
Bul "Projetos sem risco alto/critico: % dos projetos com TAP registrado (com Projeto vinculado) que nao tem nenhum risco em aberto (Status diferente de Encerrado) com severidade Alto ou Critico - a mesma matriz de severidade da Matriz de risco do TAP (Passo 3, secao 9), agregada em nivel de portfolio."
Img "51_relatorio_resultados_kpis.png" "Os quatro indicadores estrategicos do Relatorio de Resultados, com meta, barra de progresso, selo de status e tendencia." 5.8
Nota "Meta e selo de status (Normal/Atencao/Critico): cada indicador compara o realizado com a meta definida pelo Admin (secao 6.3). Sem meta definida, o cartao mostra `Meta nao definida` em vez de um selo - nunca inventa um status. Com meta definida: Normal na meta ou acima; Atencao ate 15 pontos percentuais abaixo da meta; Critico alem disso."
Nota "Tendencia: Marcos no prazo, Projetos dentro do orcamento e Projetos sem risco alto/critico mostram a evolucao dos ultimos salvamentos do Relatorio de Situacao (ex.: `81% -> 85% -> 87%`), recalculada a partir do historico de versoes que esse relatorio ja mantinha (secao 4.1) - nao e um fechamento de calendario mes a mes, e sim uma tendencia de `ultimos salvamentos`. A tendencia de risco so existe a partir do primeiro salvamento feito depois que este indicador foi lancado - salvamentos anteriores nao tem esse numero registrado."
Nota "O Planejamento guarda tanto a data-alvo (Termino previsto) quanto a Conclusao real de cada marco, entao o indicador ja distingue um marco que terminou no prazo de um que so ainda nao venceu. Projetos dentro do orcamento e Entraves resolvidos continuam aparecendo como travessao (`-`) quando nao ha dado suficiente para calcular, em vez de mostrar um numero sem base."
Bul "Plano de acao: uma linha para cada entrave e encaminhamento com status Aberto no Relatorio de Situacao, nas colunas O que?, Por que?, Quem?, Quando? e Resultado esperado."
Img "52_relatorio_resultados_plano_acao.png" "Tabela Plano de acao, com entraves e encaminhamentos em aberto." 5.8
Nota "O rotulo `Atualizado em` no topo da pagina mostra quando os dados foram lidos - esta pagina e uma foto em tempo real, assim como o Painel Executivo, nao um fechamento formal de mes."
P "No topo da pagina, dois botoes exportam o conteudo para levar a reuniao de resultados, no mesmo modelo visual de um arquivo de referencia fornecido para essa reuniao:"
Bul "Exportar PDF: abre a caixa de impressao do navegador com o conteudo da pagina pronto para salvar como PDF - mesmo mecanismo `Imprimir` ja usado nos demais relatorios (esconde a barra de autenticacao e os botoes antes de imprimir)."
Bul "Gerar Apresentacao (PPTX): monta, inteiramente no navegador, uma apresentacao de 3 slides - capa, os indicadores estrategicos (com meta, barra de progresso, selo de status e tendencia) e o Plano de Acao - usando os mesmos numeros exibidos na tela naquele momento. Uma biblioteca de geracao e carregada sob demanda na primeira vez que o botao e usado (pode levar alguns segundos)."
Img "54_relatorio_resultados_botoes.png" "Botoes Exportar PDF e Gerar Apresentacao (PPTX) no topo do Relatorio de Resultados." 5.8
Img "55_pptx_capa_gerada.png" "Slide de capa da apresentacao PPTX gerada, com a identidade visual do sistema." 5.8

# ============================================================
# 7. PERGUNTAS FREQUENTES
# ============================================================
H1 "7. Perguntas frequentes"
H3 "Como corrijo um registro que ja salvei?"
P "Va ate a aba de registros cadastrados do formulario correspondente, clique na linha da tabela para abrir o painel de detalhes e clique em `Editar dados` (ou altere o status pelo seletor no rodape). Para os relatorios SGP-11 e SGP-12, volte para `Editar dados`, ajuste os campos e clique novamente em `Salvar e ver painel`."
H3 "Como excluo um registro?"
P "Abra o registro na tela de detalhes e clique em `Excluir`, no rodape. A acao pede confirmacao e nao pode ser desfeita depois de confirmada."
H3 "Por que o botao de registrar/salvar esta desabilitado?"
P "Nos 7 formularios com restricao por equipe (secao 2.5), o botao fica desabilitado se voce nao for membro da equipe do projeto selecionado - ou nao tiver selecionado um projeto vinculado ainda. Peca a um Admin para adiciona-lo a equipe em Administracao > Equipes."
H3 "Por que nao consigo abrir um formulario?"
P "Se aparecer a tela de login (secao 1.3), voce precisa se autenticar. Se depois de logado aparecer `Acesso restrito`, essa pagina (Administracao) e exclusiva para o papel Admin."
H3 "Os dados ficam salvos onde?"
P "Em um banco de dados real (Supabase/PostgreSQL), associados a sua conta. Nao e necessario salvar manualmente em outro lugar, mas recomenda-se exportar periodicamente em CSV os registros importantes, como backup."
H3 "Preciso preencher tudo na mesma sessao?"
P "Nao. Cada formulario pode ser preenchido e salvo em etapas diferentes - o registro fica disponivel na aba de consulta para ser retomado quando for necessario."

# ============================================================
# 8. ANEXO
# ============================================================
H1 "8. Anexo - Tabela de artefatos"
$r8 = @(
  @("#","Artefato","Codigo","Diretriz","Equipe?"),
  @("1","Solicitacao de Demanda","SGP-01","D01.1","Nao"),
  @("2","Canvas de Projeto","SGP-02","D01.6","Sim"),
  @("3","TAP - Termo de Abertura","SGP-03","D02.1","Sim"),
  @("4","Planejamento e Desenvolvimento","SGP-04","D02.2-D02.10","Sim"),
  @("5","EAP - Estrutura Analitica","(sem codigo)","D02.9","Sim"),
  @("6","Ata de Reuniao","SGP-07","D01-D07 (transversal)","Nao"),
  @("7","SMP - Solicitacao de Mudanca","SGP-06","D04.4","Sim"),
  @("8","TEP - Termo de Encerramento","SGP-09","D05.1.9","Sim"),
  @("9","RLA - Licoes Aprendidas","SGP-10","D05.1.6","Sim"),
  @("10","Plano de Comunicacao","SGP-08","D03.5","Nao"),
  @("11","Relatorio de Situacao","SGP-11","D04.3 / D06.5-D06.7","Nao"),
  @("12","Relatorio de Entregas e Beneficios","SGP-12","D06.1/D06.2","Nao")
)
TableSimple $r8 @(1.0,5.2,3.2,4.6,2.0)
P "Documento gerado a partir do estado atual do codigo do sistema, com capturas de tela reais coletadas em 28 e 29/07/2026." 9 $false $true $colMuted "left" 0

# ============================================================
# FINALIZAR
# ============================================================
$doc.TablesOfContents.Item(1).Update()
$doc.Fields.Update()

if (Test-Path $OutPath) { Remove-Item $OutPath -Force }
$doc.SaveAs2($OutPath, 16)
$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output "SAVED: $OutPath"
