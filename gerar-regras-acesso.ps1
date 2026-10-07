param([string]$OutPath = (Join-Path $PSScriptRoot "Regras de Acesso e Permissoes - SGP.docx"))
$ErrorActionPreference = "Stop"

# Gera "Regras de Acesso e Permissoes - SGP.docx" via automacao COM do Microsoft Word
# (Node.js/pandoc/LibreOffice nao estao disponiveis neste ambiente).
# Ver CLAUDE.md - secao "Documentacao oficial de regras de acesso" para quando reexecutar este script.

function RGB($r,$g,$b) { return [int]($r + ($g*256) + ($b*65536)) }
$colRed   = RGB 0xB9 0x1D 0x2E
$colInk   = RGB 0x1A 0x1A 0x1A
$colMuted = RGB 0x6A 0x6A 0x70
$colWhite = RGB 0xFF 0xFF 0xFF

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()
$sel = $word.Selection

function P($text, $size=11, $bold=$false, $italic=$false, $color=$colInk, $align="left", $spaceAfter=8) {
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
  $sel.Font.Size = 16
  $sel.Font.Bold = $true
  $sel.Font.Color = $colInk
  $sel.ParagraphFormat.Alignment = 0
  $sel.ParagraphFormat.SpaceBefore = 18
  $sel.ParagraphFormat.SpaceAfter = 8
  $sel.ParagraphFormat.Borders.Item(3).LineStyle = 1  # bottom border (wdBorderBottom=3, wdLineStyleSingle=1)
  $sel.ParagraphFormat.Borders.Item(3).Color = RGB 0xE4 0xE4 0xE7
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Borders.Item(3).LineStyle = 0
  $sel.Font.Bold = $false
}

function H2($text) {
  $sel.Font.Size = 13
  $sel.Font.Bold = $true
  $sel.Font.Color = $colRed
  $sel.ParagraphFormat.Alignment = 0
  $sel.ParagraphFormat.SpaceBefore = 12
  $sel.ParagraphFormat.SpaceAfter = 6
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Font.Bold = $false
  $sel.Font.Color = $colInk
}

function Bul($text, $level=0) {
  $sel.Font.Size = 11
  $sel.Font.Bold = $false
  $sel.Font.Color = $colInk
  $sel.ParagraphFormat.SpaceAfter = 4
  $sel.ParagraphFormat.LineSpacing = 13
  if ($level -eq 0) { $sel.ParagraphFormat.LeftIndent = $word.CentimetersToPoints(0.6) }
  else { $sel.ParagraphFormat.LeftIndent = $word.CentimetersToPoints(1.2) }
  $sel.Range.ListFormat.ApplyBulletDefault()
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Range.ListFormat.RemoveNumbers()
  $sel.ParagraphFormat.LeftIndent = 0
}

function HR() {
  $sel.ParagraphFormat.SpaceBefore = 6
  $sel.ParagraphFormat.SpaceAfter = 10
  $sel.ParagraphFormat.Borders.Item(3).LineStyle = 1
  $sel.ParagraphFormat.Borders.Item(3).Color = RGB 0xD1 0xD5 0xDB
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Borders.Item(3).LineStyle = 0
}

# ---- Capa ----
P "SGP - SISTEMA DE GESTAO DE PROJETOS" 10 $true $false $colRed "left" 4
$sel.Font.Size = 26; $sel.Font.Bold = $true; $sel.Font.Color = $colInk
$sel.ParagraphFormat.SpaceAfter = 4
$sel.TypeText("Regras de Acesso e Permissoes")
$sel.TypeParagraph()
$sel.Font.Bold = $false
P "Sistema de Gestao de Projetos" 14 $false $true $colMuted "left" 20

P "Este documento descreve, de forma completa, todas as regras de acesso e permissao implementadas no Sistema de Gestao de Projetos (SGP) (login, acesso sem login, papeis de usuario, gates de aprovacao e restricao por equipe de projeto), conforme o estado atual do codigo. Este e o documento oficial de referencia: qualquer inclusao, alteracao ou remocao de regra de acesso no sistema deve ser refletida aqui." 11 $false $false $colInk "left" 16

HR

# ---- 1 ----
H1 "1. Login obrigatorio (regra geral)"
P "A maioria das 15 paginas do sistema exige login antes de carregar ou salvar qualquer dado. O login e feito sem senha:"
Bul "Link magico por e-mail - o usuario informa o e-mail e recebe um link de acesso, sem senha."
P "As paginas abertas sem exigir login sao a pagina inicial (index.html, mapa de diretrizes) e o Validador de Projetos (ferramenta de apoio a decisao) - ambas mostram o conteudo livremente e so exibem a barra `"Conectado como...`" caso ja exista uma sessao ativa. Diferente da Solicitacao de Demanda e da Ata de Reuniao (secao 2), o acesso sem login do Validador nao depende de nenhuma ativacao pelo Admin - e sempre aberto."

HR

# ---- 2 ----
H1 "2. Acesso sem login (convidado)"
P "Dois formularios oferecem uma opcao de envio sem necessidade de login, para facilitar o registro por pessoas que nao tem (ou nao querem usar) uma conta institucional:"
Bul "Solicitacao de Demanda"
Bul "Ata de Reuniao"
P "Nenhum outro formulario do sistema tem essa opcao."

H2 "2.1 Como e ativado"
P "Cada um dos dois formularios tem um interruptor independente, controlado exclusivamente pelo Admin, na aba `"Configuracoes`" da pagina de Administracao. Ou seja, e possivel ativar o `"sem login`" so na Solicitacao de Demanda, so na Ata, nas duas, ou em nenhuma - sao chaves separadas."

H2 "2.2 O que o convidado pode fazer"
P "Quando a opcao esta ativada, aparece o botao `"Continuar sem login`" na tela de entrada. A pessoa informa nome completo e e-mail e pode:"
Bul "Preencher e enviar um registro novo (uma nova solicitacao ou uma nova ata)."

H2 "2.3 O que o convidado NAO pode fazer"
P "O acesso sem login e somente para criacao. Um usuario sem login nao consegue, em nenhuma hipotese:"
Bul "Editar um registro ja existente - inclusive um que ele mesmo tenha criado."
Bul "Excluir um registro existente."
Bul "Alterar o status de um registro (ex.: aprovar, mudar etapa) - aplicavel a Ata de Reuniao."
P "Essas acoes ficam bloqueadas de duas formas: os botoes `"Editar dados`" e `"Excluir`" somem da tela para quem esta sem login, e, mesmo que a acao seja tentada por outro caminho, o sistema recusa e mostra um aviso explicando que so e permitido criar novos registros."

H2 "2.4 Identificacao do registro"
P "Todo registro criado sem login recebe um selo `"Sem login`" na listagem e no detalhe, junto com o nome e e-mail informados pela pessoa. Se depois um usuario autenticado normalmente abrir esse mesmo registro e salvar uma edicao, o selo `"Sem login`" e removido - o registro passa a valer como editado por um usuario identificado."

H2 "2.5 Caso particular: Validador de Projetos"
P "O Validador de Projetos e mais aberto que os dois formularios acima: o quadro de conexoes, o simulador `"e se?`" e o assistente de decisao (8 perguntas, com veredito) funcionam por inteiro sem nenhum login - nao ha selo, nao ha admin para ativar, e nao ha bloqueio de nenhuma acao dentro da propria ferramenta."
Bul "Login so e pedido para uma funcionalidade especifica: vincular a avaliacao a um `"Projeto vinculado`" e salvar o veredito no historico daquele projeto."
Bul "Sem login, essa area do painel mostra um aviso com um botao `"Entrar`" - a pessoa pode logar a qualquer momento sem perder as respostas ja dadas no assistente."
Bul "Depois de logada, a pessoa ve o campo de projeto normalmente, como qualquer outro usuario autenticado."

HR

# ---- 3 ----
H1 "3. Papeis de usuario"
P "Cada pessoa que faz login recebe um papel, usado para liberar ou restringir acoes especificas no sistema:"
Bul "Solicitante - papel padrao, atribuido automaticamente a todo novo usuario no primeiro login."
Bul "Gerente de Projetos"
Bul "Gestor Responsavel"
Bul "Dono do Negocio"
Bul "Alta Gestao"
Bul "Admin - papel de administracao do sistema, atribuido manualmente por quem ja e Admin."

H2 "3.1 Acoes exclusivas de Admin"
P "Somente usuarios com papel Admin podem:"
Bul "Acessar a pagina de Administracao (sempre exclusiva a Admin) e definir quais papeis podem ver o Painel Executivo, o Painel de Prazos, o Relatorio de Resultados, o Relatorio de Situacao e o Relatorio de Entregas (ver 3.4)."
Bul "Alterar o papel de outros usuarios."
Bul "Adicionar ou remover membros da equipe de um projeto."
Bul "Ativar/desativar as opcoes de `"sem login`" (Solicitacao de Demanda e Ata de Reuniao)."
Bul "Aprovar ou reprovar o Gate 1 na Solicitacao de Demanda (ver secao 4)."
Bul "Pactuar ou reabrir o Gate 2 no Relatorio de Entregas e Beneficios (ver secao 4)."
Bul "Pre-cadastrar uma pessoa que ainda nao fez login, informando nome, telefone, e-mail e papel (ver 3.2)."
Bul "Remover da lista um usuario que nunca teve nenhuma relacao com projetos nem registros no sistema (ver 3.3)."
Bul "Ligar/desligar a importacao de transcricao por IA na Ata de Reuniao e definir quais papeis podem usa-la (ver 3.5)."
Bul "Ligar/desligar, separadamente, a importacao de audio (gravacao da reuniao) por IA na Ata de Reuniao e definir quais papeis podem usa-la (ver 3.6)."
Bul "Resetar e repopular o ambiente de treinamento com os 7 projetos de exemplo, pelo botao na aba Administracao (ver 3.8)."
P "Importante: um Admin sempre e considerado `"membro`" de qualquer equipe de projeto automaticamente - nao precisa ser adicionado manualmente para poder editar registros vinculados a um projeto (ver secao 5)."

H2 "3.2 Adicionar usuario antes do primeiro login (conta criada na hora)"
P "Normalmente uma pessoa so apareceria na aba `"Usuarios`" da Administracao depois de fazer login pela primeira vez (o perfil seria criado automaticamente, com papel `"Solicitante`"). `"+ Adicionar usuario`" permite ao Admin adiantar isso, criando a conta de verdade na hora:"
Bul "Na aba Usuarios, o Admin preenche nome, telefone (opcional), e-mail e papel e clica em `"+ Adicionar usuario`"."
Bul "O sistema cria a conta de verdade no Supabase Auth (via Admin API, sem exigir senha nem confirmacao por e-mail) e o perfil correspondente (nome, telefone, papel) na hora - a pessoa ja aparece na lista de Usuarios como conta ativa, e ja pode ser adicionada a equipe de qualquer projeto (secao 5) sem esperar ela logar."
Bul "Quando essa pessoa eventualmente fizer o primeiro login (link magico), o Supabase Auth reconhece o e-mail e autentica direto na conta ja existente - nao cria uma conta duplicada, e o perfil definido pelo Admin permanece do jeito que estava."
P "Controle de acesso: a criacao de conta e feita por uma Edge Function (`criar-usuario-admin`) que confere, a cada chamada, que quem esta chamando tem papel Admin em `public.perfis` - a mesma trava ja usada em outras acoes sensiveis exclusivas de Admin (ex.: reset do ambiente de treinamento, secao 3.8). A function usa a chave de servico do Supabase (nunca exposta no navegador) para criar a conta e gravar o perfil."
Bul "Pre-cadastros antigos (feitos antes de 15/09/2026, tabela `perfis_pendentes`) continuam funcionando: a pessoa aparece com o selo `"Pendente - 1o login`" ate fazer login ou ate o Admin usar o novo botao `"Criar conta agora`" para converte-la em conta de verdade imediatamente, pelo mesmo mecanismo acima."

H2 "3.3 Remocao de usuario"
P "Na aba Usuarios, cada pessoa ja cadastrada (exceto o proprio Admin logado) tem um botao `"Remover`". Ao clicar, o sistema verifica automaticamente, na hora, se a pessoa tem algum vinculo:"
Bul "E membro da equipe de algum projeto (secao 5)."
Bul "E quem criou algum projeto (a partir de uma Solicitacao de Demanda aprovada)."
Bul "Tem algum registro criado em seu nome em um dos 9 formularios que gravam quem criou o registro: Solicitacao de Demanda, Canvas, TAP, Planejamento e Desenvolvimento, EAP, SMP, Ata de Reuniao, TEP e RLA."
P "Se qualquer um desses vinculos existir, a remocao e bloqueada e o Admin ve a lista dos motivos especificos. So quando a pessoa nao tem nenhum vinculo o sistema pede confirmacao final antes de remover."
P "Limitacao conhecida: os tres formularios que sao documentos unicos e compartilhados de todo o sistema - Plano de Comunicacao de Projeto, Relatorio de Situacao de Projetos e Relatorio de Entregas e Beneficios - nao tem o conceito de `"quem criou o registro`" (sao editados coletivamente, sem dono individual), entao nao entram nessa checagem automatica."
P "Controle de acesso (Row Level Security no Supabase): remover um perfil e uma acao restrita por politica `"perfis_delete_admin`" a usuarios com papel Admin - qualquer tentativa de exclusao por outro papel e recusada pelo proprio banco de dados, mesmo que alguem tente contornar a interface."
P "Remover aqui apaga apenas o perfil (papel, nome, telefone) do sistema de gestao de projetos - a pessoa continua podendo fazer login depois; se fizer, um perfil novo e criado do zero, com o papel padrao Solicitante."

H2 "3.4 Restricao por papel no Painel Executivo, no Painel de Prazos, no Relatorio de Resultados e nos dois relatorios finais"
P "Cinco paginas tem sua visibilidade controlada por uma lista de papeis, definida pelo Admin na aba Configuracoes da Administracao: Painel Executivo, Painel de Prazos, Relatorio de Resultados, Relatorio de Situacao de Projetos e Relatorio de Entregas e Beneficios. Para cada uma dessas paginas, o Admin marca (com caixas de selecao, podendo marcar mais de um papel) quais papeis conseguem abri-la; quem tem um papel fora da lista ve a mensagem `"Acesso restrito`" ao tentar acessar."
Bul "Painel Executivo, Painel de Prazos e Relatorio de Resultados comecam configurados apenas para Admin, mantendo o padrao mais restritivo das paginas de gestao."
Bul "Relatorio de Situacao e Relatorio de Entregas comecam configurados para todos os papeis, mantendo o comportamento original (acesso livre a qualquer usuario autenticado) ate que o Admin decida restringir."
P "Importante: um usuario com papel Admin sempre consegue acessar essas cinco paginas, mesmo que o papel Admin seja removido da lista por engano - essa trava evita que o proprio Admin fique bloqueado sem ter como reverter a configuracao."

H2 "3.5 Importacao de transcricao por IA (Ata de Reuniao)"
P "A Ata de Reuniao pode preencher automaticamente pauta, participantes, resumo, encaminhamentos e entraves a partir de uma transcricao de reuniao colada pelo usuario, analisada por IA (Claude, via uma Supabase Edge Function - a chave da API fica guardada no servidor, nunca exposta no navegador)."
Bul "Existe um interruptor geral em Administracao > Configuracoes, desligado por padrao. Diferente da restricao por papel da secao 3.4, esse interruptor vale para todo mundo sem excecao, inclusive Admin - desligado, ninguem consegue usar a importacao."
Bul "Com o interruptor ligado, uma lista de papeis (podendo marcar mais de um) define quem pode usar a importacao - Admin sempre tem acesso quando o interruptor esta ligado, mesmo sem estar marcado na lista."
Bul "Quem nao tem permissao simplesmente nao ve a opcao de importar transcricao no formulario."
P "A Edge Function que processa a transcricao confere essa mesma regra direto no banco de dados antes de chamar a IA - nao basta esconder o botao na tela, a chamada e recusada no servidor para quem nao tem permissao, mesmo que tentada por fora da interface."

H2 "3.6 Importacao de audio por IA (Ata de Reuniao)"
P "Alem de colar/anexar a transcricao em texto (secao 3.5), a Ata de Reuniao pode transcrever a gravacao de audio da reuniao (.mp3, .m4a, .aac, .wav, .ogg e outros formatos comuns) usando um segundo servico de IA (OpenAI, tambem via uma Supabase Edge Function separada, com sua propria chave guardada no servidor). O audio e dividido automaticamente em pedacos de alguns minutos no proprio navegador antes do envio, por causa de limites tecnicos da API de transcricao - o usuario nao precisa fazer nada manualmente para isso."
Bul "Regra de acesso INDEPENDENTE da importacao de transcricao em texto (secao 3.5): interruptor proprio em Administracao > Configuracoes, desligado por padrao, valendo para todo mundo sem excecao, inclusive Admin."
Bul "Com o interruptor de audio ligado, uma lista de papeis propria (podendo marcar mais de um) define quem pode usar a importacao de audio - Admin sempre tem acesso quando ligado, mesmo sem estar marcado na lista."
Bul "Na pratica, importar audio so faz sentido para quem tambem pode rodar a analise da transcricao (secao 3.5) - por isso a opcao de anexar audio so aparece para quem tem as duas permissoes ativas ao mesmo tempo."
P "Assim como a importacao de texto, a Edge Function que processa o audio confere a permissao direto no banco antes de chamar a IA, nao so a interface."

H2 "3.7 Restricao de edicao no Plano de Comunicacao de Projeto"
P "O Plano de Comunicacao de Projeto (SGP-08) e um documento unico e compartilhado (secao 5.3) - qualquer usuario autenticado sempre pode abrir a aba Painel, visualizar o conteudo e usar o botao Imprimir. Diferente da secao 3.4 (que restringe visualizacao de pagina inteira), aqui a restricao e so sobre a aba `"Editar dados`": uma lista de papeis, definida pelo Admin na aba Configuracoes da Administracao, decide quem pode usar `"Editar dados`" e salvar alteracoes."
Bul "Comeca configurado apenas para PMO/Admin - os demais papeis veem a aba `"Editar dados`" escondida e uma nota explicando a restricao no lugar dela, mas continuam vendo o Painel normalmente."
Bul "Um usuario com papel Admin sempre consegue editar, mesmo que o papel Admin seja removido da lista por engano - mesma trava de seguranca usada na secao 3.4."
Bul "A checagem tambem bloqueia a funcao de salvar caso alguem tente forcar a aba de edicao por fora da interface."

H2 "3.8 Reset do ambiente de treinamento (aba Administracao, so em producao)"
P "O sistema tem um ambiente de treinamento/demonstracao totalmente separado (banco de dados proprio, mesmo schema e regras de acesso da producao, com 7 projetos de exemplo ficticios cobrindo o ciclo completo), usado para apresentacoes e capacitacao sem tocar em dado real. A aba Administracao, quando acessada em producao (`"paramiri.github.io/unialfa-gestao-projetos`"), mostra uma aba extra `"Ambiente de Treino`" com um botao que apaga os dados atuais do ambiente de treinamento e recria os 7 projetos de exemplo do zero - util para deixar o ambiente limpo antes de uma nova turma."
Bul "A aba so aparece em producao - quem acessa a Administracao pelo proprio ambiente de treinamento nao ve essa opcao."
Bul "A acao e restrita a Admin: o botao chama uma Edge Function que confere, direto no banco de producao (nao so na tela), que quem chamou tem papel Admin antes de fazer qualquer alteracao - inclusive rejeitando um token de sessao valido cujo dono nao seja Admin."
Bul "Nao apaga nem altera nenhum dado de producao em nenhuma hipotese - a Edge Function so tem permissao de escrita no banco do ambiente de treinamento, nunca no de producao, e se recusa a rodar caso seja implantada por engano no projeto errado."
Bul "As 6 contas fixas de treinamento e as configuracoes do ambiente de treino nao sao apagadas pelo reset - so os projetos e os registros de formulario."

H2 "3.9 Assistente de preenchimento por IA (Canvas, TAP, Planejamento, SMP, TEP, RLA, Relatorio de Situacao e Relatorio de Entregas)"
P "O Canvas de Projeto, o TAP (Termo de Abertura de Projeto), o Planejamento e Desenvolvimento de Projeto, a SMP, o TEP e o RLA tem um botao `"Sugerir com IA`", que aparece ao lado do seletor de projeto vinculado depois que um projeto e escolhido. Em cada formulario, o botao le os documentos anteriores ja registrados daquele mesmo projeto - seguindo a esteira (Demanda alimenta o Canvas; Canvas e Demanda alimentam o TAP; TAP e Canvas alimentam o Planejamento; TAP e Planejamento alimentam a SMP; TAP, Planejamento e EAP alimentam o TEP; TEP e SMPs alimentam o RLA), sempre incluindo todas as Atas de Reuniao vinculadas ao projeto - e usa a API da Claude (Anthropic) para sugerir o preenchimento dos campos ainda vazios. A sugestao nunca salva nada sozinha - o usuario sempre revisa e confirma antes de usar `"Registrar`"/`"Salvar alteracoes`", exatamente como ja funciona na importacao de transcricao da Ata (secao 3.5)."
Bul "Mesma logica de permissao das secoes 3.5/3.6: interruptor mestre em Administracao > Configuracoes, desligado por padrao, valendo para todo mundo sem excecao, inclusive Admin - um unico interruptor cobre os 6 formularios."
Bul "Com o interruptor ligado, uma lista de papeis propria define quem pode usar o assistente - Admin sempre tem acesso, mesmo sem estar marcado na lista."
Bul "Alem da permissao de papel, so quem faz parte da equipe do projeto selecionado (secao 5) pode acionar o botao - a Edge Function confere isso direto no banco antes de gastar uma chamada de API, para ninguem sugerir preenchimento de projeto alheio."
Bul "So preenche campos vazios: um campo que o usuario ja digitou nunca e sobrescrito pela sugestao. Cada campo preenchido pela IA fica marcado com um selo `"IA`" ate ser editado, e uma linha no rodape lista os documentos usados como base da sugestao."
Bul "A EAP nao tem mais este assistente por IA - o botao `"Sugerir com IA`" de la foi substituido por `"Importar do Cronograma`", uma copia mecanica dos nomes dos marcos do Planejamento, sem chamada de IA e sem a restricao de interruptor/papel desta secao (exige so fazer parte da equipe do projeto, igual as demais acoes da EAP)."
Bul "Na SMP, a logica e diferente das demais: o TAP e o Planejamento so mostram o que foi combinado originalmente, entao a IA procura nas Atas mais recentes alguma mudanca de fato sendo discutida e contrasta com esse combinado original. Se as Atas nao discutirem nenhuma mudanca concreta, a sugestao devolve todos os campos vazios em vez de inventar uma mudanca hipotetica - e nunca copia os campos do TAP/Planejamento diretamente, ja que a SMP descreve o que MUDA, nao o que ja estava definido."
Bul "No TEP, a IA compara o planejado (TAP, Planejamento, EAP) com o que as Atas mais recentes confirmam como realizado, para redigir a justificativa de encerramento e as consideracoes finais. No RLA, a IA le o TEP, as Atas e todas as SMPs do projeto para preencher as 4 secoes de texto livre (Visao geral, Destaques, Desafios, Tarefas pos-projeto) - os blocos de avaliacao estruturada Sim/Nao/Parcial com placar percentual (secoes 6 a 9 do formulario) nao sao preenchidos pela IA, continuam sendo autoavaliacao manual da equipe."
Bul "No Relatorio de Situacao e no Relatorio de Entregas, a logica de acesso e diferente dos outros 7 formularios: como um unico relatorio cobre varios projetos ao mesmo tempo (um por linha, via `"Projeto vinculado (Gate 1)`"), o botao `"Sugerir com IA`" aparece por linha, nao uma vez so para o formulario inteiro - e a checagem de equipe (acima) e feita contra o projeto daquela linha especifica, nao contra o relatorio em si, que continua sem restricao por equipe (secao 3.4)."
Bul "Disponivel apenas em producao - o ambiente de treinamento nao recebe a chave de IA nem a Edge Function correspondente, de proposito, para nunca gerar cobranca real ao demonstrar o sistema."

H2 "3.10 Importacao de documento por IA (Solicitacao de Demanda)"
P "A Solicitacao de Demanda (SGP-01) tem um painel `"Importar documento preenchido`", no topo do formulario, que aceita colar texto ou anexar um arquivo (.txt, .docx ou .pdf) do formulario oficial ja preenchido a mao fora do sistema. O texto e extraido no proprio navegador (sem passar pelo servidor) e analisado pela API da Claude (Anthropic), que sugere o preenchimento dos campos ainda vazios - nome do projeto, solicitante, departamento, justificativa, objetivo, escopo, prazo, orcamento, partes interessadas e anexos. Igual as demais importacoes por IA, a sugestao nunca salva nada sozinha."
Bul "Mesma logica de permissao das secoes 3.5/3.6/3.9: interruptor mestre em Administracao > Configuracoes, desligado por padrao, valendo para todo mundo sem excecao, inclusive Admin; com o interruptor ligado, uma lista de papeis propria define quem pode usar - Admin sempre tem acesso."
Bul "Diferente das demais secoes deste capitulo, aqui nao ha restricao por equipe de projeto - a Solicitacao de Demanda e o primeiro formulario da esteira, registrado antes de existir um projeto aprovado."
Bul "Acesso sem login (secao 2 do documento institucional) NAO tem acesso a este painel, mesmo com o interruptor ligado para os papeis normais - a importacao exige um perfil de usuario autenticado com papel definido."
Bul "Os campos Departamento e Prioridade so sao preenchidos automaticamente se o valor sugerido pela IA corresponder exatamente a uma das opcoes ja existentes na lista suspensa - qualquer valor fora da lista e descartado, para nunca deixar esses dois campos num estado invalido."
Bul "Disponivel apenas em producao, pela mesma razao das demais importacoes por IA."

H2 "3.11 Backup automatizado"
P "Na aba Configuracoes da Administracao, um bloco separado deixa o Admin/PMO definir os parametros de uma rotina de backup periodico do sistema (banco de dados, arquivos anexados e codigo-fonte) para um destino fora do GitHub: destino e caminho, frequencia, retencao, e-mail de alerta e e-mails com acesso ao destino externo (campo de registro, nao concede acesso de fato)."
Bul "Mesma logica de permissao das demais listas de papeis deste capitulo: uma lista de papeis propria (`"Quem pode gerenciar as configuracoes de backup`") define quem, alem do Admin, pode ver e alterar esses campos - Admin sempre tem acesso, mesmo sem estar marcado na lista."
Bul "Diferente das secoes 3.5 a 3.10, esta nao tem um interruptor mestre `"desligado por padrao`" que bloqueia todo mundo - o campo `"Habilitar backup automatizado`" e, ele mesmo, um dos parametros salvos, nao uma trava de acesso a tela: quem tem o papel liberado sempre ve e edita a tela, ligado ou desligado."
Bul "Esta tela define apenas os parametros - quem efetivamente executa a copia dos dados e uma rotina agendada fora do sistema (GitHub Actions), sem acesso direto ao sistema nem papel de usuario."

H2 "3.12 Meu Painel e o campo Responsavel (conta)"
P "`"Meu Painel`" e uma pagina de `"Gestao`" sem nenhuma restricao por papel: qualquer usuario autenticado, de qualquer papel, pode abri-la e ve apenas o que esta atribuido a propria conta (marcos e tarefas do cronograma, riscos do TAP, entradas do Planejamento, entraves e encaminhamentos do Relatorio de Situacao), com um botao para concluir cada item direto da lista. Diferente das cinco paginas da secao 3.4, nao ha lista de papeis para configurar aqui - o controle de acesso e so o login, igual a Ficha do Projeto."
Bul "Desde 16/09/2026, o campo `"Responsavel (conta)`" (nos tres formularios abaixo) aceita marcar mais de uma pessoa ao mesmo tempo - um painel com uma caixa de selecao por conta, em vez de um `<select>` de escolha unica. Nao muda o alcance de acesso descrito nesta secao, so quantas contas podem ficar atribuidas a um mesmo item; o item passa a aparecer no Meu Painel de cada conta marcada, nao so de uma."
Bul "Em Riscos (TAP) e Entradas (Planejamento), o seletor `"Responsavel (conta)`" continua restrito a quem faz parte da equipe do projeto selecionado (mesma fonte de dados da restricao por equipe, secao 5) - ninguem de fora da equipe aparece na lista."
Bul "Em Entraves e Encaminhamentos (Relatorio de Situacao), o seletor lista todas as contas do sistema, nao so uma equipe - esse relatorio nao tem um projeto vinculado unico ao qual amarrar uma equipe. Para viabilizar essa lista, uma funcao no banco de dados (`listar_emails_usuarios`, SECURITY DEFINER) devolve e-mail e nome de todas as contas cadastradas para qualquer usuario autenticado - uma excecao deliberada e estreita a regra geral de `perfis` (secao abaixo), que continua sem devolver telefone nem papel de outra pessoa."
Bul "O nome devolvido por `listar_emails_usuarios` e usado para mostrar, nos tres seletores acima, o nome cadastrado de cada conta em vez do e-mail cru (caindo para o e-mail so quando a conta nao tem nome cadastrado) - desde 17/09/2026. Antes dessa data a funcao devolvia so o e-mail; a extensao para incluir o nome foi deliberada, ja que o nome ja aparece como texto livre em varios outros lugares do sistema (Responsavel, Gerente do projeto, Solicitante) - telefone e papel continuam de fora, por serem os dados mais sensiveis do perfil."
P "Controle de acesso (Row Level Security no Supabase): a politica `perfis_select` normalmente so libera a propria linha de `perfis` para quem nao e Admin ((auth.uid() = id) OR is_admin()) - por isso um SELECT direto na tabela nao serve para montar a lista de Entraves/Encaminhamentos acima. A funcao `listar_emails_usuarios` contorna essa politica de proposito (SECURITY DEFINER), mas devolve deliberadamente so as colunas e-mail e nome, preservando a mesma protecao de privacidade (telefone/papel de outros usuarios) que ja existia antes."

H2 "3.13 Gestao de Atividades"
P "`"Gestao de Atividades`" e uma pagina de `"Gestao`" que agrega, por conta responsavel, os mesmos itens do Meu Painel (secao 3.12) - dando a quem tem acesso uma visao da carga de trabalho de toda a equipe, nao so da propria conta. Diferente do Meu Painel, esta pagina TEM restricao por papel, no mesmo esquema da secao 3.4: uma lista de papeis, definida pelo Admin na aba Configuracoes da Administracao (chave `papeis_gestao_atividades`), decide quem mais - alem do PMO/Admin, que sempre tem acesso - pode abrir a pagina. Comeca configurada so para Admin, igual ao Painel Executivo e ao Painel de Prazos."
P "Diferenca importante em relacao as outras listas de papeis da secao 3.4: aqui a permissao nao libera so uma tela, libera tambem ENXERGAR NOME E PAPEL DE TODOS OS USUARIOS DO SISTEMA - e o dado que a pagina usa para agrupar as pendencias por pessoa. Por isso essa checagem nao fica so na interface: uma funcao no banco de dados (`listar_perfis_gestao_atividades`, SECURITY DEFINER) confere, ela mesma, se quem chamou e Admin OU tem um papel presente em `papeis_gestao_atividades` antes de devolver qualquer linha - uma segunda funcao auxiliar, `pode_ver_gestao_atividades`, concentra essa mesma checagem para ser reaproveitada se novas funcoes precisarem dela no futuro. Sem essa dupla camada, bastaria alguem inspecionar o codigo da pagina para descobrir a chamada e contornar a tela, obtendo a lista de usuarios sem ter sido liberado para isso."
P "Os dados de atividade em si (marcos, riscos, entradas, entraves/encaminhamentos, guardados na tabela `kv_store`) ja sao de leitura ampla para qualquer usuario autenticado - e assim que o proprio Meu Painel consegue ler projetos que nao sao do usuario logado, para achar os itens atribuidos a ele em qualquer lugar do sistema. A parte nova e so a lista de nome/papel de todo mundo, que antes so o Admin conseguia ver."
P "De passagem, ao criar esta pagina foi encontrada e corrigida uma lacuna na configuracao do Painel de Prazos (secao 3.4): a linha `papeis_painel_prazos` nunca tinha sido inserida na tabela de configuracoes do sistema - os campos apareciam normalmente na tela da Administracao, mas o Admin nao conseguia de fato salvar uma mudanca ali, porque a gravacao so atualiza uma linha ja existente, nunca cria uma nova."

HR

# ---- 4 ----
H1 "4. Gates de aprovacao"
P "O sistema tem dois pontos de decisao formal (gates) no ciclo de vida de um projeto, ambos aparecendo no fluxo de diretrizes da pagina inicial:"

H2 "4.1 Gate 1 - Triagem"
P "Ocorre logo apos o registro da demanda (D01.4/D01.5). Decide se a demanda entra ou nao no portfolio de projetos."
Bul "Quem decide: Admin (atuando como Gestor Responsavel no sistema)."
Bul "Onde: no status da Solicitacao de Demanda - so o Admin consegue alterar o campo de status; qualquer outro papel ve o campo travado, com um aviso explicando que so o PMO/Admin pode aprovar ou reprovar."

H2 "4.2 Gate 2 - Pactuacao"
P "Ocorre ao final do Planejamento (D02), antes do inicio da Execucao (D03). E o gate mais critico: autoriza formalmente o inicio da execucao do projeto."
Bul "Quem decide: Dono do Negocio, perante a Alta Gestao."
Bul "Onde: no Relatorio de Entregas e Beneficios (SGP-12), aba Editar dados - campo Status (Pendente de pactuacao / Pactuado). So o Admin consegue alterar esse campo; qualquer outro papel ve o controle travado, com um aviso explicando que so o PMO/Admin pode pactuar."
Bul "Ao marcar `"Pactuado`", o sistema preenche automaticamente a Data de pactuacao (hoje) e o Aprovador (nome de quem esta logado), ambos editaveis pelo Admin."
Bul "Enquanto o status estiver `"Pactuado`", todos os demais campos do relatorio (ficha do programa, indicadores, projetos vinculados) ficam bloqueados para edicao - inclusive para o Admin - ate que o Gate 2 seja reaberto (status voltar para `"Pendente de pactuacao`")."

HR

# ---- 5 ----
H1 "5. Restricao por equipe do projeto"
P "Varios formularios exigem vincular o registro a um `"Projeto vinculado`" (um projeto ja Aprovado no Gate 1). Nesses formularios, apenas quem faz parte da equipe daquele projeto especifico - ou um Admin - pode criar ou editar um registro vinculado a ele."

H2 "5.1 Formularios com essa restricao"
P "Aplicada nos 7 formularios que usam o conceito de `"equipe por projeto`":"
Bul "Canvas de Projeto"
Bul "TAP - Termo de Abertura de Projeto"
Bul "Planejamento e Desenvolvimento de Projeto"
Bul "EAP - Estrutura Analitica de Projeto"
Bul "SMP - Solicitacao de Mudanca de Projeto"
Bul "TEP - Termo de Encerramento de Projeto"
Bul "RLA - Registro de Licoes Aprendidas"

H2 "5.2 Como funciona"
P "Ao selecionar um projeto no campo `"Projeto vinculado`":"
Bul "Se a pessoa nao for da equipe daquele projeto (e nao for Admin), aparece um aviso na tela e o botao de enviar/registrar fica desabilitado."
Bul "Mesmo que o botao seja habilitado por algum outro meio, o envio e bloqueado no momento de salvar, com a mensagem `"Voce nao faz parte da equipe deste projeto.`""
Bul "A equipe de cada projeto e definida pelo Admin, na pagina de Administracao, aba Equipes."
Bul "A mesma restricao vale para os botoes Imprimir e Exportar Word de um registro ja salvo, nesses 7 formularios (detalhado na secao 5.5)."

H2 "5.3 Formularios sem essa restricao"
P "Os demais formularios nao usam o conceito de equipe por projeto - qualquer usuario autenticado pode criar/editar/excluir registros neles, sujeito apenas as regras de papel e (quando aplicavel) as regras especificas ja descritas nas secoes 2 e 4:"
Bul "Solicitacao de Demanda (tem o Gate 1 e a opcao de sem login, descritos acima)."
Bul "Ata de Reuniao (tem a opcao de sem login, descrita acima)."
Bul "Plano de Comunicacao de Projeto"
Bul "Relatorio de Situacao de Projetos"
Bul "Relatorio de Entregas e Beneficios"
P "O Relatorio de Situacao e o Relatorio de Entregas tem, cada um, um campo opcional `"Projeto vinculado (Gate 1)`" em cada projeto listado - diferente do conceito desta secao, ele nao aplica nenhuma restricao de edicao por equipe. Serve apenas para ligar aquele item ao historico compartilhado do projeto (a mesma trilha de auditoria usada pelos 7 formularios com restricao por equipe), visivel pelo botao `"Ver historico`". Deixar sem selecionar mantem o comportamento anterior: qualquer usuario autenticado continua podendo criar/editar esses registros livremente."

H2 "5.4 Quem pode editar ou excluir um registro ja criado"
P "Alem da restricao por equipe (quem pode criar um registro vinculado a um projeto), o sistema tambem restringe quem pode editar ou excluir um registro depois de criado, nos 9 formularios que tem essa acao: Solicitacao de Demanda, Canvas de Projeto, TAP, Planejamento e Desenvolvimento, EAP, SMP, Ata de Reuniao, TEP e RLA."
Bul "Editar: apenas quem criou o registro, quem tem o papel Gerente de Projetos, ou o Admin."
Bul "Excluir: apenas quem tem o papel Gerente de Projetos, ou o Admin."
Bul "O papel Gerente de Projetos e um atributo global da conta (Administracao > Usuarios), nao um vinculo especifico a um projeto - qualquer conta com esse papel pode editar/excluir registros de qualquer projeto, nao so dos que participa. Nao existe hoje um vinculo estruturado de `"GP responsavel por este projeto`" separado do papel da conta."
Bul "Os botoes `"Editar dados`" e `"Excluir`" ficam ocultos para quem nao tem permissao, em vez de aparecerem e barrarem a acao so no clique."
Bul "Nao se aplica ao Plano de Comunicacao de Projeto (secao 3.7 ja cobre a edicao desse formulario separadamente) nem ao Relatorio de Situacao/Relatorio de Entregas, que sao documentos unicos por projeto sem o conceito de `"registro individual criado por alguem`"."
P "Desde 15/09/2026, essa regra tambem e aplicada diretamente no banco de dados (politicas de RLS na tabela kv_store), nao so no JavaScript do navegador - alguem tentando editar ou excluir um registro sem permissao direto pela API, contornando a tela do sistema, tambem e bloqueado. Fica documentado no arquivo `supabase/migrations/20260915143738_kv_store_row_level_access_control.sql` do repositorio. Duas coisas ficaram de fora dessa camada de banco, so continuam no JavaScript: a decisao especifica de quem aprova o Gate 1/Gate 2 (mais restrita que a regra geral de edicao) e as regras dos 3 relatorios de chave unica (secao 3.7 e o Gate 2 do Relatorio de Entregas)."

H2 "5.5 Quem pode imprimir ou exportar um registro"
P "Nos mesmos 7 formularios com restricao por equipe (secao 5.1), os botoes Imprimir e Exportar Word de um registro ja salvo tambem passam a exigir fazer parte da equipe daquele projeto, ou ser Admin."
Bul "Diferente da regra de editar/excluir (secao 5.4), aqui nao ha excecao para o papel Gerente de Projetos - precisa realmente estar na equipe do projeto (ou ser Admin) para imprimir ou exportar um registro."
Bul "Os botoes `"Imprimir`" e `"Exportar Word`" ficam ocultos para quem nao faz parte da equipe, da mesma forma que os botoes de editar/excluir."
Bul "Nao se aplica aos outros 5 formularios sem restricao por equipe (secao 5.3) nem aos relatorios agregados - continuam sem restricao de impressao/exportacao por equipe."
P "Essa regra e aplicada apenas no JavaScript do navegador, nao no banco de dados - diferente da trava de editar/excluir (secao 5.4), que desde 15/09/2026 tambem e reforcada por politicas de RLS na tabela kv_store."

HR

# ---- 6: tabela resumo ----
H1 "6. Resumo por formulario"

$rows = @(
  @("Formulario","Requer login","Restrito por equipe","Aprovacao especial"),
  @("Solicitacao de Demanda","Sim (ou sem login, se ativado)","Nao","Gate 1 (Admin)"),
  @("Canvas de Projeto","Sim","Sim","-"),
  @("TAP","Sim","Sim","-"),
  @("Planejamento e Desenvolvimento","Sim","Sim","-"),
  @("EAP","Sim","Sim","-"),
  @("SMP","Sim","Sim","-"),
  @("Ata de Reuniao","Sim (ou sem login, se ativado)","Nao","-"),
  @("Plano de Comunicacao","Sim","Nao","Edicao por papel (padrao Admin)"),
  @("TEP","Sim","Sim","-"),
  @("RLA","Sim","Sim","-"),
  @("Relatorio de Situacao","Sim (por papel)","Nao","-"),
  @("Relatorio de Entregas","Sim (por papel)","Nao","Gate 2 (Admin)"),
  @("Administracao","Sim (so Admin acessa)","-","-"),
  @("Painel Executivo","Sim (por papel, padrao Admin)","-","-"),
  @("Painel de Prazos","Sim (por papel, padrao Admin)","-","-"),
  @("Relatorio de Resultados","Sim (por papel, padrao Admin)","-","-"),
  @("Validador de Projetos","Nao (so p/ vincular projeto)","Nao","-")
)

$nRows = $rows.Count
$nCols = 4
$tableRange = $sel.Range
$table = $doc.Tables.Add($tableRange, $nRows, $nCols)
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
$table.Columns.Item(1).Width = $word.CentimetersToPoints(5.2)
$table.Columns.Item(2).Width = $word.CentimetersToPoints(4.2)
$table.Columns.Item(3).Width = $word.CentimetersToPoints(3.6)
$table.Columns.Item(4).Width = $word.CentimetersToPoints(3.0)

$sel.EndKey(6) | Out-Null  # wdStory
$sel.TypeParagraph()
P "Documento gerado a partir do estado atual do codigo do sistema." 9 $false $true $colMuted "left" 0

if (Test-Path $OutPath) { Remove-Item $OutPath -Force }
$doc.SaveAs2($OutPath, 16)
$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output "SAVED: $OutPath"
