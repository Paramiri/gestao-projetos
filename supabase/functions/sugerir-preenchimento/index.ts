// Edge Function: sugerir-preenchimento
// Le os documentos ja registrados de um projeto (Solicitacao de Demanda,
// Canvas de Projeto, TAP, Planejamento, EAP, SMP, TEP e Atas de Reuniao) e
// usa a API da Claude (Anthropic) para sugerir o preenchimento dos campos
// do proximo formulario da esteira (Fase 1: Canvas e TAP; Fase 2:
// Planejamento; Fase 3: SMP, TEP e RLA; Fase B dos relatorios de
// portfolio: Relatorio de Situacao e Relatorio de Entregas, sugeridos POR
// LINHA de projeto, ja que um unico relatorio cobre varios projetos). O
// usuario sempre revisa
// antes de salvar - esta funcao so sugere, nunca salva nada sozinha. Mesma
// arquitetura da funcao ja em producao `analisar-transcricao-ata`,
// generalizada para ler o historico do projeto em vez de uma transcricao
// colada.
//
// A EAP (Passo 5) NAO usa mais este assistente - o "Sugerir com IA" de la foi
// substituido por uma importacao mecanica dos marcos do Cronograma do
// Planejamento (funcao `importarDoCronograma` em 05, sem chamada de IA),
// porque o casamento Entrega x Marco na Ficha do Projeto exige nome IDENTICO
// e uma sugestao de IA reformulava o texto, quebrando esse casamento.
//
// Segredo necessario (ja configurado nesta base para analisar-transcricao-ata):
//   ANTHROPIC_API_KEY - chave da API da Claude (console.anthropic.com)
// SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY sao injetados automaticamente.
//
// Verificacao de JWT permanece ligada. Alem disso, esta funcao confere no
// banco: (1) interruptor mestre + papel permitido em `configuracoes`
// (Admin sempre passa), e (2) se quem chamou faz parte da equipe do
// projeto informado (`projeto_equipe`) - para que ninguem gaste uma
// chamada de API (custo real) sugerindo preenchimento de projeto alheio.

const ANTHROPIC_API_KEY = Deno.env.get("ANTHROPIC_API_KEY");
const SUPABASE_URL = Deno.env.get("SUPABASE_URL");
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
}

function svcHeaders(extra?: Record<string, string>) {
  return { apikey: SUPABASE_SERVICE_ROLE_KEY!, Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`, ...(extra || {}) };
}

const FORMULARIOS = ["canvas", "tap", "planejamento", "smp", "tep", "rla", "relatorio-situacao", "relatorio-entregas"] as const;
type Formulario = typeof FORMULARIOS[number];

const SYSTEM_PROMPTS: Record<Formulario, string> = {
  canvas: `Você é um assistente que ajuda a preencher o Canvas de Projeto do Sistema de Gestão de Projetos (SGP), em português do Brasil, a partir de documentos já registrados do mesmo projeto (Solicitação de Demanda e, quando houver, Atas de Reunião).

O Canvas segue o modelo Por quê / O quê / Quem / Como / Quando e quanto, com estes campos:
- justificativas: as dores que originam o projeto (2 a 4 frases)
- objetivo: o objetivo do projeto (1 a 2 frases, formato SMART quando possível)
- beneficios: benefícios esperados
- produto: o produto/resultado do projeto (1 a 2 frases)
- parceiros: parceiros/áreas envolvidas no projeto
- entregas: grupos de entrega macro
- restricoes: restrições do projeto
- riscos: principais riscos identificáveis
- custos: estimativa de custos por grupo de entrega, somente se os documentos derem base para isso

Regras importantes:
- Baseie-se SOMENTE nos documentos fornecidos — nunca invente informação que não esteja neles.
- Se não houver informação suficiente para um campo, retorne null para ele — não escreva texto genérico de preenchimento.
- Escreva no mesmo tom dos documentos de origem: direto, objetivo, português formal de gestão de projetos.
- Responda APENAS com JSON válido, sem markdown, sem texto explicativo antes ou depois.
Formato exato: {"justificativas":"...","objetivo":"...","beneficios":"...","produto":"...","parceiros":"...","entregas":"...","restricoes":"...","riscos":"...","custos":"..."}`,

  tap: `Você é um assistente que ajuda a preencher o TAP (Termo de Abertura de Projeto) do Sistema de Gestão de Projetos (SGP), em português do Brasil, a partir de documentos já registrados do mesmo projeto (Solicitação de Demanda, Canvas de Projeto e, quando houver, Atas de Reunião).

Extraia/redija:
- justificativa: justificativa do projeto
- objetivos: objetivos do projeto
- publico: público-alvo
- beneficios: benefícios esperados
- exclusoes: exclusões — o que fica fora do escopo
- premissas: premissas assumidas
- restricoes: restrições
- criterios: critérios de aceitação
- riscos: array com até 5 riscos, cada item {"risco":"...","status":"Aberto","resp":""} — status sempre "Aberto" para riscos novos; "resp" vazio se os documentos não indicarem um responsável
- custos: array com até 6 itens, cada item {"item":"...","espec":"","unid":"","qtd":"","valor":""} — deixe campos vazios quando o documento não der o detalhe, nunca invente números
- interessadas: array de partes interessadas já nomeadas nos documentos, cada item {"nome":"","unidade":""}

Regras importantes:
- Baseie-se SOMENTE nos documentos fornecidos — nunca invente nomes, valores, datas ou responsáveis que não estejam neles.
- Campo de texto sem informação suficiente: retorne null. Tabela sem item identificável: retorne [].
- Responda APENAS com JSON válido, sem markdown, sem texto explicativo antes ou depois.
Formato exato: {"justificativa":"...","objetivos":"...","publico":"...","beneficios":"...","exclusoes":"...","premissas":"...","restricoes":"...","criterios":"...","riscos":[...],"custos":[...],"interessadas":[...]}`,

  planejamento: `Você é um assistente que ajuda a preencher o dossiê de Planejamento e Desenvolvimento de Projeto do Sistema de Gestão de Projetos (SGP), em português do Brasil, a partir de documentos já registrados do mesmo projeto (Solicitação de Demanda, TAP, Canvas de Projeto e, quando houver, Atas de Reunião).

Extraia/redija, para a capa (identificação):
- solicitante: nome de quem solicitou o projeto — copie da Solicitação de Demanda, se houver

Para a aba Pré-projeto:
- produtos: produtos impactados pelo projeto
- contexto: contexto e problemática que originou o projeto
- objGeral: objetivo geral do projeto
- objEspec: objetivos específicos do projeto
- escIncluido: o que está incluído no escopo
- escExcluido: o que fica fora do escopo
- entregaveis: entregáveis do projeto
- premissas: premissas assumidas — parta das já descritas no TAP, detalhando se possível
- restricoes: restrições — parta das já descritas no TAP, detalhando se possível
- partes: partes interessadas — parta das já descritas no TAP, detalhando se possível
- macroFases: macro fases do projeto (ex.: levantamento, desenvolvimento, implantação...) — parta dos entregáveis e do escopo já descritos no TAP, se houver, resumindo-os como fases em texto corrido

Para a aba Viabilidade — atenção: estes campos têm um propósito diferente dos campos do Pré-projeto acima, mesmo baseando-se nos mesmos fatos. Redija cada um com a redação própria do que é pedido, nunca copiando literalmente o texto de outro campo:
- introducao: introdução breve ao problema/oportunidade, para quem vai avaliar a viabilidade
- situacao: situação atual — o cenário/problema de hoje, antes do projeto (mesma base factual do campo "contexto" do Pré-projeto, mas com foco em descrever o estado atual das coisas, não em justificar o projeto)
- proposta: proposta de mudança — o que está sendo proposto para resolver a situação atual (mesma base factual do campo "objGeral" do Pré-projeto, mas redigida como uma proposta de solução, não como um objetivo)
- beneficios: benefícios do projeto — parta dos benefícios esperados já descritos no TAP, detalhando se possível

Para a aba Cronograma:
- cronograma: array com até 6 marcos de entrega macro, cada item {"marco":"...","resp":"","ini":null,"fim":null,"custo":""} — datas em YYYY-MM-DD SOMENTE se houver menção explícita, senão null. Baseie-se nos entregáveis/escopo do TAP e do próprio dossiê e, quando houver, nas Atas de Reunião.

Regras importantes:
- Baseie-se SOMENTE nos documentos fornecidos — nunca invente informação que não esteja neles.
- Campo de texto sem informação suficiente: retorne null. Tabela sem item identificável: retorne [].
- Responda APENAS com JSON válido, sem markdown, sem texto explicativo antes ou depois.
Formato exato: {"solicitante":"...","produtos":"...","contexto":"...","objGeral":"...","objEspec":"...","escIncluido":"...","escExcluido":"...","entregaveis":"...","premissas":"...","restricoes":"...","partes":"...","macroFases":"...","introducao":"...","situacao":"...","proposta":"...","beneficios":"...","cronograma":[...]}`,

  smp: `Você é um assistente que ajuda a preencher a SMP (Solicitação de Mudança de Projeto) do Sistema de Gestão de Projetos (SGP), em português do Brasil, a partir de documentos já registrados do mesmo projeto (TAP, Planejamento e Desenvolvimento e, quando houver, Atas de Reunião).

Diferente dos demais formulários, a SMP não tem um documento anterior que já descreva a mudança em si — o TAP e o Planejamento só mostram o que foi COMBINADO ORIGINALMENTE. Sua tarefa é procurar, nas Atas de Reunião mais recentes, alguma mudança de objetivo, cronograma, escopo, custo ou outro aspecto que esteja sendo discutida e que ainda não esteja refletida no TAP/Planejamento, e redigir um rascunho da SMP contrastando o que muda em relação ao combinado original.

Extraia/redija:
- descricao: descrição da mudança proposta
- beneficios: benefícios esperados da mudança
- naoImplantacao: impacto de NÃO implementar a mudança
- impObjetivo: como a mudança afeta o objetivo do projeto, comparado ao TAP/Planejamento
- impCronograma: como afeta o cronograma
- impEscopo: como afeta o escopo
- impCusto: como afeta o custo
- impEstrategico: como afeta o alinhamento estratégico
- impQualidade: como afeta a qualidade
- impRiscos: riscos decorrentes da mudança
- impOutros: outros impactos

Regras importantes:
- Só preencha um campo se as Atas realmente discutirem uma mudança concreta em relação ao TAP/Planejamento — se não houver nenhuma mudança clara sendo discutida, retorne null para TODOS os campos. Nunca invente uma mudança hipotética.
- NUNCA copie os campos do TAP/Planejamento diretamente — eles são o "antes", não o "depois". A SMP descreve o que MUDA.
- Responda APENAS com JSON válido, sem markdown, sem texto explicativo antes ou depois.
Formato exato: {"descricao":"...","beneficios":"...","naoImplantacao":"...","impObjetivo":"...","impCronograma":"...","impEscopo":"...","impCusto":"...","impEstrategico":"...","impQualidade":"...","impRiscos":"...","impOutros":"..."}`,

  tep: `Você é um assistente que ajuda a preencher o TEP (Termo de Encerramento de Projeto) do Sistema de Gestão de Projetos (SGP), em português do Brasil, a partir de documentos já registrados do mesmo projeto (TAP, Planejamento e Desenvolvimento, EAP e, quando houver, Atas de Reunião).

Extraia/redija, comparando o que foi planejado (TAP/Planejamento/EAP) com o que as Atas mais recentes indicam ter sido de fato realizado:
- justificativa: justificativa para o encerramento do projeto
- atividades: resumo das atividades/entregas encerradas, com base na EAP e no que as Atas confirmam como concluído
- analise: considerações finais — comparação entre planejado e entregue, principais desvios (se houver) e recomendações

Regras importantes:
- Baseie-se SOMENTE nos documentos fornecidos — nunca invente resultados, datas ou números que não estejam neles.
- Se as Atas não confirmarem o que foi de fato entregue, baseie "atividades" no que a EAP e o Planejamento definiram como escopo, deixando claro no texto que é o planejado — nunca invente confirmação de entrega que não esteja nos documentos.
- Se não houver informação suficiente para um campo, retorne null.
- Responda APENAS com JSON válido, sem markdown, sem texto explicativo antes ou depois.
Formato exato: {"justificativa":"...","atividades":"...","analise":"..."}`,

  rla: `Você é um assistente que ajuda a preencher o RLA (Registro de Lições Aprendidas) do Sistema de Gestão de Projetos (SGP), em português do Brasil, a partir de documentos já registrados do mesmo projeto (TEP, Atas de Reunião e SMPs — Solicitações de Mudança de Projeto).

O RLA tem várias seções. Você só preenche as 4 primeiras, todas de texto livre — NUNCA os blocos de avaliação estruturada Sim/Não/Parcial das seções seguintes (Fase de planejamento, Execução, Fatores humanos, Geral), que são uma autoavaliação manual da equipe e não fazem parte da sua tarefa:

Seção "Visão geral do projeto":
- q21: metas e objetivos do projeto, a partir do TEP
- q22: critérios de sucesso definidos
- q23: conclusão vs. expectativa — o resultado final bateu com o esperado?
- q24: comentários adicionais sobre metas e resultados

Seção "Destaques do projeto" (o que funcionou bem):
- q31: principais realizações do projeto
- q32: métodos ou práticas que deram certo
- q33: o que foi útil e vale repetir em outros projetos
- q34: comentários adicionais

Seção "Desafios do projeto" (o que pode melhorar):
- q41: o que deu errado — baseie-se nos entraves recorrentes das Atas e nos motivos de mudança das SMPs
- q42: processos que precisam melhorar
- q43: como melhorar
- q44: áreas de problemas recorrentes
- q45: desafios técnicos enfrentados
- q46: comentários adicionais

Seção "Tarefas pós-projeto / considerações futuras":
- q51: objetivos contínuos após o encerramento
- q52: ações pendentes
- q53: itens pendentes
- q54: comentários adicionais

Regras importantes:
- Baseie-se SOMENTE nos documentos fornecidos — nunca invente lições ou problemas sem base neles.
- Se não houver informação suficiente para um campo, retorne null — não escreva generalidades vazias como "o projeto correu bem".
- Responda APENAS com JSON válido, sem markdown, sem texto explicativo antes ou depois.
Formato exato: {"q21":"...","q22":"...","q23":"...","q24":"...","q31":"...","q32":"...","q33":"...","q34":"...","q41":"...","q42":"...","q43":"...","q44":"...","q45":"...","q46":"...","q51":"...","q52":"...","q53":"...","q54":"..."}`,

  "relatorio-situacao": `Você é um assistente que ajuda a preencher, por linha de projeto, o Relatório de Situação de Projetos (SGP-11) do Sistema de Gestão de Projetos (SGP), em português do Brasil — uma visão consolidada de portfólio, atualizada durante a execução do projeto.

Esta sugestão é para UMA linha específica do relatório, referente a um projeto já vinculado — não para o relatório inteiro. Baseie-se nos documentos desse projeto (TAP e, quando houver, Atas de Reunião) e, se fornecida, na situação já digitada nesta linha (status, % de execução, datas, sinalização automática de atraso).

Redija:
- atencao: o que merece atenção nesta linha — riscos em aberto do TAP, entraves recentes das Atas ainda sem solução, ou a sinalização automática de atraso/prazo vencido já calculada, se houver. Seja específico e breve (1 a 3 frases).
- destaque: o que merece destaque nesta linha — ausência de riscos em aberto, entraves recentes resolvidos, ou benefícios esperados relevantes do TAP. Seja específico e breve (1 a 3 frases).

Regras importantes:
- Baseie-se SOMENTE nos documentos e na situação fornecidos — nunca invente riscos, entraves ou conquistas que não estejam neles.
- Se não houver nada de fato relevante para um campo, retorne null — não escreva generalidades vazias como "o projeto está indo bem".
- Nunca sugira ou corrija o percentual de execução, o status ou as datas da linha — esses continuam sendo preenchidos manualmente por quem acompanha o projeto; sua tarefa é só os dois campos de texto abaixo.
- Responda APENAS com JSON válido, sem markdown, sem texto explicativo antes ou depois.
Formato exato: {"atencao":"...","destaque":"..."}`,

  "relatorio-entregas": `Você é um assistente que ajuda a preencher, por linha de projeto, o Relatório de Entregas e Benefícios (SGP-12) do Sistema de Gestão de Projetos (SGP), em português do Brasil — a ficha que embasa o Gate 2, pactuado com o Dono do Negócio ANTES da execução do projeto começar.

Esta sugestão é para UMA linha específica do relatório, referente a um projeto já vinculado — não para o relatório inteiro. Baseie-se no TAP desse projeto e, quando houver, nas Atas de Reunião.

Redija:
- impacto: o impacto/resultado ESPERADO do projeto para a organização — parta dos benefícios esperados e dos objetivos já descritos no TAP, com redação própria e focada em impacto organizacional, não uma cópia literal do campo de benefícios do TAP (que já é copiado automaticamente em outro campo desta mesma linha).

Regras importantes:
- Como o Gate 2 ocorre antes da execução, NUNCA descreva isso como um resultado já alcançado ou medido — é sempre uma expectativa, redija no futuro/condicional.
- Baseie-se SOMENTE no TAP e nas Atas fornecidos — nunca invente benefícios que não estejam neles.
- Se não houver informação suficiente, retorne null.
- Responda APENAS com JSON válido, sem markdown, sem texto explicativo antes ou depois.
Formato exato: {"impacto":"..."}`,
};

async function papelEquipePermitido(authHeader: string, projetoId: string): Promise<{ ok: boolean; motivo?: string }> {
  if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
    return { ok: false, motivo: "Configuração do servidor incompleta" };
  }
  try {
    const userResp = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
      headers: { apikey: SUPABASE_SERVICE_ROLE_KEY, Authorization: authHeader },
    });
    if (!userResp.ok) return { ok: false, motivo: "Sessão inválida" };
    const user = await userResp.json();
    const email = user?.email;
    const userId = user?.id;
    if (!email || !userId) return { ok: false, motivo: "Sessão inválida" };

    const perfilResp = await fetch(
      `${SUPABASE_URL}/rest/v1/perfis?email=eq.${encodeURIComponent(email)}&select=papel`,
      { headers: svcHeaders() }
    );
    const perfilRows = perfilResp.ok ? await perfilResp.json() : [];
    const papel = perfilRows[0]?.papel;
    if (!papel) return { ok: false, motivo: "Perfil não encontrado" };
    const isAdmin = papel === "admin";

    const cfgResp = await fetch(
      `${SUPABASE_URL}/rest/v1/configuracoes?chave=in.(permitir_assistente_preenchimento,papeis_assistente_preenchimento)&select=chave,valor`,
      { headers: svcHeaders() }
    );
    const cfgRows = cfgResp.ok ? await cfgResp.json() : [];
    const habilitado = cfgRows.find((r: any) => r.chave === "permitir_assistente_preenchimento")?.valor === true;
    if (!habilitado) return { ok: false, motivo: "O assistente de preenchimento por IA está desativado" };
    if (!isAdmin) {
      const papeis = cfgRows.find((r: any) => r.chave === "papeis_assistente_preenchimento")?.valor;
      if (!Array.isArray(papeis) || !papeis.includes(papel)) {
        return { ok: false, motivo: "Seu papel não tem permissão para usar o assistente de preenchimento" };
      }
    }

    if (!isAdmin) {
      const equipeResp = await fetch(
        `${SUPABASE_URL}/rest/v1/projeto_equipe?projeto_id=eq.${encodeURIComponent(projetoId)}&usuario_id=eq.${encodeURIComponent(userId)}&select=usuario_id`,
        { headers: svcHeaders() }
      );
      const equipeRows = equipeResp.ok ? await equipeResp.json() : [];
      if (!equipeRows.length) return { ok: false, motivo: "Você não faz parte da equipe deste projeto" };
    }

    return { ok: true };
  } catch (e) {
    console.error("Falha ao checar permissão:", e);
    return { ok: false, motivo: "Falha ao checar permissão" };
  }
}

async function fetchDocsByPrefix(prefix: string, projetoId: string): Promise<any[]> {
  const r = await fetch(
    `${SUPABASE_URL}/rest/v1/kv_store?key=like.${encodeURIComponent(prefix)}*&select=value`,
    { headers: svcHeaders() }
  );
  if (!r.ok) return [];
  const rows = await r.json();
  const out: any[] = [];
  for (const row of rows) {
    try {
      const parsed = JSON.parse(row.value);
      if (parsed && parsed.projetoId === projetoId) out.push(parsed);
    } catch { /* ignora registro malformado */ }
  }
  return out;
}

async function fetchDemanda(solicitacaoId: string | null): Promise<any | null> {
  if (!solicitacaoId) return null;
  const r = await fetch(
    `${SUPABASE_URL}/rest/v1/kv_store?key=eq.${encodeURIComponent("demanda:" + solicitacaoId)}&select=value`,
    { headers: svcHeaders() }
  );
  if (!r.ok) return null;
  const rows = await r.json();
  if (!rows.length) return null;
  try { return JSON.parse(rows[0].value); } catch { return null; }
}

function bloco(titulo: string, campos: Array<[string, string | null | undefined]>): string | null {
  const linhas = campos
    .filter(([, v]) => v != null && String(v).trim() !== "")
    .map(([label, v]) => `${label}: ${v}`);
  if (!linhas.length) return null;
  return `${titulo}\n${linhas.join("\n")}`;
}

function formatDemanda(d: any): string | null {
  return bloco(`SOLICITAÇÃO DE DEMANDA (protocolo ${d.protocolo || "—"})`, [
    ["Projeto", d.projeto],
    ["Solicitante", d.solicitante],
    ["Departamento/unidade", d.departamento],
    ["Justificativa/necessidade", d.justificativa],
    ["Objetivo/resultado esperado", d.objetivo],
    ["Escopo", d.escopo],
    ["Prazo desejado", d.prazo],
    ["Prioridade", d.prioridade],
    ["Orçamento estimado", d.orcamento],
    ["Partes interessadas", d.partes],
  ]);
}

function formatCanvas(c: any): string | null {
  return bloco(`CANVAS DE PROJETO (protocolo ${c.protocolo || "—"})`, [
    ["Justificativas", c.justificativas],
    ["Objetivo", c.objetivo],
    ["Benefícios", c.beneficios],
    ["Produto", c.produto],
    ["Parceiros", c.parceiros],
    ["Grupo de entregas", c.entregas],
    ["Restrições", c.restricoes],
    ["Riscos", c.riscos],
    ["Custos", c.custos],
  ]);
}

function formatTap(t: any): string | null {
  return bloco(`TAP - TERMO DE ABERTURA DE PROJETO (protocolo ${t.protocolo || "—"})`, [
    ["Alinhamento estratégico", t.alinhamento],
    ["Programa vinculado", t.programa],
    ["Justificativa", t.justificativa],
    ["Objetivos", t.objetivos],
    ["Público-alvo", t.publico],
    ["Benefícios esperados", t.beneficios],
    ["Exclusões (fora do escopo)", t.exclusoes],
    ["Premissas", t.premissas],
    ["Restrições", t.restricoes],
    ["Critérios de aceitação", t.criterios],
  ]);
}

function formatPlanejamento(p: any): string | null {
  const saidasResumo = Array.isArray(p.saidas) && p.saidas.length
    ? p.saidas.map((s: any) => `- ${s.item || ""}`).filter((l: string) => l !== "- ").join("\n")
    : "";
  const cronogramaResumo = Array.isArray(p.cronograma) && p.cronograma.length
    ? p.cronograma.map((c: any) => `- ${c.marco || ""}${c.resp ? ` (responsável: ${c.resp})` : ""}${c.fim ? ` (término: ${c.fim})` : ""}`).join("\n")
    : "";
  return bloco(`PLANEJAMENTO E DESENVOLVIMENTO DE PROJETO (protocolo ${p.protocolo || "—"})`, [
    ["Produtos impactados", p.produtos],
    ["Contexto e problemática", p.contexto],
    ["Objetivo geral", p.objGeral],
    ["Objetivos específicos", p.objEspec],
    ["Escopo incluído", p.escIncluido],
    ["Escopo excluído", p.escExcluido],
    ["Entregáveis", p.entregaveis],
    ["Benefícios do projeto", p.beneficios],
    ["Saídas / entregáveis detalhados", saidasResumo || null],
    ["Cronograma de entregas macro", cronogramaResumo || null],
  ]);
}

function formatEap(e: any): string | null {
  if (!Array.isArray(e.n1s) || !e.n1s.length) return null;
  const linhas: string[] = [];
  e.n1s.forEach((n1: any, i: number) => {
    if (!n1.name) return;
    linhas.push(`Pacote ${i + 1}: ${n1.name}`);
    (n1.n2s || []).forEach((n2: any, j: number) => {
      if (n2.name) linhas.push(`  Entrega ${i + 1}.${j + 1}: ${n2.name}`);
    });
  });
  if (!linhas.length) return null;
  return `EAP - ESTRUTURA ANALÍTICA DE PROJETO (protocolo ${e.protocolo || "—"})\n${linhas.join("\n")}`;
}

function formatSmp(s: any): string | null {
  return bloco(`SMP - SOLICITAÇÃO DE MUDANÇA (protocolo ${s.protocolo || "—"})`, [
    ["Descrição da mudança", s.descricao],
    ["Decisão", s.decisao],
    ["Justificativa da decisão", s.justDecisao],
    ["Impacto no objetivo", s.impObjetivo],
    ["Impacto no cronograma", s.impCronograma],
    ["Impacto no escopo", s.impEscopo],
    ["Impacto no custo", s.impCusto],
  ]);
}

function formatTep(t: any): string | null {
  return bloco(`TEP - TERMO DE ENCERRAMENTO DE PROJETO (protocolo ${t.protocolo || "—"})`, [
    ["Justificativa", t.justificativa],
    ["Atividades encerradas", t.atividades],
    ["Considerações finais", t.analise],
  ]);
}

function formatAta(a: any): string | null {
  const encaminhamentos = Array.isArray(a.saidas) && a.saidas.length
    ? a.saidas.map((s: any) => `- ${s.encam || ""}${s.resp ? ` (responsável: ${s.resp})` : ""}${s.prazo ? ` (prazo: ${s.prazo})` : ""}`).join("\n")
    : "";
  const entraves = Array.isArray(a.entraves) && a.entraves.length
    ? a.entraves.map((e: any) => `- ${e.entrave || ""}`).join("\n")
    : "";
  return bloco(`ATA DE REUNIÃO (protocolo ${a.protocolo || "—"}${a.data ? `, ${a.data}` : ""})`, [
    ["Pauta", a.pauta],
    ["Resumo", a.descricao],
    ["Encaminhamentos", encaminhamentos || null],
    ["Entraves", entraves || null],
  ]);
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  const authHeader = req.headers.get("Authorization");
  if (!authHeader) return json({ error: "Não autenticado" }, 401);

  let payload: { projetoId?: unknown; formulario?: unknown; contextoLinha?: unknown };
  try {
    payload = await req.json();
  } catch {
    return json({ error: "Corpo da requisição não é um JSON válido" }, 400);
  }

  const projetoId = typeof payload.projetoId === "string" ? payload.projetoId.trim() : "";
  const formulario = typeof payload.formulario === "string" ? payload.formulario.trim() : "";
  // Situação já digitada nesta linha do relatório (opcional) - só usado pelo
  // formulario 'relatorio-situacao', para o assistente não repetir uma
  // sinalização de atraso que a própria pessoa já vê na tela.
  const CONTEXTO_LINHA_CAMPOS: Record<string, string> = {
    status: "Status", pctExecucao: "% Execução", inicioPrevisto: "Início previsto",
    terminoPrevisto: "Término previsto", sinalizacao: "Sinalização automática",
  };
  let contextoLinhaTexto: string | null = null;
  if (payload.contextoLinha && typeof payload.contextoLinha === "object") {
    const linhas: string[] = [];
    for (const [chave, rotulo] of Object.entries(CONTEXTO_LINHA_CAMPOS)) {
      const v = (payload.contextoLinha as Record<string, unknown>)[chave];
      if (typeof v === "string" && v.trim()) linhas.push(`${rotulo}: ${v.trim()}`);
    }
    if (linhas.length) contextoLinhaTexto = `SITUAÇÃO JÁ REGISTRADA NESTA LINHA DO RELATÓRIO\n${linhas.join("\n")}`;
  }
  if (!projetoId) return json({ error: "Envie o campo 'projetoId'" }, 400);
  if (!FORMULARIOS.includes(formulario as Formulario)) {
    return json({ error: "Campo 'formulario' inválido — use 'canvas', 'tap', 'planejamento', 'smp', 'tep', 'rla', 'relatorio-situacao' ou 'relatorio-entregas'" }, 400);
  }

  const permissao = await papelEquipePermitido(authHeader, projetoId);
  if (!permissao.ok) return json({ error: permissao.motivo || "Sem permissão" }, 403);

  if (!ANTHROPIC_API_KEY) {
    return json({ error: "ANTHROPIC_API_KEY não configurada nos secrets da function" }, 500);
  }

  try {
    const projetoResp = await fetch(
      `${SUPABASE_URL}/rest/v1/projetos?id=eq.${encodeURIComponent(projetoId)}&select=id,nome,solicitacao_id`,
      { headers: svcHeaders() }
    );
    const projetoRows = projetoResp.ok ? await projetoResp.json() : [];
    const projeto = projetoRows[0];
    if (!projeto) return json({ error: "Projeto não encontrado" }, 404);

    // Quais documentos anteriores cada formulario-alvo le, seguindo a esteira documental
    const FONTES_POR_FORMULARIO: Record<Formulario, { demanda?: boolean; canvas?: boolean; tap?: boolean; planejamento?: boolean; eap?: boolean; smp?: boolean; tep?: boolean; atas?: boolean }> = {
      canvas: { demanda: true, atas: true },
      tap: { demanda: true, canvas: true, atas: true },
      planejamento: { demanda: true, canvas: true, tap: true, atas: true },
      smp: { tap: true, planejamento: true, atas: true },
      tep: { tap: true, planejamento: true, eap: true, atas: true },
      rla: { tep: true, smp: true, atas: true },
      "relatorio-situacao": { tap: true, atas: true },
      "relatorio-entregas": { tap: true, atas: true },
    };
    const fontesConfig = FONTES_POR_FORMULARIO[formulario as Formulario];

    const demanda = fontesConfig.demanda ? await fetchDemanda(projeto.solicitacao_id || null) : null;
    const canvas = fontesConfig.canvas ? (await fetchDocsByPrefix("canvas:", projetoId))[0] || null : null;
    const tap = fontesConfig.tap ? (await fetchDocsByPrefix("tap:", projetoId))[0] || null : null;
    const planejamento = fontesConfig.planejamento ? (await fetchDocsByPrefix("plan:", projetoId))[0] || null : null;
    const eap = fontesConfig.eap ? (await fetchDocsByPrefix("eap:", projetoId))[0] || null : null;
    const tep = fontesConfig.tep ? (await fetchDocsByPrefix("tep:", projetoId))[0] || null : null;
    const smps = fontesConfig.smp ? await fetchDocsByPrefix("smp:", projetoId) : [];
    const atas = fontesConfig.atas ? await fetchDocsByPrefix("ata:", projetoId) : [];

    const blocos: string[] = [];
    const fontes: Array<{ tipo: string; protocolo: string }> = [];
    if (demanda) {
      const b = formatDemanda(demanda);
      if (b) { blocos.push(b); fontes.push({ tipo: "Solicitação de Demanda", protocolo: demanda.protocolo || "—" }); }
    }
    if (canvas) {
      const b = formatCanvas(canvas);
      if (b) { blocos.push(b); fontes.push({ tipo: "Canvas de Projeto", protocolo: canvas.protocolo || "—" }); }
    }
    if (tap) {
      const b = formatTap(tap);
      if (b) { blocos.push(b); fontes.push({ tipo: "TAP", protocolo: tap.protocolo || "—" }); }
    }
    if (planejamento) {
      const b = formatPlanejamento(planejamento);
      if (b) { blocos.push(b); fontes.push({ tipo: "Planejamento e Desenvolvimento", protocolo: planejamento.protocolo || "—" }); }
    }
    if (eap) {
      const b = formatEap(eap);
      if (b) { blocos.push(b); fontes.push({ tipo: "EAP", protocolo: eap.protocolo || "—" }); }
    }
    if (tep) {
      const b = formatTep(tep);
      if (b) { blocos.push(b); fontes.push({ tipo: "TEP", protocolo: tep.protocolo || "—" }); }
    }
    for (const smp of smps) {
      const b = formatSmp(smp);
      if (b) { blocos.push(b); fontes.push({ tipo: "SMP", protocolo: smp.protocolo || "—" }); }
    }
    for (const ata of atas) {
      const b = formatAta(ata);
      if (b) { blocos.push(b); fontes.push({ tipo: "Ata de Reunião", protocolo: ata.protocolo || "—" }); }
    }

    if (!blocos.length && !contextoLinhaTexto) {
      return json({ error: "Nenhum documento anterior com conteúdo foi encontrado para este projeto ainda" }, 422);
    }
    if (contextoLinhaTexto) blocos.push(contextoLinhaTexto);

    let contexto = blocos.join("\n\n");
    if (contexto.length > 60000) contexto = contexto.slice(0, 60000) + "\n\n[...conteúdo adicional omitido por limite de tamanho...]";

    const resp = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-5",
        max_tokens: 16000,
        system: SYSTEM_PROMPTS[formulario as Formulario],
        messages: [{ role: "user", content: `Documentos do projeto "${projeto.nome || ""}":\n\n${contexto}` }],
      }),
    });

    if (!resp.ok) {
      const errBody = await resp.text();
      console.error("Anthropic API retornou erro:", resp.status, errBody);
      return json({ error: "Falha ao gerar a sugestão (erro na API de IA)" }, 502);
    }

    const data = await resp.json();
    if (data.stop_reason === "max_tokens") {
      console.error("Resposta da IA cortada por atingir o limite de tokens.");
      return json({ error: "O contexto do projeto é longo demais para sugerir de uma vez" }, 502);
    }

    const text = (data.content || [])
      .filter((b: any) => b.type === "text")
      .map((b: any) => b.text)
      .join("");
    const clean = text.replace(/```json|```/g, "").trim();

    let parsed: unknown;
    try {
      parsed = JSON.parse(clean);
    } catch {
      console.error("Resposta da IA nao era JSON valido:", clean.slice(0, 500));
      return json({ error: "A IA não conseguiu estruturar a sugestão — tente novamente" }, 502);
    }

    return json({ ok: true, dados: parsed, fontes });
  } catch (e) {
    console.error("Falha ao gerar sugestão de preenchimento:", e);
    return json({ error: "Falha ao gerar a sugestão" }, 502);
  }
});
