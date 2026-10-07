// Edge Function: analisar-documento-demanda
// Recebe o texto extraido de um documento Word/PDF/txt (o formulario oficial
// SGP-01 - Solicitacao de Demanda, ja preenchido a mao fora do sistema)
// colado ou anexado pelo usuario, e usa a API da Claude (Anthropic) para
// extrair os campos estruturados e devolver JSON pronto para preencher o
// formulario online. O usuario sempre revisa e completa antes de salvar -
// esta funcao so sugere, nunca salva nada sozinha. Mesma arquitetura das
// funcoes ja em producao `analisar-transcricao-ata` e `sugerir-preenchimento`.
//
// Segredo necessario (ja configurado nesta base):
//   ANTHROPIC_API_KEY - chave da API da Claude (console.anthropic.com)
// SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY sao injetados automaticamente.
//
// Verificacao de JWT permanece ligada. Alem disso, confere no banco (nao so
// na interface) se a importacao esta habilitada (tabela configuracoes) e se
// o papel de quem chamou esta na lista permitida - Admin sempre passa.

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

const UNIDADES = [
  "Superintendência Reitoria", "Ouvidoria", "CPA", "Comitê da Qualidade",
  "Diretoria Educação Executiva", "Diretoria de Escola Business", "Diretoria de Escola de Direito",
  "Diretoria de Operações EAD", "Diretoria de Marketing e Vendas", "Coord. Curso Lato Sensu",
  "Coord. Pós-Graduação Stricto Sensu", "Coord. Curso Graduação, Técnico e Ensino Médio",
  "Coord. Curso Graduação", "Coord. Operações Acadêmicas EAD", "Coord. Gestão de Polos",
  "Corpo Docente", "Corpo Docente (PJ)", "Marketing", "Vendas", "Relacionamento",
  "Planej. e Operações", "RH", "Secretaria Geral", "Asseg. Qualidade", "Segurança",
  "ADM e Infraestrutura", "GPTI", "Controladoria", "Gerência de Controladoria",
];

const SYSTEM_PROMPT = `Você é um assistente que extrai dados estruturados do texto de uma Solicitação de Demanda de Projetos (formulário oficial SGP-01), em português do Brasil. O texto vem de um documento Word ou PDF que a pessoa preencheu manualmente fora do sistema e agora está transcrevendo para o formulário online.

O formulário oficial tem estas seções numeradas, nesta ordem: 1 Nome do projeto, 2 Solicitante, 3 Justificativa/Necessidade, 4 Objetivo/Resultado esperado, 5 Escopo, 6 Prazo desejado, 7 Orçamento estimado, 8 Partes interessadas, 9 Anexos.

Extraia:
- projeto: nome do projeto (seção 1)
- solicitante: SOMENTE o nome completo de quem solicita — sem o departamento, mesmo que estejam juntos no documento (seção 2)
- departamento: a unidade/departamento do solicitante, se mencionado (também costuma estar na seção 2, junto do nome). Escolha o valor EXATAMENTE como escrito nesta lista fixa, ou null se nada corresponder com confiança: ${JSON.stringify(UNIDADES)}
- justificativa: justificativa/necessidade (seção 3)
- objetivo: objetivo/resultado esperado (seção 4)
- escopo: escopo — o que está incluído e o que não está (seção 5)
- prazo: prazo desejado em formato YYYY-MM-DD, SOMENTE se houver uma data explícita e completa (seção 6) — senão null
- prioridade: urgência da demanda — use "Crítica", "Alta", "Média" ou "Baixa" SOMENTE se o documento indicar isso claramente (um campo de prioridade preenchido, ou palavras inequívocas como "urgente"); senão null — nunca infira prioridade a partir do conteúdo da justificativa
- orcamento: orçamento estimado, como texto (seção 7)
- partes: partes interessadas — quem será impactado (seção 8)
- anexos: documentos auxiliares mencionados (seção 9)

Regras importantes:
- Baseie-se SOMENTE no texto fornecido — nunca invente informação que não esteja nele.
- Se um campo não tiver informação suficiente, retorne null.
- Responda APENAS com JSON válido, sem markdown, sem texto explicativo antes ou depois.
Formato exato: {"projeto":"...","solicitante":"...","departamento":"...","justificativa":"...","objetivo":"...","escopo":"...","prazo":"...","prioridade":"...","orcamento":"...","partes":"...","anexos":"..."}`;

async function papelPermitidoNoBanco(authHeader: string): Promise<{ ok: boolean; motivo?: string }> {
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
    if (!email) return { ok: false, motivo: "Sessão inválida" };

    const perfilResp = await fetch(
      `${SUPABASE_URL}/rest/v1/perfis?email=eq.${encodeURIComponent(email)}&select=papel`,
      { headers: { apikey: SUPABASE_SERVICE_ROLE_KEY, Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}` } }
    );
    const perfilRows = perfilResp.ok ? await perfilResp.json() : [];
    const papel = perfilRows[0]?.papel;
    if (!papel) return { ok: false, motivo: "Perfil não encontrado" };

    const cfgResp = await fetch(
      `${SUPABASE_URL}/rest/v1/configuracoes?chave=in.(permitir_importar_documento_demanda,papeis_importar_documento_demanda)&select=chave,valor`,
      { headers: { apikey: SUPABASE_SERVICE_ROLE_KEY, Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}` } }
    );
    const cfgRows = cfgResp.ok ? await cfgResp.json() : [];
    const habilitado = cfgRows.find((r: any) => r.chave === "permitir_importar_documento_demanda")?.valor === true;
    if (!habilitado) return { ok: false, motivo: "A importação de documento está desativada" };
    if (papel === "admin") return { ok: true };

    const papeis = cfgRows.find((r: any) => r.chave === "papeis_importar_documento_demanda")?.valor;
    if (!Array.isArray(papeis) || !papeis.includes(papel)) {
      return { ok: false, motivo: "Seu papel não tem permissão para importar documento" };
    }
    return { ok: true };
  } catch (e) {
    console.error("Falha ao checar permissão:", e);
    return { ok: false, motivo: "Falha ao checar permissão" };
  }
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  const authHeader = req.headers.get("Authorization");
  if (!authHeader) return json({ error: "Não autenticado" }, 401);

  const permissao = await papelPermitidoNoBanco(authHeader);
  if (!permissao.ok) return json({ error: permissao.motivo || "Sem permissão" }, 403);

  if (!ANTHROPIC_API_KEY) {
    return json({ error: "ANTHROPIC_API_KEY não configurada nos secrets da function" }, 500);
  }

  let payload: { texto?: unknown };
  try {
    payload = await req.json();
  } catch {
    return json({ error: "Corpo da requisição não é um JSON válido" }, 400);
  }

  const texto = typeof payload.texto === "string" ? payload.texto.trim() : "";
  if (!texto) return json({ error: "Envie o campo 'texto' com o conteúdo do documento" }, 400);
  if (texto.length > 100000) {
    return json({ error: "Documento muito longo (máximo de ~100.000 caracteres)" }, 400);
  }

  try {
    const resp = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-5",
        max_tokens: 4000,
        system: SYSTEM_PROMPT,
        messages: [{ role: "user", content: `Texto do documento:\n\n${texto}` }],
      }),
    });

    if (!resp.ok) {
      const errBody = await resp.text();
      console.error("Anthropic API retornou erro:", resp.status, errBody);
      return json({ error: "Falha ao analisar o documento (erro na API de IA)" }, 502);
    }

    const data = await resp.json();
    if (data.stop_reason === "max_tokens") {
      console.error("Resposta da IA cortada por atingir o limite de tokens.");
      return json({ error: "O documento é longo demais para ser analisado de uma vez" }, 502);
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
      return json({ error: "A IA não conseguiu estruturar os dados desse documento — tente novamente" }, 502);
    }

    return json({ ok: true, dados: parsed });
  } catch (e) {
    console.error("Falha ao chamar a API da Claude:", e);
    return json({ error: "Falha ao analisar o documento" }, 502);
  }
});
