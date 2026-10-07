// Edge Function: verificar-prazos-cronograma
// Roda 1x por dia via pg_cron (ver supabase/migrations/<data>_prazos_cronograma_cron.sql),
// sem depender de ninguem abrir o sistema. Varre o cronograma (marcos) de cada projeto
// ainda ativo (Planejamento e Desenvolvimento, chave `plan:` no kv_store) e dispara dois
// tipos de aviso por e-mail/push (reaproveitando a function send-notification), sem
// Conclusao real preenchida:
//
// 1) Aviso de prazo (existente):
//   - N1 dias antes do Termino previsto (padrao 5, faixa entre N2+1 e N1 dias)
//   - N2 dias antes (padrao 2, faixa de 1 a N2 dias)
//   - no dia do vencimento
//   - atrasado (repete a cada N3 dias enquanto continuar sem Conclusao real — padrao 7)
//
// 2) Aviso de silencio (novo — fecha o ciclo do campo Execucao, ver 04/20/19): marco cujo
//   Inicio previsto ja chegou e que esta ha N4 dias ou mais sem nenhuma mudanca no campo
//   Execucao (comparando com execStatusAtualizadoEm, ou com o proprio Inicio previsto se o
//   campo nunca foi preenchido — padrao N4 = 5) — repete a cada N5 dias enquanto continuar
//   parado (padrao N5 = 5). Existe porque o aviso de prazo acima so olha a data final: um
//   marco com o Termino previsto ainda longe (ex.: 44 dias) podia ficar semanas parado, sem
//   Execucao preenchida, sem ninguem ser avisado. Nao dispara no mesmo dia em que o aviso de
//   prazo ja foi enviado para o mesmo marco, para nao mandar dois e-mails de uma vez.
//
// N1..N5 sao parametrizaveis em Administracao > Configuracoes > Regras de aviso automatico
// (tabela `configuracoes`, chaves aviso_prazo_dias_antes1/antes2/repeticao_atrasado e
// aviso_silencio_dias_sem_atualizar/repeticao — ver migracao
// 20260917160000_config_avisos_prazo_execucao.sql). Lidos a cada execucao no inicio da
// function; se uma chave nao existir ou tiver valor invalido, cai no mesmo padrao que era
// fixo no codigo antes desta migracao — a function nunca fica sem rodar por configuracao
// ausente.
//
// Destinatarios de cada aviso:
//   - os e-mails vinculados no campo "Responsavel (conta)" do marco (respConta —
//     pode ter mais de uma pessoa), se preenchido;
//   - + toda conta com papel 'gerente_projetos' que estiver na equipe daquele projeto
//     (projeto_equipe + perfis.papel — nao usa o campo de texto livre "Gerente do projeto").
// Se nenhum dos dois existir, o marco fica sem destinatario e nenhum e-mail e enviado
// (o indicador visual do Painel a Vista continua mostrando o atraso normalmente).
//
// Projetos com TEP ja registrado (kv_store tipo='tep') sao tratados como encerrados e
// ignorados, mesmo que o Planejamento correspondente ainda tenha marcos sem Conclusao real.
//
// Controle de "ja avisado" (evita repetir o mesmo aviso todo dia dentro da mesma janela):
// gravado no proprio kv_store, prefixo `alertaprazo:<marcoId>` (marcoId e o `id` estavel
// gerado no formulario 04 para cada marco — ver novoId() em
// "04 - planejamento-desenvolvimento-projeto.html"; marcos antigos sem esse id usam um
// identificador de fallback baseado em posicao, best-effort).
//
// Tecnicas de engajamento aplicadas nos dois e-mails (17/09/2026), depois de revisar o
// layout inicial (so estrutura visual, sem essas tecnicas):
//   - Saudacao pessoal ("Ola, {primeiro nome}") — por isso agora o envio e feito uma vez por
//     destinatario (enviarParaCadaDestinatario), nao um unico envio com todos no "to"; de
//     passagem, tambem para de expor o e-mail de um destinatario aos outros.
//   - Link direto para o marco especifico dentro do Meu Painel (?marco=<id>, ver
//     meuPainelUrlMarco e a leitura do parametro em 20 - meu-painel.html), nao so a tela
//     generica — reduz a ficcao entre clicar no e-mail e agir.
//   - Numero grande em destaque (dias atrasado / dias para o prazo / dias sem atualizar) antes
//     de qualquer frase, para o dado mais acionavel ser lido em menos de 1 segundo.
//   - Tom que escalona por severidade: lembrete tranquilo quando falta tempo (antes5), mais
//     direto no vencimento (antes2/nodia), e no atrasado/segunda cobranca de silencio soma o
//     framing de visibilidade (accountability) — o status ja esta sendo visto pela Gerencia de
//     Projetos no Painel de Prazos, nao e so um pedido isolado.
//   - Preheader (texto escondido que a caixa de entrada mostra ao lado do assunto, ver
//     `preheader` em emailTemplate) — sem isso a pre-visualizacao pegava a barra "UNIALFA...",
//     que nao diz nada.
//   - Microcopy do botao ligada ao resultado esperado ("Regularizar agora", "Atualizar
//     execucao agora"), nao a navegacao ("Abrir o Meu Painel").
//
// A URL do Meu Painel muda com o ambiente (producao vs treino) verificando se SUPABASE_URL
// contem o ref de producao, mesmo padrao ja usado em reset-treino/index.ts (PRODUCAO_REF).
//
// Cada e-mail tambem termina indicando a Gerencia de Projetos como alternativa de contato,
// caso a pessoa prefira ajustar o prazo diretamente em vez de seguir só pelo aviso automatico
// — o nome do Gerente vem do campo "Gerente do projeto" (rec.gestor) do proprio Planejamento,
// nao e fixo.
//
// Segredos necessarios: nenhum novo — reaproveita SUPABASE_URL e
// SUPABASE_SERVICE_ROLE_KEY (injetados automaticamente pelo runtime das Edge Functions)
// para ler kv_store/perfis/projeto_equipe direto (sem RLS) e chamar a send-notification.
//
// Verificacao de JWT permanece ligada (padrao do Supabase, igual as demais functions) — o
// cron chama esta function com a anon key (publica, ja embutida em todo o front-end),
// exatamente como o restante da API REST.

const SUPABASE_URL = Deno.env.get("SUPABASE_URL");
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
const PRODUCAO_REF = "fiarntunpqteopwjkhjg";
const SITE_URL = SUPABASE_URL && SUPABASE_URL.includes(PRODUCAO_REF)
  ? "https://paramiri.github.io/unialfa-gestao-projetos"
  : "https://paramiri.github.io/unialfa-gestao-projetos-treino";
const MEU_PAINEL_URL = `${SITE_URL}/20%20-%20meu-painel.html`;
// Link direto para o marco especifico (nao so o Meu Painel generico) — quem abre ja chega no
// item certo, destacado (ver 20 - meu-painel.html, leitura do parametro ?marco=). Reduz a
// fricção entre clicar no e-mail e agir: sem isso, a pessoa precisava achar o item na lista.
function meuPainelUrlMarco(marcoId: string): string {
  return `${MEU_PAINEL_URL}?marco=${encodeURIComponent(marcoId)}`;
}

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
  return {
    apikey: SUPABASE_SERVICE_ROLE_KEY!,
    Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
    "Content-Type": "application/json",
    ...(extra || {}),
  };
}

function esc(s: unknown): string {
  return String(s ?? "").replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c] as string)
  );
}

// Layout compartilhado dos e-mails automaticos (aviso de prazo e de silencio) — tabela com
// estilo embutido em cada elemento, fonte de sistema (Arial/Helvetica): e-mail nao renderiza
// CSS moderno nem fontes externas de forma confiavel (Gmail remove <style> do <head>, Outlook
// classico ignora boa parte do CSS), entao esse e o formato seguro para e-mail transacional.
// Mesma funcao (duplicada, sem import entre arquivos — convencao ja usada no resto do
// projeto) tambem existe em cada um dos 6 formularios que disparam notificacao por e-mail.
type EmailBadgeTipo = "ok" | "bad" | "warn" | "slate";
const EMAIL_CORES: Record<EmailBadgeTipo, { bg: string; fg: string }> = {
  ok: { bg: "#E6F4EC", fg: "#1B7F4B" },
  bad: { bg: "#FBE7EC", fg: "#9F1239" },
  warn: { bg: "#FBF1E3", fg: "#95530A" },
  slate: { bg: "#EEF1F4", fg: "#475569" },
};
function emailTemplate(opts: {
  preheader?: string;
  saudacao?: string;
  badgeTexto: string;
  badgeTipo: EmailBadgeTipo;
  titulo: string;
  subtitulo?: string;
  statNumero?: string;
  statLabel?: string;
  corpo?: string;
  linhas?: { label: string; valor: string }[];
  ctaTexto?: string;
  ctaUrl?: string;
  rodape?: string;
}): string {
  const c = EMAIL_CORES[opts.badgeTipo];
  const linhasHtml = (opts.linhas || [])
    .map(
      (l) =>
        `<tr><td style="padding:10px 14px;font-size:12.5px;color:#8A8E94;width:40%;border-bottom:1px solid #E4E4E7">${l.label}</td><td style="padding:10px 14px;font-size:12.5px;color:#1A1A1A;font-weight:600;border-bottom:1px solid #E4E4E7">${l.valor}</td></tr>`
    )
    .join("");
  return (
    // Preheader: texto que a caixa de entrada mostra ao lado do assunto (Gmail/Apple Mail) —
    // sem isso, ela pega o primeiro texto visível do e-mail (a barra "UNIALFA..."), que nao diz
    // nada. Escondido visualmente, nao aparece ao abrir o e-mail.
    (opts.preheader
      ? `<div style="display:none;max-height:0;overflow:hidden;mso-hide:all;font-size:1px;line-height:1px;color:#F3F4F6;opacity:0">${opts.preheader}&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;</div>`
      : "") +
    `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#F3F4F6;padding:24px 0"><tr><td>` +
    `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:520px;margin:0 auto;background:#FFFFFF;border-radius:10px;overflow:hidden;border:1px solid #E4E4E7;font-family:Arial,Helvetica,sans-serif">` +
    `<tr><td style="background:#B91D2E;padding:16px 24px"><span style="color:#FFFFFF;font-size:12.5px;font-weight:800;letter-spacing:.05em">UNIALFA · GESTÃO DE PROJETOS</span></td></tr>` +
    `<tr><td style="padding:26px 24px 8px">` +
    (opts.saudacao ? `<p style="margin:0 0 14px;font-size:13px;color:#52525B">${opts.saudacao}</p>` : "") +
    `<span style="display:inline-block;font-size:10.5px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;padding:4px 11px;border-radius:20px;margin-bottom:14px;background:${c.bg};color:${c.fg}">${opts.badgeTexto}</span>` +
    `<p style="margin:0 0 4px;font-size:19px;color:#1A1A1A;font-weight:700">${opts.titulo}</p>` +
    (opts.subtitulo
      ? `<p style="margin:0 0 18px;font-size:12.5px;color:#8A8E94">${opts.subtitulo}</p>`
      : `<div style="height:10px;line-height:10px">&nbsp;</div>`) +
    // Numero grande: o dado mais acionavel (ha quantos dias) lido em menos de 1 segundo, antes
    // de qualquer frase — hierarquia visual em vez de so texto corrido.
    (opts.statNumero
      ? `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 18px"><tr><td align="center" style="padding:14px 0;background:${c.bg};border-radius:8px">` +
        `<div style="font-size:38px;font-weight:800;color:${c.fg};line-height:1">${opts.statNumero}</div>` +
        `<div style="font-size:10.5px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:${c.fg};margin-top:4px">${opts.statLabel}</div>` +
        `</td></tr></table>`
      : "") +
    (opts.corpo ? `<p style="margin:0 0 20px;font-size:13.5px;color:#3F3F46;line-height:1.6">${opts.corpo}</p>` : "") +
    (linhasHtml
      ? `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#FAFAFB;border:1px solid #E4E4E7;border-radius:8px;margin:0 0 22px">${linhasHtml}</table>`
      : "") +
    (opts.ctaUrl
      ? `<a href="${opts.ctaUrl}" style="display:inline-block;background:#B91D2E;color:#FFFFFF;font-size:13.5px;font-weight:700;padding:12px 24px;border-radius:8px;text-decoration:none;margin-bottom:22px">${opts.ctaTexto} →</a>`
      : "") +
    `</td></tr>` +
    (opts.rodape ? `<tr><td style="padding:16px 24px;border-top:1px solid #E4E4E7">${opts.rodape}</td></tr>` : "") +
    `</table></td></tr></table>`
  );
}

// "Hoje" em America/Sao_Paulo (UTC-3, sem horario de verao desde 2019) — o cron roda as
// 11:00 UTC (08:00 BRT), horario bem longe da virada de meia-noite local.
function hojeDataBR(): string {
  return new Date(Date.now() - 3 * 3600 * 1000).toISOString().slice(0, 10);
}
function diasEntre(dataA: string, dataB: string): number {
  const a = Date.parse(dataA + "T00:00:00Z");
  const b = Date.parse(dataB + "T00:00:00Z");
  return Math.round((a - b) / 86400000);
}
// Responsável (conta) pode ter mais de uma pessoa, guardado como string separada
// por vírgula (mesmo formato usado nos formulários) — um e-mail só, sem vírgula,
// continua funcionando normalmente (formato anterior, sem precisar migrar dado).
function parseRespConta(v?: string): string[] {
  return String(v || "").split(",").map((s) => s.trim()).filter(Boolean);
}

type Marco = {
  id?: string;
  marco?: string;
  resp?: string;
  respConta?: string;
  ini?: string;
  fim?: string;
  concl?: string;
  execStatus?: string;
  execStatusAtualizadoEm?: string;
};
type PlanoRec = {
  id?: string;
  nomeProjeto?: string;
  protocolo?: string;
  projetoId?: string;
  gestor?: string;
  cronograma?: Marco[];
};

type Bucket = "antes5" | "antes2" | "nodia" | "atrasado";
const BUCKET_LABEL: Record<Bucket, string> = {
  antes5: "vence em até 5 dias",
  antes2: "vence em até 2 dias",
  nodia: "vence hoje",
  atrasado: "está atrasado",
};

// dAntes2/dAntes1 vêm de Administração > Configurações — normaliza a ordem (min/max) para
// nunca abrir um buraco na faixa de dias se alguém configurar dAntes2 maior que dAntes1.
function bucketParaDias(dias: number, dAntes2: number, dAntes1: number): Bucket | null {
  if (dias < 0) return "atrasado";
  if (dias === 0) return "nodia";
  const antes2 = Math.min(dAntes2, dAntes1);
  const antes1 = Math.max(dAntes2, dAntes1);
  if (dias >= 1 && dias <= antes2) return "antes2";
  if (dias > antes2 && dias <= antes1) return "antes5";
  return null;
}

// Limite de taxa do Resend (plano atual): enviar rapido demais em sequencia derruba varios
// envios de uma vez com RateLimitError. Desde que o envio passou a ser um por destinatario
// (para personalizar a saudacao — ver enviarParaCadaDestinatario abaixo), uma execucao com
// muitos marcos pendentes de uma vez (ex.: primeira ativacao, ou o dia em que o cron tinha
// parado de rodar) facilmente ultrapassa 2 envios/segundo. Um intervalo minimo global entre
// CADA envio (nao so entre destinatarios do mesmo marco) resolve na raiz, em vez de tentar de
// novo depois de tomar o rate limit — descoberto em producao em 17/09/2026 (21 de 33 avisos
// enviados, 12 falharam por RateLimitError antes desta correcao).
let ultimoEnvioEmailTs = 0;
async function aguardarRateLimit(minIntervaloMs = 600) {
  const espera = ultimoEnvioEmailTs + minIntervaloMs - Date.now();
  if (espera > 0) await new Promise((r) => setTimeout(r, espera));
  ultimoEnvioEmailTs = Date.now();
}

// Envia uma copia do e-mail para cada destinatario, em vez de um so envio com todos no
// mesmo "to" — permite personalizar a saudacao por pessoa (buildHtml recebe o e-mail de
// quem vai receber) e, de passagem, para de expor o e-mail de um destinatario aos outros.
async function enviarParaCadaDestinatario(
  destinatarios: Set<string>,
  subject: string,
  buildHtml: (destinatario: string) => string
): Promise<{ enviouAlgum: boolean; erros: string[] }> {
  let enviouAlgum = false;
  const erros: string[] = [];
  for (const destinatario of destinatarios) {
    try {
      await aguardarRateLimit();
      const r = await fetch(`${SUPABASE_URL}/functions/v1/send-notification`, {
        method: "POST",
        headers: svcHeaders(),
        body: JSON.stringify({ to: [destinatario], subject, html: buildHtml(destinatario) }),
      });
      if (!r.ok) throw new Error(await r.text());
      enviouAlgum = true;
    } catch (e) {
      erros.push(`${destinatario}: ${String(e)}`);
    }
  }
  return { enviouAlgum, erros };
}

async function restGet(path: string) {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, { headers: svcHeaders() });
  if (!r.ok) throw new Error(`GET ${path} -> ${r.status}: ${await r.text()}`);
  return r.json();
}

// Lê um numero configuravel em Administracao > Configuracoes (tabela `configuracoes`),
// caindo no padrao se a chave nao existir, nao ser numero, ou nao for positiva.
async function getConfigDias(chave: string, padrao: number): Promise<number> {
  try {
    const rows: { valor: unknown }[] = await restGet(
      `configuracoes?chave=eq.${chave}&select=valor`
    );
    const v = rows[0]?.valor;
    return typeof v === "number" && v > 0 ? v : padrao;
  } catch {
    return padrao;
  }
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }
  if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
    return json({ error: "SUPABASE_URL/SUPABASE_SERVICE_ROLE_KEY não disponíveis" }, 500);
  }

  const hoje = hojeDataBR();
  const resultado = { avaliados: 0, avisosEnviados: 0, erros: [] as string[] };

  try {
    // Dias configuraveis das duas regras de disparo (Administracao > Configuracoes).
    const [diasAntes1, diasAntes2, diasRepeticaoAtrasado, diasSilencioLimite, diasSilencioRepeticao] =
      await Promise.all([
        getConfigDias("aviso_prazo_dias_antes1", 5),
        getConfigDias("aviso_prazo_dias_antes2", 2),
        getConfigDias("aviso_prazo_dias_repeticao_atrasado", 7),
        getConfigDias("aviso_silencio_dias_sem_atualizar", 5),
        getConfigDias("aviso_silencio_dias_repeticao", 5),
      ]);

    // Projetos ja encerrados (TEP registrado) — ficam de fora.
    const teps: { projeto_id: string | null }[] = await restGet(
      "kv_store?tipo=eq.tep&select=projeto_id"
    );
    const projetosEncerrados = new Set(teps.map((t) => t.projeto_id).filter(Boolean));

    // Nome de cada conta (para a saudacao pessoal do e-mail — "Ola, {primeiro nome}") e papel
    // Gerente de Projetos (para resolver o segundo destinatario, GP da equipe do projeto, sem
    // uma consulta por projeto) — uma unica leitura de perfis para as duas coisas.
    const todosPerfis: { id: string; email: string; nome: string | null; papel: string | null }[] =
      await restGet("perfis?select=id,email,nome,papel");
    const nomePorEmail = new Map<string, string>();
    for (const p of todosPerfis) {
      if (p.email && p.nome) nomePorEmail.set(p.email.toLowerCase(), p.nome);
    }
    function saudacaoPara(email: string): string {
      const nome = nomePorEmail.get(email.toLowerCase());
      return nome ? `Olá, ${esc(nome.split(" ")[0])}.` : "Olá.";
    }
    const gps = todosPerfis.filter((p) => p.papel === "gerente_projetos");
    const idsGP = new Set(gps.map((g) => g.id));
    const equipe: { projeto_id: string; usuario_id: string; usuario_email: string }[] =
      await restGet("projeto_equipe?select=projeto_id,usuario_id,usuario_email");
    const gpsPorProjeto = new Map<string, Set<string>>();
    for (const m of equipe) {
      if (!idsGP.has(m.usuario_id) || !m.usuario_email) continue;
      if (!gpsPorProjeto.has(m.projeto_id)) gpsPorProjeto.set(m.projeto_id, new Set());
      gpsPorProjeto.get(m.projeto_id)!.add(m.usuario_email);
    }

    // Todos os Planejamentos (cronogramas).
    const planos: { key: string; value: string }[] = await restGet(
      "kv_store?tipo=eq.plan&select=key,value"
    );

    for (const row of planos) {
      let rec: PlanoRec;
      try {
        rec = JSON.parse(row.value);
      } catch {
        continue;
      }
      if (!rec.projetoId || projetosEncerrados.has(rec.projetoId)) continue;
      const marcos = Array.isArray(rec.cronograma) ? rec.cronograma : [];
      const planoIdSuffix = row.key.replace(/^plan:/, "");

      for (let idx = 0; idx < marcos.length; idx++) {
        const marco = marcos[idx];
        if (!marco.marco || !marco.fim || marco.concl) continue;
        resultado.avaliados++;

        const marcoId = marco.id || `noid_${planoIdSuffix}_${idx}`;

        const destinatarios = new Set<string>();
        for (const email of parseRespConta(marco.respConta)) destinatarios.add(email);
        for (const email of gpsPorProjeto.get(rec.projetoId) || []) destinatarios.add(email);
        if (!destinatarios.size) continue; // nada a enviar, mas nao marca como avisado

        const alertaKey = `alertaprazo:${marcoId}`;
        let jaAvisado: { buckets?: Partial<Record<Bucket, string>> & { silencio?: string } } = {};
        try {
          const existentes = await restGet(
            `kv_store?key=eq.${encodeURIComponent(alertaKey)}&select=value`
          );
          if (existentes[0]) jaAvisado = JSON.parse(existentes[0].value) || {};
        } catch {
          // sem registro anterior — segue como se nunca tivesse avisado
        }
        const buckets = jaAvisado.buckets || {};
        let mudouBuckets = false;

        const projetoNome = esc(rec.nomeProjeto || "Projeto sem nome");
        const marcoNome = esc(marco.marco);
        const subtituloProjeto = projetoNome + (rec.protocolo ? ` · ${esc(rec.protocolo)}` : "");
        const rodapeComum =
          `<p style="margin:0 0 6px;font-size:11.5px;color:#8A8E94;line-height:1.55">Avaliação automática do Painel de Prazos.</p>` +
          `<p style="margin:0;font-size:11.5px;color:#8A8E94;line-height:1.55">Prefere ajustar direto? Procure a Gerência de Projetos${
            rec.gestor ? ` — Gerente <b style="color:#52525B">${esc(rec.gestor)}</b>` : ""
          }.</p>`;
        const linhasComuns = [
          { label: "Projeto", valor: projetoNome },
          ...(marco.resp ? [{ label: "Responsável", valor: esc(marco.resp) }] : []),
        ];

        // --- 1) Aviso de prazo ---
        let avisouPrazoAgora = false;
        const dias = diasEntre(marco.fim, hoje);
        const bucket = bucketParaDias(dias, diasAntes2, diasAntes1);
        if (bucket) {
          const ultimoEnvioBucket = buckets[bucket];
          const deveAvisar =
            bucket === "atrasado"
              ? !ultimoEnvioBucket ||
                diasEntre(hoje, ultimoEnvioBucket.slice(0, 10)) >= diasRepeticaoAtrasado
              : !ultimoEnvioBucket;
          if (deveAvisar) {
            const dataFim = new Date(marco.fim + "T00:00:00Z").toLocaleDateString("pt-BR", {
              timeZone: "UTC",
            });
            const diasAbs = Math.abs(dias);
            const projNomeRaw = rec.nomeProjeto || rec.protocolo || "um projeto";
            // Tom escalona por severidade: lembrete tranquilo quando falta tempo, direto no dia,
            // e no atrasado soma o framing de visibilidade (accountability) — o status ja esta
            // sendo visto pela Gerencia de Projetos, nao e so um pedido, e uma pendencia exposta.
            let badgeTexto: string, statNumero: string, statLabel: string, corpoBase: string, ctaTexto: string, subject: string, preheader: string;
            if (bucket === "atrasado") {
              badgeTexto = "Atrasado";
              statNumero = String(diasAbs);
              statLabel = diasAbs === 1 ? "dia atrasado" : "dias atrasado";
              corpoBase = `Este marco passou do término previsto (${dataFim}) e já aparece como <b>atrasado</b> no Painel de Prazos, visível para a Gerência de Projetos.`;
              ctaTexto = "Regularizar agora";
              subject = `⚠ Atrasado: marco "${marco.marco}" — ${projNomeRaw}`;
              preheader = `${statNumero} ${statLabel} em ${projNomeRaw}. Um ajuste rápido evita que a pendência continue se acumulando.`;
            } else if (bucket === "nodia") {
              badgeTexto = "Vence hoje";
              statNumero = "HOJE";
              statLabel = "é o prazo";
              corpoBase = `O término previsto deste marco é <b>hoje</b> (${dataFim}). Se já concluiu, marque agora; se ainda está em andamento, um status atualizado evita virar atraso amanhã.`;
              ctaTexto = "Concluir ou atualizar agora";
              subject = `Hoje é o prazo do marco "${marco.marco}" — ${projNomeRaw}`;
              preheader = `O prazo deste marco em ${projNomeRaw} é hoje — só leva um minuto para atualizar.`;
            } else {
              badgeTexto = "Vencendo";
              statNumero = String(dias);
              statLabel = dias === 1 ? "dia para o prazo" : "dias para o prazo";
              corpoBase =
                bucket === "antes2"
                  ? `Faltam poucos dias para o término previsto (${dataFim}). Vale conferir se ainda dá tempo ou se o prazo precisa ser revisto.`
                  : `Ainda dá tempo — um lembrete antecipado para manter a execução em dia até o término previsto (${dataFim}).`;
              ctaTexto = "Atualizar execução agora";
              subject = `${bucket === "antes2" ? "Faltam" : "Lembrete: faltam"} ${dias} dia${dias === 1 ? "" : "s"} — marco "${marco.marco}" — ${projNomeRaw}`;
              preheader = `${dias} dia${dias === 1 ? "" : "s"} até o prazo em ${projNomeRaw}.`;
            }
            const { enviouAlgum, erros } = await enviarParaCadaDestinatario(
              destinatarios,
              subject,
              (destinatario) =>
                emailTemplate({
                  preheader,
                  saudacao: saudacaoPara(destinatario),
                  badgeTexto,
                  badgeTipo: bucket === "atrasado" ? "bad" : "warn",
                  titulo: marcoNome,
                  subtitulo: subtituloProjeto,
                  statNumero,
                  statLabel,
                  corpo: corpoBase,
                  linhas: linhasComuns,
                  ctaTexto,
                  ctaUrl: meuPainelUrlMarco(marcoId),
                  rodape: rodapeComum,
                })
            );
            if (enviouAlgum) {
              buckets[bucket] = new Date().toISOString();
              mudouBuckets = true;
              avisouPrazoAgora = true;
              resultado.avisosEnviados++;
            }
            erros.forEach((e) => resultado.erros.push(`Falha ao notificar marco ${marcoId} (prazo): ${e}`));
          }
        }

        // --- 2) Aviso de silêncio (Execução parada há 5+ dias, marco já iniciado) ---
        if (!avisouPrazoAgora && marco.ini && diasEntre(hoje, marco.ini) >= 0) {
          const referencia = marco.execStatusAtualizadoEm
            ? marco.execStatusAtualizadoEm.slice(0, 10)
            : marco.ini;
          const diasSemAtualizar = diasEntre(hoje, referencia);
          if (diasSemAtualizar >= diasSilencioLimite) {
            const ultimoSilencio = buckets.silencio;
            const deveAvisarSilencio =
              !ultimoSilencio ||
              diasEntre(hoje, ultimoSilencio.slice(0, 10)) >= diasSilencioRepeticao;
            if (deveAvisarSilencio) {
              const projNomeRaw = rec.nomeProjeto || rec.protocolo || "um projeto";
              // Repeticao (buckets.silencio ja tinha um valor) = ja avisamos antes e continua
              // parado — tom mais direto que o primeiro aviso, mesma logica de escalonar por
              // severidade usada no aviso de prazo acima.
              const jaAvisadoAntes = !!ultimoSilencio;
              const subject = `${jaAvisadoAntes ? "Ainda sem novidade" : "Sem atualização"}: marco "${marco.marco}" — ${projNomeRaw}`;
              const preheader = `${diasSemAtualizar} dias sem atualização de Execução em ${projNomeRaw} — um clique resolve.`;
              const corpoBase = jaAvisadoAntes
                ? `Este marco continua sem nenhuma atualização de <b>Execução</b> — já são <b>${diasSemAtualizar} dias</b>. Ninguém sinalizou estar trabalhando nele, e isso já apareceu antes no Painel de Prazos.`
                : `Este marco já começou e está sem nenhuma atualização de <b>Execução</b> há <b>${diasSemAtualizar} dias</b>. Um status rápido evita que ele apareça como parado para o resto da equipe.`;
              const { enviouAlgum, erros } = await enviarParaCadaDestinatario(
                destinatarios,
                subject,
                (destinatario) =>
                  emailTemplate({
                    preheader,
                    saudacao: saudacaoPara(destinatario),
                    badgeTexto: jaAvisadoAntes ? "Ainda sem novidade" : "Sem atualização",
                    badgeTipo: "slate",
                    titulo: marcoNome,
                    subtitulo: subtituloProjeto,
                    statNumero: String(diasSemAtualizar),
                    statLabel: diasSemAtualizar === 1 ? "dia sem atualizar" : "dias sem atualizar",
                    corpo: corpoBase,
                    linhas: linhasComuns,
                    ctaTexto: "Atualizar status agora",
                    ctaUrl: meuPainelUrlMarco(marcoId),
                    rodape: rodapeComum,
                  })
              );
              if (enviouAlgum) {
                buckets.silencio = new Date().toISOString();
                mudouBuckets = true;
                resultado.avisosEnviados++;
              }
              erros.forEach((e) => resultado.erros.push(`Falha ao notificar marco ${marcoId} (silêncio): ${e}`));
            }
          }
        }

        if (mudouBuckets) {
          try {
            await fetch(`${SUPABASE_URL}/rest/v1/kv_store?on_conflict=key`, {
              method: "POST",
              headers: svcHeaders({ Prefer: "resolution=merge-duplicates" }),
              body: JSON.stringify({
                key: alertaKey,
                value: JSON.stringify({ buckets }),
                updated_at: new Date().toISOString(),
              }),
            });
          } catch (e) {
            resultado.erros.push(`Falha ao gravar controle de aviso ${marcoId}: ${String(e)}`);
          }
        }
      }
    }
  } catch (e) {
    return json({ error: String(e), ...resultado }, 500);
  }

  return json({ ok: true, hoje, ...resultado });
});
