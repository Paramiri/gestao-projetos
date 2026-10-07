// Edge Function: send-notification
// Envia e-mails transacionais via SMTP do Gmail e, para quem tiver ativado, lembretes
// push no celular via Web Push, para as notificacoes de transicao de estado
// do Sistema de Gestao de Projetos (SGP) (Gate 1, Gate 2, SMP, Canvas, TAP, TEP).
//
// Segredos necessarios (definidos via `supabase secrets set NOME=valor`):
//   SMTP_USER           - conta Gmail remetente (ex.: fulano@gmail.com)
//   SMTP_PASS           - senha de app do Gmail (16 letras, gerada em
//                         myaccount.google.com/apppasswords — nao a senha normal)
//   VAPID_PUBLIC_KEY     )
//   VAPID_PRIVATE_KEY    )  chaves do Web Push (par gerado uma unica vez)
//   VAPID_SUBJECT         )  ex.: mailto:fulano@gmail.com (padrao: mailto:SMTP_USER)
// O envio usa a porta 465 (SMTPS) porque as Edge Functions bloqueiam 25 e 587.
// SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY sao injetados automaticamente pelo
// runtime das Edge Functions — usados para ler push_subscriptions.
//
// Verificacao de JWT permanece ligada (padrao do Supabase): só aceita chamadas
// autenticadas com a anon key ou um token de sessao valido, igual ao resto da API REST.
//
// O envio push é best-effort: se nenhuma chave VAPID estiver configurada, ou se
// um destinatário não tiver assinatura push, a função continua funcionando
// normalmente só com e-mail (comportamento anterior preservado).

import webpush from "npm:web-push@3.6.7";
import nodemailer from "npm:nodemailer@6.9.16";

const SMTP_USER = Deno.env.get("SMTP_USER");
const SMTP_PASS = Deno.env.get("SMTP_PASS");
const FROM_ADDRESS = `"Sistema de Gestão de Projetos" <${SMTP_USER}>`;
const smtp = SMTP_USER && SMTP_PASS
  ? nodemailer.createTransport({ host: "smtp.gmail.com", port: 465, secure: true, auth: { user: SMTP_USER, pass: SMTP_PASS } })
  : null;

const VAPID_PUBLIC_KEY = Deno.env.get("VAPID_PUBLIC_KEY");
const VAPID_PRIVATE_KEY = Deno.env.get("VAPID_PRIVATE_KEY");
const VAPID_SUBJECT = Deno.env.get("VAPID_SUBJECT") || `mailto:${SMTP_USER}`;

const SUPABASE_URL = Deno.env.get("SUPABASE_URL");
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

const PUSH_ATIVO = Boolean(VAPID_PUBLIC_KEY && VAPID_PRIVATE_KEY && SUPABASE_URL && SUPABASE_SERVICE_ROLE_KEY);
if (VAPID_PUBLIC_KEY && VAPID_PRIVATE_KEY) {
  webpush.setVapidDetails(VAPID_SUBJECT, VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY);
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

function htmlParaTexto(html: string): string {
  const texto = html
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<br\s*\/?>/gi, " ")
    .replace(/<\/p>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
  return texto.length > 160 ? texto.slice(0, 157) + "..." : texto;
}

async function buscarAssinaturasPush(emails: string[]): Promise<{ endpoint: string; p256dh: string; auth: string }[]> {
  if (!PUSH_ATIVO || !emails.length) return [];
  try {
    const lista = emails.map((e) => `"${e.replace(/"/g, '\\"')}"`).join(",");
    const r = await fetch(
      `${SUPABASE_URL}/rest/v1/push_subscriptions?user_email=in.(${lista})&select=endpoint,p256dh,auth`,
      { headers: { apikey: SUPABASE_SERVICE_ROLE_KEY!, Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}` } }
    );
    if (!r.ok) return [];
    return await r.json();
  } catch {
    return [];
  }
}

async function removerAssinatura(endpoint: string) {
  try {
    await fetch(`${SUPABASE_URL}/rest/v1/push_subscriptions?endpoint=eq.${encodeURIComponent(endpoint)}`, {
      method: "DELETE",
      headers: { apikey: SUPABASE_SERVICE_ROLE_KEY!, Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}` },
    });
  } catch {
    // melhor esforço — se falhar, a assinatura morta só será limpa na próxima tentativa
  }
}

async function enviarPush(subs: { endpoint: string; p256dh: string; auth: string }[], titulo: string, corpo: string) {
  const payload = JSON.stringify({ title: titulo, body: corpo, url: "./index.html" });
  await Promise.all(
    subs.map(async (s) => {
      try {
        await webpush.sendNotification({ endpoint: s.endpoint, keys: { p256dh: s.p256dh, auth: s.auth } }, payload);
      } catch (e: any) {
        if (e && (e.statusCode === 404 || e.statusCode === 410)) {
          await removerAssinatura(s.endpoint);
        } else {
          console.error("Falha ao enviar push:", e);
        }
      }
    })
  );
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  let payload: { to?: unknown; subject?: unknown; html?: unknown };
  try {
    payload = await req.json();
  } catch {
    return json({ error: "Corpo da requisição não é um JSON válido" }, 400);
  }

  const rawTo = Array.isArray(payload.to) ? payload.to : [payload.to];
  const recipients = Array.from(
    new Set(rawTo.filter((e): e is string => typeof e === "string" && e.includes("@")))
  );
  const subject = typeof payload.subject === "string" ? payload.subject.trim() : "";
  const html = typeof payload.html === "string" ? payload.html : "";

  if (!recipients.length || !subject || !html) {
    return json({ error: "Parâmetros inválidos — 'to' (com ao menos 1 e-mail), 'subject' e 'html' são obrigatórios" }, 400);
  }

  let emailId: unknown = null;
  let emailError: unknown = null;
  if (!smtp) {
    emailError = "SMTP_USER/SMTP_PASS não configurados nos secrets da function";
  } else {
    try {
      const info = await smtp.sendMail({ from: FROM_ADDRESS, to: recipients, subject, html });
      emailId = info.messageId;
    } catch (e) {
      console.error("Falha ao enviar e-mail via SMTP do Gmail:", e);
      emailError = String(e);
    }
  }

  let pushEnviados = 0;
  try {
    const subs = await buscarAssinaturasPush(recipients);
    if (subs.length) {
      await enviarPush(subs, subject, htmlParaTexto(html));
      pushEnviados = subs.length;
    }
  } catch (e) {
    console.error("Falha ao enviar notificações push:", e);
  }

  return json({ ok: !emailError, emailId, emailError, pushEnviados });
});
