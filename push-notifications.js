// Shim de notificações push (lembretes no celular) do Sistema de Gestão de Projetos (SGP).
// Independente do unialfaAuth (usado hoje só em index.html) — recebe email/token de quem chamar.
window.unialfaPush = (function () {
  var SUPABASE_URL = 'https://fiarntunpqteopwjkhjg.supabase.co';
  var SUPABASE_ANON_KEY = 'sb_publishable_RO-UPexCYhZ0rVZiIYWunA_a7YqodnQ';
  var VAPID_PUBLIC_KEY = 'BFCF2XQYAGZ_eB0qk56DfEXMS-Psi7GQFkFPdkXy26VcZ1fRc-YNBSN1BhxJnIo24s2djYBRHmfb-AVTjFpuv9g';

  function suportado() {
    return 'serviceWorker' in navigator && 'PushManager' in window && 'Notification' in window;
  }

  // iOS só habilita Push/Notification API depois que o site foi adicionado à Tela de Início.
  function precisaInstalarNoIOS() {
    var iOS = /iphone|ipad|ipod/i.test(navigator.userAgent);
    var standalone = window.navigator.standalone === true || window.matchMedia('(display-mode: standalone)').matches;
    return iOS && !standalone;
  }

  function urlBase64ToUint8Array(base64String) {
    var padding = '='.repeat((4 - (base64String.length % 4)) % 4);
    var base64 = (base64String + padding).replace(/-/g, '+').replace(/_/g, '/');
    var rawData = atob(base64);
    var outputArray = new Uint8Array(rawData.length);
    for (var i = 0; i < rawData.length; ++i) outputArray[i] = rawData.charCodeAt(i);
    return outputArray;
  }

  async function registrarSW() {
    return navigator.serviceWorker.register('service-worker.js');
  }

  async function assinaturaAtual() {
    if (!suportado()) return null;
    var reg = await navigator.serviceWorker.getRegistration();
    if (!reg) return null;
    return reg.pushManager.getSubscription();
  }

  async function status() {
    if (!suportado()) return precisaInstalarNoIOS() ? 'precisa_instalar_ios' : 'indisponivel';
    if (Notification.permission === 'denied') return 'negado';
    var sub = await assinaturaAtual();
    return sub ? 'ativo' : 'inativo';
  }

  async function ativar(email, accessToken) {
    if (!suportado()) {
      throw new Error(
        precisaInstalarNoIOS()
          ? 'No iPhone/iPad, primeiro adicione este site à Tela de Início (compartilhar → Adicionar à Tela de Início) para poder ativar notificações.'
          : 'Este navegador não é compatível com notificações push.'
      );
    }
    var permissao = await Notification.requestPermission();
    if (permissao !== 'granted') throw new Error('Permissão de notificação negada.');

    var reg = await registrarSW();
    await navigator.serviceWorker.ready;
    var sub = await reg.pushManager.getSubscription();
    if (!sub) {
      sub = await reg.pushManager.subscribe({
        userVisibleOnly: true,
        applicationServerKey: urlBase64ToUint8Array(VAPID_PUBLIC_KEY),
      });
    }
    var keys = sub.toJSON().keys;
    await fetch(SUPABASE_URL + '/rest/v1/push_subscriptions', {
      method: 'POST',
      headers: {
        apikey: SUPABASE_ANON_KEY,
        Authorization: 'Bearer ' + (accessToken || SUPABASE_ANON_KEY),
        'Content-Type': 'application/json',
        Prefer: 'resolution=merge-duplicates',
      },
      body: JSON.stringify({ user_email: email, endpoint: sub.endpoint, p256dh: keys.p256dh, auth: keys.auth }),
    });
    return sub;
  }

  async function desativar(accessToken) {
    var sub = await assinaturaAtual();
    if (!sub) return;
    var endpoint = sub.endpoint;
    await sub.unsubscribe();
    await fetch(SUPABASE_URL + '/rest/v1/push_subscriptions?endpoint=eq.' + encodeURIComponent(endpoint), {
      method: 'DELETE',
      headers: { apikey: SUPABASE_ANON_KEY, Authorization: 'Bearer ' + (accessToken || SUPABASE_ANON_KEY) },
    });
  }

  return { suportado: suportado, precisaInstalarNoIOS: precisaInstalarNoIOS, status: status, ativar: ativar, desativar: desativar };
})();
