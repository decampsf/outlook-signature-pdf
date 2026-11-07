/* Accès Graph uniquement sur action de l'utilisateur. Jeton conservé en mémoire. */
window.OneDrivePDF = (() => {
  let session = null, connecting = false;
  const root = 'https://graph.microsoft.com/v1.0';
  const pendingKey = 'outlook-signature-onedrive-cleanup-v1';
  const guid = /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/i;
  const configured = () => guid.test(window.M365_CONFIG?.clientId) && guid.test(window.M365_CONFIG?.tenantId);
  function auth() {
    if (!session || session.expiresAt < Date.now() + 30000) throw new Error('Connectez-vous à Microsoft 365 pour convertir en PDF.');
    return session;
  }
  function pending() { try { const items=JSON.parse(localStorage.getItem(pendingKey) || '[]');return Array.isArray(items)?items.filter(e=>typeof e.accountId==='string'&&/^Signature-Outlook-temp-[0-9a-f-]{36}\.(docx|xlsx)$/.test(e.name)):[]; } catch { return []; } }
  function savePending(items) { localStorage.setItem(pendingKey, JSON.stringify(items)); }
  const brief = status => status === 401 ? 'La connexion Microsoft 365 a expiré. Reconnectez-vous.' : status === 403 ? 'Accès OneDrive refusé. Vérifiez les autorisations accordées au complément.' : `Microsoft 365 a refusé l’opération (code ${status}).`;
  async function request(path, options = {}, current = auth()) {
    if (!path.startsWith('/') || path.startsWith('//')) throw new Error('Chemin Microsoft invalide.');
    const response = await fetch(root + path, {...options, credentials: 'omit', cache: 'no-store',
      signal: AbortSignal.timeout(45000), headers: {...options.headers, Authorization: 'Bearer ' + current.accessToken}});
    if (!response.ok) { const error = new Error(brief(response.status)); error.status = response.status; throw error; }
    return response;
  }
  function connect() {
    if (!configured()) return Promise.reject(new Error('Configuration Microsoft 365 nécessaire : voir le guide de configuration.'));
    if (connecting) return Promise.reject(new Error('Une fenêtre de connexion est déjà ouverte.'));
    connecting = true;
    return new Promise((resolve, reject) => {
      const channel = crypto.randomUUID(), url = new URL('auth.html', location.href);
      url.searchParams.set('channel', channel);
      let dialog, finished = false;
      const finish = (error, value) => { if (finished) return; finished = true; clearTimeout(timer); connecting = false; dialog?.close(); error ? reject(error) : resolve(value); };
      const timer = setTimeout(() => finish(new Error('Délai de connexion dépassé. Réessayez.')), 5 * 60 * 1000);
      try { Office.context.ui.displayDialogAsync(url.href, {height: 65, width: 40, displayInIframe: false}, result => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) { finish(new Error('Fenêtre Microsoft inaccessible : ' + (result.error?.code || 'erreur'))); return; }
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, event => {
          if (event.origin && event.origin !== location.origin) return;
          let data; try { data = JSON.parse(event.message); } catch { return; }
          if (data.channel !== channel) return;
          if (data.type === 'error') { finish(new Error(data.message)); return; }
          if (data.type !== 'authenticated' || typeof data.accessToken !== 'string' || !data.accountId || !Number.isFinite(data.expiresAt) || data.expiresAt < Date.now()) { finish(new Error('Réponse de connexion invalide.')); return; }
          session = {accessToken: data.accessToken, accountId: data.accountId, username: data.username, expiresAt: data.expiresAt};
          finish(null, {username: session.username});
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => finish(new Error('Connexion fermée.')));
      }); } catch(error) { finish(new Error('Fenêtre Microsoft inaccessible. ' + error.message)); }
    });
  }
  async function cleanup(entry, current) {
    try {
      await request('/me/drive/root:/' + encodeURIComponent(entry.name), {method: 'DELETE'}, current);
    } catch (error) { if (error.status !== 404) return false; }
    savePending(pending().filter(e => !(e.accountId === entry.accountId && e.name === entry.name)));
    return true;
  }
  async function retryCleanup() {
    const current = auth(); let remaining = 0;
    for (const entry of pending().filter(e => e.accountId === current.accountId)) if (!await cleanup(entry, current)) remaining++;
    return remaining;
  }
  async function convertedBytes(id, current) {
    // La redirection 302 est reçue dans le JSON du batch, sans prévol CORS
    // sur l'URL de téléchargement. Aucun jeton Graph n'est envoyé à cette URL.
    let result;
    for (let attempt = 0; attempt < 4; attempt++) {
      const response = await request('/$batch', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({requests: [{id: 'pdf', method: 'GET', url: '/me/drive/items/' + encodeURIComponent(id) + '/content?format=pdf'}]})}, current);
      result = (await response.json()).responses?.find(r => r.id === 'pdf');
      if (![404,423,429,503,504].includes(result?.status)) break;
      await new Promise(resolve => setTimeout(resolve, 1000 * (attempt + 1)));
    }
    if (!result || ![200,302].includes(result.status)) throw new Error(brief(result?.status || 'inconnu') + ' La conversion PDF n’a pas abouti.');
    let bytes;
    if (result.status === 200 && typeof result.body === 'string') {
      try { bytes = Uint8Array.from(atob(result.body), c => c.charCodeAt(0)); } catch { throw new Error('Réponse PDF invalide.'); }
    } else {
      const locationValue = Object.entries(result.headers || {}).find(([name]) => name.toLowerCase() === 'location')?.[1];
      if (!locationValue) throw new Error('Microsoft 365 n’a pas fourni de lien de conversion.');
      const url = new URL(locationValue);
      if (url.protocol !== 'https:' || url.username || url.password) throw new Error('Lien de conversion invalide.');
      const headers = url.origin === 'https://graph.microsoft.com' ? {Authorization: 'Bearer ' + current.accessToken} : {};
      const response = await fetch(url.href, {headers, credentials: 'omit', cache: 'no-store', referrerPolicy: 'no-referrer', signal: AbortSignal.timeout(60000)});
      if (!response.ok) throw new Error('Le PDF converti n’a pas pu être téléchargé.');
      bytes = new Uint8Array(await response.arrayBuffer());
    }
    if (String.fromCharCode(...bytes.slice(0, 5)) !== '%PDF-') throw new Error('Microsoft 365 n’a pas retourné un PDF valide.');
    return bytes;
  }
  async function convert(bytes, extension) {
    if (!['docx','xlsx'].includes(extension)) throw new Error('Format de conversion non pris en charge.');
    if (bytes.length > 2 * 1024 * 1024) throw new Error('La conversion Microsoft est limitée à 2 Mo. Choisissez le format d’origine ou exportez le PDF dans Word/Excel.');
    const current = auth();
    const entry = {accountId: current.accountId, name: 'Signature-Outlook-temp-' + crypto.randomUUID() + '.' + extension};
    // Enregistrer le nom avant le téléversement permet de nettoyer même si
    // la réponse réseau de l'envoi est perdue. Ne jamais toucher à un original.
    savePending([...pending(), entry]);
    let pdfBytes, primaryError, cleaned;
    try {
      const response = await request('/me/drive/root:/' + encodeURIComponent(entry.name) + ':/content', {method: 'PUT', headers: {'Content-Type': 'application/octet-stream'}, body: bytes}, current);
      const file = await response.json();
      if (!file.id) throw new Error('Copie temporaire non confirmée par OneDrive.');
      pdfBytes = await convertedBytes(file.id, current);
    } catch (error) { primaryError = error; }
    finally { cleaned = await cleanup(entry, current); }
    const warning = cleaned ? '' : 'Une copie temporaire reste dans OneDrive. Reconnectez-vous pour réessayer son nettoyage.';
    if (primaryError) throw new Error(primaryError.message + (warning ? ' ' + warning : ''));
    return {bytes: pdfBytes, warning};
  }
  return {configured, connect, convert, retryCleanup, disconnect: () => {session = null;}, account: () => session?.username || null};
})();
