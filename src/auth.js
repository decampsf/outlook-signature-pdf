import { PublicClientApplication } from '@azure/msal-browser';
(async () => {
  await Office.onReady();
  const label = document.getElementById('authStatus');
  const guid = /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/i;
  let channel;
  const send = payload => Office.context.ui.messageParent(JSON.stringify({...payload, channel}), {targetOrigin: location.origin});
  try {
    const query = new URLSearchParams(location.search);
    if (query.has('channel')) sessionStorage.setItem('signature-dialog-channel', query.get('channel'));
    channel = sessionStorage.getItem('signature-dialog-channel');
    if (!channel) throw new Error('Ouvrez la connexion depuis le complément Outlook.');
    const {clientId, tenantId} = window.M365_CONFIG || {};
    if (!guid.test(clientId) || !guid.test(tenantId)) throw new Error('La connexion Microsoft 365 doit être configurée.');
    const client = new PublicClientApplication({auth: {
      clientId, authority: `https://login.microsoftonline.com/${tenantId}`,
      redirectUri: new URL('auth.html', location.href).href.split('?')[0].split('#')[0],
      navigateToLoginRequestUrl: false
    }, cache: {cacheLocation: 'sessionStorage'}});
    await client.initialize();
    const result = await client.handleRedirectPromise();
    if (result?.accessToken) {
      if (result.account?.tenantId?.toLowerCase() !== tenantId.toLowerCase()) throw new Error('Le compte appartient à une autre organisation.');
      send({type: 'authenticated', accessToken: result.accessToken, expiresAt: result.expiresOn?.getTime(),
        accountId: result.account.homeAccountId, username: result.account.username});
      label.textContent = 'Connexion terminée. Vous pouvez fermer cette fenêtre.';
    } else {
      label.textContent = 'Connexion à votre compte professionnel…';
      await client.loginRedirect({scopes: ['Files.ReadWrite'], prompt: 'select_account'});
    }
  } catch (error) {
    label.textContent = 'Connexion impossible. ' + (error.errorCode || error.message || 'Réessayez depuis Outlook.');
    if (channel) send({type: 'error', message: label.textContent});
  }
})();
