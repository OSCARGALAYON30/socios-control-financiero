const GRAPH = 'https://graph.microsoft.com/v1.0';

function need(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing environment variable: ${name}`);
  return value;
}

export function graphConfig() {
  return {
    tenantId: need('MS_TENANT_ID'),
    clientId: need('MS_CLIENT_ID'),
    clientSecret: need('MS_CLIENT_SECRET'),
    driveId: process.env.GRAPH_DRIVE_ID || 'b!KYEcvbbv6k6sa9p2DMpR3S7qvuh9x0lKmtPSLxGAbt2bLn10WrsBQYYjvtsA2Myh',
    folderPath: process.env.GRAPH_FOLDER_PATH || 'DOCUMENTOS EMPRESAS AREA FINANCIERA/SOCIOS INFORMES Y DASHBOARD',
    clientState: need('GRAPH_CLIENT_STATE')
  };
}

export async function getGraphToken() {
  const { tenantId, clientId, clientSecret } = graphConfig();
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });
  const r = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body
  });
  if (!r.ok) throw new Error(`Graph token failed: ${r.status} ${await r.text()}`);
  return (await r.json()).access_token;
}

export async function graphFetch(path, options = {}) {
  const token = await getGraphToken();
  const r = await fetch(path.startsWith('http') ? path : GRAPH + path, {
    ...options,
    headers: {
      authorization: `Bearer ${token}`,
      ...(options.body ? { 'content-type': 'application/json' } : {}),
      ...(options.headers || {})
    }
  });
  if (!r.ok) throw new Error(`Graph ${options.method || 'GET'} ${path} failed: ${r.status} ${await r.text()}`);
  if (r.status === 204) return null;
  const type = r.headers.get('content-type') || '';
  return type.includes('json') ? r.json() : r;
}

export async function listSourceFolder() {
  const { driveId, folderPath } = graphConfig();
  const encodedPath = folderPath.split('/').map(encodeURIComponent).join('/');
  let url = `/drives/${encodeURIComponent(driveId)}/root:/${encodedPath}:/children?$top=200&$select=id,name,lastModifiedDateTime,size,file,eTag`;
  const out = [];
  while (url) {
    const data = await graphFetch(url);
    out.push(...(data.value || []));
    url = data['@odata.nextLink'] || null;
  }
  return out.filter(x => x.file);
}

export async function downloadDriveItem(itemId) {
  const { driveId } = graphConfig();
  const token = await getGraphToken();
  const r = await fetch(`${GRAPH}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`, {
    headers: { authorization: `Bearer ${token}` }
  });
  if (!r.ok) throw new Error(`Graph download failed: ${r.status} ${await r.text()}`);
  return Buffer.from(await r.arrayBuffer());
}

export function parseNameDate(name) {
  const m = String(name || '').trim().match(/^(\d{2})(\d{2})(\d{4})_/);
  if (!m) return 0;
  const [,dd,mm,yyyy] = m;
  const ms = Date.UTC(Number(yyyy), Number(mm)-1, Number(dd));
  return Number.isFinite(ms) ? ms : 0;
}

export const SOURCE_RULES = {
  as: n => /_ASF_ASW\s*\.xlsx$/i.test(n),
  colegios: n => /_(?:Resumen Colegios|COLEGIOS)\.xlsx$/i.test(n),
  planes: n => /_ALIANZA PLANES EMPRESARIALES\.xlsx$/i.test(n),
  boracay: n => /_BORACAY\.xlsx$/i.test(n),
  liquitty: n => /_LIQUITTY\.xlsx$/i.test(n),
  cableado: n => /_CABLEADO ESTRUCTURADO\.xlsx$/i.test(n),
  sentencias: n => /PORTAFOLIO DE SENT(?:E|)NCIAS\.xlsx$/i.test(n) || /PORTAFOLIO DE SENTENCIAS\.xlsx$/i.test(n)
};

export function latestForProject(items, project) {
  const match = SOURCE_RULES[project];
  return items.filter(x => match?.(x.name)).sort((a,b) => {
    const d = parseNameDate(b.name) - parseNameDate(a.name);
    if (d) return d;
    return new Date(b.lastModifiedDateTime || 0) - new Date(a.lastModifiedDateTime || 0);
  })[0] || null;
}

function expiration27d() {
  return new Date(Date.now() + 27 * 24 * 60 * 60 * 1000).toISOString();
}

export async function ensureGraphSubscription(notificationUrl) {
  const { driveId, clientState } = graphConfig();
  const resource = `/drives/${driveId}/root`;
  const listed = await graphFetch('/subscriptions');
  const existing = (listed.value || []).find(s => s.resource === resource && s.notificationUrl === notificationUrl);
  const expirationDateTime = expiration27d();

  if (existing) {
    const expires = Date.parse(existing.expirationDateTime || 0);
    if (expires - Date.now() > 7 * 24 * 60 * 60 * 1000) return { action: 'unchanged', subscription: existing };
    const subscription = await graphFetch(`/subscriptions/${existing.id}`, {
      method: 'PATCH',
      body: JSON.stringify({ expirationDateTime })
    });
    return { action: 'renewed', subscription };
  }

  const subscription = await graphFetch('/subscriptions', {
    method: 'POST',
    body: JSON.stringify({
      changeType: 'updated',
      notificationUrl,
      resource,
      expirationDateTime,
      clientState,
      latestSupportedTlsVersion: 'v1_2'
    })
  });
  return { action: 'created', subscription };
}
