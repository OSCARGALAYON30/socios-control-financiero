import { get, put } from '@vercel/blob';

const prefix = 'socios-live';

async function readText(pathname) {
  const result = await get(pathname, { access: 'private', useCache: false });
  if (!result) return null;
  return new Response(result.stream).text();
}

export async function readJson(pathname) {
  try {
    const text = await readText(pathname);
    return text ? JSON.parse(text) : null;
  } catch (e) {
    if (/not found|404/i.test(String(e))) return null;
    throw e;
  }
}

export async function writeJson(pathname, value) {
  return put(pathname, JSON.stringify(value), {
    access: 'private',
    contentType: 'application/json',
    allowOverwrite: true
  });
}

export const projectPath = project => `${prefix}/projects/${project}.json`;
export const metaPath = `${prefix}/meta.json`;

export async function readProject(project) {
  return readJson(projectPath(project));
}
export async function writeProject(project, data) {
  return writeJson(projectPath(project), data);
}
export async function readMeta() {
  return (await readJson(metaPath)) || { projects: {}, updatedAt: null };
}
export async function writeMeta(meta) {
  return writeJson(metaPath, meta);
}
