import { waitUntil } from '@vercel/functions';
import { syncAll } from '../lib/sync.js';

export default async function handler(req,res){
  const validationToken=req.query?.validationToken;
  if(validationToken){
    res.setHeader('Content-Type','text/plain');
    return res.status(200).send(String(validationToken));
  }
  if(req.method!=='POST')return res.status(405).end();
  const notifications=Array.isArray(req.body?.value)?req.body.value:[];
  const expected=process.env.GRAPH_CLIENT_STATE;
  const valid=expected&&notifications.some(n=>n.clientState===expected);
  if(!valid)return res.status(202).end();
  waitUntil(syncAll().catch(e=>console.error('SOCIOS webhook sync failed',e)));
  return res.status(202).end();
}
