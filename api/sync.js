import { syncAll } from '../lib/sync.js';

function authorized(req){
  const admin=process.env.SOCIOS_ADMIN_SECRET;
  const cron=process.env.CRON_SECRET; return (admin&&req.headers['x-socios-admin-secret']===admin)||(cron&&req.headers.authorization===`Bearer ${cron}`);
}
export default async function handler(req,res){
  if(req.method!=='POST'&&req.method!=='GET')return res.status(405).json({error:'Method not allowed'});
  if(!authorized(req))return res.status(401).json({error:'Unauthorized'});
  try{return res.status(200).json(await syncAll({force:req.query?.force==='1'}));}
  catch(e){return res.status(500).json({ok:false,error:String(e.message||e)});}
}
