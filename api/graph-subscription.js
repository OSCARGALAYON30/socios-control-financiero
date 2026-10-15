import { ensureGraphSubscription } from '../lib/graph.js';

function authorized(req){
  const auth=req.headers.authorization||'';
  const cron=process.env.CRON_SECRET;
  const admin=process.env.SOCIOS_ADMIN_SECRET;
  const supplied=req.headers['x-socios-admin-secret'];
  return (cron&&auth===`Bearer ${cron}`)||(admin&&supplied===admin);
}

export default async function handler(req,res){
  if(!authorized(req))return res.status(401).json({error:'Unauthorized'});
  try{
    const notificationUrl=process.env.GRAPH_NOTIFICATION_URL||`https://${req.headers.host}/api/graph-webhook`;
    const result=await ensureGraphSubscription(notificationUrl);
    return res.status(200).json({ok:true,notificationUrl,action:result.action,subscription:{id:result.subscription.id,resource:result.subscription.resource,expirationDateTime:result.subscription.expirationDateTime}});
  }catch(e){
    return res.status(500).json({ok:false,error:String(e.message||e)});
  }
}
