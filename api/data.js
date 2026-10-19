import { readProject, readMeta } from '../lib/store.js';

export default async function handler(req,res){
  if(req.method!=='GET')return res.status(405).json({error:'Method not allowed'});
  res.setHeader('Cache-Control','no-store, max-age=0');
  try{
    const names=['as','colegios','boracay','liquitty','sentencias','cableado'];
    const values=await Promise.all(names.map(async name=>[name,await readProject(name)]));
    const projects=Object.fromEntries(values.filter(([,v])=>v));
    const meta=await readMeta();
    return res.status(200).json({projects,meta});
  }catch(e){
    return res.status(503).json({error:'Live data store not configured',detail:String(e.message||e)});
  }
}
