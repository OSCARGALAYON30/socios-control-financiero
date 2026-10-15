export default function handler(req,res){
  res.setHeader('Cache-Control','no-store');
  const env=['MS_TENANT_ID','MS_CLIENT_ID','MS_CLIENT_SECRET','GRAPH_CLIENT_STATE','SOCIOS_ADMIN_SECRET'];
  return res.status(200).json({
    ok:true,
    service:'SOCIOS live backend',
    configured:Object.fromEntries(env.map(k=>[k,Boolean(process.env[k])])),
    driveId:Boolean(process.env.GRAPH_DRIVE_ID)||true,
    folderPath:Boolean(process.env.GRAPH_FOLDER_PATH)||true,
    blobHint:'Requires a private Vercel Blob store connected to this project.',
    now:new Date().toISOString()
  });
}
