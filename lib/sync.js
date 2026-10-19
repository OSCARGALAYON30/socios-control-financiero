import { listSourceFolder, latestForProject, downloadDriveItem } from './graph.js';
import { importWorkbook } from './importers.js';
import { readMeta, writeMeta, writeProject } from './store.js';

const PROJECTS=['as','colegios','boracay','liquitty','sentencias','cableado'];

function versionOf(item){
  return [item.id,item.eTag||'',item.lastModifiedDateTime||'',item.size||0].join('|');
}

export async function syncAll({force=false}={}) {
  const items=await listSourceFolder();
  const meta=await readMeta();
  const results=[];
  for(const project of PROJECTS){
    const latest=latestForProject(items,project);
    if(!latest){results.push({project,status:'missing'});continue;}
    const version=versionOf(latest);
    const previous=meta.projects?.[project];
    if(!force&&previous?.version===version){results.push({project,status:'unchanged',source:latest.name});continue;}
    try{
      const buffer=await downloadDriveItem(latest.id);
      const data=importWorkbook(project,buffer,latest.name);
      data.live={sourceId:latest.id,version,lastModifiedDateTime:latest.lastModifiedDateTime,syncedAt:new Date().toISOString()};
      await writeProject(project,data);
      meta.projects=meta.projects||{};
      meta.projects[project]={source:latest.name,version,itemId:latest.id,lastModifiedDateTime:latest.lastModifiedDateTime,syncedAt:data.live.syncedAt,status:'ok'};
      results.push({project,status:'updated',source:latest.name});
    }catch(e){
      meta.projects=meta.projects||{};
      meta.projects[project]={...(previous||{}),source:latest.name,attemptedVersion:version,lastError:String(e.message||e),lastAttemptAt:new Date().toISOString(),status:'error'};
      results.push({project,status:'error',source:latest.name,error:String(e.message||e)});
    }
  }
  meta.updatedAt=new Date().toISOString();
  await writeMeta(meta);
  return {ok:results.every(x=>x.status!=='error'),results,updatedAt:meta.updatedAt};
}
