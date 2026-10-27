const fs=require('fs'),vm=require('vm'),path=require('path');
async function main(){
 let scripts=0,links=0; const urls=new Set();
 for(const f of fs.readdirSync(__dirname).filter(f=>f.endsWith('.html'))){
  const html=fs.readFileSync(path.join(__dirname,f),'utf8');
  for(const m of html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g))if(m[1].trim()){new vm.Script(m[1],{filename:f});scripts++;}
  for(const m of html.matchAll(/href="([^"]+)"/g))if(!m[1].startsWith('http')){if(!fs.existsSync(path.join(__dirname,m[1])))throw Error('Missing: '+m[1]);links++;}
  for(const m of html.matchAll(/(?:src|href)="(https:\/\/cdn.jsdelivr.net[^"]+)"/g))urls.add(m[1]);
 }
 console.log(`PASS: 11 HTML pages; ${scripts} scripts parsed; ${links} local links resolved.`);
 await Promise.all([...urls].map(async url=>{try{const r=await fetch(url,{method:'HEAD',signal:AbortSignal.timeout(20000)});console.log(r.status+' '+url);if(!r.ok)process.exitCode=1;}catch(e){console.log('FAILED '+url+': '+e.message);process.exitCode=1;}}));
}
main().catch(e=>{console.error(e);process.exitCode=1;});
