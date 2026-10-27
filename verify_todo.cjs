const fs=require('fs'),vm=require('vm'),assert=require('assert');
const elements=new Map;const el=id=>{if(!elements.has(id))elements.set(id,{value:'',innerHTML:'',textContent:'',className:'',dataset:{},addEventListener(){},setAttribute(){},classList:{add(){},remove(){}}});return elements.get(id)};
const context={console,TextDecoder,Date,Blob,URL,setTimeout,document:{getElementById:el,querySelectorAll:()=>[],createElement:()=>({click(){}})},window:{addEventListener(){},echarts:{init(){return {off(){},on(){},resize(){},setOption(o){assert(o.series)}}}}}};context.echarts=context.window.echarts;vm.createContext(context);
vm.runInContext(fs.readFileSync('todo_charts.html','utf8').match(/<script>([\s\S]*)<\/script>/)[1],context);
vm.runInContext(`
if(ATTACHMENTS.length!==3)throw Error('Expected three attachments');
$('demo').onclick();
for(active=0;active<sets.length;active++){
 for(const t of ['mix','quality','volume']){tab=t;path=[];render();if(tableRows.length!==sets[active].rows.length)throw Error('Row count');}
 tab='mix';drill('cat',sets[active].rows[0].cat);if(tableRows.some(r=>r.cat!==sets[active].rows[0].cat))throw Error('Drill filter');
 $('back').onclick();if(path.length)throw Error('Back failed');
 console.log(sets[active].name,sets[active].rows.length,'rows',sets[active].warnings,'date warnings');
}
if(parseCSV('a,b\\r\\n"x,y","line1\\nline2"')[1][0]!=='x,y')throw Error('Quoted CSV');
if(date('2025-02-30')!==null)throw Error('Invalid date accepted');
if(date('2025-01-01')===null)throw Error('Valid date rejected');
if(delivery({end:date('2025-01-01 23:59:59'),due:date('2025-01-01')})!=='準時完成')throw Error('Due day cutoff');
active=0;path=[];$('from').value='2025-02-01';$('to').value='2025-02-28';render();if(tableRows.some(r=>month(r.created)!=='2025-02'))throw Error('Date filter');
$('from').value='2026-01-01';$('to').value='2025-01-01';render();if(tableRows.length)throw Error('Invalid range');
console.log('PASS: all tabs, drill/back, CSV quoting, dates, due-day logic and date filters');
`,context);
