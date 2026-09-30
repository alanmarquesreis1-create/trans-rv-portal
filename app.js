const KEY='transrv.portal.v2';
const DEFAULT_RULES=[
 {id:'AMZ-01',client:'AMAZON',priority:'1',rule:'OPB se subtransportadora começa com OB/OPB'},
 {id:'AMZ-02',client:'AMAZON',priority:'2',rule:'RTP se Travel ID começa com T- ou conta contém OutboundReturns'},
 {id:'AMZ-03',client:'AMAZON',priority:'3',rule:'LH quando não for OPB nem RTP'},
 {id:'MELI-01',client:'MERCADO LIVRE',priority:'1',rule:'CPL quando serviço identifica CPL'},
 {id:'MELI-02',client:'MERCADO LIVRE',priority:'2',rule:'C195/C173/C18/C14 identifica família DEDICADO'},
 {id:'MELI-03',client:'MERCADO LIVRE',priority:'3',rule:'MULTISTOP quando há múltiplas localidades e não é dedicado'},
 {id:'MELI-04',client:'MERCADO LIVRE',priority:'4',rule:'LH como fallback'},
 {id:'SHP-01',client:'SHOPEE',priority:'1',rule:'Todo serviço elegível Shopee = LINE_HAUL'}
];
const DEFAULT_LOCALITIES=[
 ['CORDOVIL - RJ','RIO DE JANEIRO - RJ'],['JURUBATUBA - SP','SAO PAULO - SP'],['PETROPOLIS - RJ','RIO DE JANEIRO - RJ'],['TAQUARA - RJ','RIO DE JANEIRO - RJ'],['PARQUE NOVO MUNDO - SP','SAO PAULO - SP'],['CAMPO GRANDE - RJ','RIO DE JANEIRO - RJ'],['JACAREPAGUA - RJ','RIO DE JANEIRO - RJ'],['ZONA SUL - SP','SAO PAULO - SP'],['VILA GUILHERME - SP','SAO PAULO - SP'],['VILA_GUILHERME - SP','SAO PAULO - SP'],['ARTHUR ALVIM - SP','SAO PAULO - SP'],['ARTHUR ALVIN - SP','SAO PAULO - SP'],['ARENA - SP','BARUERI - SP'],['VIX1','CARIACICA - ES'],['SCH9','CACHOEIRO DE ITAPEMIRIM - ES']
];
const DEFAULT_TYPES=[['CARRETA','CAVALO MECANICO'],['CARRETA 32 PALLETS','CAVALO MECANICO']];
const state={files:[],rows:[],minutes:[],freight:[],nodes:[],treatments:[],history:load('history',[])};

function load(k,d){try{const x=JSON.parse(localStorage.getItem(KEY+'.'+k));return x??d}catch{return d}}
function save(k,v){localStorage.setItem(KEY+'.'+k,JSON.stringify(v))}
let cfg={rules:load('rules',DEFAULT_RULES),localities:load('localities',DEFAULT_LOCALITIES),types:load('types',DEFAULT_TYPES)};
const $=s=>document.querySelector(s), $$=s=>document.querySelectorAll(s);
const norm=v=>String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/\s+/g,' ').trim();
const compact=v=>norm(v).replace(/[^A-Z0-9]/g,'');
function toast(msg){const t=$('#toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),2400)}
function audit(action,detail){state.history.unshift({data:new Date().toLocaleString('pt-BR'),acao:action,detalhe:detail});state.history=state.history.slice(0,300);save('history',state.history);renderHistory()}
function setView(id){$$('.view').forEach(v=>v.classList.toggle('active',v.id===id));$$('.nav').forEach(b=>b.classList.toggle('active',b.dataset.view===id));render()}
$$('.nav').forEach(b=>b.onclick=()=>setView(b.dataset.view));
function readFile(file){return file.arrayBuffer().then(buf=>{const wb=XLSX.read(buf,{type:'array',cellDates:false});const sheets=wb.SheetNames;const tokens=['CODIGO','CNPJ','NOME DA TABELA','TIPO VEICULO','STATUS','MINUTA','LH TRIP','ID DA VIAGEM','SUBTRANSPORTADORA','CONTA DO EXPEDIDOR','NODE','CIDADE','UF','SERVICO','SERVIÇO','TRAVEL ID','ROTA','ORIGEM','DESTINO','TIPO DE SERVICO'];let bestSheet='',bestMatrix=[],bestHeader=0,bestScore=-1;for(const sheetName of sheets){const matrix=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,defval:'',raw:true});for(let i=0;i<Math.min(matrix.length,15);i++){const row=matrix[i].map(v=>norm(v));const score=tokens.reduce((n,t)=>n+(row.some(v=>v===t||v.includes(t))?1:0),0);if(score>bestScore){bestScore=score;bestSheet=sheetName;bestMatrix=matrix;bestHeader=i}}}const rawHeaders=bestMatrix[bestHeader]||[];const headers=rawHeaders.map((h,i)=>{const s=String(h??'').trim();return s||('COLUNA_'+(i+1))});const data=bestMatrix.slice(bestHeader+1).filter(r=>r.some(v=>String(v??'').trim()!=='')).map(r=>Object.fromEntries(headers.map((h,i)=>[h,r[i]??''])));return {name:file.name,sheets,selectedSheet:bestSheet,headerIndex:bestHeader,headers,data}})}
function detectFileRole(name,rows){const s=norm(name+' '+Object.keys(rows[0]||{}).join(' '));const keys=Object.keys(rows[0]||{}).map(norm);const has=t=>keys.some(k=>k===t||k.includes(t));if(has('CODIGO')&&has('NOME DA TABELA')&&has('TIPO VEICULO'))return 'FREIGHT';if(has('MINUTA')||has('NUMERO MINUTA')||has('CODIGO TABELA'))return 'MINUTES';if(has('NODE')&&has('CIDADE')&&has('UF'))return 'NODES';if(s.includes('RELAY')||has('SUBTRANSPORTADORA')&&has('CONTA DO EXPEDIDOR')&&has('ID DA VIAGEM'))return 'AMAZON_FINAL';if(s.includes('AMAZON')||has('TIPO DE SERVICO')&&has('ID DA CARGA'))return 'AMAZON_DEMANDA';if(s.includes('MONITOING BILLING')||has('ROSTERING ID')&&has('TRANSPORTADOR')&&has('SERVICO'))return 'MELI_FINAL';if(s.includes('MERCADO LIVRE')||has('TRAVEL ID')&&has('SERVICO')&&has('PARCEIRO'))return 'MELI_DEMANDA';if(s.includes('SHOPEE')||has('LH TRIP')&&has('3PL')&&has('ENDERECO ORIGEM'))return 'SHOPEE_DEMANDA';return 'DEMANDA'}
function fileClient(name,rows){const role=detectFileRole(name,rows);if(role.startsWith('AMAZON'))return 'AMAZON';if(role.startsWith('MELI'))return 'MERCADO LIVRE';if(role.startsWith('SHOPEE'))return 'SHOPEE';return 'OUTRO'}
function val(row,aliases){const keys=Object.keys(row);for(const a of aliases){const na=norm(a);const k=keys.find(x=>norm(x)===na);if(k)return row[k]}for(const a of aliases){const na=compact(a);const k=keys.find(x=>compact(x).includes(na));if(k)return row[k]}return ''}
function locality(raw){const n=norm(raw).replace(/Á/g,'A');const pair=cfg.localities.find(x=>norm(x[0])===n);return pair?pair[1]:n}
function typeMap(raw){const n=norm(raw);const p=cfg.types.find(x=>norm(x[0])===n);return p?p[1]:n||''}
function detectOperation(client,row){
 const service=norm(val(row,['ROTA','SERVICO','SERVIÇO','SERVICE','ACCOUNT','CONTA DO EXPEDIDOR','SERVICE NAME']));
 const sub=norm(val(row,['SUBTRANSPORTADORA','3PL','SUBCARRIER']));
 const travel=norm(val(row,['ID DA VIAGEM RELAY','TRAVEL ID','TRAVELID','ID VIAGEM']));
 const account=norm(val(row,['CONTA DO EXPEDIDOR','ACCOUNT','INDICADOR']));
 if(client==='SHOPEE')return {op:'LINE_HAUL',reason:'SHOPEE-R01'};
 if(client==='AMAZON'){
  if(/^OB|^OPB/.test(sub))return {op:'OPB',reason:'AMZ-01'};
  if(/^T-/.test(travel)||account.includes('OUTBOUNDRETURNS'))return {op:'RTP',reason:'AMZ-02'};
  return {op:'LH',reason:'AMZ-03'};
 }
 if(client==='MERCADO LIVRE'){
  if(/CPL/.test(service))return {op:'CPL',reason:'MELI-01'};
  const m=service.match(/C(195|173|18|14)(?:[.][0-9]+)?/);
  if(m)return {op:'DEDICADO '+m[1],reason:'MELI-02'};
  const locations=(service.match(/[A-Z]{2,}[0-9]{1,}/g)||[]).length;
  if(service.includes('@')||locations>2)return {op:'MULTISTOP',reason:'MELI-03'};
  return {op:'LH',reason:'MELI-04'};
 }
 return {op:'ANALISE',reason:'SEM_REGRA'};
}
function processRow(row,fileName){
 const client=fileClient(fileName,[row]);
 const op=detectOperation(client,row);
 const rawOrigin=val(row,['ORIGEM','CIDADE ORIGEM','ORIGIN','ORIGIN CITY','ENDEREÇO ORIGEM']);
 const rawDest=val(row,['DESTINO','CIDADE DESTINO','DESTINATION','DESTINATION CITY','ENDEREÇO DESTINO']);
 const origin=locality(rawOrigin),dest=locality(rawDest);
 const rawType=val(row,['TIPO VEICULO','TIPO VEÍCULO','VEICULO','VEÍCULO PLANEJADO','TIPOLOGIA','VEHICLE TYPE']);
 const typology=typeMap(rawType);
 const trip=String(val(row,['VRID','LH TRIP','TRIP ID','ID VIAGEM','VIAGEM','LOAD ID','ID']));
 const minute=String(val(row,['MINUTA','MINUTE','NUMERO MINUTA','Nº MINUTA','MINUTA TMS']));
 const node=String(val(row,['NODE','SITE','ORIGEM NODE']));
 let nodeCity='',nodeAlert='';
 if(client==='AMAZON'&&node){const n=state.nodes.find(x=>compact(x.NODE||x.Node||x.node)===compact(node));if(n){nodeCity=(n.CIDADE||n.Cidade||n.city)+' - '+(n.UF||n.Uf||n.uf)}else nodeAlert='NODE_AMAZON_NAO_MAPEADO'}
 let table=findFreight(client,op.op,origin,dest,typology,serviceText(row));
 let result=op==='ANALISE'?'BLOQUEIO':nodeAlert?'BLOQUEIO':table.code?'CLASSIFICADO':'BLOQUEIO';
 return {...row,CLIENTE:client,ID_VIAGEM:trip,MINUTA:minute,OPERACAO:op.op,REGRA:op.reason,ORIGEM_NORMALIZADA:origin,DESTINO_NORMALIZADO:dest,TIPOLOGIA_RAW:rawType,TIPOLOGIA_NORMALIZADA:typology,NODE:node,NODE_CIDADE_UF:nodeCity,TABELA_ESPERADA:table.code,TABELAS_ALTERNATIVAS:table.alt||'',RESULTADO:result,MINUTA_STATUS:minute?'PENDENTE_REVALIDACAO':'MINUTA_PENDENTE',MOTIVO:nodeAlert||(!table.code?'TABELA_NAO_LOCALIZADA':'OK'),ARQUIVO_ORIGEM:fileName};
}
function serviceText(row){return norm(val(row,['ROTA','SERVICO','SERVIÇO','SERVICE','NOME DA TABELA','SERVICE NAME']))}
function freightClientMatches(raw,client){const c=norm(raw);if(!c)return true;const groups={AMAZON:['AMAZON LOGISTICA DO BRASIL LTDA','AMAZON SERVICOS DE VAREJO DO BRASIL LTDA'],'MERCADO LIVRE':['EBAZAR.COM.BR.LTDA'],SHOPEE:['SHPX LOGISTICA LTDA','SHPX LOGISTICA LTDA.']};return (groups[client]||[client]).some(x=>c===norm(x))}
function operationTableMatches(name,op,service){const n=norm(name);if(!op||op==='ANALISE')return true;if(op==='OPB')return /^OPB\\b/.test(n);if(op==='RTP')return /^RTP\\b/.test(n);if(op==='LH')return /^LH\\b/.test(n);if(op==='CPL')return /^CPL\\b/.test(n);if(op==='MULTISTOP')return n.includes('MULT');if(op==='LINE_HAUL')return /^LH\\b/.test(n);if(op.startsWith('DEDICADO ')){const family=op.split(' ')[1];return n.includes('DED')&&n.includes('C'+family)}return false}
function findFreight(client,op,origin,dest,type,service){
 if(!state.freight.length)return {code:'',alt:''};
 const today=new Date();
 const cand=state.freight.filter(r=>{
  const rawClient=r.CLIENTE||r.Cliente||r.client;const ro=norm(r['CIDADE ORIGEM']||r.ORIGEM||r.CidadeOrigem);const rd=norm(r['CIDADE DESTINO']||r.DESTINO||r.CidadeDestino);const rt=norm(r['TIPO VEICULO']||r['TIPO VEÍCULO']||r.TIPOLOGIA||r['Tipo Veículo']);const name=norm(r['NOME DA TABELA']||r['Nome da Tabela']||r.NOME||r.TABELA);const status=norm(r.STATUS||r.Status);const expiry=String(r['DATA VALIDADE']||r['Data Validade']||'');
  if(!freightClientMatches(rawClient,client))return false;
  if(ro&&origin&&locality(ro)!==origin)return false;
  if(rd&&dest&&locality(rd)!==dest)return false;
  if(rt&&type&&!rt.includes(type))return false;
  if(status&&status!=='ATIVO')return false;
  if(expiry){const m=expiry.match(/(\\d{2})[\\/.-](\\d{2})[\\/.-](\\d{4})/);if(m&&new Date(+m[3],+m[2]-1,+m[1])<today)return false;}
  if(!operationTableMatches(name,op,service))return false;
  if(op.startsWith('DEDICADO ')&&service){const family=op.split(' ')[1];if(!service.includes('C'+family))return false;}
  return true;
 });
 cand.sort((a,b)=>Number(a.CODIGO||a.Código||a.CODIGO_TABELA||999999)-Number(b.CODIGO||b.Código||b.CODIGO_TABELA||999999));
 const codes=[...new Set(cand.map(r=>String(r.CODIGO||r.Código||r.CODIGO_TABELA||r.codigo||'')).filter(Boolean))];
 return {code:codes[0]||'',alt:codes.slice(1).join(',')};
}
function applyMinutes(){
 const byTrip=new Map(),byMinute=new Map();
 state.minutes.forEach(m=>{const trip=String(val(m,['VRID','LH TRIP','TRIP ID','ID VIAGEM','VIAGEM','LOAD ID','ID']));const min=String(val(m,['MINUTA','MINUTE','NUMERO MINUTA','Nº MINUTA']));if(trip)byTrip.set(compact(trip),m);if(min)byMinute.set(compact(min),m)});
 state.rows.forEach(r=>{
  const m=byTrip.get(compact(r.ID_VIAGEM))||byMinute.get(compact(r.MINUTA));
  if(m){r.MINUTA=String(val(m,['MINUTA','MINUTE','NUMERO MINUTA','Nº MINUTA']))||r.MINUTA;r.MINUTA_STATUS='MINUTA_LOCALIZADA';r.TABELA_TMS=String(val(m,['CODIGO TABELA','TABELA','TABLE','COD TABELA']));r.POS_VALIDACAO=r.TABELA_TMS&&r.TABELA_ESPERADA&&r.TABELA_TMS===r.TABELA_ESPERADA?'OK':'DIVERGENTE';}
  else r.MINUTA_STATUS='MINUTA_PENDENTE';
 });
 state.outside=state.minutes.filter(m=>{const id=String(val(m,['VRID','LH TRIP','TRIP ID','ID VIAGEM','VIAGEM','LOAD ID','ID']));return id&&!state.rows.some(r=>compact(r.ID_VIAGEM)===compact(id))}).map(m=>({...m,SITUACAO:'MINUTA_FORA_DA_DEMANDA'}));
 state.treatments=state.rows.filter(r=>r.RESULTADO==='BLOQUEIO'||r.POS_VALIDACAO==='DIVERGENTE'||r.MOTIVO==='NODE_AMAZON_NAO_MAPEADO').map(r=>({ID_VIAGEM:r.ID_VIAGEM,CLIENTE:r.CLIENTE,MINUTA:r.MINUTA,OPERACAO:r.OPERACAO,TABELA_ESPERADA:r.TABELA_ESPERADA,TABELA_TMS:r.TABELA_TMS||'',MOTIVO:r.POS_VALIDACAO==='DIVERGENTE'?'TABELA_TMS_DIVERGENTE':r.MOTIVO,CATEGORIA:r.MOTIVO==='NODE_AMAZON_NAO_MAPEADO'?'DADO_CADASTRAL':'CLASSIFICACAO',AREA_SUGERIDA:r.MOTIVO==='NODE_AMAZON_NAO_MAPEADO'?'CADASTRO':'FATURAMENTO',STATUS:'PENDENTE'}));
}
function analyze(){
 if(!state.files.length)return toast('Selecione os arquivos primeiro.');
 Promise.all(state.files.map(readFile)).then(results=>{
  state.rows=[];state.minutes=[];state.freight=[];state.outside=[];state.treatments=[];
  let demand=0,summary=[];
  results.forEach(x=>{
   const role=detectFileRole(x.name,x.data);summary.push(x.name+' → '+role);
   if(role==='MINUTES'){state.minutes.push(...x.data);return}
   if(role==='FREIGHT'){state.freight.push(...x.data);return}
   if(role==='NODES'){state.nodes=x.data;return}
   if(role==='AMAZON_DEMANDA'&&results.some(y=>detectFileRole(y.name,y.data)==='AMAZON_FINAL'))return;
   if(role==='MELI_DEMANDA'&&results.some(y=>detectFileRole(y.name,y.data)==='MELI_FINAL'))return;
   if(role.endsWith('_FINAL')||role.endsWith('_DEMANDA')||role==='SHOPEE_DEMANDA'||role==='DEMANDA'){
    demand+=x.data.length;
    x.data.forEach(r=>state.rows.push(processRow(r,x.name)));
   }
  });
  applyMinutes();const now=new Date();$('#periodBadge').textContent=now.toLocaleDateString('pt-BR',{month:'2-digit',day:'2-digit'});save('rows',state.rows);save('minutes',state.minutes);save('freight',state.freight);save('nodes',state.nodes);audit('ANALISE',demand+' registros processados • '+summary.join(' | '));toast('Análise concluída');render();
 }).catch(e=>{console.error(e);toast('Erro ao ler arquivo: '+e.message)});
}
function table(el,rows,limit=500){
 const t=$(el);if(!rows.length){t.innerHTML='<tr><td class="empty">Nenhum registro.</td></tr>';return}
 const cols=Object.keys(rows[0]).slice(0,18);const data=rows.slice(0,limit);
 t.innerHTML='<thead><tr>'+cols.map(c=>'<th>'+escapeHtml(c)+'</th>').join('')+'</tr></thead><tbody>'+data.map(r=>'<tr>'+cols.map(c=>'<td>'+cell(r[c])+'</td>').join('')+'</tr>').join('')+'</tbody>';
}
function cell(v){const s=String(v??'');if(['CLASSIFICADO','OK','MINUTA_LOCALIZADA'].includes(s))return '<span class="badge ok">'+escapeHtml(s)+'</span>';if(['PENDENCIA OPERACIONAL','PENDENTE_REVALIDACAO','MINUTA_PENDENTE','DIVERGENTE'].includes(s))return '<span class="badge warn">'+escapeHtml(s)+'</span>';if(['BLOQUEIO','NODE_AMAZON_NAO_MAPEADO','TABELA_NAO_LOCALIZADA'].includes(s))return '<span class="badge bad">'+escapeHtml(s)+'</span>';return escapeHtml(s)}
function escapeHtml(s){return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}
function renderKpis(){const n=state.rows.length,classified=state.rows.filter(r=>r.RESULTADO==='CLASSIFICADO').length,pending=state.rows.filter(r=>r.MINUTA_STATUS==='MINUTA_PENDENTE').length,blocked=state.rows.filter(r=>r.RESULTADO==='BLOQUEIO').length,ok=state.rows.filter(r=>r.POS_VALIDACAO==='OK').length,div=state.rows.filter(r=>r.POS_VALIDACAO==='DIVERGENTE').length;
 $('#kpis').innerHTML=[['VIAGENS',n],['CLASSIFICADAS',classified],['MINUTA PENDENTE',pending],['BLOQUEIOS',blocked],['PÓS OK',ok],['PÓS DIVERGENTE',div]].map(x=>'<div class="kpi"><div class="value">'+x[1]+'</div><div class="label">'+x[0]+'</div></div>').join('');
 $('#tmsKpis').innerHTML=[['PRONTOS TMS',state.rows.filter(r=>r.RESULTADO==='CLASSIFICADO'&&r.MINUTA_STATUS==='MINUTA_LOCALIZADA').length],['PENDENTES',pending],['BLOQUEADOS',blocked],['DIVERGÊNCIAS',div]].map(x=>'<div class="kpi"><div class="value">'+x[1]+'</div><div class="label">'+x[0]+'</div></div>').join('');
}
function bars(id,items){const max=Math.max(1,...items.map(x=>x[1]));$(id).innerHTML=items.length?items.map(x=>'<div class="barrow"><span>'+escapeHtml(x[0])+'</span><div class="bartrack"><div class="barfill" style="width:'+(x[1]/max*100)+'%"></div></div><b>'+x[1]+'</b></div>').join(''):'<div class="empty">Sem dados.</div>'}
function renderDashboard(){const c={},s={};state.rows.forEach(r=>{c[r.CLIENTE]=(c[r.CLIENTE]||0)+1;s[r.RESULTADO]=(s[r.RESULTADO]||0)+1});bars('#clientChart',Object.entries(c));bars('#statusChart',Object.entries(s));$('#lastRun').textContent=state.history[0]?.data||'—';$('#runSummary').innerHTML='<span><b>'+state.rows.length+'</b> viagens</span><span><b>'+state.minutes.length+'</b> minutas importadas</span><span><b>'+state.freight.length+'</b> tabelas de frete</span><span><b>'+state.nodes.length+'</b> Nodes</span>'}
function renderClass(){const clients=[...new Set(state.rows.map(r=>r.CLIENTE))].filter(Boolean),ops=[...new Set(state.rows.map(r=>r.OPERACAO))].filter(Boolean);$('#clientFilter').innerHTML='<option value="">Todos os clientes</option>'+clients.map(x=>'<option>'+escapeHtml(x)+'</option>').join('');$('#opFilter').innerHTML='<option value="">Todas as operações</option>'+ops.map(x=>'<option>'+escapeHtml(x)+'</option>').join('');filterClass()}
function filterClass(){const c=$('#clientFilter').value,o=$('#opFilter').value,z=$('#resultFilter').value,q=norm($('#searchFilter').value);const rows=state.rows.filter(r=>(!c||r.CLIENTE===c)&&(!o||r.OPERACAO===o)&&(!z||r.RESULTADO===z)&&(!q||Object.values(r).some(v=>norm(v).includes(q))));table('#classTable',rows)}
function render(){renderKpis();renderDashboard();renderClass();table('#minuteTable',state.minutes);table('#outsideTable',state.outside);table('#treatTable',state.treatments);const ready=state.rows.filter(r=>r.RESULTADO==='CLASSIFICADO'&&r.MINUTA_STATUS==='MINUTA_LOCALIZADA');table('#tmsTable',ready.map(r=>({MINUTA:r.MINUTA,TABELA:r.TABELA_ESPERADA,CLIENTE:r.CLIENTE,ID_VIAGEM:r.ID_VIAGEM,OPERACAO:r.OPERACAO})));renderConfig();renderFreight();renderNodes();renderHistory();$('#importSummary').innerHTML=state.rows.length?'<div class="panel summary"><span><b>'+state.rows.length+'</b> viagens</span><span><b>'+ready.length+'</b> prontas para TMS</span><span><b>'+state.outside.length+'</b> minutas fora da demanda</span></div>':'';$('#minuteSummary').innerHTML=state.minutes.length?'<div class="panel summary"><span><b>'+state.minutes.length+'</b> minutas importadas</span><span><b>'+state.outside.length+'</b> fora da demanda</span></div>':'';$('#freightSummary').innerHTML='<div class="panel summary"><span><b>'+state.freight.length+'</b> registros de tabela</span></div>';$('#nodeSummary').innerHTML='<div class="panel summary"><span><b>'+state.nodes.length+'</b> Nodes cadastrados</span></div>'}
function renderConfig(){const rl=$('#ruleList');rl.innerHTML=cfg.rules.map(r=>'<div class="rule"><b>'+r.id+'</b><span>'+escapeHtml(r.rule)+'</span><code>P'+r.priority+'</code></div>').join('');$('#localityList').innerHTML=cfg.localities.map((x,i)=>'<div class="configrow"><input data-loc="'+i+'" value="'+escapeHtml(x[0])+'"><input data-dest="'+i+'" value="'+escapeHtml(x[1])+'"><button class="btn small" onclick="removeLoc('+i+')">×</button></div>').join('');$('#typologyList').innerHTML=cfg.types.map((x,i)=>'<div class="configrow"><input data-typraw="'+i+'" value="'+escapeHtml(x[0])+'"><input data-typcanon="'+i+'" value="'+escapeHtml(x[1])+'"><button class="btn small" onclick="removeTyp('+i+')">×</button></div>').join('')}
function renderFreight(){table('#freightTable',state.freight)}
function renderNodes(){table('#nodeTable',state.nodes)}
function renderHistory(){table('#historyTable',state.history)}
function parseImport(input,callback){const f=input.files[0];if(!f)return;readFile(f).then(x=>callback(x.data,f.name)).catch(e=>toast('Erro: '+e.message))}
function download(rows,name){if(!rows.length)return toast('Não há dados para exportar.');const ws=XLSX.utils.json_to_sheet(rows);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Dados');XLSX.writeFile(wb,name+'.xlsx');audit('EXPORTACAO',name)}
function exportType(type){if(type==='classification')download(state.rows,'classificacoes_trans_rv');if(type==='minutes')download(state.minutes,'minutas_tms');if(type==='outside')download(state.outside,'minutas_fora_demanda');if(type==='treatments')download(state.treatments,'tratativas');if(type==='tms')download(state.rows.filter(r=>r.RESULTADO==='CLASSIFICADO'&&r.MINUTA_STATUS==='MINUTA_LOCALIZADA').map(r=>({MINUTA:r.MINUTA,TABELA:r.TABELA_ESPERADA})).filter(r=>r.MINUTA&&r.TABELA),'classificacao_tms');if(type==='freight')download(state.freight,'tabelas_frete');if(type==='nodes')download(state.nodes,'nodes_amazon')}
$('#fileInput').onchange=e=>{state.files=[...e.target.files];renderFiles()};
function renderFiles(){$('#fileList').innerHTML=state.files.map(f=>'<div class="fileitem"><span>'+escapeHtml(f.name)+'</span><span>'+Math.round(f.size/1024)+' KB</span></div>').join('')}
$('#dropzone').addEventListener('dragover',e=>{e.preventDefault();$('#dropzone').classList.add('drag')});$('#dropzone').addEventListener('dragleave',()=>$('#dropzone').classList.remove('drag'));$('#dropzone').addEventListener('drop',e=>{e.preventDefault();$('#dropzone').classList.remove('drag');state.files=[...e.dataTransfer.files].filter(f=>/\.(xlsx|xls|csv)$/i.test(f.name));renderFiles()});
$('#analyzeBtn').onclick=analyze;$('#clearBtn').onclick=()=>{state.files=[];state.rows=[];state.minutes=[];state.treatments=[];state.outside=[];$('#fileInput').value='';audit('LIMPEZA','Execução limpa pelo usuário');render();toast('Execução limpa')};
$('#minuteInput').onchange=e=>parseImport(e,(data,name)=>{state.minutes=data;applyMinutes();save('minutes',data);audit('IMPORTACAO_MINUTAS',name+' • '+data.length+' registros');render();toast('Minutas importadas')});
$('#freightInput').onchange=e=>parseImport(e,(data,name)=>{state.freight=data;save('freight',data);state.rows=state.rows.map(r=>({...r,...processRow(r,r.ARQUIVO_ORIGEM)}));applyMinutes();audit('IMPORTACAO_TABELAS',name+' • '+data.length+' registros');render();toast('Tabelas importadas')});
$('#nodeInput').onchange=e=>parseImport(e,(data,name)=>{state.nodes=data;save('nodes',data);state.rows=state.rows.map(r=>({...r,...processRow(r,r.ARQUIVO_ORIGEM)}));applyMinutes();audit('IMPORTACAO_NODES',name+' • '+data.length+' registros');render();toast('Nodes importados')});
['clientFilter','opFilter','resultFilter','searchFilter'].forEach(id=>$('#'+id).addEventListener(id==='searchFilter'?'input':'change',filterClass));
$$('[data-export]').forEach(b=>b.onclick=()=>exportType(b.dataset.export));$('#exportAll').onclick=()=>download(state.rows,'resultado_completo_trans_rv');
$('#resetRules').onclick=()=>{cfg.rules=DEFAULT_RULES;save('rules',cfg.rules);audit('CONFIG_REGRAS','Regras padrão restauradas');renderConfig();toast('Regras restauradas')};
$('#addLocality').onclick=()=>{cfg.localities.push(['NOVA LOCALIDADE','LOCALIDADE DESTINO']);save('localities',cfg.localities);audit('CONFIG_LOCALIDADE','Nova linha adicionada');renderConfig()};
$('#addTypology').onclick=()=>{cfg.types.push(['NOVA TIPOLOGIA','TIPOLOGIA CANONICA']);save('types',cfg.types);audit('CONFIG_TIPOLOGIA','Nova linha adicionada');renderConfig()};
window.removeLoc=i=>{cfg.localities.splice(i,1);save('localities',cfg.localities);audit('CONFIG_LOCALIDADE','Linha removida');renderConfig()};
window.removeTyp=i=>{cfg.types.splice(i,1);save('types',cfg.types);audit('CONFIG_TIPOLOGIA','Linha removida');renderConfig()};
document.addEventListener('change',e=>{if(e.target.dataset.loc!==undefined){const i=+e.target.dataset.loc;cfg.localities[i][0]=e.target.value;save('localities',cfg.localities)}if(e.target.dataset.dest!==undefined){const i=+e.target.dataset.dest;cfg.localities[i][1]=e.target.value;save('localities',cfg.localities)}if(e.target.dataset.typraw!==undefined){const i=+e.target.dataset.typraw;cfg.types[i][0]=e.target.value;save('types',cfg.types)}if(e.target.dataset.typcanon!==undefined){const i=+e.target.dataset.typcanon;cfg.types[i][1]=e.target.value;save('types',cfg.types)}});
$('#clearHistory').onclick=()=>{state.history=[];save('history',[]);renderHistory();toast('Histórico local limpo')};
state.rows=load('rows',[]);state.minutes=load('minutes',[]);state.freight=load('freight',[]);state.nodes=load('nodes',[]);applyMinutes();render();