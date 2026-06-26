// Lógica compartilhada do dashboard admin e do painel público.
// Define window.DASH_ENDPOINT antes de carregar este arquivo.
let chartTipo, chartProc, chartDia;
const ESTADOS={
  PREPARANDO:{t:'Preparando',c:'#F59E0B',i:'hourglass-split'},
  PREENCHER:{t:'Preencher',c:'#64748b',i:'pencil-square'},
  FILA_BANHO:{t:'Na fila',c:'#7C5CFC',i:'list-ol'},
  EM_BANHO:{t:'Em banho',c:'#2563EB',i:'droplet-fill'},
};

function getRange(){
  const de=document.getElementById('fDe')?.value||'';
  const ate=document.getElementById('fAte')?.value||'';
  const p=new URLSearchParams();
  if(de)p.set('de',de); if(ate)p.set('ate',ate);
  return p.toString()?('?'+p.toString()):'';
}

async function atualizar(){
  const r=await fetch(window.DASH_ENDPOINT+getRange());
  const d=await r.json();
  set('kTotal',d.total); set('kAnd',d.em_andamento);
  set('kNormais',d.normais); set('kRetrab',d.retrabalhos);
  set('kPrep',d.media_prep); set('kBanho',d.media_banho);
  renderAtivos(d.ativos);
  renderCharts(d);
  if(document.getElementById('tbody')) renderTabela(d.registros);
}
function set(id,v){const e=document.getElementById(id);if(e)e.textContent=v;}

function renderAtivos(ativos){
  const c=document.getElementById('andamento'); if(!c)return;
  if(!ativos.length){c.innerHTML='<div class="empty"><i class="bi bi-clipboard-check"></i>Nenhum cesto em andamento.</div>';return;}
  window._ativosCache=ativos;
  c.innerHTML=ativos.map(a=>{
    const e=ESTADOS[a.estado]||{t:a.estado,c:'#888',i:'circle'};
    // cronômetro: em banho conta do início do banho; preparando conta do início da prep
    let ini='';
    if(a.estado==='EM_BANHO') ini=a.banho_inicio_iso;
    else if(a.estado==='PREPARANDO'&&!a.pausado) ini=a.prep_inicio_iso;
    else if(a.estado==='FILA_BANHO') ini=a.prep_fim_iso;
    const cron=ini?`<div class="and-cron cron-el" data-ini="${ini}">${fmtCron(segCron(ini))}</div>`:'';
    const tipoTag=`<span class="pill ${a.tipo==='Retrabalho'?'pill-retrab':'pill-normal'}" style="font-size:.62rem;">${a.tipo}</span>`;
    return `<div class="and-chip" style="border-left:4px solid ${e.c};cursor:pointer;" onclick="verDetalhes(${a.id})">
      <div class="and-top"><span class="and-cesto">${a.numero_cesto}</span><i class="bi bi-${e.i}" style="color:${e.c}"></i></div>
      <div class="and-st" style="color:${e.c}">${e.t}</div>
      <div class="and-op">OP ${a.ordem||'—'} ${tipoTag}</div>
      ${cron}
    </div>`;
  }).join('');
}

// cronômetro sincronizado com o servidor
let _offsetPainel=0;
async function sincPainel(){try{const t0=Date.now();const j=await (await fetch('/api/agora')).json();const t1=Date.now();_offsetPainel=new Date(j.agora_iso).getTime()+(t1-t0)/2-t1;}catch(_){}}
function agoraPainel(){return Date.now()+_offsetPainel;}
function segCron(iso){return iso?Math.max(0,Math.floor((agoraPainel()-new Date(iso).getTime())/1000)):0;}
function fmtCron(s){const h=Math.floor(s/3600),m=Math.floor((s%3600)/60),x=s%60;const mm=String(m).padStart(2,'0'),xx=String(x).padStart(2,'0');return h>0?`${h}:${mm}:${xx}`:`${mm}:${xx}`;}
function tickPainel(){document.querySelectorAll('#andamento .cron-el[data-ini]').forEach(e=>{if(e.dataset.ini)e.textContent=fmtCron(segCron(e.dataset.ini));});}

// detalhes do cesto (modal)
function verDetalhes(id){
  const a=(window._ativosCache||[]).find(x=>x.id===id);
  if(!a)return;
  const itens=(a.itens&&a.itens.length)?a.itens:[{ordem:a.ordem,material:a.material,texto_breve:a.texto_breve,quantidade:a.quantidade}];
  const linhas=itens.map(it=>`<tr>
    <td><strong>${it.ordem||'—'}</strong></td><td>${it.material||'—'}</td>
    <td><span class="small">${it.texto_breve||'—'}</span></td><td class="mono">${it.quantidade||0}</td>
  </tr>`).join('');
  const e=ESTADOS[a.estado]||{t:a.estado};
  document.getElementById('detTitulo').textContent='Cesto '+a.numero_cesto;
  document.getElementById('detCorpo').innerHTML=`
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px;">
      <span class="pill" style="background:${e.c||'#888'}1a;color:${e.c||'#888'};">${e.t}</span>
      <span class="pill ${a.tipo==='Retrabalho'?'pill-retrab':'pill-normal'}">${a.tipo}</span>
      <span class="pill pill-banho">${a.processo||'Sem processo'}</span>
    </div>
    <div class="table-wrap"><table class="tbl">
      <thead><tr><th>OP</th><th>Código</th><th>Descrição</th><th>Qtd</th></tr></thead>
      <tbody>${linhas}</tbody>
    </table></div>
    <div style="display:flex;gap:18px;flex-wrap:wrap;margin-top:14px;">
      <div><span class="small muted">Total de peças</span><br><strong style="font-size:1.2rem;">${a.qtd_total||0}</strong></div>
      ${a.peso_total!==undefined?`<div><span class="small muted">Peso total</span><br><strong style="font-size:1.2rem;">${a.peso_total} kg</strong></div>`:''}
      ${a.area_total!==undefined?`<div><span class="small muted">Área total</span><br><strong style="font-size:1.2rem;">${a.area_total} m²</strong></div>`:''}
    </div>`;
  document.getElementById('modalDetalhes').classList.add('open');
}
function fecharDetalhes(){document.getElementById('modalDetalhes').classList.remove('open');}

function renderCharts(d){
  const labels=Object.keys(d.por_processo), data=Object.values(d.por_processo);
  const diaL=Object.keys(d.por_dia), diaD=Object.values(d.por_dia);
  if(!chartTipo){
    chartTipo=new Chart(document.getElementById('chartTipo'),{type:'doughnut',
      data:{labels:['Normal','Retrabalho'],datasets:[{data:[d.normais,d.retrabalhos],backgroundColor:['#16A34A','#F59E0B'],borderWidth:0}]},
      options:{plugins:{legend:{position:'bottom'}},cutout:'66%'}});
    chartProc=new Chart(document.getElementById('chartProc'),{type:'bar',
      data:{labels,datasets:[{data,backgroundColor:'#0F3D5C',borderRadius:7}]},
      options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{stepSize:1}}}}});
    if(document.getElementById('chartDia'))
      chartDia=new Chart(document.getElementById('chartDia'),{type:'line',
        data:{labels:diaL,datasets:[{data:diaD,borderColor:'#16B8A6',backgroundColor:'rgba(22,184,166,.15)',fill:true,tension:.3,borderWidth:3,pointRadius:4,pointBackgroundColor:'#16B8A6'}]},
        options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{stepSize:1}}}}});
  }else{
    chartTipo.data.datasets[0].data=[d.normais,d.retrabalhos];chartTipo.update();
    chartProc.data.labels=labels;chartProc.data.datasets[0].data=data;chartProc.update();
    if(chartDia){chartDia.data.labels=diaL;chartDia.data.datasets[0].data=diaD;chartDia.update();}
  }
}

function renderTabela(regs){
  const tb=document.getElementById('tbody');
  if(!regs.length){tb.innerHTML='<tr><td colspan="11" class="empty">Nenhum registro no período.</td></tr>';return;}
  tb.innerHTML=regs.map(r=>`<tr class="${r.tipo==='Retrabalho'?'retrab':''}"
    data-txt="${(r.numero_cesto+' '+r.ordem+' '+r.material+' '+(r.texto_breve||'')+' '+r.operador_prep+' '+r.operador_banho).toLowerCase()}"
    data-tipo="${r.tipo}" data-proc="${r.processo}">
    <td>${r.id}</td><td><strong>${r.numero_cesto}</strong></td><td>${r.n_itens>1?(r.ordem+' +'+(r.n_itens-1)):(r.ordem||'—')}</td>
    <td>${r.material||'—'}</td><td><span class="small">${r.texto_breve||'—'}</span></td><td>${r.qtd_total}</td>
    <td>${r.processo||'—'}</td><td><span class="pill ${r.tipo==='Retrabalho'?'pill-retrab':'pill-normal'}">${r.tipo}</span></td>
    <td class="mono">${r.prep_minutos}</td><td class="mono">${r.banho_minutos}</td>
    <td><span class="small">${r.banho_fim||''}</span></td>
  </tr>`).join('');
  aplicarFiltroTexto();
}

function aplicarFiltroTexto(){
  const txt=(document.getElementById('fTexto')?.value||'').toLowerCase();
  const tipo=document.getElementById('fTipo')?.value||'';
  const proc=document.getElementById('fProc')?.value||'';
  document.querySelectorAll('#tbody tr[data-txt]').forEach(tr=>{
    const ok=(!txt||tr.dataset.txt.includes(txt))&&(!tipo||tr.dataset.tipo===tipo)&&(!proc||tr.dataset.proc===proc);
    tr.style.display=ok?'':'none';
  });
}

document.addEventListener('DOMContentLoaded',()=>{
  ['fTexto','fTipo','fProc'].forEach(id=>{const e=document.getElementById(id);if(e)e.addEventListener(id==='fTexto'?'input':'change',aplicarFiltroTexto);});
  ['fDe','fAte'].forEach(id=>{const e=document.getElementById(id);if(e)e.addEventListener('change',atualizar);});
  if(document.getElementById('andamento')){
    sincPainel().then(()=>{atualizar();setInterval(atualizar,8000);setInterval(tickPainel,1000);});
    setInterval(sincPainel,60000);
  }else{
    atualizar();
    setInterval(atualizar,8000);
  }
});
