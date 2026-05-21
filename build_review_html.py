"""
ESCMID 2027 Proposal Review — HTML generator.
Importable as a module or run directly.

Usage:
    python build_review_html.py               # reads data/output/proposals_tagged.json
    import build_review_html; html = build_review_html.generate_html(data_dict)
"""
import json
from pathlib import Path

CRITERIA = [
    ("C1","Hot / Timely / Controversial","Is the topic hot, timely, or controversial? Does it reflect current scientific priorities?"),
    ("C2","Not Duplicated","Not duplicated from ESCMID Global 2024/2025/2026, unless new developments justify resubmission."),
    ("C3","Cross-Disciplinary","Combines CM/ID/IC aspects; engages different disciplines (nurses, pharmacists) or age groups."),
    ("C4","Basic + Clinical Integration","Combines basic, translational, applied, and clinical aspects where applicable."),
    ("C5","Appropriate Format","The chosen session format is appropriate to meet the stated session goals."),
    ("C6","Relevant Collaborators","Study Groups, Affiliated Societies, or other organisations are appropriately involved."),
    ("C7","Adheres to Format Rules","Strictly adheres to format guidelines (number of talks, duration, etc.)."),
    ("C8","Proper Description","Proper description highlighting importance, relevance, and educational goals."),
    ("C9","Gender & Geographic Balance","Adequate balance. Underrepresented: young/female speakers, Eastern Europe, low-resource countries."),
    ("C10","Best Speakers / Not Self-Serving","Best speakers chosen; free from self-serving bias."),
    ("C11","Engages Young Investigators","Engages young investigators or new faces where applicable."),
]


def generate_html(data) -> str:
    """Generate the full review dashboard HTML.
    Accepts either a dict {proposals:[...], ...} or a bare list of proposals.
    """
    if isinstance(data, list):
        data = {"proposals": data, "tag_palette": {}, "total_hist": 0}

    proposals = data.get("proposals", [])
    n_prop    = len(proposals)
    n_hist    = data.get("total_hist", 0)

    # Clean surrogates before JSON serialisation
    raw       = json.dumps(data, ensure_ascii=True)
    data_json = raw

    crit_json   = json.dumps(CRITERIA, ensure_ascii=True)
    tag_pal_json = json.dumps(data.get("tag_palette", {}), ensure_ascii=True)

    return (_TEMPLATE
        .replace("__DATA_JSON__",   data_json)
        .replace("__CRIT_JSON__",   crit_json)
        .replace("__TAG_PAL_JSON__", tag_pal_json)
        .replace("__N_PROP__",      str(n_prop))
        .replace("__N_HIST__",      str(n_hist))
    )


_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>ESCMID 2027 Proposal Review</title>
<style>
:root{--bg:#080e1a;--surf:#0f1d2e;--surf2:#162336;--bdr:#1e3450;--text:#c8d8ec;--dim:#617d9b;--acc:#4fc3f7;--green:#34d399;--yellow:#fbbf24;--red:#f87171}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:system-ui,sans-serif;font-size:13px;line-height:1.6;height:100vh;display:flex;flex-direction:column;overflow:hidden}
header{padding:10px 24px;border-bottom:1px solid var(--bdr);background:#0a1830;flex-shrink:0;display:flex;align-items:center;gap:14px}
.logo{font-size:15px;font-weight:700;color:#e8f4ff}.logo span{color:var(--acc)}
.hstats{display:flex;gap:14px;margin-left:auto}
.stat{display:flex;flex-direction:column}
.sn{font-family:monospace;font-size:17px;font-weight:700;line-height:1}
.sl{font-size:10px;color:var(--dim);text-transform:uppercase;letter-spacing:.1em}
.tabbar{display:flex;background:#080e1a;border-bottom:1px solid var(--bdr);flex-shrink:0;padding:0 24px}
.tab{padding:9px 13px;font-size:11px;font-weight:600;color:var(--dim);background:none;border:none;border-bottom:2px solid transparent;cursor:pointer;letter-spacing:.04em;text-transform:uppercase;white-space:nowrap;font-family:inherit;transition:color .15s,border-color .15s}
.tab:hover{color:var(--text)}.tab.active{color:var(--acc);border-color:var(--acc)}
.content{flex:1;min-height:0;position:relative}
.pane{display:none;position:absolute;inset:0;overflow-y:auto;padding:18px 24px}
.pane.active{display:block}
.card{background:var(--surf);border:1px solid var(--bdr);border-radius:8px;padding:14px 18px;margin-bottom:12px}
.filters{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px;align-items:center}
.sb{flex:1;min-width:140px;max-width:240px;background:var(--surf2);border:1px solid var(--bdr);color:var(--text);padding:5px 11px;border-radius:16px;font-size:12px;font-family:inherit;outline:none}
.sb:focus,.esel:focus{border-color:var(--acc)}
.esel{background:var(--surf2);border:1px solid var(--bdr);color:var(--text);padding:5px 9px;border-radius:16px;font-size:11px;font-family:inherit;outline:none;cursor:pointer}
.tc{font-family:monospace;font-size:11px;color:var(--dim);margin-left:auto}
.ptable{width:100%;border-collapse:collapse;font-size:11px}
.ptable th{background:var(--surf2);color:var(--dim);font-weight:600;font-size:10px;text-transform:uppercase;padding:7px 9px;border:1px solid var(--bdr);text-align:left;white-space:nowrap}
.ptable td{padding:8px 9px;border:1px solid var(--bdr);vertical-align:top}
.ptable tr.cr{cursor:pointer}.ptable tr.cr:hover td{background:rgba(79,195,247,.05)}
.tp{display:inline-block;font-size:8px;padding:1px 5px;border-radius:8px;font-family:monospace;margin:1px 1px 1px 0;white-space:nowrap}
.badge{font-family:monospace;font-size:10px;padding:2px 6px;border-radius:8px;white-space:nowrap;display:inline-block}
#pane-review{display:none;position:absolute;inset:0;flex-direction:column;overflow:hidden;padding:0}
#pane-review.active{display:flex}
.rev-top{display:flex;align-items:center;gap:8px;padding:7px 18px;border-bottom:1px solid var(--bdr);background:var(--surf);flex-shrink:0;flex-wrap:wrap}
.rnbtn{background:var(--surf2);border:1px solid var(--bdr);color:var(--text);padding:4px 12px;border-radius:5px;font-size:11px;cursor:pointer;font-family:inherit;font-weight:600}
.rnbtn:hover{border-color:var(--acc);color:var(--acc)}
.rpos{font-family:monospace;font-size:11px;color:var(--dim);min-width:80px;text-align:center}
.rprog{font-size:11px;color:var(--green);margin-left:auto;font-family:monospace}
.savest{font-size:11px;color:var(--dim);font-family:monospace}
.dlbtn{background:rgba(52,211,153,.15);border:1px solid rgba(52,211,153,.4);color:var(--green);padding:4px 12px;border-radius:5px;font-size:11px;cursor:pointer;font-family:inherit;font-weight:600}
.dlbtn:hover{background:rgba(52,211,153,.25)}
.rev-body{display:grid;grid-template-columns:58% 42%;flex:1;overflow:hidden;min-height:0}
.rleft{overflow-y:auto;padding:16px 20px;border-right:1px solid var(--bdr)}
.rright{overflow-y:auto;padding:14px 16px;background:var(--surf2)}
.rleft::-webkit-scrollbar,.rright::-webkit-scrollbar{width:4px}
.rleft::-webkit-scrollbar-thumb,.rright::-webkit-scrollbar-thumb{background:var(--bdr)}
.sech{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.1em;color:var(--dim);margin:12px 0 5px}
.sech:first-child{margin-top:0}
.fg{display:grid;grid-template-columns:120px 1fr;gap:4px 8px;margin-bottom:8px}
.fl{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.07em;color:var(--dim);padding-top:2px}
.fv{font-size:12px;color:var(--text);line-height:1.5}
.trow{padding:5px 0;border-bottom:1px solid var(--bdr)}
.ttit{font-size:12px;color:#e8f4ff}.tspk{font-size:10px;color:var(--dim);margin-top:2px}
.crow{display:flex;align-items:center;gap:5px;padding:6px 0;border-bottom:1px solid rgba(30,52,80,.6)}
.clbl{position:relative;font-family:monospace;font-size:11px;color:var(--acc);cursor:help;min-width:30px;font-weight:700;user-select:none}
.ctip{display:none;position:fixed;background:#0a1525;border:1px solid var(--acc);border-radius:6px;padding:9px 13px;font-size:11px;color:var(--text);width:260px;z-index:9999;line-height:1.6;pointer-events:none;box-shadow:0 8px 24px rgba(0,0,0,.7)}
.clbl:hover .ctip{display:block}
.rgrp{display:flex;gap:3px}
.rb{width:30px;height:24px;border:1px solid var(--bdr);background:var(--surf);color:var(--dim);font-size:10px;font-family:monospace;border-radius:4px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .12s;flex-shrink:0;user-select:none}
.rb:hover{border-color:var(--acc);color:var(--acc)}
.rb.sel{background:var(--acc);border-color:var(--acc);color:#000;font-weight:700}
.rb.na{font-size:8px}.rb.na.sel{background:var(--dim);border-color:var(--dim);color:#fff}
.airow{margin-left:auto;display:flex;align-items:center;gap:2px;flex-shrink:0}
.aidot{width:7px;height:7px;border-radius:50%}
.ailbl{font-size:8px;color:var(--dim);font-family:monospace;min-width:14px;text-align:right}
.stars{display:flex;gap:6px;margin:10px 0 4px}
.star{font-size:22px;cursor:pointer;user-select:none;line-height:1;transition:transform .1s}
.star:hover{transform:scale(1.2)}
.nota{width:100%;background:var(--surf);border:1px solid var(--bdr);color:var(--text);padding:8px 10px;border-radius:6px;font-size:12px;font-family:inherit;resize:vertical;min-height:80px;outline:none;margin-top:6px}
.nota:focus{border-color:var(--acc)}
.jumpsel{background:var(--surf2);border:1px solid var(--bdr);color:var(--text);padding:4px 8px;border-radius:5px;font-size:11px;font-family:monospace;outline:none;cursor:pointer;max-width:260px}
.sbarwrap{display:flex;align-items:center;gap:8px;margin:3px 0;cursor:pointer;padding:3px 4px;border-radius:4px;transition:background .15s}
.sbarwrap:hover{background:rgba(79,195,247,.07)}
.sbarlbl{font-size:11px;min-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sbar{flex:1;height:8px;background:var(--surf2);border-radius:4px;overflow:hidden}
.sbarfill{height:100%;border-radius:4px}
.ccard{background:var(--surf2);border:1px solid var(--bdr);border-radius:7px;padding:13px 17px;margin-bottom:8px;display:grid;grid-template-columns:50px 180px 1fr;gap:10px;align-items:start}
.ccode{font-family:monospace;font-size:14px;font-weight:700;color:var(--acc)}
.cname{font-size:12px;font-weight:600;color:#e8f4ff;line-height:1.4}
.cdesc{font-size:12px;color:var(--dim);line-height:1.6}
</style>
</head>
<body>
<header>
  <div class="logo">ESCMID 2027 <span>Proposal Review</span></div>
  <div class="hstats">
    <div class="stat"><span class="sn" style="color:var(--acc)" id="hn">__N_PROP__</span><span class="sl">Proposals</span></div>
    <div class="stat"><span class="sn" style="color:var(--yellow)">__N_HIST__</span><span class="sl">Hist. sessions</span></div>
    <div class="stat"><span class="sn" style="color:var(--green)" id="hrev">0</span><span class="sl">Reviewed</span></div>
  </div>
</header>
<div class="tabbar">
  <button class="tab active" data-tab="proposals">Proposals</button>
  <button class="tab" data-tab="review">Review Form</button>
  <button class="tab" data-tab="scores">Scores</button>
  <button class="tab" data-tab="criteria">Criteria</button>
</div>
<div class="content">
  <div class="pane active" id="pane-proposals">
    <h2 style="font-size:14px;font-weight:700;color:#e8f4ff;margin-bottom:4px">All Proposals</h2>
    <p style="font-size:12px;color:var(--dim);margin-bottom:12px">Click any row to open the Review Form. N=Novelty, T=Trend, AI=heuristic score.</p>
    <div class="card" style="padding:12px 16px">
      <div class="filters">
        <input class="sb" id="ps" placeholder="Search...">
        <select class="esel" id="pc"><option value="">All categories</option></select>
        <select class="esel" id="pt"><option value="">All types</option></select>
        <select class="esel" id="po">
          <option value="id">Sort: ID</option>
          <option value="novelty">Sort: Novelty</option>
          <option value="trend">Sort: Trend</option>
          <option value="ai">Sort: AI score</option>
        </select>
        <span class="tc" id="pct"></span>
      </div>
      <div style="overflow-x:auto">
        <table class="ptable">
          <thead><tr><th>ID</th><th>Title</th><th>Category</th><th>Type</th><th>Tags</th><th>N</th><th>T</th><th>AI</th><th>Done</th></tr></thead>
          <tbody id="ptb"></tbody>
        </table>
      </div>
    </div>
  </div>
  <div id="pane-review">
    <div class="rev-top">
      <button class="rnbtn" id="btnprev">&#8592; Prev</button>
      <div class="rpos" id="rpos">1 / __N_PROP__</div>
      <button class="rnbtn" id="btnnext">Next &#8594;</button>
      <select class="jumpsel" id="rjump"><option value="">Jump to...</option></select>
      <span class="savest" id="savest"></span>
      <span class="rprog" id="rprog">0 / __N_PROP__ reviewed</span>
      <button class="dlbtn" id="btncsv">Download CSV</button>
    </div>
    <div class="rev-body">
      <div class="rleft" id="rleft">Select a proposal to begin.</div>
      <div class="rright" id="rright"></div>
    </div>
  </div>
  <div class="pane" id="pane-scores">
    <h2 style="font-size:14px;font-weight:700;color:#e8f4ff;margin-bottom:4px">Scores</h2>
    <p style="font-size:12px;color:var(--dim);margin-bottom:12px">Click any bar to open that proposal in the Review Form.</p>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
      <div class="card" style="display:flex;flex-direction:column">
        <div style="font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.1em;color:var(--dim);margin-bottom:8px;flex-shrink:0">Novelty (highest first)</div>
        <div id="nb" style="overflow-y:auto;flex:1;max-height:calc(100vh - 230px)"></div>
      </div>
      <div class="card" style="display:flex;flex-direction:column">
        <div style="font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.1em;color:var(--dim);margin-bottom:8px;flex-shrink:0">Trend (positive = growing)</div>
        <div id="tb" style="overflow-y:auto;flex:1;max-height:calc(100vh - 230px)"></div>
      </div>
    </div>
  </div>
  <div class="pane" id="pane-criteria">
    <h2 style="font-size:14px;font-weight:700;color:#e8f4ff;margin-bottom:4px">Scoring Criteria</h2>
    <p style="font-size:12px;color:var(--dim);margin-bottom:12px">Score each 1-5 or N/A in the Review Form. Hover labels in the form for descriptions.</p>
    <div id="critlist"></div>
  </div>
</div>
<script>
var D=__DATA_JSON__;
var PP=D.proposals||[];
var CR=__CRIT_JSON__;
var TP=__TAG_PAL_JSON__;
var STORE='escmid_2027_review';
function tclr(t){var k=Object.keys(TP);for(var i=0;i<k.length;i++){if(t.indexOf(k[i])===0)return TP[k[i]];}return'#617d9b';}
function h2r(h,a){var r=parseInt(h.slice(1,3),16),g=parseInt(h.slice(3,5),16),b=parseInt(h.slice(5,7),16);return'rgba('+r+','+g+','+b+','+a+')';}
function nc(n){return n>0.7?'var(--green)':n>0.4?'var(--yellow)':'var(--red)';}
function tc2(t){return t>0.1?'var(--green)':t<-0.1?'var(--red)':'var(--yellow)';}
function sc(n){return n>=4?'var(--green)':n>=3?'var(--yellow)':'var(--red)';}
function loadAll(){try{var s=localStorage.getItem(STORE);return s?JSON.parse(s):{};}catch(e){return{};}}
function saveAll(d){try{localStorage.setItem(STORE,JSON.stringify(d));}catch(e){}}
function loadR(pid){var a=loadAll();return a[String(pid)]||{};}
function saveR(pid,r){var a=loadAll();a[String(pid)]=r;saveAll(a);updateProgress();var el=document.getElementById('savest');if(el)el.textContent='Saved';}
function countReviewed(){return PP.filter(function(p){return loadR(p.proposal_id).overall;}).length;}
function updateProgress(){var n=countReviewed();var h=document.getElementById('hrev');if(h)h.textContent=n;var rp=document.getElementById('rprog');if(rp)rp.textContent=n+' / '+PP.length+' reviewed';}
updateProgress();
document.querySelectorAll('.tab').forEach(function(b){b.addEventListener('click',function(){
  document.querySelectorAll('.tab').forEach(function(x){x.classList.remove('active');});b.classList.add('active');
  var tab=b.getAttribute('data-tab');
  document.querySelectorAll('.pane').forEach(function(p){p.classList.remove('active');});
  var rv=document.getElementById('pane-review');rv.classList.remove('active');
  if(tab==='review'){rv.classList.add('active');if(!revInited)initReview();renderReview();}
  else{var pane=document.getElementById('pane-'+tab);if(pane)pane.classList.add('active');if(tab==='scores'&&!barsBuilt)buildBars();if(tab==='criteria'&&!critBuilt)buildCriteria();}
});});
var cats=[],types=[];
PP.forEach(function(p){if(p.category&&cats.indexOf(p.category)<0)cats.push(p.category);if(p.session_type&&types.indexOf(p.session_type)<0)types.push(p.session_type);});
cats.sort();types.sort();
var pcEl=document.getElementById('pc');cats.forEach(function(c){pcEl.innerHTML+='<option>'+c+'</option>';});
var ptEl=document.getElementById('pt');types.forEach(function(t){ptEl.innerHTML+='<option>'+t+'</option>';});
var pS='',pC='',pT='',pO='id';
function trClick(el){ openInReview(parseInt(el.getAttribute('data-idx'))); }
function rbClick(el){ setScore(el.getAttribute('data-pid'), el.getAttribute('data-code'), el.getAttribute('data-val')); }
function starClick(el){ setScore(el.getAttribute('data-pid'), 'overall', parseInt(el.getAttribute('data-val'))); }
function renderTable(){
  var q=pS.toLowerCase();
  var f=PP.filter(function(p){
    if(pC&&(p.category||'')!==pC)return false;
    if(pT&&(p.session_type||'')!==pT)return false;
    if(q&&[(p.session_title||''),(p.category||''),(p.motivation||'')].join(' ').toLowerCase().indexOf(q)<0)return false;
    return true;
  });
  if(pO==='novelty')f.sort(function(a,b){return(b.score_novelty||0)-(a.score_novelty||0);});
  else if(pO==='trend')f.sort(function(a,b){return(b.score_trend||0)-(a.score_trend||0);});
  else if(pO==='ai')f.sort(function(a,b){return((b.ai_rating&&b.ai_rating.overall)||0)-((a.ai_rating&&a.ai_rating.overall)||0);});
  else f.sort(function(a,b){return String(a.proposal_id).localeCompare(String(b.proposal_id));});
  document.getElementById('pct').textContent=f.length+' proposals';
  var rows='';
  f.forEach(function(p){
    var idx=PP.indexOf(p);
    var tags=(p.tags||[]).slice(0,3).map(function(t){return'<span class="tp" style="background:'+h2r(tclr(t),.18)+';color:'+tclr(t)+';border:1px solid '+h2r(tclr(t),.3)+'">'+t.replace(/^[a-z_]+_/,'')+'</span>';}).join('');
    var n=p.score_novelty||0,tr=p.score_trend||0,ai=p.ai_rating&&p.ai_rating.overall;
    var nb='<span class="badge" style="background:'+h2r(nc(n),.18)+';color:'+nc(n)+';border:1px solid '+h2r(nc(n),.35)+'">'+n.toFixed(2)+'</span>';
    var tb='<span class="badge" style="background:'+h2r(tc2(tr),.18)+';color:'+tc2(tr)+';border:1px solid '+h2r(tc2(tr),.35)+'">'+(tr>0?'+':'')+tr.toFixed(2)+'</span>';
    var aib=ai?'<span class="badge" style="background:'+h2r(sc(ai),.18)+';color:'+sc(ai)+';border:1px solid '+h2r(sc(ai),.35)+'">'+ai+'/5</span>':'--';
    var done=loadR(p.proposal_id).overall?'&#10003;':'';
    rows+='<tr class="cr" data-idx="'+idx+'" onclick="trClick(this)"><td style="font-family:monospace;font-size:10px;color:var(--acc);white-space:nowrap">#'+p.proposal_id+'</td>';
    rows+='<td style="max-width:280px"><strong style="color:#e8f4ff">'+(p.session_title||'(no title)')+'</strong></td>';
    rows+='<td style="max-width:130px;color:var(--dim);font-size:11px">'+((p.category||'').replace(/^\\d+\.\\s*/,''))+'</td>';
    rows+='<td style="white-space:nowrap;font-size:11px;color:var(--dim)">'+(p.session_type||'')+'</td>';
    rows+='<td>'+tags+'</td><td>'+nb+'</td><td>'+tb+'</td><td>'+aib+'</td>';
    rows+='<td style="text-align:center;color:var(--green);font-weight:700">'+done+'</td></tr>';
  });
  document.getElementById('ptb').innerHTML=rows;
}
renderTable();
document.getElementById('ps').addEventListener('input',function(e){pS=e.target.value;renderTable();});
document.getElementById('pc').addEventListener('change',function(e){pC=e.target.value;renderTable();});
document.getElementById('pt').addEventListener('change',function(e){pT=e.target.value;renderTable();});
document.getElementById('po').addEventListener('change',function(e){pO=e.target.value;renderTable();});
var curIdx=0,revInited=false;
function initReview(){
  revInited=true;
  var sel=document.getElementById('rjump');
  PP.forEach(function(p,i){sel.innerHTML+='<option value="'+i+'">#'+p.proposal_id+' '+(p.session_title||'').substring(0,50)+'</option>';});
  document.getElementById('btnprev').addEventListener('click',function(){curIdx=Math.max(0,curIdx-1);renderReview();});
  document.getElementById('btnnext').addEventListener('click',function(){curIdx=Math.min(PP.length-1,curIdx+1);renderReview();});
  document.getElementById('rjump').addEventListener('change',function(e){if(e.target.value!==''){curIdx=parseInt(e.target.value);renderReview();}});
  document.getElementById('btncsv').addEventListener('click',downloadCSV);
  document.addEventListener('keydown',function(e){
    var rv=document.getElementById('pane-review');
    if(rv&&rv.classList.contains('active')){if(e.key==='ArrowRight'){curIdx=Math.min(PP.length-1,curIdx+1);renderReview();}if(e.key==='ArrowLeft'){curIdx=Math.max(0,curIdx-1);renderReview();}}
  });
}
function openInReview(idx){
  curIdx=idx;
  document.querySelectorAll('.tab').forEach(function(x){x.classList.remove('active');});
  document.querySelectorAll('.pane').forEach(function(x){x.classList.remove('active');});
  var rv=document.getElementById('pane-review');rv.classList.add('active');
  document.querySelector('.tab[data-tab="review"]').classList.add('active');
  if(!revInited)initReview();
  renderReview();
}
function renderReview(){
  var p=PP[curIdx]; if(!p) return;
  document.getElementById('rpos').textContent=(curIdx+1)+' / '+PP.length;
  document.getElementById('rjump').value=curIdx;
  var r=loadR(p.proposal_id), ai=p.ai_rating||{}, n=p.score_novelty||0, tr=p.score_trend||0;

  // Tags
  var tags=(p.tags||[]).map(function(t){
    return '<span class="tp" style="background:'+h2r(tclr(t),.18)+';color:'+tclr(t)+';border:1px solid '+h2r(tclr(t),.3)+'">'+t+'</span>';
  }).join(' ');

  // Topics
  var topicsHtml='';
  if(p.topic_titles){
    var tits=p.topic_titles.split('|'), spks=(p.topic_speakers||'').split('|');
    tits.filter(function(t){return t.trim();}).forEach(function(tit,i){
      var spk=spks[i]?spks[i].trim():'';
      topicsHtml+='<div class="trow"><div class="ttit">'+tit.trim()+'</div>'+(spk?'<div class="tspk">'+spk+'</div>':'')+'</div>';
    });
  }

  // Similar sessions
  var simsHtml='';
  (p.similar_sessions||[]).forEach(function(s){
    var pct=(s.similarity*100).toFixed(0)+'%';
    var col=s.similarity>0.5?'var(--red)':s.similarity>0.25?'var(--yellow)':'var(--green)';
    simsHtml+='<div style="display:flex;gap:8px;padding:5px 0;border-bottom:1px solid var(--bdr);font-size:11px">'
      +'<span style="font-family:monospace;color:'+col+';min-width:44px">'+pct+'</span>'
      +'<span style="font-family:monospace;color:var(--acc);min-width:75px">'+s.year+' '+s.code+'</span>'
      +'<span>'+s.title+'</span></div>';
  });

  // Left panel
  var lh='';
  lh+='<div style="font-size:11px;color:var(--dim);margin-bottom:4px">#'+p.proposal_id
    +(p.session_type?' &middot; '+p.session_type:'')+(p.category?' &middot; '+p.category:'')+'</div>';
  lh+='<div style="font-size:17px;font-weight:700;color:#e8f4ff;line-height:1.35;margin-bottom:10px">'+(p.session_title||'(no title)')+'</div>';
  lh+='<div style="margin-bottom:10px">'+tags+'</div>';
  lh+='<span class="badge" style="background:'+h2r(nc(n),.18)+';color:'+nc(n)+';border:1px solid '+h2r(nc(n),.35)+'">Novelty '+n.toFixed(2)+'</span> ';
  lh+='<span class="badge" style="background:'+h2r(tc2(tr),.18)+';color:'+tc2(tr)+';border:1px solid '+h2r(tc2(tr),.35)+'">Trend '+(tr>0?'+':'')+tr.toFixed(2)+'</span>';
  if(ai.overall) lh+=' <span class="badge" style="background:'+h2r(sc(ai.overall),.18)+';color:'+sc(ai.overall)+';border:1px solid '+h2r(sc(ai.overall),.35)+'">AI '+ai.overall+'/5</span>';
  if(ai.rationale) lh+='<div style="font-size:11px;color:var(--dim);font-style:italic;margin:10px 0;padding:8px 12px;background:var(--surf2);border-radius:6px;border-left:3px solid var(--acc)">AI: '+ai.rationale+'</div>';
  if(p.chairs||p.reserve_chairs||p.champion||p.proposing_entities){
    lh+='<div class="sech">Details</div><div class="fg">';
    if(p.chairs) lh+='<div class="fl">Chairs</div><div class="fv">'+p.chairs.replace(/\\s*\|\\s*/g,' &middot; ')+'</div>';
    if(p.reserve_chairs) lh+='<div class="fl">Reserve</div><div class="fv">'+p.reserve_chairs+'</div>';
    if(p.champion) lh+='<div class="fl">Champion</div><div class="fv">'+p.champion+'</div>';
    if(p.proposing_entities) lh+='<div class="fl">Proposer</div><div class="fv">'+p.proposing_entities+'</div>';
    lh+='</div>';
  }
  if(topicsHtml) lh+='<div class="sech">Topics ('+p.num_topics+')</div>'+topicsHtml;
  if(p.motivation) lh+='<div class="sech">Description</div><div style="font-size:12px;color:var(--dim);line-height:1.7">'+p.motivation+'</div>';
  if(simsHtml) lh+='<div class="sech">Similar Historical Sessions</div>'+simsHtml;
  document.getElementById('rleft').innerHTML=lh;

  // Right panel - use data-attributes everywhere to avoid quoting issues
  var pid=String(p.proposal_id);
  var rh='<div style="font-size:12px;font-weight:600;color:#e8f4ff;margin-bottom:12px">Your Rating</div>';
  CR.forEach(function(crit){
    var code=crit[0], name=crit[1], desc=crit[2], aiVal=ai[code], saved=r[code];
    var opts=[1,2,3,4,5,'N/A'].map(function(v){
      var isSel=String(saved)===String(v);
      var cls='rb'+(v==='N/A'?' na':'')+(isSel?' sel':'');
      // Use data-attributes — no quoting issues
      return '<div class="'+cls+'" data-pid="'+pid+'" data-code="'+code+'" data-val="'+v+'" onclick="rbClick(this)">'+v+'</div>';
    }).join('');
    var aiHtml='';
    if(typeof aiVal==='number'){
      var dots='';
      for(var di=1;di<=5;di++) dots+='<div class="aidot" style="background:'+(di<=aiVal?'var(--acc)':'rgba(255,255,255,.15)')+'"></div>';
      aiHtml='<div class="airow">'+dots+'<div class="ailbl">'+aiVal+'</div></div>';
    }
    rh+='<div class="crow">'
      +'<div class="clbl">'+code+'<div class="ctip"><strong>'+name+'</strong><br>'+desc+'</div></div>'
      +'<div class="rgrp">'+opts+'</div>'+aiHtml+'</div>';
  });

  // Stars — use data-attributes
  var ov=r.overall||0, starsHtml='';
  for(var si=1;si<=5;si++){
    starsHtml+='<div class="star" data-pid="'+pid+'" data-val="'+si+'" onclick="starClick(this)">'+(si<=ov?'&#9733;':'&#9734;')+'</div>';
  }
  rh+='<div style="font-size:11px;font-weight:600;color:var(--dim);text-transform:uppercase;letter-spacing:.1em;margin-top:10px">Overall Rating</div>';
  rh+='<div class="stars">'+starsHtml+'</div>';
  rh+='<div style="font-size:11px;font-weight:600;color:var(--dim);text-transform:uppercase;letter-spacing:.1em;margin-top:10px">Notes</div>';
  rh+='<textarea class="nota" id="nota_'+pid+'">'+(r.notes||'')+'</textarea>';
  document.getElementById('rright').innerHTML=rh;
  var ta=document.getElementById('nota_'+pid);
  if(ta) ta.addEventListener('input',function(){ saveNotes(pid); });
}

function setScore(pid,code,val){var r=loadR(pid);r[code]=isNaN(Number(val))?val:Number(val)||val;saveR(pid,r);renderReview();}
function saveNotes(pid){var ta=document.getElementById('nota_'+pid);if(!ta)return;var r=loadR(pid);r.notes=ta.value;saveR(pid,r);}
var barsBuilt=false;
function buildBars(){
  barsBuilt=true;
  var byN=PP.slice().sort(function(a,b){return(b.score_novelty||0)-(a.score_novelty||0);});
  var nh='';
  byN.forEach(function(p){var idx=PP.indexOf(p),n=p.score_novelty||0,col=nc(n);nh+='<div class="sbarwrap" data-idx="'+idx+'" onclick="trClick(this)"><div class="sbarlbl">#'+p.proposal_id+' '+(p.session_title||'').substring(0,42)+'</div><div class="sbar"><div class="sbarfill" style="width:'+(n*100).toFixed(0)+'%;background:'+col+'"></div></div><span style="font-family:monospace;font-size:10px;color:'+col+';min-width:34px">'+n.toFixed(2)+'</span></div>';});
  document.getElementById('nb').innerHTML=nh;
  var byT=PP.slice().sort(function(a,b){return(b.score_trend||0)-(a.score_trend||0);});
  var th='';
  byT.forEach(function(p){var idx=PP.indexOf(p),t=p.score_trend||0,col=tc2(t),pct=Math.min(Math.abs(t)*50,50);th+='<div class="sbarwrap" data-idx="'+idx+'" onclick="trClick(this)"><div class="sbarlbl">#'+p.proposal_id+' '+(p.session_title||'').substring(0,42)+'</div><div class="sbar" style="display:flex"><div style="flex:1;display:flex;justify-content:flex-end"><div style="width:'+(t<0?pct:0)+'%;background:var(--red);height:8px;border-radius:2px 0 0 2px"></div></div><div style="width:1px;background:var(--bdr)"></div><div style="flex:1"><div style="width:'+(t>0?pct:0)+'%;background:var(--green);height:8px;border-radius:0 2px 2px 0"></div></div></div><span style="font-family:monospace;font-size:10px;color:'+col+';min-width:44px">'+(t>0?'+':'')+t.toFixed(2)+'</span></div>';});
  document.getElementById('tb').innerHTML=th;
}
var critBuilt=false;
function buildCriteria(){critBuilt=true;var h='';CR.forEach(function(c){h+='<div class="ccard"><div class="ccode">'+c[0]+'</div><div class="cname">'+c[1]+'</div><div class="cdesc">'+c[2]+'</div></div>';});document.getElementById('critlist').innerHTML=h;}
function downloadCSV(){
  var hdrs=['ID','Title','Category','Type','Chairs','Tags','Novelty','Trend','AI Overall'];
  CR.forEach(function(c){hdrs.push(c[0]);});hdrs.push('Overall','Notes');
  var rows=[hdrs];
  PP.forEach(function(p){var r=loadR(p.proposal_id),ai=p.ai_rating||{};var row=[p.proposal_id,p.session_title||'',p.category||'',p.session_type||'',(p.chairs||'').replace(/\\s*\|\\s*/g,' | '),(p.tags||[]).join('; '),p.score_novelty||'',p.score_trend||'',ai.overall||''];CR.forEach(function(c){row.push(r[c[0]]||'');});row.push(r.overall||'',r.notes||'');rows.push(row);});
  var csv=rows.map(function(r){return r.map(function(v){return'"'+String(v==null?'':v).replace(/"/g,'""')+'"';}).join(',');}).join('\n');
  var blob=new Blob([csv],{type:'text/csv;charset=utf-8;'});var url=URL.createObjectURL(blob);var a=document.createElement('a');a.href=url;a.download='proposal_ratings_escmid_2027.csv';document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
}
</script>
</body>
</html>"""


if __name__ == "__main__":
    import sys
    candidates = [
        Path("data/output/proposals_tagged.json"),
        Path("data/output/Proposal_Review_data.json"),
    ]
    src = next((p for p in candidates if p.exists()), None)
    if not src:
        print("No proposals JSON found. Run proposal_reviewer.py first.")
        sys.exit(1)
    with open(src) as f:
        raw = json.load(f)
    html = generate_html(raw)
    out = Path("data/output/Proposal_Review.html")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")
    print(f"Written {len(html):,} bytes to {out}")
