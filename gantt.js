// Gantt chart renderer for Atelier 5 plan tab (full-width version)
function drawGanttChartSVG(container, actions) {
  if (!container) return;
  const tooltip = document.getElementById('gantt-tooltip');
  container.innerHTML = '';
  if (tooltip) tooltip.style.opacity = '0';
  if (!actions || actions.length === 0) return;

  const rowHeight = 26;
  const rowGap = 8;
  const leftPad = 200;
  const rightPad = 24;
  const topPad = 40;
  const bottomPad = 38;
  const pxPerDay = 26;

  function parseDate(s){ const [y,m,d]=s.split('-').map(Number); return new Date(y, m-1, d); }
  function ymd(d){ const y=d.getFullYear(); const m=String(d.getMonth()+1).padStart(2,'0'); const day=String(d.getDate()).padStart(2,'0'); return `${y}-${m}-${day}`; }
  function clampTo00(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  function addDays(d,n){ const x=new Date(d); x.setDate(x.getDate()+n); return x; }
  function daysBetween(a,b){ return Math.round((clampTo00(b)-clampTo00(a))/86400000); }
  function weekNumber(d){
    const dt=new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    const dayNum = (dt.getUTCDay() || 7);
    dt.setUTCDate(dt.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(dt.getUTCFullYear(),0,1));
    return Math.ceil((((dt - yearStart)/86400000)+1)/7);
  }
  function toMonday(d){ const day=(d.getDay()||7); return addDays(d,1-day); }
  function toSunday(d){ const day=(d.getDay()||7); return addDays(d,7-day); }

  const actData = actions.map((a, idx)=>({
    id: a.id || `A-${String(idx+1).padStart(2,'0')}`,
    title: a.name || a.title || `Action ${idx+1}`,
    start: a.start,
    end: a.end,
    responsable: a.responsable || '',
    description: a.description || ''
  })).filter(a=>a.start && a.end);
  if (actData.length === 0) return;

  const dates = actData.flatMap(a=>[parseDate(a.start), parseDate(a.end)]);
  let minDate = new Date(Math.min(...dates));
  let maxDate = new Date(Math.max(...dates));
  minDate = addDays(minDate, -1);
  maxDate = addDays(maxDate, 1);
  minDate = toMonday(minDate);
  maxDate = toSunday(maxDate);

  const totalDays = Math.max(1, daysBetween(minDate, maxDate));
  const chartWidth = leftPad + rightPad + totalDays * pxPerDay;
  const chartHeight = topPad + bottomPad + actData.length * (rowHeight + rowGap);

  const svgNS = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(svgNS,'svg');
  svg.setAttribute('width', chartWidth);
  svg.setAttribute('height', chartHeight);
  svg.setAttribute('role','img');
  svg.setAttribute('aria-label','Diagramme de Gantt');

  const grid=document.createElementNS(svgNS,'g'); grid.setAttribute('class','grid');
  for(let d=new Date(minDate); d<=maxDate; d=addDays(d,7)){
    const x = leftPad + daysBetween(minDate, d) * pxPerDay + 0.5;
    const line=document.createElementNS(svgNS,'line');
    line.setAttribute('x1',x); line.setAttribute('x2',x);
    line.setAttribute('y1',topPad); line.setAttribute('y2',chartHeight-bottomPad);
    grid.appendChild(line);
  }
  svg.appendChild(grid);

  const axisTop=document.createElementNS(svgNS,'g'); axisTop.setAttribute('class','axis month');
  let cursor=new Date(minDate.getFullYear(), minDate.getMonth(), 1);
  while(cursor <= maxDate){
    const monthStart = new Date(cursor);
    const monthEnd = new Date(cursor.getFullYear(), cursor.getMonth()+1, 1);
    const x1 = leftPad + Math.max(0, daysBetween(minDate, monthStart)) * pxPerDay + 2;
    const x2 = leftPad + Math.min(totalDays, daysBetween(minDate, monthEnd)) * pxPerDay - 2;
    const labelX = x1 + Math.max(0,(x2-x1))/2;
    const t=document.createElementNS(svgNS,'text');
    t.setAttribute('x', labelX);
    t.setAttribute('y', 16);
    t.setAttribute('text-anchor','middle');
    const mm = String(monthStart.getMonth()+1).padStart(2,'0');
    t.textContent = `${monthStart.getFullYear()}-${mm}`;
    axisTop.appendChild(t);
    cursor = monthEnd;
  }
  svg.appendChild(axisTop);

  const axisBot=document.createElementNS(svgNS,'g'); axisBot.setAttribute('class','axis');
  for(let d=new Date(minDate); d<=maxDate; d=addDays(d,7)){
    const x = leftPad + daysBetween(minDate, d) * pxPerDay + 4;
    const t=document.createElementNS(svgNS,'text');
    t.setAttribute('x', x);
    t.setAttribute('y', chartHeight - 14);
    t.textContent = `W${String(weekNumber(d)).padStart(2,'0')}`;
    axisBot.appendChild(t);
  }
  svg.appendChild(axisBot);

  const ylabels=document.createElementNS(svgNS,'g'); ylabels.setAttribute('class','ylabels');
  const barIndexById = {};
  actData.forEach((a, idx)=>{
    const y = topPad + idx*(rowHeight+rowGap) + rowHeight/2 + 5;
    const t=document.createElementNS(svgNS,'text');
    t.setAttribute('x', leftPad - 10);
    t.setAttribute('y', y);
    t.setAttribute('text-anchor','end');
    t.setAttribute('data-id', a.id);
    t.textContent = a.title;
    t.addEventListener('click', ()=> scrollToAction(a.id, true));
    ylabels.appendChild(t);
  });
  svg.appendChild(ylabels);

  const today=new Date();
  if(today>=minDate && today<=maxDate){
    const x = leftPad + daysBetween(minDate, new Date(today.getFullYear(), today.getMonth(), today.getDate())) * pxPerDay + 0.5;
    const tl=document.createElementNS(svgNS,'line');
    tl.setAttribute('class','today-line');
    tl.setAttribute('x1',x); tl.setAttribute('y1',topPad);
    tl.setAttribute('x2',x); tl.setAttribute('y2',chartHeight-bottomPad);
    svg.appendChild(tl);
  }

  const bars=document.createElementNS(svgNS,'g');
  actData.forEach((a, idx)=>{
    const start=parseDate(a.start), end=parseDate(a.end);
    const x = leftPad + daysBetween(minDate, start) * pxPerDay;
    const w = Math.max(pxPerDay * Math.max(1, daysBetween(start, end)+1) - 8, 8);
    const y = topPad + idx*(rowHeight+rowGap) + 4;
    const h = rowHeight - 8;

    const r=document.createElementNS(svgNS,'rect');
    r.setAttribute('class','bar');
    r.setAttribute('x',x); r.setAttribute('y',y);
    r.setAttribute('width',w); r.setAttribute('height',h);
    r.setAttribute('id', `bar-${a.id}`);
    r.setAttribute('data-id', a.id);
    r.setAttribute('data-title', a.title);
    r.setAttribute('data-resp', a.responsable);
    r.setAttribute('data-deb', a.start);
    r.setAttribute('data-fin', a.end);
    r.setAttribute('data-desc', a.description || '');

    r.addEventListener('mouseenter', ev=>showTip(ev,r));
    r.addEventListener('mouseleave', hideTip);
    r.addEventListener('mousemove', moveTip);

    bars.appendChild(r);
    barIndexById[a.id] = {x, y, w, h};
  });
  svg.appendChild(bars);

  container.appendChild(svg);

  function showTip(ev, el){
    if(!tooltip) return;
    tooltip.querySelector('.title').textContent = el.getAttribute('data-title') || '';
    tooltip.querySelector('.resp').textContent  = el.getAttribute('data-resp') || '';
    tooltip.querySelector('.deb').textContent   = el.getAttribute('data-deb') || '';
    tooltip.querySelector('.fin').textContent   = el.getAttribute('data-fin') || '';
    tooltip.querySelector('.desc').textContent  = el.getAttribute('data-desc') || '';
    tooltip.style.opacity='1';
    moveTip(ev);
  }
  function moveTip(ev){
    if(!tooltip) return;
    const pad=16, vw=window.innerWidth, vh=window.innerHeight;
    const rect=tooltip.getBoundingClientRect();
    let x=ev.clientX, y=ev.clientY;
    if(x-rect.width/2<pad) x=pad+rect.width/2;
    if(x+rect.width/2>vw-pad) x=vw-pad-rect.width/2;
    if(y-rect.height-12<pad) tooltip.style.transform='translate(-50%, 12px)'; else tooltip.style.transform='translate(-50%, calc(-100% - 10px))';
    tooltip.style.left=x+'px'; tooltip.style.top=y+'px';
  }
  function hideTip(){ if(tooltip) tooltip.style.opacity='0'; }

  function scrollToAction(id, hi=false){
    const meta=barIndexById[id]; if(!meta) return;
    const targetLeft = Math.max(0, meta.x - leftPad/2);
    const targetTop  = Math.max(0, meta.y - topPad/2);
    container.scrollTo({ left: targetLeft, top: targetTop, behavior:'smooth' });
    if(hi){
      const bar=document.getElementById(`bar-${id}`);
      if(bar){
        bar.classList.add('hi');
        setTimeout(()=> bar.classList.remove('hi'), 1200);
      }
    }
  }

  const sr=document.createElement('div');
  sr.className='sr-only';
  sr.textContent = `Gantt de ${actData.length} actions, de ${ymd(minDate)} a ${ymd(maxDate)}. Axe: mois en haut, semaines en bas.`;
  container.appendChild(sr);

  window.addEventListener('resize', ()=>{ /* svg width is content-based */ });
}
