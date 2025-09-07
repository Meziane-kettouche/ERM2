// Gantt chart renderer for Atelier 5 plan tab
function drawGanttChartSVG(container, actions) {
  if (!container) return;
  const tooltip = document.getElementById('gantt-tooltip');
  container.innerHTML = '';
  if (tooltip) tooltip.style.opacity = '0';
  if (!actions || actions.length === 0) return;

  const rowHeight = 34;
  const rowGap = 10;
  const leftPad = 160;
  const rightPad = 24;
  const topPad = 30;
  const bottomPad = 30;

  function parseDate(s){ const [y,m,d] = s.split('-').map(Number); return new Date(y, m-1, d); }
  function formatDate(d){
    const y = d.getFullYear();
    const m = String(d.getMonth()+1).padStart(2,'0');
    const day = String(d.getDate()).padStart(2,'0');
    return `${y}-${m}-${day}`;
  }
  function daysBetween(a,b){ return Math.round((b - a) / 86400000); }

  const actData = actions.map((a, idx) => ({
    id: `A-${idx+1}`,
    title: a.name || a.title || `Action ${idx+1}`,
    start: a.start,
    end: a.end,
    responsable: a.responsable || '',
    description: a.description || ''
  })).filter(a => a.start && a.end);
  if (actData.length === 0) return;

  const dates = actData.flatMap(a => [parseDate(a.start), parseDate(a.end)]);
  let minDate = new Date(Math.min(...dates.map(d=>d.getTime())));
  let maxDate = new Date(Math.max(...dates.map(d=>d.getTime())));
  const padDays = 1;
  minDate = new Date(minDate.getFullYear(), minDate.getMonth(), minDate.getDate()-padDays);
  maxDate = new Date(maxDate.getFullYear(), maxDate.getMonth(), maxDate.getDate()+padDays);

  const totalDays = Math.max(1, daysBetween(minDate, maxDate));
  const pxPerDay = 28;
  const chartWidth = leftPad + rightPad + totalDays * pxPerDay;
  const chartHeight = topPad + bottomPad + actData.length * (rowHeight + rowGap);

  const svgNS = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(svgNS,'svg');
  svg.setAttribute('width', chartWidth);
  svg.setAttribute('height', chartHeight);
  svg.setAttribute('role','img');
  svg.setAttribute('aria-label','Diagramme de Gantt');

  const grid = document.createElementNS(svgNS,'g'); grid.setAttribute('class','grid');
  for(let i=0;i<=totalDays;i++){
    const x = leftPad + i*pxPerDay + 0.5;
    const l = document.createElementNS(svgNS,'line');
    l.setAttribute('x1', x); l.setAttribute('y1', topPad);
    l.setAttribute('x2', x); l.setAttribute('y2', chartHeight - bottomPad);
    l.setAttribute('stroke-dasharray', i%7===0 ? '' : '2 4');
    grid.appendChild(l);
  }
  svg.appendChild(grid);

  const axis = document.createElementNS(svgNS,'g'); axis.setAttribute('class','axis');
  for(let i=0;i<=totalDays;i++){
    const d = new Date(minDate.getFullYear(), minDate.getMonth(), minDate.getDate()+i);
    if(i % 2 === 0 || i===0 || i===totalDays){
      const tx = document.createElementNS(svgNS,'text');
      tx.setAttribute('x', leftPad + i*pxPerDay + 2);
      tx.setAttribute('y', topPad - 8);
      tx.textContent = formatDate(d);
      axis.appendChild(tx);
    }
  }
  svg.appendChild(axis);

  const ylabels = document.createElementNS(svgNS,'g'); ylabels.setAttribute('class','ylabels');
  actData.forEach((a, idx)=>{
    const y = topPad + idx*(rowHeight+rowGap) + rowHeight/2 + 6;
    const t = document.createElementNS(svgNS,'text');
    t.setAttribute('x', leftPad - 8);
    t.setAttribute('y', y);
    t.setAttribute('text-anchor','end');
    t.textContent = a.title;
    ylabels.appendChild(t);
  });
  svg.appendChild(ylabels);

  const today = new Date();
  if(today >= minDate && today <= maxDate){
    const x = leftPad + daysBetween(minDate, new Date(today.getFullYear(), today.getMonth(), today.getDate())) * pxPerDay + 0.5;
    const tl = document.createElementNS(svgNS,'line');
    tl.setAttribute('class','today-line');
    tl.setAttribute('x1', x); tl.setAttribute('y1', topPad);
    tl.setAttribute('x2', x); tl.setAttribute('y2', chartHeight - bottomPad);
    svg.appendChild(tl);
  }

  const bars = document.createElementNS(svgNS,'g');
  actData.forEach((a, idx)=>{
    const start = parseDate(a.start);
    const end = parseDate(a.end);
    const x = leftPad + daysBetween(minDate, start) * pxPerDay;
    const w = Math.max(pxPerDay* Math.max(1, daysBetween(start, end)+1) - 6, 8);
    const y = topPad + idx*(rowHeight+rowGap) + 2;
    const h = rowHeight - 4;

    const r = document.createElementNS(svgNS,'rect');
    r.setAttribute('class','bar');
    r.setAttribute('x', x); r.setAttribute('y', y);
    r.setAttribute('width', w); r.setAttribute('height', h);
    r.setAttribute('data-title', a.title);
    r.setAttribute('data-resp', a.responsable);
    r.setAttribute('data-deb', a.start);
    r.setAttribute('data-fin', a.end);
    r.setAttribute('data-desc', a.description);

    r.addEventListener('mouseenter', ev => showTip(ev, r));
    r.addEventListener('mouseleave', hideTip);
    r.addEventListener('mousemove', moveTip);

    bars.appendChild(r);
  });
  svg.appendChild(bars);

  container.appendChild(svg);

  const sr = document.createElement('div');
  sr.className = 'sr-only';
  sr.textContent = `Diagramme de Gantt de ${actData.length} actions, du ${formatDate(minDate)} au ${formatDate(maxDate)}.`;
  container.appendChild(sr);

  function showTip(ev, el){
    if (!tooltip) return;
    tooltip.querySelector('.title').textContent = el.getAttribute('data-title') || '';
    tooltip.querySelector('.resp').textContent  = el.getAttribute('data-resp') || '';
    tooltip.querySelector('.deb').textContent   = el.getAttribute('data-deb') || '';
    tooltip.querySelector('.fin').textContent   = el.getAttribute('data-fin') || '';
    tooltip.querySelector('.desc').textContent  = el.getAttribute('data-desc') || '';
    tooltip.style.opacity = '1';
    moveTip(ev);
  }
  function moveTip(ev){
    if (!tooltip) return;
    const pad = 16;
    const vw = window.innerWidth, vh = window.innerHeight;
    const rect = tooltip.getBoundingClientRect();
    let x = ev.clientX, y = ev.clientY;
    if(x - rect.width/2 < pad) x = pad + rect.width/2;
    if(x + rect.width/2 > vw - pad) x = vw - pad - rect.width/2;
    if(y - rect.height - 16 < pad) tooltip.style.transform = 'translate(-50%, 12px)'; else tooltip.style.transform = 'translate(-50%, calc(-100% - 12px))';
    tooltip.style.left = x + 'px';
    tooltip.style.top  = y + 'px';
  }
  function hideTip(){ if (tooltip) tooltip.style.opacity = '0'; }
}
