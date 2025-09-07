function renderAtelier1StaticGraph() {
  const svg = document.getElementById('viz');
  const tip = document.getElementById('tip');
  if (!svg || !tip) return;
  svg.innerHTML = '';

  const nodes = [
    {id:'V1', type:'valeur', label:'Gestion RH'},
    {id:'V2', type:'valeur', label:'Facturation'},
    {id:'V3', type:'valeur', label:'Suivi Clients'},
    {id:'S1', type:'support', label:'Base RH'},
    {id:'S2', type:'support', label:'Serveur Finance'},
    {id:'S3', type:'support', label:'CRM'},
    {id:'E1', type:'event', label:'Fuite données', severity:4},
    {id:'E2', type:'event', label:'Indispo appli', severity:3},
    {id:'E3', type:'event', label:'Fraude paiement', severity:2},
    {id:'E4', type:'event', label:'Erreur saisie', severity:1}
  ];
  const links = [
    {id:'L1', source:'V1', target:'S1', weight:2},
    {id:'L2', source:'V2', target:'S2', weight:3},
    {id:'L3', source:'V3', target:'S3', weight:2},
    {id:'L4', source:'S1', target:'E1', weight:1},
    {id:'L5', source:'S1', target:'E2', weight:1},
    {id:'L6', source:'S2', target:'E2', weight:2},
    {id:'L7', source:'S2', target:'E3', weight:1},
    {id:'L8', source:'S3', target:'E4', weight:1},
    {id:'L9', source:'S3', target:'E1', weight:1}
  ];

  const sevColor = s => ({1:'#22c55e',2:'#eab308',3:'#f97316',4:'#ef4444'}[s] || '#6ea8fe');
  const byId = Object.fromEntries(nodes.map(n => [n.id, n]));
  const outEdges = {}, inEdges = {};
  links.forEach(l => {
    (outEdges[l.source] ||= []).push(l);
    (inEdges[l.target] ||= []).push(l);
  });
  function maxSevSupport(sId){
    let max = 0;
    (outEdges[sId] || []).forEach(l => {
      const t = byId[l.target];
      if (t.type === 'event') max = Math.max(max, t.severity || 0);
    });
    return max;
  }
  function maxSevValue(vId){
    let max = 0;
    (outEdges[vId] || []).forEach(l => {
      const s = byId[l.target];
      max = Math.max(max, maxSevSupport(s.id));
    });
    return max;
  }
  function fillForNode(n){
    if (n.type === 'event') return sevColor(n.severity || 0);
    if (n.type === 'support') return sevColor(maxSevSupport(n.id) || 0);
    if (n.type === 'valeur') return sevColor(maxSevValue(n.id) || 0);
    return '#6ea8fe';
  }

  const X = {valeur:160, support:540, event:900}, groups = {valeur:[], support:[], event:[]};
  nodes.forEach(n => groups[n.type].push(n));
  const top = 80, gap = 100;
  Object.keys(groups).forEach(t => groups[t].forEach((n,i) => { n.x = X[t]; n.y = top + i * gap; }));

  const NS = 'http://www.w3.org/2000/svg';
  const nodeEls = {}, linkEls = {};
  function bindTooltip(el, txtFn){
    el.addEventListener('mousemove', e => {
      tip.style.left = e.clientX + 'px';
      tip.style.top = e.clientY + 'px';
      tip.textContent = txtFn();
      tip.style.opacity = 1;
    });
    el.addEventListener('mouseleave', () => tip.style.opacity = 0);
  }
  function collectConnected(id){
    const N = new Set([id]);
    const L = new Set();
    const q = [id];
    const seen = new Set([id]);
    while (q.length){
      const cur = q.shift();
      (outEdges[cur] || []).forEach(l => {
        L.add(l.id);
        N.add(l.target);
        if (!seen.has(l.target)) { seen.add(l.target); q.push(l.target); }
      });
      (inEdges[cur] || []).forEach(l => {
        L.add(l.id);
        N.add(l.source);
        if (!seen.has(l.source)) { seen.add(l.source); q.push(l.source); }
      });
    }
    return {nodes:N, links:L};
  }
  function applyHighlight(keep){
    Object.values(nodeEls).forEach(g => g.classList.add('dim'));
    Object.values(linkEls).forEach(p => p.classList.add('dim'));
    keep.nodes.forEach(id => {
      const el = nodeEls[id];
      if (el){ el.classList.remove('dim'); el.classList.add('hi-node'); }
    });
    keep.links.forEach(id => {
      const el = linkEls[id];
      if (el){ el.classList.remove('dim'); el.classList.add('hi-link'); }
    });
  }
  function clearHighlight(){
    Object.values(nodeEls).forEach(g => g.classList.remove('dim','hi-node'));
    Object.values(linkEls).forEach(p => p.classList.remove('dim','hi-link'));
  }
  function drawLink(l){
    const a = byId[l.source], b = byId[l.target];
    const path = document.createElementNS(NS, 'path');
    const bend = 160, x1 = a.x, y1 = a.y, x2 = b.x, y2 = b.y;
    path.setAttribute('class','link');
    path.setAttribute('id', l.id);
    path.setAttribute('stroke-width', Math.max(3, (l.weight || 1) * 4));
    path.setAttribute('d', `M ${x1+40} ${y1} C ${x1+40+bend} ${y1}, ${x2-40-bend} ${y2}, ${x2-40} ${y2}`);
    bindTooltip(path, () => `${a.label} → ${b.label}`);
    svg.appendChild(path);
    linkEls[l.id] = path;
  }
  function drawNode(n){
    const g = document.createElementNS(NS, 'g');
    g.setAttribute('id', n.id);
    g.setAttribute('transform', `translate(${n.x},${n.y})`);
    let shape;
    if (n.type === 'valeur') {
      shape = document.createElementNS(NS, 'circle');
      shape.setAttribute('r', 26);
    } else if (n.type === 'support') {
      shape = document.createElementNS(NS, 'polygon');
      shape.setAttribute('points', '-30,0 -15,-26 15,-26 30,0 15,26 -15,26');
    } else {
      shape = document.createElementNS(NS, 'polygon');
      shape.setAttribute('points', '-34,-26 -34,26 24,0');
    }
    shape.setAttribute('class','shape');
    shape.setAttribute('fill', fillForNode(n));
    g.appendChild(shape);
    const t = document.createElementNS(NS, 'text');
    t.setAttribute('class','lbl');
    t.setAttribute('x',40);
    t.setAttribute('y',4);
    t.textContent = (n.type === 'event' && n.severity) ? `${n.label} [G${n.severity}]` : n.label;
    g.appendChild(t);
    g.addEventListener('mouseenter', () => {
      const keep = collectConnected(n.id);
      applyHighlight(keep);
    });
    g.addEventListener('mouseleave', clearHighlight);
    bindTooltip(g, () => {
      if (n.type === 'event') return `${n.label} (G${n.severity})`;
      if (n.type === 'support') return `${n.label} — max sev: ${maxSevSupport(n.id) || '-'}`;
      if (n.type === 'valeur') return `${n.label} — max sev: ${maxSevValue(n.id) || '-'}`;
      return n.label;
    });
    svg.appendChild(g);
    nodeEls[n.id] = g;
  }

  links.forEach(drawLink);
  nodes.forEach(drawNode);
}

