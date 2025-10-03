/*
  Client side script for the EBIOS RM offline risk management tool.
  The script handles data persistence via localStorage, dynamic form
  rendering for each of the five workshops, simple chart drawing
  routines, and import/export functionality.
*/

(function() {
  // ----- Data model and persistence
  let analyses = [];
  let currentIndex = -1;
  let stakeholderCanvas;
  let stakeholderCtx;
  let stakeholderChartResizeBound = false;
  let stakeholderPoints = [];
  let stakeholderTooltipEl;

  let atelier4ChartInstance;
  let risquesChartInstance;
  let risquesScoreChartInstance;
  let atelier4DragHandlers = null;
  let atelier5DragHandlers = null;

  if (typeof window !== 'undefined') {
    window.addEventListener('resize', () => {
      if (atelier4ChartInstance) {
        atelier4ChartInstance.resize();
      }
      if (risquesChartInstance) {
        risquesChartInstance.resize();
      }
      if (risquesScoreChartInstance) {
        risquesScoreChartInstance.resize();
      }
    });
  }

  function loadAnalyses() {
    try {
      const data = localStorage.getItem('ebiosAnalyses');
      analyses = data ? JSON.parse(data) : [];
      // Ensure every analysis has a stable identifier for persistence.
      // Older saved analyses may lack an `id` field, so assign one when
      // loading and immediately save back to storage so future loads keep it.
      let changed = false;
      analyses.forEach(a => {
        if (a && !a.id) {
          a.id = uid();
          changed = true;
        }
      });
      if (changed) {
        saveAnalyses();
      }
    } catch (e) {
      console.warn('Failed to parse localStorage data, resetting.', e);
      analyses = [];
    }
  }

  function saveAnalyses() {
    localStorage.setItem('ebiosAnalyses', JSON.stringify(analyses));
  }

  // Persist the currently selected analysis in localStorage so that
  // navigating between separate workshop pages restores the same
  // analysis automatically.  Both the stable `id` and the numerical
  // index are saved.  Storing the index avoids having to search for the
  // ID on the next page load, which previously could fail and caused the
  // first analysis to be selected by default.
  function persistCurrentAnalysisId() {
    try {
      const sel = analyses[currentIndex];
      if (sel && sel.id) {
        // Save any pending changes and remember the current analysis
        saveAnalyses();
        localStorage.setItem('ebiosCurrentAnalysisId', sel.id);
        localStorage.setItem('ebiosCurrentAnalysisIndex', String(currentIndex));
      }
    } catch (e) {
      // Ignore storage errors (e.g., private browsing)
    }
  }

  // ----- Utility: generate a simple UID
  function uid() {
    return 'id-' + Math.random().toString(36).substring(2, 10) + Date.now().toString(36);
  }

  // Map a level (1-4) to a background colour used across ateliers.
  function levelColor(lvl) {
    const num = parseFloat(lvl);
    if (!Number.isFinite(num)) return 'transparent';
    const bucket = Math.max(1, Math.min(4, Math.round(num)));
    switch (bucket) {
      case 1: return '#2a9d8f';
      case 2: return '#e9c46a';
      case 3: return '#f4a261';
      case 4: return '#e63946';
      default: return 'transparent';
    }
  }

  // Map SSI level (1-4) to colour where 1=red and 4=green.
  function ssiColor(lvl) {
    switch (parseInt(lvl, 10)) {
      case 1: return '#e63946';
      case 2: return '#f4a261';
      case 3: return '#e9c46a';
      case 4: return '#2a9d8f';
      default: return '#9aa0a6';
    }
  }

  // Determine the qualitative zone associated with a threat index.
  function threatZoneMeta(indice) {
    const value = Number.isFinite(indice) ? indice : 0;
    if (value >= 3) return { key: 'danger', label: 'Zone de danger', bucket: 4 };
    if (value >= 2) return { key: 'controle', label: 'Zone de contrôle', bucket: 3 };
    if (value >= 1) return { key: 'veille', label: 'Zone de veille', bucket: 2 };
    return { key: 'confiance', label: 'Zone de confiance', bucket: 1 };
  }

  // ----- Rendering functions
  function renderAnalysisList() {
    const listEl = document.getElementById('analysis-list');
    listEl.innerHTML = '';
    analyses.forEach((analysis, index) => {
      const item = document.createElement('div');
      item.className = 'analysis-item' + (index === currentIndex ? ' active' : '');
      const span = document.createElement('span');
      span.textContent = analysis.title || 'Nouvelle analyse';
      item.appendChild(span);
      item.addEventListener('click', () => {
        selectAnalysis(index);
      });
      listEl.appendChild(item);
    });
    // If there are no analyses yet, show a hint
    if (analyses.length === 0) {
      const hint = document.createElement('p');
      hint.textContent = 'Aucune analyse créée.';
      hint.style.fontStyle = 'italic';
      hint.style.color = 'var(--text-secondary)';
      listEl.appendChild(hint);
    }
  }

  function selectAnalysis(index) {
    currentIndex = index;
    // Persist selection so that changing workshops does not require the
    // user to pick the analysis again.
    persistCurrentAnalysisId();
    renderAnalysisList();
    const analysis = analyses[currentIndex];
    if (!analysis) return;
    document.getElementById('analysis-title').value = analysis.title || '';
    // Populate each atelier
    renderMissionDescription();
    renderMissionsTable();
    // Gap analysis (atelier1 second sub‑tab)
    renderGapTable();
    renderSROV();
    // Render PP list and cartography tables
    renderPP();
    renderPPCarto();
    renderStrategies();
    renderSS();
    renderSO();
    renderRisques();
    // Update charts/graph
    updateAtelier1Graph();
    updateGapChart();
    updateAtelier2Chart();
    updateAtelier3Chart();
    updateAtelier4Chart();
    updateAtelier5Chart();
    // Atelier 5 actions and plan
    renderGapActions();
    renderSupportActions();
    renderPartiesActions();
    renderRisquesActions();
    renderPlanActions();
    renderRisquesChart();

  }

  // ----- Rendering helper for generic items
  function createInput(labelText, type, value, onInput) {
    const container = document.createElement('div');
    const label = document.createElement('label');
    label.textContent = labelText;
    const input = type === 'textarea' ? document.createElement('textarea') : document.createElement('input');
    if (type !== 'textarea') input.type = type;
    input.value = value || '';
    input.addEventListener('input', (e) => {
      onInput(e.target.value);
    });
    container.appendChild(label);
    container.appendChild(input);
    return container;
  }

  // Create a select input with a set of options. Options should be an array
  // of objects: { value: string, label: string }. The selectedValue may be
  // empty or undefined. onChange is called with the selected value.
  function createSelect(labelText, options, selectedValue, onChange) {
    const container = document.createElement('div');
    const label = document.createElement('label');
    label.textContent = labelText;
    const select = document.createElement('select');
    // If there are no options, add a disabled placeholder
    if (!options || options.length === 0) {
      const opt = document.createElement('option');
      opt.value = '';
      opt.textContent = '— Aucune —';
      select.appendChild(opt);
    } else {
      options.forEach(opt => {
        const optionEl = document.createElement('option');
        optionEl.value = opt.value;
        optionEl.textContent = opt.label;
        if (opt.value === selectedValue) optionEl.selected = true;
        select.appendChild(optionEl);
      });
    }
    select.addEventListener('change', (e) => {
      onChange(e.target.value);
    });
    container.appendChild(label);
    container.appendChild(select);
    return container;
  }

  // Create a multi‑select input using a <select multiple> element. Options
  // should be an array of { value, label }. selectedValues is an array of
  // values. onChange receives an array of selected values.
  function createMultiSelect(labelText, options, selectedValues, onChange) {
    const container = document.createElement('div');
    const label = document.createElement('label');
    label.textContent = labelText;
    const select = document.createElement('select');
    select.multiple = true;
    if (!Array.isArray(selectedValues)) selectedValues = [];
    options.forEach(opt => {
      const optionEl = document.createElement('option');
      optionEl.value = opt.value;
      optionEl.textContent = opt.label;
      optionEl.selected = selectedValues.includes(opt.value);
      select.appendChild(optionEl);
    });
    select.addEventListener('change', (e) => {
      const selected = Array.from(e.target.selectedOptions).map(o => o.value);
      onChange(selected);
    });
    container.appendChild(label);
    container.appendChild(select);
    return container;
  }

  // Format a risk identifier by keeping only the characters before the first dash.
  function formatRiskIdentifier(identifier) {
    if (!identifier) return '';
    const str = String(identifier).trim();
    const dashIndex = str.search(/[-–—]/);
    if (dashIndex === -1) return str;
    return str.slice(0, dashIndex).trim();
  }

  function getCssVar(name, fallback) {
    const styles = getComputedStyle(document.documentElement);
    const value = styles.getPropertyValue(name);
    return value && value.trim() ? value.trim() : fallback;
  }

  function parseLevel(value, fallback) {
    const num = parseFloat(value);
    if (!Number.isFinite(num)) return fallback;
    return Math.max(0, Math.min(4, num));
  }

  function spreadInSquare(points, radius = 0.12) {
    const normalized = [];
    points.forEach(point => {
      if (!point || !Array.isArray(point.value) || point.value.length < 2) return;
      const x = parseFloat(point.value[0]);
      const y = parseFloat(point.value[1]);
      if (!Number.isFinite(x) || !Number.isFinite(y)) return;
      normalized.push({
        ...point,
        value: [Math.max(0, Math.min(4, x)), Math.max(0, Math.min(4, y))]
      });
    });

    const groups = new Map();
    normalized.forEach(point => {
      const key = `${Math.round(point.value[0] * 1000) / 1000}|${Math.round(point.value[1] * 1000) / 1000}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(point);
    });

    const result = [];
    groups.forEach(group => {
      if (group.length === 1) {
        const p = group[0];
        result.push({ ...p, value: [Number(p.value[0].toFixed(3)), Number(p.value[1].toFixed(3))] });
        return;
      }
      const baseX = group[0].value[0];
      const baseY = group[0].value[1];
      const cellBaseX = Math.min(Math.max(Math.floor(Math.min(baseX, 3)), 0), 3);
      const cellBaseY = Math.min(Math.max(Math.floor(Math.min(baseY, 3)), 0), 3);
      const minX = cellBaseX;
      const maxX = cellBaseX + 1;
      const minY = cellBaseY;
      const maxY = cellBaseY + 1;
      group.forEach((point, idx) => {
        const angle = (2 * Math.PI * idx) / group.length;
        let nx = baseX + radius * Math.cos(angle);
        let ny = baseY + radius * Math.sin(angle);
        nx = Math.max(minX + 0.05, Math.min(maxX - 0.05, nx));
        ny = Math.max(minY + 0.05, Math.min(maxY - 0.05, ny));
        nx = Math.max(0, Math.min(4, nx));
        ny = Math.max(0, Math.min(4, ny));
        result.push({ ...point, value: [Number(nx.toFixed(3)), Number(ny.toFixed(3))] });
      });
    });
    return result;
  }

  // ----- Atelier 1: Mission description
    function renderMissionDescription() {
      const textarea = document.getElementById('mission-description');
      const directionInput = document.getElementById('direction');
      const respMetierInput = document.getElementById('responsable-metier');
      const chefProjetInput = document.getElementById('chef-projet');
      if (!textarea || !directionInput || !respMetierInput || !chefProjetInput) return;
      const analysis = analyses[currentIndex];
      if (!analysis.data) analysis.data = {};
      if (typeof analysis.data.missionDescription !== 'string') analysis.data.missionDescription = '';
      if (typeof analysis.data.direction !== 'string') analysis.data.direction = '';
      if (typeof analysis.data.responsableMetier !== 'string') analysis.data.responsableMetier = '';
      if (typeof analysis.data.chefProjet !== 'string') analysis.data.chefProjet = '';
      textarea.value = analysis.data.missionDescription;
      directionInput.value = analysis.data.direction;
      respMetierInput.value = analysis.data.responsableMetier;
      chefProjetInput.value = analysis.data.chefProjet;
      textarea.oninput = (e) => {
        analysis.data.missionDescription = e.target.value;
        saveAnalyses();
      };
      directionInput.oninput = (e) => {
        analysis.data.direction = e.target.value;
        saveAnalyses();
      };
      respMetierInput.oninput = (e) => {
        analysis.data.responsableMetier = e.target.value;
        saveAnalyses();
      };
      chefProjetInput.oninput = (e) => {
        analysis.data.chefProjet = e.target.value;
        saveAnalyses();
      };
    }

  // ----- Atelier 1: Valeurs et supports (table rendering)
  function renderMissionsTable() {
    const tbody = document.getElementById('missions-body');
    if (!tbody) return;
    tbody.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis.data) analysis.data = {};
    if (!analysis.data.missions) analysis.data.missions = [];
    analysis.data.missions.forEach((mission, idx) => {
      // Ensure unique id and supports array with objects {name, description, responsable}
      if (!mission.id) mission.id = uid();
      // Normalise legacy data formats where supports might be stored as
      // strings or arrays of strings instead of the expected array of
      // objects.  This previously failed when mission.supports was already
      // an array because a guarding condition prevented the conversion.
      // Remove that guard and handle each case explicitly.
      if (Array.isArray(mission.supports)) {
        mission.supports = mission.supports.map(s => {
          if (typeof s === 'string') {
            return { name: s, description: '', responsable: '' };
          }
          return s;
        });
      } else if (typeof mission.supports === 'string' && mission.supports.trim() !== '') {
        mission.supports = mission.supports
          .split(',')
          .map(s => ({ name: s.trim(), description: '', responsable: '' }));
      } else {
        mission.supports = [];
      }
      const tr = document.createElement('tr');
      // Denomination
      let td = document.createElement('td');
      const denomInput = document.createElement('input');
      denomInput.type = 'text';
      denomInput.value = mission.denom || '';
      denomInput.oninput = (e) => {
        mission.denom = e.target.value;
        saveAnalyses();
        updateAtelier1Graph();
      };
      td.appendChild(denomInput);
      tr.appendChild(td);
      // Nature select
      td = document.createElement('td');
      const natureSelect = document.createElement('select');
      ['information','processus','fonction'].forEach(optVal => {
        const opt = document.createElement('option');
        opt.value = optVal;
        opt.textContent = optVal.charAt(0).toUpperCase() + optVal.slice(1);
        if ((mission.nature || '') === optVal) opt.selected = true;
        natureSelect.appendChild(opt);
      });
      natureSelect.onchange = (e) => {
        mission.nature = e.target.value;
        saveAnalyses();
      };
      td.appendChild(natureSelect);
      tr.appendChild(td);
      // Description (use textarea for long text)
      td = document.createElement('td');
      const descInput = document.createElement('textarea');
      descInput.rows = 2;
      descInput.value = mission.description || '';
      descInput.style.width = '100%';
      descInput.oninput = (e) => {
        mission.description = e.target.value;
        saveAnalyses();
      };
      td.appendChild(descInput);
      tr.appendChild(td);
      // Responsable
      td = document.createElement('td');
      const respInput = document.createElement('input');
      respInput.type = 'text';
      respInput.value = mission.responsable || '';
      respInput.oninput = (e) => {
        mission.responsable = e.target.value;
        saveAnalyses();
      };
      td.appendChild(respInput);
      tr.appendChild(td);
      // Supports cell
      td = document.createElement('td');
      const supportsCell = document.createElement('div');
      supportsCell.className = 'supports-cell';
      mission.supports.forEach((support, sIdx) => {
        if (!support.id) support.id = uid();
        const sItem = document.createElement('div');
        sItem.className = 'support-item';
        const nameInput = document.createElement('input');
        nameInput.type = 'text';
        nameInput.placeholder = 'Nom';
        nameInput.value = support.name || '';
        nameInput.oninput = (e) => {
          support.name = e.target.value;
          saveAnalyses();
          renderSupportsQualifTable();
          renderSupportActions();
          updateAtelier1Graph();
        };
        const supDescInput = document.createElement('textarea');
        supDescInput.rows = 2;
        supDescInput.placeholder = 'Description';
        supDescInput.value = support.description || '';
        supDescInput.style.width = '100%';
        supDescInput.oninput = (e) => {
          support.description = e.target.value;
          saveAnalyses();
          renderSupportsQualifTable();
        };
        const supRespInput = document.createElement('input');
        supRespInput.type = 'text';
        supRespInput.placeholder = 'Responsable';
        supRespInput.value = support.responsable || '';
        supRespInput.oninput = (e) => {
          support.responsable = e.target.value;
          saveAnalyses();
          renderSupportsQualifTable();
        };
        const rmBtn = document.createElement('button');
        rmBtn.textContent = '×';
        rmBtn.title = 'Supprimer ce bien support';
        rmBtn.addEventListener('click', () => {
          mission.supports.splice(sIdx, 1);
          saveAnalyses();
          renderMissionsTable();
          renderSupportsQualifTable();
          updateAtelier1Graph();
        });
        sItem.appendChild(nameInput);
        sItem.appendChild(supDescInput);
        sItem.appendChild(supRespInput);
        sItem.appendChild(rmBtn);
        supportsCell.appendChild(sItem);
      });
      const addSupBtn = document.createElement('button');
      addSupBtn.className = 'add-support-btn';
      addSupBtn.textContent = '+ Support';
      addSupBtn.addEventListener('click', () => {
        mission.supports.push({ id: uid(), name: '', description: '', responsable: '' });
        saveAnalyses();
        renderMissionsTable();
        renderSupportsQualifTable();
        updateAtelier1Graph();
      });
      supportsCell.appendChild(addSupBtn);
      // Button to add an existing support from other missions
      const addExistingBtn = document.createElement('button');
      addExistingBtn.className = 'add-support-btn';
      addExistingBtn.textContent = '+ Support existant';
      addExistingBtn.addEventListener('click', () => {
        // Collect all unique supports across missions and vulnerability table
        const allSupports = [];
        const addUnique = (s) => {
          const name = (s.name || '').trim();
          if (!name) return;
          if (!allSupports.some(ss => ss.name === name)) {
            const id = s.id || s.refId;
            allSupports.push({ id: id, name: name, description: s.description || '', responsable: s.responsable || '' });
          }
        };
        (analysis.data.missions || []).forEach(m2 => {
          (m2.supports || []).forEach(addUnique);
        });
        (analysis.data.supportsQualif || []).forEach(addUnique);
        // Exclude supports already associated with the current mission
        const currentNames = mission.supports.map(s => (s.name || '').trim());
        const available = allSupports.filter(s => !currentNames.includes(s.name));
        if (available.length === 0) {
          alert('Aucun bien support existant disponible.');
          return;
        }
        const msg = 'Sélectionnez un bien support existant:\n' + available.map((s, i) => `${i + 1}. ${s.name}`).join('\n');
        const input = prompt(msg);
        if (input === null) return;
        const index = parseInt(input, 10) - 1;
        if (!isNaN(index) && index >= 0 && index < available.length) {
          // Clone the selected support so edits in this mission do not
          // modify the original object from another mission.
          const selected = Object.assign({}, available[index]);
          mission.supports.push(selected);
          saveAnalyses();
          renderMissionsTable();
          renderSupportsQualifTable();
          updateAtelier1Graph();
        }
      });
      supportsCell.appendChild(addExistingBtn);
      td.appendChild(supportsCell);
      tr.appendChild(td);
      // Events cell: associate events with this mission
      td = document.createElement('td');
      const eventsCell = document.createElement('div');
      eventsCell.className = 'events-cell';
      // Gather events tied to this mission
      const events = (analysis.data.events || []).filter(ev => ev.missionId === mission.id);
      // Helper to colour impact select
      function colorImpact(selectEl) {
        const lvl = parseInt(selectEl.value, 10);
        let color;
        switch (lvl) {
          case 1: color = '#2a9d8f'; break;
          case 2: color = '#e9c46a'; break;
          case 3: color = '#f4a261'; break;
          case 4: color = '#e63946'; break;
          default:
            color = getComputedStyle(document.documentElement).getPropertyValue('--bg-light').trim() || '#12203a';
        }
        selectEl.style.backgroundColor = color;
      }
      events.forEach((event, eIdx) => {
        const evItem = document.createElement('div');
        evItem.className = 'event-item';
        // Event description
        const evDesc = document.createElement('textarea');
        evDesc.rows = 2;
        evDesc.placeholder = 'Évènement';
        evDesc.value = event.evenement || '';
        evDesc.oninput = (e) => {
          event.evenement = e.target.value;
          saveAnalyses();
        };
        evItem.appendChild(evDesc);
        // Impact description
        const impDesc = document.createElement('textarea');
        impDesc.rows = 2;
        impDesc.placeholder = 'Description des impacts';
        impDesc.value = event.impactDescription || '';
        impDesc.oninput = (e) => {
          event.impactDescription = e.target.value;
          saveAnalyses();
        };
        evItem.appendChild(impDesc);
        // Impact level select
        const impSelect = document.createElement('select');
        [1,2,3,4].forEach(num => {
          const opt = document.createElement('option');
          opt.value = num;
          opt.textContent = num;
          if (event.impact === num) opt.selected = true;
          impSelect.appendChild(opt);
        });
        // initial colour
        colorImpact(impSelect);
        impSelect.onchange = (e) => {
          event.impact = parseInt(e.target.value, 10);
          saveAnalyses();
          colorImpact(impSelect);
          updateAtelier1Graph();
        };
        evItem.appendChild(impSelect);
        // Remove event button
        const rmEvBtn = document.createElement('button');
        rmEvBtn.textContent = '×';
        rmEvBtn.title = 'Supprimer cet évènement';
        rmEvBtn.addEventListener('click', () => {
          // Remove this event from analysis.data.events
          const idxEv = analysis.data.events.findIndex(ev => ev.id === event.id);
          if (idxEv >= 0) {
            analysis.data.events.splice(idxEv, 1);
            saveAnalyses();
            renderMissionsTable();
            updateAtelier1Graph();
          }
        });
        evItem.appendChild(rmEvBtn);
        eventsCell.appendChild(evItem);
      });
      // Button to add a new event for this mission
      const addEvBtn = document.createElement('button');
      addEvBtn.className = 'add-event-btn';
      addEvBtn.textContent = '+ Évènement';
      addEvBtn.addEventListener('click', () => {
        if (!analysis.data.events) analysis.data.events = [];
        analysis.data.events.push({ id: uid(), missionId: mission.id, evenement: '', impactDescription: '', impact: 1 });
        saveAnalyses();
        renderMissionsTable();
        updateAtelier1Graph();
      });
      eventsCell.appendChild(addEvBtn);
      td.appendChild(eventsCell);
      tr.appendChild(td);
      // Actions: delete mission
      td = document.createElement('td');
      const delBtn = document.createElement('button');
      delBtn.className = 'delete-item';
      delBtn.textContent = '×';
      delBtn.title = 'Supprimer cette valeur';
      delBtn.addEventListener('click', () => {
        if (!confirm('Supprimer cette mission ?')) return;
        analysis.data.missions.splice(idx, 1);
        // Remove associated events when mission deleted
        if (analysis.data.events) {
          analysis.data.events = analysis.data.events.filter(ev => ev.missionId !== mission.id);
        }
        saveAnalyses();
        renderMissionsTable();
        renderSupportsQualifTable();
        updateAtelier1Graph();
      });
      td.appendChild(delBtn);
      tr.appendChild(td);
      tbody.appendChild(tr);
    });
    // After rendering rows, set up resizable columns on the missions table
    addMissionTableResizers();
    renderSupportsQualifTable();
    updateAtelier1Graph();
  }

  // Add resizer handles to the header of the missions table.  Users
  // can drag these handles to adjust the width of each column.  This
  // function is called after renderMissionsTable() creates the table.
  function addMissionTableResizers() {
    const table = document.getElementById('missions-table');
    if (!table) return;
    const ths = table.querySelectorAll('thead th');
    const rows = table.querySelectorAll('tbody tr');
    ths.forEach((th, index) => {
      // Remove any existing resizer to avoid duplicates
      const existing = th.querySelector('.col-resizer');
      if (existing) th.removeChild(existing);
      const resizer = document.createElement('div');
      resizer.className = 'col-resizer';
      th.appendChild(resizer);
      resizer.addEventListener('mousedown', (e) => {
        e.preventDefault();
        const startX = e.clientX;
        const startWidth = th.offsetWidth;
        // Determine minimum width from CSS or fallback
        const computedStyle = window.getComputedStyle(th);
        const minWidth = parseInt(computedStyle.minWidth) || 80;
        function onMouseMove(ev) {
          const delta = ev.clientX - startX;
          let newWidth = startWidth + delta;
          if (newWidth < minWidth) newWidth = minWidth;
          th.style.width = newWidth + 'px';
          th.style.minWidth = newWidth + 'px';
          // Apply width to all corresponding cells in the column
          rows.forEach(row => {
            const cell = row.children[index];
            if (cell) {
              cell.style.width = newWidth + 'px';
              cell.style.minWidth = newWidth + 'px';
            }
          });
        }
        function onMouseUp() {
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
        }
        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
      });
    });
  }

  // ----- Atelier 1: Qualification des biens supports -----
  function vulnLevelColor(lvl) {
    switch ((lvl || '').toLowerCase()) {
      case 'info':
        return '#6c757d';
      case 'faible':
        return '#2a9d8f';
      case 'moderee':
        return '#e9c46a';
      case 'forte':
      case 'fort':
        return '#f4a261';
      case 'critique':
        return '#e63946';
      default:
        return '#9aa0a6';
    }
  }

  function setVulnLevelColor(selectEl) {
    const lvl = (selectEl.value || '').toLowerCase();
    const color = vulnLevelColor(lvl);
    if (lvl === 'forte' || lvl === 'fort' || lvl === 'critique') selectEl.style.color = '#fff';
    else selectEl.style.color = '';
    selectEl.style.backgroundColor = color;
  }

  function renderSupportsQualifTable() {
    const tbody = document.getElementById('supports-qualif-body');
    if (!tbody) return;
    tbody.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis.data) analysis.data = {};
    if (!Array.isArray(analysis.data.supportsQualif)) analysis.data.supportsQualif = [];

    // Synchronize supports from missions to avoid duplicates when editing
    const sq = analysis.data.supportsQualif;
    const validIds = new Set();
    let changed = false;
    (analysis.data.missions || []).forEach(m => {
      (m.supports || []).forEach(s => {
        if (!s.id) { s.id = uid(); changed = true; }
        validIds.add(s.id);
        let entry = sq.find(e => e.refId === s.id);
        if (entry) {
          if (entry.name !== (s.name || '') || entry.description !== (s.description || '') || entry.responsable !== (s.responsable || '')) {
            entry.name = s.name || '';
            entry.description = s.description || '';
            entry.responsable = s.responsable || '';
            changed = true;
          }
        } else {
          sq.push({ refId: s.id, name: s.name || '', description: s.description || '', responsable: s.responsable || '', vulnerabilities: [] });
          changed = true;
        }
      });
    });
    // Remove supports that no longer exist in missions
    for (let i = sq.length - 1; i >= 0; i--) {
      const entry = sq[i];
      if (entry.refId && !validIds.has(entry.refId)) {
        sq.splice(i, 1);
        changed = true;
      }
    }
    // Deduplicate by name
    const seenNames = new Set();
    for (let i = sq.length - 1; i >= 0; i--) {
      const n = sq[i].name || '';
      if (seenNames.has(n)) {
        sq.splice(i, 1);
        changed = true;
      } else {
        seenNames.add(n);
      }
    }
    if (changed) saveAnalyses();

    analysis.data.supportsQualif.forEach((support, idx) => {
      const tr = document.createElement('tr');
      // name
      let td = document.createElement('td');
      const nameInput = document.createElement('input');
      nameInput.type = 'text';
      nameInput.value = support.name || '';
      nameInput.oninput = (e) => {
        support.name = e.target.value;
        saveAnalyses();
        renderSupportActions();
      };
      td.appendChild(nameInput);
      tr.appendChild(td);
      // description
      td = document.createElement('td');
      const descInput = document.createElement('textarea');
      descInput.rows = 2;
      descInput.style.width = '100%';
      descInput.value = support.description || '';
      descInput.oninput = (e) => {
        support.description = e.target.value;
        saveAnalyses();
      };
      td.appendChild(descInput);
      tr.appendChild(td);
      // responsable
      td = document.createElement('td');
      const respInput = document.createElement('input');
      respInput.type = 'text';
      respInput.value = support.responsable || '';
      respInput.oninput = (e) => {
        support.responsable = e.target.value;
        saveAnalyses();
      };
      td.appendChild(respInput);
      tr.appendChild(td);
      // vulnerabilities
      td = document.createElement('td');
      const vulnDiv = document.createElement('div');
      vulnDiv.className = 'vuln-cell';
      if (!Array.isArray(support.vulnerabilities)) support.vulnerabilities = [];
      support.vulnerabilities.forEach((v, vIdx) => {
        const vItem = document.createElement('div');
        vItem.className = 'vuln-item';
        const vName = document.createElement('input');
        vName.type = 'text';
        vName.placeholder = 'Nom';
        vName.value = v.name || '';
        vName.oninput = (e) => {
          v.name = e.target.value;
          saveAnalyses();
          renderSupportActions();
        };
        const vDesc = document.createElement('textarea');
        vDesc.rows = 2;
        vDesc.placeholder = 'Description';
        vDesc.value = v.description || '';
        vDesc.oninput = (e) => {
          v.description = e.target.value;
          saveAnalyses();
        };
        const vLevel = document.createElement('select');
        ['info','faible','moderee','forte','critique'].forEach(optVal => {
          const opt = document.createElement('option');
          opt.value = optVal;
          opt.textContent = optVal.charAt(0).toUpperCase() + optVal.slice(1);
          if ((v.level || '') === optVal) opt.selected = true;
          vLevel.appendChild(opt);
        });
        setVulnLevelColor(vLevel);
        vLevel.onchange = (e) => {
          v.level = e.target.value;
          saveAnalyses();
          setVulnLevelColor(vLevel);
          renderVulnChart();
        };
        const rmV = document.createElement('button');
        rmV.textContent = '×';
        rmV.title = 'Supprimer cette vulnérabilité';
        rmV.addEventListener('click', () => {
          support.vulnerabilities.splice(vIdx, 1);
          saveAnalyses();
          renderSupportsQualifTable();
          renderSupportActions();
        });
        vItem.appendChild(vName);
        vItem.appendChild(vDesc);
        vItem.appendChild(vLevel);
        vItem.appendChild(rmV);
        vulnDiv.appendChild(vItem);
      });
      const addVBtn = document.createElement('button');
      addVBtn.className = 'add-support-btn';
      addVBtn.textContent = '+ Vulnérabilité';
      addVBtn.addEventListener('click', () => {
        support.vulnerabilities.push({ name:'', description:'', level:'info' });
        saveAnalyses();
        renderSupportsQualifTable();
        renderSupportActions();
      });
      vulnDiv.appendChild(addVBtn);
      td.appendChild(vulnDiv);
      tr.appendChild(td);
      // actions
      td = document.createElement('td');
      const delSup = document.createElement('button');
      delSup.className = 'delete-item';
      delSup.textContent = '×';
      delSup.title = 'Supprimer ce bien support';
      delSup.addEventListener('click', () => {
        if (!confirm('Supprimer ce bien support ?')) return;
        analysis.data.supportsQualif.splice(idx, 1);
        saveAnalyses();
        renderSupportsQualifTable();
        renderSupportActions();
      });
      td.appendChild(delSup);
      tr.appendChild(td);
      tbody.appendChild(tr);
    });
    addDataTableResizers('supports-qualif-table');
    renderVulnChart();
  }

  function renderVulnChart() {
    const canvas = document.getElementById('vuln-level-chart');
    if (!canvas) return;
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) {
      clearCanvas(canvas);
      return;
    }
    const levels = ['info','faible','moderee','forte','critique'];
    const counts = { info:0, faible:0, moderee:0, forte:0, critique:0 };
    (analysis.data.supportsQualif || []).forEach(s => {
      (s.vulnerabilities || []).forEach(v => {
        const lvl = (v.level || '').toLowerCase();
        if (counts.hasOwnProperty(lvl)) counts[lvl]++;
      });
    });
    const labels = levels.map(l => l.charAt(0).toUpperCase() + l.slice(1));
    const data = levels.map(l => counts[l]);
    const colors = levels.map(l => vulnLevelColor(l));
    if (data.some(v => v > 0)) drawBarChart(canvas, labels, data, colors);
    else clearCanvas(canvas);
  }

  // ----- Atelier 1: Évènements (table rendering)
  function renderEventsTable() {
    const tbody = document.getElementById('events-body');
    if (!tbody) return;
    tbody.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis.data) analysis.data = {};
    if (!analysis.data.events) analysis.data.events = [];
    const missions = analysis.data.missions || [];
    analysis.data.events.forEach((ev, idx) => {
      if (!ev.id) ev.id = uid();
      // ensure impact is numeric 1-4
      if (!ev.impact || typeof ev.impact !== 'number') {
        const parsed = parseInt(ev.impact, 10);
        ev.impact = (parsed >= 1 && parsed <= 4) ? parsed : 1;
      }
      const tr = document.createElement('tr');
      // mission select
      let td = document.createElement('td');
      const missionSelect = document.createElement('select');
      missions.forEach((m, i) => {
        const opt = document.createElement('option');
        opt.value = m.id;
        opt.textContent = m.denom || `Valeur ${i + 1}`;
        if ((ev.missionId || '') === m.id) opt.selected = true;
        missionSelect.appendChild(opt);
      });
      missionSelect.onchange = (e) => {
        ev.missionId = e.target.value;
        saveAnalyses();
        updateAtelier1Graph();
      };
      td.appendChild(missionSelect);
      tr.appendChild(td);
      // Event description (textarea for long text)
      td = document.createElement('td');
      const evInput = document.createElement('textarea');
      evInput.rows = 2;
      evInput.value = ev.evenement || '';
      evInput.style.width = '100%';
      evInput.oninput = (e) => {
        ev.evenement = e.target.value;
        saveAnalyses();
      };
      td.appendChild(evInput);
      tr.appendChild(td);
      // Impact description (textarea)
      td = document.createElement('td');
      const impactDescInput = document.createElement('textarea');
      impactDescInput.rows = 2;
      impactDescInput.value = ev.impactDescription || '';
      impactDescInput.style.width = '100%';
      impactDescInput.oninput = (e) => {
        ev.impactDescription = e.target.value;
        saveAnalyses();
      };
      td.appendChild(impactDescInput);
      tr.appendChild(td);
      // Impact select (1-4)
      td = document.createElement('td');
      const impactSelect = document.createElement('select');
      [1,2,3,4].forEach(num => {
        const opt = document.createElement('option');
        opt.value = num;
        opt.textContent = num;
        if (ev.impact === num) opt.selected = true;
        impactSelect.appendChild(opt);
      });
      // Color code the cell based on impact
      function updateImpactStyle() {
        // Determine the background colour based on the selected
        // impact level.  When the level is outside 1–4, fall back
        // to the interface's secondary dark colour.  Because CSS
        // variables are not directly usable in JS, retrieve the
        // computed value from the document root.
        const lvl = parseInt(impactSelect.value, 10);
        let color;
        switch (lvl) {
          case 1: color = '#2a9d8f'; break; // vert
          case 2: color = '#e9c46a'; break; // jaune
          case 3: color = '#f4a261'; break; // orange
          case 4: color = '#e63946'; break; // rouge
          default:
            // fallback to the theme's light background colour
            color = getComputedStyle(document.documentElement).getPropertyValue('--bg-light').trim() || '#12203a';
        }
        impactSelect.style.backgroundColor = color;
      }
      impactSelect.onchange = (e) => {
        ev.impact = parseInt(e.target.value, 10);
        saveAnalyses();
        updateImpactStyle();
        updateAtelier1Graph();
      };
      // initial style
      updateImpactStyle();
      td.appendChild(impactSelect);
      tr.appendChild(td);
      // Actions: delete event
      td = document.createElement('td');
      const delBtn = document.createElement('button');
      delBtn.className = 'delete-item';
      delBtn.textContent = '×';
      delBtn.title = 'Supprimer cet évènement';
      delBtn.addEventListener('click', () => {
        if (!confirm('Supprimer cet évènement ?')) return;
        analysis.data.events.splice(idx, 1);
        saveAnalyses();
        renderEventsTable();
        updateAtelier1Graph();
      });
      td.appendChild(delBtn);
      tr.appendChild(td);
      tbody.appendChild(tr);
    });
  }

  // ----- Atelier 1: GAP analysis table rendering
  function renderGapTable() {
    const tbody = document.getElementById('gap-body');
    if (!tbody) return;
    tbody.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis) return;
    if (!analysis.data) analysis.data = {};
    if (!analysis.data.gap) analysis.data.gap = [];
    analysis.data.gap.forEach((req, idx) => {
      if (!req.id) req.id = uid();
      const tr = document.createElement('tr');
      // Domaine
      let td = document.createElement('td');
      const domInput = document.createElement('input');
      domInput.type = 'text';
      domInput.value = req.domaine || '';
      domInput.oninput = (e) => {
        req.domaine = e.target.value;
        saveAnalyses();
      };
      td.appendChild(domInput);
      tr.appendChild(td);
      // Titre
      td = document.createElement('td');
      const titreInput = document.createElement('input');
      titreInput.type = 'text';
      titreInput.value = req.titre || '';
      titreInput.oninput = (e) => {
        req.titre = e.target.value;
        saveAnalyses();
      };
      td.appendChild(titreInput);
      tr.appendChild(td);
      // Description
      td = document.createElement('td');
      const descInput = document.createElement('textarea');
      descInput.rows = 2;
      descInput.value = req.description || '';
      descInput.style.width = '100%';
      descInput.oninput = (e) => {
        req.description = e.target.value;
        saveAnalyses();
      };
      td.appendChild(descInput);
      tr.appendChild(td);
      // Application
      td = document.createElement('td');
      const appSelect = document.createElement('select');
      const options = [
        { value:'Appliqué', label:'Appliqué' },
        { value:'Partiellement appliqué', label:'Partiellement appliqué' },
        { value:'Non appliqué', label:'Non appliqué' },
        { value:'Non applicable', label:'Non applicable' }
      ];
      options.forEach(opt => {
        const optEl = document.createElement('option');
        optEl.value = opt.value;
        optEl.textContent = opt.label;
        if ((req.application || '') === opt.value) optEl.selected = true;
        appSelect.appendChild(optEl);
      });
      appSelect.onchange = (e) => {
        req.application = e.target.value;
        saveAnalyses();
        updateGapChart();
      };
      td.appendChild(appSelect);
      tr.appendChild(td);
      // Justification
      td = document.createElement('td');
      const justInput = document.createElement('textarea');
      justInput.rows = 2;
      justInput.value = req.justification || '';
      justInput.style.width = '100%';
      justInput.oninput = (e) => {
        req.justification = e.target.value;
        saveAnalyses();
      };
      td.appendChild(justInput);
      tr.appendChild(td);
      // Actions: delete requirement
      td = document.createElement('td');
      const delBtn = document.createElement('button');
      delBtn.className = 'delete-item';
      delBtn.textContent = '×';
      delBtn.title = 'Supprimer cette exigence';
      delBtn.addEventListener('click', () => {
        if (!confirm('Supprimer cette exigence ?')) return;
        analysis.data.gap.splice(idx, 1);
        saveAnalyses();
        renderGapTable();
        updateGapChart();
      });
      td.appendChild(delBtn);
      tr.appendChild(td);
      tbody.appendChild(tr);
    });
    // After rendering the GAP table, make its columns resizable like the missions table
    addGapTableResizers();
  }

  // Add resizer handles to the header of the GAP analysis table.  Users
  // can drag these handles to adjust the width of each column.  This
  // helper mirrors the behaviour of addMissionTableResizers() but
  // targets the GAP table instead of the missions table.
  function addGapTableResizers() {
    const table = document.getElementById('gap-table');
    if (!table) return;
    const ths = table.querySelectorAll('thead th');
    const rows = table.querySelectorAll('tbody tr');
    ths.forEach((th, index) => {
      // Remove existing resizer to avoid duplicates
      const existing = th.querySelector('.col-resizer');
      if (existing) th.removeChild(existing);
      const resizer = document.createElement('div');
      resizer.className = 'col-resizer';
      th.appendChild(resizer);
      resizer.addEventListener('mousedown', (e) => {
        e.preventDefault();
        const startX = e.clientX;
        const startWidth = th.offsetWidth;
        const computedStyle = window.getComputedStyle(th);
        const minWidth = parseInt(computedStyle.minWidth) || 80;
        function onMouseMove(ev) {
          const delta = ev.clientX - startX;
          let newWidth = startWidth + delta;
          if (newWidth < minWidth) newWidth = minWidth;
          th.style.width = newWidth + 'px';
          th.style.minWidth = newWidth + 'px';
          rows.forEach(row => {
            const cell = row.children[index];
            if (cell) {
              cell.style.width = newWidth + 'px';
              cell.style.minWidth = newWidth + 'px';
            }
          });
        }
        function onMouseUp() {
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
        }
        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
      });
    });
  }

  // Add resizer handles to the header of the SROV table.  This function
  // mirrors the behaviour of the missions and GAP tables, allowing
  // column widths to be adjusted by dragging the edges of header cells.
  function addSrovTableResizers() {
    const table = document.getElementById('srov-table');
    if (!table) return;
    const ths = table.querySelectorAll('thead th');
    const rows = table.querySelectorAll('tbody tr');
    ths.forEach((th, index) => {
      const existing = th.querySelector('.col-resizer');
      if (existing) th.removeChild(existing);
      const resizer = document.createElement('div');
      resizer.className = 'col-resizer';
      th.appendChild(resizer);
      resizer.addEventListener('mousedown', (e) => {
        e.preventDefault();
        const startX = e.clientX;
        const startWidth = th.offsetWidth;
        const computedStyle = window.getComputedStyle(th);
        const minWidth = parseInt(computedStyle.minWidth) || 80;
        function onMouseMove(ev) {
          const delta = ev.clientX - startX;
          let newWidth = startWidth + delta;
          if (newWidth < minWidth) newWidth = minWidth;
          th.style.width = newWidth + 'px';
          th.style.minWidth = newWidth + 'px';
          rows.forEach(row => {
            const cell = row.children[index];
            if (cell) {
              cell.style.width = newWidth + 'px';
              cell.style.minWidth = newWidth + 'px';
            }
          });
        }
        function onMouseUp() {
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
        }
        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
      });
    });
  }

  // Add resizer handles to the cartography table columns.  This mirrors
  // the behaviour implemented for missions and SROV tables, allowing
  // users to adjust column widths with the mouse.  The last column
  // (actions) is excluded from resizing.
  function addPPCartoTableResizers() {
    const table = document.getElementById('ppc-table');
    if (!table) return;
    const ths = table.querySelectorAll('thead th');
    const rows = table.querySelectorAll('tbody tr');
    ths.forEach((th, index) => {
      const existing = th.querySelector('.col-resizer');
      if (existing) th.removeChild(existing);
      const resizer = document.createElement('div');
      resizer.className = 'col-resizer';
      th.appendChild(resizer);
      resizer.addEventListener('mousedown', (e) => {
        e.preventDefault();
        const startX = e.clientX;
        const startWidth = th.offsetWidth;
        const computedStyle = window.getComputedStyle(th);
        const minWidth = parseInt(computedStyle.minWidth) || 80;
        function onMouseMove(ev) {
          const delta = ev.clientX - startX;
          let newWidth = startWidth + delta;
          if (newWidth < minWidth) newWidth = minWidth;
          th.style.width = newWidth + 'px';
          th.style.minWidth = newWidth + 'px';
          rows.forEach(row => {
            const cell = row.children[index];
            if (cell) {
              cell.style.width = newWidth + 'px';
              cell.style.minWidth = newWidth + 'px';
            }
          });
        }
        function onMouseUp() {
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
        }
        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
      });
    });
  }

  // Add resizer handles to the operational scenarios table columns.  This mirrors
  // the behaviour implemented for other tables such as missions and SROV,
  // allowing users to adjust column widths by dragging the header edges. The
  // last column (actions) is excluded from resizing.
  function addOpsTableResizers() {
    const table = document.getElementById('ops-table');
    if (!table) return;
    const ths = table.querySelectorAll('thead th');
    const rows = table.querySelectorAll('tbody tr');
    ths.forEach((th, index) => {
      const existing = th.querySelector('.col-resizer');
      if (existing) th.removeChild(existing);
      const resizer = document.createElement('div');
      resizer.className = 'col-resizer';
      th.appendChild(resizer);
      resizer.addEventListener('mousedown', (e) => {
        e.preventDefault();
        const startX = e.clientX;
        const startWidth = th.offsetWidth;
        const computedStyle = window.getComputedStyle(th);
        const minWidth = parseInt(computedStyle.minWidth) || 80;
        function onMouseMove(ev) {
          const delta = ev.clientX - startX;
          let newWidth = startWidth + delta;
          if (newWidth < minWidth) newWidth = minWidth;
          th.style.width = newWidth + 'px';
          th.style.minWidth = newWidth + 'px';
          rows.forEach(row => {
            const cell = row.children[index];
            if (cell) {
              cell.style.width = newWidth + 'px';
              cell.style.minWidth = newWidth + 'px';
            }
          });
        }
        function onMouseUp() {
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
        }
        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
      });
    });
  }

  // Store the scenario currently being edited when the risk modal is opened.
  let riskModalTarget = null;

  // Store the current MITRE ATT&CK library loaded from a CSV.  Each
  // element has the shape { id: 'Txxxx', title: 'Technique name',
  // description: '...', mitigations: [{ id, mitigation, description }] }.
  let mitreLibrary = [];

  // On initialization, load any persisted MITRE library from localStorage.
  (function(){
    const stored = localStorage.getItem('ebiosMitreLibrary');
    if (stored) {
      try {
        const arr = JSON.parse(stored);
        if (Array.isArray(arr)) {
          mitreLibrary = arr;
        }
      } catch (e) {
        console.warn('Failed to parse stored MITRE library:', e);
      }
    }
  })();

  // Parse a MITRE CSV content into the mitreLibrary array.  The CSV is
  // expected to have a header row with at least the columns:
  // Technique ID, Technique Name, Technique Description, Mitigation ID,
  // Mitigation Name, Mitigation Description.  Additional columns will be
  // ignored.  Rows with the same technique ID will be grouped.
  function parseMitreCsv(text) {
    // Split the CSV text into logical rows while honouring quoted newlines
    function splitRows(str) {
      const rows = [];
      let cur = '';
      let inQuotes = false;
      for (let i = 0; i < str.length; i++) {
        const c = str[i];
        if (c === '"') {
          if (inQuotes && str[i + 1] === '"') { cur += '"'; i++; }
          else { inQuotes = !inQuotes; cur += c; }
          continue;
        }
        if ((c === '\n' || c === '\r') && !inQuotes) {
          if (c === '\r' && str[i + 1] === '\n') i++;
          rows.push(cur);
          cur = '';
          continue;
        }
        cur += c;
      }
      if (cur) rows.push(cur);
      return rows;
    }

    // Split a row into columns handling quoted commas
    function splitCols(row) {
      const cols = [];
      let cur = '';
      let inQuotes = false;
      for (let i = 0; i < row.length; i++) {
        const c = row[i];
        if (c === '"') {
          if (inQuotes && row[i + 1] === '"') { cur += '"'; i++; }
          else inQuotes = !inQuotes;
          continue;
        }
        if (c === ',' && !inQuotes) {
          cols.push(cur);
          cur = '';
          continue;
        }
        cur += c;
      }
      cols.push(cur);
      return cols;
    }

    const rows = splitRows(text);
    if (rows.length < 2) return [];
    const header = splitCols(rows[0]);
    const idIndex = header.findIndex(h => /technique id/i.test(h));
    const nameIndex = header.findIndex(h => /technique name/i.test(h));
    const descIndex = header.findIndex(h => /technique description/i.test(h));
    const mitIdIndex = header.findIndex(h => /mitigation id/i.test(h));
    const mitNameIndex = header.findIndex(h => /mitigation name/i.test(h));
    const mitDescIndex = header.findIndex(h => /mitigation description/i.test(h));
    const map = new Map();
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i].trim()) continue;
      const cols = splitCols(rows[i]);
      const tid = cols[idIndex] ? cols[idIndex].trim() : '';
      const tname = cols[nameIndex] ? cols[nameIndex].trim() : '';
      const tdesc = cols[descIndex] ? cols[descIndex].replace(/\s+/g, ' ').trim() : '';
      const mid = cols[mitIdIndex] ? cols[mitIdIndex].trim() : '';
      const mname = cols[mitNameIndex] ? cols[mitNameIndex].trim() : '';
      const mdesc = cols[mitDescIndex] ? cols[mitDescIndex].replace(/\s+/g, ' ').trim() : '';
      if (!tid) continue;
      if (!map.has(tid)) {
        map.set(tid, { id: tid, title: tname, description: tdesc, mitigations: [] });
      }
      const entry = map.get(tid);
      if (mid) {
        entry.mitigations.push({ id: mid, mitigation: mname, description: mdesc });
      }
    }
    return Array.from(map.values());
  }


  /*
   * ----- Import modal setup -----
   * This modal allows the user to import multiple entities (GAP requirements,
   * supports, parties, or risks) in one action.  When the user clicks an
   * import button, the modal is populated with all available items of the
   * requested type.  The user can select one or more items, then confirm to
   * create rows in the corresponding actions tables.  The modal uses the
   * existing modal classes and search bar from the risk modal.  Internal
   * state (currentImportType and selected items) is managed here.
   */
  let currentImportType = null;
  let importSelections = new Set();
  let importItems = [];
  let strategyEventTarget = null;

  function setupActionImport() {
    const confirmBtn = document.getElementById('import-confirm');
    if (confirmBtn) confirmBtn.addEventListener('click', () => applyImportSelection());
    const cancelBtn = document.getElementById('import-cancel');
    if (cancelBtn) cancelBtn.addEventListener('click', () => closeImportModal());
    const selectAllBtn = document.getElementById('import-select-all');
    if (selectAllBtn) selectAllBtn.addEventListener('click', () => importAllVisibleItems());
    const searchInput = document.getElementById('import-search');
    if (searchInput) searchInput.addEventListener('input', () => filterImportList(searchInput.value));
  }

  // Populate and show the import modal for the given type
  function openImportModal(type) {
    const modal = document.getElementById('import-modal');
    if (!modal) return;
    currentImportType = type;
    if (type !== 'strategyEvents') {
      strategyEventTarget = null;
    }
    const titleEl = document.getElementById('import-modal-title');
    const listEl = document.getElementById('import-list');
    const searchInput = document.getElementById('import-search');
    const selectedDiv = document.getElementById('import-selected');
    if (!titleEl || !listEl || !searchInput || !selectedDiv) return;
    // Set modal title based on type
    switch (type) {
      case 'gap':
        titleEl.textContent = 'Importer des exigences';
        break;
      case 'supports':
        titleEl.textContent = 'Importer des supports';
        break;
      case 'vulns':
        titleEl.textContent = 'Importer des vulnérabilités';
        break;
      case 'parties':
        titleEl.textContent = 'Importer des parties';
        break;
      case 'risques':
        titleEl.textContent = 'Importer des risques';
        break;
      case 'strategyEvents':
        titleEl.textContent = 'Sélectionner des évènements redoutés';
        break;
      default:
        titleEl.textContent = 'Importer';
    }
    // Build items list based on type
    const analysis = analyses[currentIndex];
    let items = [];
    if (!analysis || !analysis.data) items = [];
    else if (type === 'gap') {
      (analysis.data.gap || []).forEach(req => {
        // filter to non fully applied as in renderGapActions
        let app = (req.application || '').toLowerCase();
        try { app = app.normalize('NFD').replace(/\p{Diacritic}/gu, ''); } catch (e) {}
        const prefix = app.replace(/\s+/g, '');
        if (prefix.startsWith('applique')) return;
        const id = req.id || (req.id = uid());
        const labelParts = [];
        if (req.domaine) labelParts.push(req.domaine);
        if (req.titre) labelParts.push(req.titre);
        const label = labelParts.join(' – ');
        const desc = req.description || '';
        const application = req.application || '';
        items.push({ id, label, desc, extra: application });
      });
    } else if (type === 'supports') {
      const supportsMap = new Map();
      (analysis.data.missions || []).forEach(mis => {
        (mis.supports || []).forEach(s => {
          const name = s.name || s.denom;
          if (!name) return;
          if (!supportsMap.has(name)) supportsMap.set(name, { name, desc: s.description || '', resp: s.responsable || '' });
        });
      });
      supportsMap.forEach((obj) => {
        items.push({ id: obj.name, label: obj.name, desc: obj.desc || '', extra: obj.resp || '' });
      });
    } else if (type === 'vulns') {
      (analysis.data.supportsQualif || []).forEach(s => {
        (s.vulnerabilities || []).forEach(v => {
          const id = `${s.name}||${v.name}`;
          items.push({ id, label: v.name || 'Vulnérabilité', desc: v.description || '', extra: s.name || '' });
        });
      });
    } else if (type === 'parties') {
      (analysis.data.ppc || []).forEach(pp => {
        const id = pp.id || (pp.id = uid());
        const label = pp.nom || pp.name || 'Partie';
        const desc = pp.categorie || pp.categorie === '' ? (pp.categorie) : '';
        // maybe show dependance / penetration but keep simple
        items.push({ id, label, desc: desc || '', extra: '' });
      });
    } else if (type === 'risques') {
      // Gather risks from the global list and from operational scenarios
      const riskMap = new Map();
      (analysis.data.risques || []).forEach(riskObj => {
        const label = riskObj.libelle || riskObj.titre || riskObj.indice || riskObj.id || '';
        if (!label) return;
        const id = riskObj.id || label;
        const desc = riskObj.description || '';
        riskMap.set(id, { id, label, desc });
      });
      (analysis.data.so || []).forEach(so => {
        (so.risks || []).forEach(rk => {
          const label = rk.name || rk.id;
          if (!label) return;
          const id = rk.id || rk.name;
          const desc = rk.description || '';
          if (!riskMap.has(id)) riskMap.set(id, { id, label, desc });
        });
      });
      riskMap.forEach(obj => {
        items.push({ id: obj.id, label: obj.label, desc: obj.desc || '', extra: '' });
      });
    } else if (type === 'strategyEvents') {
      const missionMap = new Map();
      (analysis.data.missions || []).forEach(m => {
        if (m && m.id) {
          const label = m.nom || m.name || m.titre || m.description || '';
          missionMap.set(m.id, label);
        }
      });
      (analysis.data.events || []).forEach(ev => {
        if (!ev || !ev.id) return;
        const label = ev.evenement || ev.ref || 'Évènement';
        const desc = ev.impactDescription || '';
        const parts = [];
        const missionLabel = missionMap.get(ev.missionId);
        if (missionLabel) parts.push(missionLabel);
        const impact = parseInt(ev.impact, 10);
        if (!isNaN(impact)) parts.push(`Gravité G${impact}`);
        items.push({ id: ev.id, label, desc, extra: parts.join(' – ') });
      });
    }
    // Render list
    listEl.innerHTML = '';
    if (type === 'strategyEvents' && strategyEventTarget && Array.isArray(strategyEventTarget.eventIds)) {
      importSelections = new Set(strategyEventTarget.eventIds);
    } else {
      importSelections = new Set();
    }
    importItems = items;
    items.forEach(item => {
      const div = document.createElement('div');
      div.className = 'import-item';
      div.dataset.id = item.id;
      // Compose HTML with bold label and small description
      const title = document.createElement('div');
      title.innerHTML = `<strong>${item.label}</strong>`;
      const descEl = document.createElement('div');
      descEl.className = 'import-desc';
      const parts = [];
      if (item.desc) parts.push(item.desc);
      if (item.extra) parts.push(`<i>${item.extra}</i>`);
      descEl.innerHTML = parts.join(' – ');
      div.appendChild(title);
      div.appendChild(descEl);
      if (importSelections.has(item.id)) {
        div.classList.add('selected');
      }
      div.addEventListener('click', () => {
        const id = div.dataset.id;
        if (importSelections.has(id)) {
          importSelections.delete(id);
          div.classList.remove('selected');
        } else {
          importSelections.add(id);
          div.classList.add('selected');
        }
        renderImportSelected();
      });
      listEl.appendChild(div);
    });
    renderImportSelected();
    // Show modal
    modal.style.display = 'flex';
    searchInput.value = '';
  }

  // Apply filter to import list based on search term
  function filterImportList(term) {
    term = term.toLowerCase().trim();
    const listEl = document.getElementById('import-list');
    if (!listEl) return;
    Array.from(listEl.children).forEach(div => {
      const text = div.textContent.toLowerCase();
      if (!term || text.includes(term)) {
        div.style.display = '';
      } else {
        div.style.display = 'none';
      }
    });
  }

  function importAllVisibleItems() {
    const listEl = document.getElementById('import-list');
    if (!listEl) return;
    let hasVisible = false;
    Array.from(listEl.children).forEach(div => {
      if (div.style.display === 'none') return;
      const id = div.dataset.id;
      if (!id) return;
      hasVisible = true;
      importSelections.add(id);
      div.classList.add('selected');
    });
    if (!hasVisible || !importSelections.size) return;
    renderImportSelected();
    applyImportSelection();
  }

  // Close modal without applying changes
  function closeImportModal() {
    const modal = document.getElementById('import-modal');
    if (modal) modal.style.display = 'none';
    importSelections = new Set();
    currentImportType = null;
    strategyEventTarget = null;
    renderImportSelected();
  }

  // Apply import selection to the current analysis
  function applyImportSelection() {
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data || !currentImportType) {
      closeImportModal();
      return;
    }
    if (currentImportType === 'gap') {
      // For each selected requirement, ensure an entry exists in actionsGap
      if (!Array.isArray(analysis.data.actionsGap)) analysis.data.actionsGap = [];
      importSelections.forEach(id => {
        if (!analysis.data.actionsGap.some(entry => entry.sourceId === id)) {
          analysis.data.actionsGap.push({ sourceId: id, actions: [] });
        }
      });
      saveAnalyses();
      renderGapActions();
    } else if (currentImportType === 'supports') {
      if (!Array.isArray(analysis.data.actionsSupports)) analysis.data.actionsSupports = [];
      importSelections.forEach(name => {
        if (!analysis.data.actionsSupports.some(row => row.supportName === name)) {
          analysis.data.actionsSupports.push({ supportName: name, vulnName: '', initialLevel: '', residualLevel: '', actions: [] });
        }
      });
      saveAnalyses();
      renderSupportActions();
    } else if (currentImportType === 'vulns') {
      if (!Array.isArray(analysis.data.actionsSupports)) analysis.data.actionsSupports = [];
      importSelections.forEach(id => {
        const parts = id.split('||');
        const supportName = parts[0] || '';
        const vulnName = parts[1] || '';
        const support = (analysis.data.supportsQualif || []).find(s => s.name === supportName);
        const vuln = support ? (support.vulnerabilities || []).find(v => v.name === vulnName) : null;
        const lvl = vuln ? (vuln.level || '') : '';
        if (!analysis.data.actionsSupports.some(row => row.supportName === supportName && row.vulnName === vulnName)) {
          analysis.data.actionsSupports.push({ supportName, vulnName, initialLevel: lvl, residualLevel: lvl, actions: [] });
        }
      });
      saveAnalyses();
      renderSupportActions();
    } else if (currentImportType === 'parties') {
      if (!Array.isArray(analysis.data.actionsParties)) analysis.data.actionsParties = [];
      importSelections.forEach(id => {
        if (!analysis.data.actionsParties.some(row => row.ppId === id)) {
          analysis.data.actionsParties.push({ ppId: id, actions: [] });
        }
      });
      saveAnalyses();
      renderPartiesActions();
    } else if (currentImportType === 'risques') {
      if (!Array.isArray(analysis.data.actionsRisques)) analysis.data.actionsRisques = [];
      // Build risk map again to get current levels
      const riskMap = new Map();
      (analysis.data.risques || []).forEach(riskObj => {
        const label = riskObj.libelle || riskObj.titre || riskObj.indice || riskObj.id || '';
        if (!label) return;
        const id = riskObj.id || label;
        const current = riskMap.get(id) || { id, name: label, vraisemblance: parseInt(riskObj.vraisemblance, 10) || 1, gravite: parseInt(riskObj.gravite, 10) || 1 };
        current.vraisemblance = Math.max(current.vraisemblance, parseInt(riskObj.vraisemblance, 10) || 1);
        current.gravite = Math.max(current.gravite, parseInt(riskObj.gravite, 10) || 1);
        riskMap.set(id, current);
      });
      (analysis.data.so || []).forEach(so => {
        (so.risks || []).forEach(rk => {
          const label = rk.name || rk.id;
          if (!label) return;
          const id = rk.id || rk.name;
          const current = riskMap.get(id) || { id, name: label, vraisemblance: parseInt(rk.vraisemblance, 10) || 1, gravite: parseInt(rk.gravite, 10) || 1 };
          current.vraisemblance = Math.max(current.vraisemblance, parseInt(rk.vraisemblance, 10) || 1);
          current.gravite = Math.max(current.gravite, parseInt(rk.gravite, 10) || 1);
          riskMap.set(id, current);
        });
      });
      importSelections.forEach(id => {
        const risk = riskMap.get(id);
        if (!risk) return;
        if (!analysis.data.actionsRisques.some(row => row.riskId === id)) {
          analysis.data.actionsRisques.push({ riskId: id, riskName: risk.name, residualV: 1, residualG: 1, actions: [] });
        }
      });
      saveAnalyses();
      renderRisquesActions();
    } else if (currentImportType === 'strategyEvents') {
      if (strategyEventTarget) {
        const events = analysis.data.events || [];
        const ordered = [];
        events.forEach(ev => {
          if (ev && importSelections.has(ev.id)) {
            ordered.push(ev.id);
          }
        });
        importSelections.forEach(id => {
          if (!ordered.includes(id)) ordered.push(id);
        });
        strategyEventTarget.eventIds = ordered;
        saveAnalyses();
        renderStrategies();
      }
    }
    // Update plan actions after any import
    renderPlanActions();
    closeImportModal();
  }

  function renderImportSelected() {
    const container = document.getElementById('import-selected');
    if (!container) return;
    container.innerHTML = '';
    importSelections.forEach(id => {
      const item = importItems.find(i => i.id === id);
      if (item) {
        const div = document.createElement('div');
        div.textContent = item.label;
        container.appendChild(div);
      }
    });
  }

  // Open the risk modal and display MITRE techniques from the CSV or allow manual entry
  function openRiskModal(targetItem) {
    riskModalTarget = targetItem;
    const modal = document.getElementById('risk-modal');
    if (!modal) return;
    const searchInput = document.getElementById('risk-search');
    const tableBody = document.getElementById('risk-table-body');
    const selectedDiv = document.getElementById('risk-selected');
    const closeBtn = document.getElementById('risk-close-btn');
    const applyBtn = document.getElementById('risk-select-apply');
    const manualIdInput = document.getElementById('manual-risk-id');
    const manualNameInput = document.getElementById('manual-risk-name');
    const manualAddBtn = document.getElementById('manual-risk-add');
    let selectedIds = new Set();

    function renderTable(filter) {
      const term = (filter || '').toLowerCase();
      tableBody.innerHTML = '';
      const filtered = (mitreLibrary || []).filter(obj =>
        (obj.id + ' ' + obj.title + ' ' + obj.description).toLowerCase().includes(term)
      );
      filtered.forEach(obj => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${obj.id}</td><td>${obj.title}</td><td>${obj.description}</td>`;
        if (selectedIds.has(obj.id)) tr.classList.add('selected');
        tr.addEventListener('click', () => {
          if (selectedIds.has(obj.id)) {
            selectedIds.delete(obj.id);
          } else {
            selectedIds.add(obj.id);
          }
          renderTable(searchInput.value);
          updateSelected();
        });
        tableBody.appendChild(tr);
      });
      addDataTableResizers('risk-table');
    }

    function updateSelected() {
      selectedDiv.innerHTML = '';
      selectedIds.forEach(id => {
        const item = mitreLibrary.find(t => t.id === id);
        if (item) {
          const div = document.createElement('div');
          div.textContent = `${item.id} – ${item.title}`;
          selectedDiv.appendChild(div);
        }
      });
    }

    searchInput.value = '';
    searchInput.oninput = e => {
      renderTable(e.target.value);
    };

    if (manualIdInput) manualIdInput.value = '';
    if (manualNameInput) manualNameInput.value = '';

    closeBtn.onclick = () => {
      closeRiskModal();
    };

    if (applyBtn) {
      applyBtn.onclick = () => {
        if (!riskModalTarget) return;
        selectedIds.forEach(id => {
          const item = mitreLibrary.find(t => t.id === id);
          if (item) {
            const name = item.id + ' ' + item.title;
            riskModalTarget.risks.push({ name: name, vraisemblance: 1, gravite: 1 });
          }
        });
        saveAnalyses();
        renderSO();
        closeRiskModal();
      };
    }

    if (manualAddBtn) {
      manualAddBtn.onclick = () => {
        if (!riskModalTarget) return;
        const id = manualIdInput ? manualIdInput.value.trim() : '';
        const name = manualNameInput ? manualNameInput.value.trim() : '';
        if (!name) return;
        const label = id ? id + ' ' + name : name;
        riskModalTarget.risks.push({ name: label, vraisemblance: 1, gravite: 1 });
        saveAnalyses();
        renderSO();
        if (manualIdInput) manualIdInput.value = '';
        if (manualNameInput) manualNameInput.value = '';
      };
    }

    function ensureMitreLoaded(callback) {
      if (mitreLibrary && mitreLibrary.length > 0) { callback(); return; }
      try {
        const stored = localStorage.getItem('ebiosMitreLibrary');
        if (stored) {
          const parsed = JSON.parse(stored);
          if (Array.isArray(parsed) && parsed.length > 0) {
            mitreLibrary = parsed;
            callback();
            return;
          }
        }
      } catch (e) {}
      fetch('mitre_attack.csv').then(res => {
        if (!res.ok) throw new Error('Cannot load MITRE CSV');
        return res.text();
      }).then(text => {
        const parsed = parseMitreCsv(text);
        if (parsed && parsed.length > 0) {
          mitreLibrary = parsed;
          localStorage.setItem('ebiosMitreLibrary', JSON.stringify(mitreLibrary));
        }
      }).catch(err => {
        console.warn('Failed to load MITRE CSV', err);
        alert('Impossible de charger le fichier MITRE. Utilisez un serveur local ou ajoutez un risque manuel.');
      }).finally(() => {
        callback();
      });
    }

    ensureMitreLoaded(() => {
      modal.style.display = 'flex';
      renderTable(searchInput.value);
      updateSelected();
    });
  }

  // Close the risk modal and reset state
  function closeRiskModal() {
    const modal = document.getElementById('risk-modal');
    if (!modal) return;
    modal.style.display = 'none';
    // reset target
    riskModalTarget = null;
  }

  // ----- GAP analysis: compliance chart drawing
  function updateGapChart() {
    const container = document.getElementById('gap-overview-chart');
    if (!container) return;
    const analysis = analyses[currentIndex];
    // Clear existing content
    container.innerHTML = '';
    if (!analysis || !analysis.data) {
      return;
    }
    const gap = analysis.data.gap || [];
    // Count statuses
    const counts = {
      applique: 0,
      partiel: 0,
      non: 0,
      nonApp: 0
    };
    gap.forEach(req => {
      const val = (req.application || '').toLowerCase().trim();
      if (val === 'appliqué' || val === 'applique') counts.applique += 1;
      else if (val === 'partiellement appliqué' || val === 'partiellement applique') counts.partiel += 1;
      else if (val === 'non appliqué' || val === 'non applique') counts.non += 1;
      else if (val === 'non applicable' || val === 'non applicable') counts.nonApp += 1;
    });
    const total = counts.applique + counts.partiel + counts.non + counts.nonApp;
    if (total === 0) {
      const msg = document.createElement('div');
      msg.style.color = '#657A93';
      msg.style.fontSize = '1rem';
      msg.textContent = 'Aucune exigence';
      container.appendChild(msg);
      return;
    }
    // Define categories with colours
    const categories = [
      { key:'applique', label:'Appliqué', color:'#2a9d8f' },
      { key:'partiel', label:'Partiellement appliqué', color:'#e9c46a' },
      { key:'non', label:'Non appliqué', color:'#e63946' },
      { key:'nonApp', label:'Non applicable', color:'#9aa0a6' }
    ];
    // Build gradient stops for the conic gradient
    let offset = 0;
    const stops = [];
    categories.forEach(cat => {
      const value = counts[cat.key];
      const angle = (value / total) * 360;
      const start = offset;
      const end = offset + angle;
      stops.push(`${cat.color} ${start}deg ${end}deg`);
      offset = end;
    });
    const gradientStr = stops.join(', ');
    // Create donut element
    const donut = document.createElement('div');
    donut.className = 'donut-chart';
    // Determine size of the donut based on available space.  Some
    // browsers may report zero height for an absolutely positioned
    // container, so fall back to the width or a default value.
    let cw = container.clientWidth || container.offsetWidth;
    let ch = container.clientHeight || container.offsetHeight;
    // If height is zero (e.g. flex child with no intrinsic height), use
    // a portion of the width.  Provide a sensible default when both
    // values are zero.
    if (!ch || ch < 20) ch = cw;
    if (!cw || cw < 20) cw = 400;
    const base = Math.min(cw, ch);
    // Set donut diameter to 60% of the smallest dimension but clamp
    // within reasonable bounds to ensure readability.
    const size = Math.max(120, Math.min(300, base * 0.6));
    const hole = size * 0.5;
    donut.style.width = size + 'px';
    donut.style.height = size + 'px';
    donut.style.background = `conic-gradient(${gradientStr})`;
    // Create inner hole via pseudo-element by adjusting after height
    const inner = document.createElement('div');
    inner.style.position = 'absolute';
    inner.style.top = '50%';
    inner.style.left = '50%';
    inner.style.transform = 'translate(-50%, -50%)';
    inner.style.width = hole + 'px';
    inner.style.height = hole + 'px';
    inner.style.borderRadius = '50%';
    inner.style.backgroundColor = 'var(--bg-panel)';
    donut.appendChild(inner);
    container.appendChild(donut);
    // Build legend
    const legend = document.createElement('div');
    legend.className = 'legend';
    categories.forEach(cat => {
      const value = counts[cat.key];
      const perc = Math.round((value / total) * 100);
      const item = document.createElement('div');
      item.className = 'legend-item';
      const colorBox = document.createElement('div');
      colorBox.className = 'legend-color';
      colorBox.style.backgroundColor = cat.color;
      const label = document.createElement('span');
      label.textContent = `${cat.label}: ${value} (${perc}%)`;
      item.appendChild(colorBox);
      item.appendChild(label);
      legend.appendChild(item);
    });
    container.appendChild(legend);
  }

  // ----- Atelier 1: Graph -----
  // Build nodes and links from missions/supports/events and forward them to
  // the rendering helper defined in atelier1_graph.js.
  let atelier1Chart = null; // placeholder for backward compatibility
  function updateAtelier1Graph() {
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) {
      if (typeof renderAtelier1Graph === 'function') renderAtelier1Graph([], []);
      return;
    }

    const nodes = [];
    const links = [];
    const supportMap = new Map();
    const eventMap = new Map();

    (analysis.data.missions || []).forEach(mission => {
      if (!mission.id) mission.id = uid();
      nodes.push({ id: mission.id, type: 'valeur', label: mission.denom || 'Valeur' });

      (mission.supports || []).forEach(support => {
        if (!support.id) support.id = uid();
        if (!supportMap.has(support.id)) {
          supportMap.set(support.id, true);
          nodes.push({ id: support.id, type: 'support', label: support.name || 'Support' });
        }
        links.push({ id: `${mission.id}-${support.id}`, source: mission.id, target: support.id, weight: 1 });
      });

      const events = (analysis.data.events || []).filter(ev => ev.missionId === mission.id);
      events.forEach(ev => {
        if (!ev.id) ev.id = uid();
        if (!eventMap.has(ev.id)) {
          eventMap.set(ev.id, true);
          nodes.push({
            id: ev.id,
            type: 'event',
            label: ev.evenement || 'Évènement',
            severity: parseInt(ev.impact, 10) || 0
          });
        }
        (mission.supports || []).forEach(support => {
          links.push({ id: `${support.id}-${ev.id}`, source: support.id, target: ev.id, weight: 1 });
        });
      });
    });

    if (typeof renderAtelier1Graph === 'function') {
      renderAtelier1Graph(nodes, links);
    }
  }

  // ----- Atelier 1: Missions
  function renderMissions() {
    const listEl = document.getElementById('missions-list');
    listEl.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis.data) analysis.data = {};
    if (!analysis.data.missions) analysis.data.missions = [];
    analysis.data.missions.forEach((item, idx) => {
      // Ensure each mission has a unique id
      if (!item.id) item.id = uid();
      // Normalise supports: convert string to array if needed
      if (!Array.isArray(item.supports)) {
        if (typeof item.supports === 'string' && item.supports.trim() !== '') {
          item.supports = item.supports.split(',').map(s => s.trim()).filter(Boolean);
        } else {
          item.supports = [];
        }
      }
      const el = document.createElement('div');
      el.className = 'item';
      // Delete button
      const delBtn = document.createElement('button');
      delBtn.className = 'delete-item';
      delBtn.innerHTML = '×';
      delBtn.title = 'Supprimer cette mission';
      delBtn.addEventListener('click', () => {
        if (!confirm('Supprimer cette mission ?')) return;
        analysis.data.missions.splice(idx, 1);
        saveAnalyses();
        renderMissions();
        updateAtelier1Chart();
      });
      el.appendChild(delBtn);
      // Fields
      el.appendChild(createInput('Dénomination', 'text', item.denom, (v) => {
        item.denom = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Nature', 'text', item.nature, (v) => {
        item.nature = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Description', 'textarea', item.description, (v) => {
        item.description = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Responsable', 'text', item.responsable, (v) => {
        item.responsable = v;
        saveAnalyses();
      }));
      // Supports list
      const supportsContainer = document.createElement('div');
      const supportsLabel = document.createElement('label');
      supportsLabel.textContent = 'Biens supports';
      supportsContainer.appendChild(supportsLabel);
      const supportsList = document.createElement('div');
      supportsList.className = 'supports-list';
      item.supports.forEach((support, sIdx) => {
        const sItem = document.createElement('div');
        sItem.className = 'supports-item';
        const input = document.createElement('input');
        input.type = 'text';
        input.value = support;
        input.addEventListener('input', (e) => {
          item.supports[sIdx] = e.target.value;
          saveAnalyses();
        });
        const rmBtn = document.createElement('button');
        rmBtn.innerHTML = '×';
        rmBtn.title = 'Supprimer ce bien support';
        rmBtn.addEventListener('click', () => {
          item.supports.splice(sIdx, 1);
          saveAnalyses();
          renderMissions();
        });
        sItem.appendChild(input);
        sItem.appendChild(rmBtn);
        supportsList.appendChild(sItem);
      });
      supportsContainer.appendChild(supportsList);
      const addSupportBtn = document.createElement('button');
      addSupportBtn.className = 'add-support-btn';
      addSupportBtn.textContent = '+ Bien support';
      addSupportBtn.addEventListener('click', () => {
        item.supports.push('');
        saveAnalyses();
        renderMissions();
      });
      supportsContainer.appendChild(addSupportBtn);
      el.appendChild(supportsContainer);
      listEl.appendChild(el);
    });
  }

  function renderEvents() {
    const listEl = document.getElementById('er-list');
    listEl.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis.data) analysis.data = {};
    if (!analysis.data.events) analysis.data.events = [];
    analysis.data.events.forEach((item, idx) => {
      // Ensure each event has a unique id
      if (!item.id) item.id = uid();
      const el = document.createElement('div');
      el.className = 'item';
      const delBtn = document.createElement('button');
      delBtn.className = 'delete-item';
      delBtn.innerHTML = '×';
      delBtn.title = 'Supprimer cet évènement';
      delBtn.addEventListener('click', () => {
        if (!confirm('Supprimer cet évènement ?')) return;
        analysis.data.events.splice(idx, 1);
        saveAnalyses();
        renderEvents();
        updateAtelier1Chart();
      });
      el.appendChild(delBtn);
      // Mission selector
      const missionOptions = (analysis.data.missions || []).map((m, i) => ({ value: m.id, label: m.denom || `Mission ${i + 1}` }));
      el.appendChild(createSelect('Mission', missionOptions, item.missionId || '', (v) => {
        item.missionId = v;
        saveAnalyses();
      }));
      // Fields
      el.appendChild(createInput('Référence', 'text', item.ref, (v) => {
        item.ref = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Évènement redouté', 'textarea', item.evenement, (v) => {
        item.evenement = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Impact', 'text', item.impact, (v) => {
        item.impact = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Gravité', 'text', item.gravite, (v) => {
        item.gravite = v;
        saveAnalyses();
        updateAtelier1Chart();
      }));
      listEl.appendChild(el);
    });
  }

  // ----- Atelier 2: SROV
  function renderSROV() {
    // Render SROV entries in a table with columns for source, objectif, motivation,
    // ressources, pertinence, priorité, retenue, justification and actions.
    const tbody = document.getElementById('srov-body');
    if (!tbody) return;
    tbody.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis.data) analysis.data = {};
    if (!analysis.data.srov) analysis.data.srov = [];
    const levelColor = (lvl) => {
      switch (parseInt(lvl, 10)) {
        case 1: return '#2a9d8f'; // green
        case 2: return '#e9c46a'; // yellow
        case 3: return '#f4a261'; // orange
        case 4: return '#e63946'; // red
        default: return '#9aa0a6';
      }
    };
    const pertinenceBucket = (p) => {
      if (p >= 13) return 4;
      if (p >= 9) return 3;
      if (p >= 5) return 2;
      return 1;
    };
    analysis.data.srov.forEach((item, idx) => {
      if (!item.id) item.id = uid();
      // Ensure numeric fields exist and default to 1
      item.motivation = parseInt(item.motivation, 10) || 1;
      item.ressources = parseInt(item.ressources, 10) || 1;
      item.priorite = parseInt(item.priorite, 10) || 1;
      // Default retenue to true unless explicitly false
      if (typeof item.retenue !== 'boolean') item.retenue = true;
      const tr = document.createElement('tr');
      // Source
      let td = document.createElement('td');
      const srcInput = document.createElement('input');
      srcInput.type = 'text';
      srcInput.value = item.source || '';
      srcInput.oninput = (e) => {
        item.source = e.target.value;
        saveAnalyses();
        updateAtelier2Chart();
      };
      td.appendChild(srcInput);
      tr.appendChild(td);
      // Objectif
      td = document.createElement('td');
      const objInput = document.createElement('input');
      objInput.type = 'text';
      objInput.value = item.objectif || '';
      objInput.oninput = (e) => {
        item.objectif = e.target.value;
        saveAnalyses();
        updateAtelier2Chart();
      };
      td.appendChild(objInput);
      tr.appendChild(td);
      // Motivation
      td = document.createElement('td');
      const motSelect = document.createElement('select');
      [1, 2, 3, 4].forEach(val => {
        const opt = document.createElement('option');
        opt.value = val;
        // Shorter labels for motivation to avoid truncation.  The number
        // conveys the level and the descriptor gives a quick sense of
        // intensity.
        let label;
        switch (val) {
          case 4:
            label = '4 – Fortement';
            break;
          case 3:
            label = '3 – Assez';
            break;
          case 2:
            label = '2 – Peu';
            break;
          default:
            label = '1 – Très peu';
        }
        opt.textContent = label;
        if (val === item.motivation) opt.selected = true;
        motSelect.appendChild(opt);
      });
      // Apply initial colour
      motSelect.style.backgroundColor = levelColor(item.motivation);
      motSelect.onchange = (e) => {
        item.motivation = parseInt(e.target.value, 10);
        // Update colour and pertinence cell
        e.target.style.backgroundColor = levelColor(item.motivation);
        // Update pertinence cell below
        updatePertinenceCell(tr, item);
        saveAnalyses();
        updateAtelier2Chart();
      };
      td.appendChild(motSelect);
      tr.appendChild(td);
      // Ressources
      td = document.createElement('td');
      const resSelect = document.createElement('select');
      [1, 2, 3, 4].forEach(val => {
        const opt = document.createElement('option');
        opt.value = val;
        // Shorter labels so the select fits within the table without truncation.
        // Each level indicates the available resources from limited to unlimited.
        let label;
        switch (val) {
          case 4:
            label = '4 – Illimitées';
            break;
          case 3:
            label = '3 – Importantes';
            break;
          case 2:
            label = '2 – Significatives';
            break;
          default:
            label = '1 – Limitées';
        }
        opt.textContent = label;
        if (val === item.ressources) opt.selected = true;
        resSelect.appendChild(opt);
      });
      resSelect.style.backgroundColor = levelColor(item.ressources);
      resSelect.onchange = (e) => {
        item.ressources = parseInt(e.target.value, 10);
        e.target.style.backgroundColor = levelColor(item.ressources);
        updatePertinenceCell(tr, item);
        saveAnalyses();
        updateAtelier2Chart();
      };
      td.appendChild(resSelect);
      tr.appendChild(td);
      // Pertinence (computed)
      td = document.createElement('td');
      td.className = 'pertinence-cell';
      const setPertinence = () => {
        const p = (item.motivation || 1) * (item.ressources || 1);
        const bucket = pertinenceBucket(p);
        td.textContent = `${p} (niv ${bucket})`;
        td.style.backgroundColor = levelColor(bucket);
      };
      setPertinence();
      // store helper to update later
      td.dataset.update = setPertinence;
      tr.appendChild(td);
      // Priorité (1 à 4) : utilise des libellés explicites pour que
      // l’utilisateur comprenne l’échelle.  Le niveau 1 est une
      // priorité faible et le niveau 4 correspond à une priorité
      // absolue.  Les couleurs suivent la fonction levelColor().
      td = document.createElement('td');
      const priSelect = document.createElement('select');
      [1, 2, 3, 4].forEach(val => {
        const opt = document.createElement('option');
        opt.value = val;
        let label;
        switch (val) {
          case 4:
            label = '4 – Priorité absolue';
            break;
          case 3:
            label = '3 – Priorité élevée';
            break;
          case 2:
            label = '2 – Priorité modérée';
            break;
          default:
            label = '1 – Priorité faible';
        }
        opt.textContent = label;
        if (val === item.priorite) opt.selected = true;
        priSelect.appendChild(opt);
      });
      // Colour by priority
      priSelect.style.backgroundColor = levelColor(item.priorite);
      priSelect.onchange = (e) => {
        item.priorite = parseInt(e.target.value, 10);
        e.target.style.backgroundColor = levelColor(item.priorite);
        saveAnalyses();
        updateAtelier2Chart();
      };
      td.appendChild(priSelect);
      tr.appendChild(td);
      // Retenue (checkbox)
      td = document.createElement('td');
      const check = document.createElement('input');
      check.type = 'checkbox';
      check.checked = item.retenue;
      check.onchange = (e) => {
        item.retenue = e.target.checked;
        saveAnalyses();
        updateAtelier2Chart();
      };
      td.appendChild(check);
      tr.appendChild(td);
      // Justification
      td = document.createElement('td');
      const justInput = document.createElement('textarea');
      justInput.rows = 2;
      justInput.value = item.justification || '';
      justInput.oninput = (e) => {
        item.justification = e.target.value;
        saveAnalyses();
        updateAtelier2Chart();
      };
      td.appendChild(justInput);
      tr.appendChild(td);
      // Actions
      td = document.createElement('td');
      const delBtn = document.createElement('button');
      delBtn.className = 'delete-item';
      delBtn.textContent = '×';
      delBtn.title = 'Supprimer ce couple';
      delBtn.addEventListener('click', () => {
        if (!confirm('Supprimer ce couple ?')) return;
        analysis.data.srov.splice(idx, 1);
        saveAnalyses();
        renderSROV();
        updateAtelier2Chart();
      });
      td.appendChild(delBtn);
      tr.appendChild(td);
      tbody.appendChild(tr);
    });
    // After rendering, add resizer handles to the SROV table
    addSrovTableResizers();
    // Helper to update pertinence cell when motivation or resources change
    function updatePertinenceCell(row, entry) {
      const pCell = row.querySelector('.pertinence-cell');
      if (!pCell) return;
      const p = (entry.motivation || 1) * (entry.ressources || 1);
      const bucket = pertinenceBucket(p);
      pCell.textContent = `${p} (niv ${bucket})`;
      pCell.style.backgroundColor = levelColor(bucket);
    }
  }

  // ----- Atelier 3: Parties prenantes
  function renderPP() {
    const listEl = document.getElementById('pp-list');
    if (!listEl) return;
    listEl.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis.data) analysis.data = {};
    if (!analysis.data.pp) analysis.data.pp = [];
    analysis.data.pp.forEach((item, idx) => {
      if (!item.id) item.id = uid();
      const el = document.createElement('div');
      el.className = 'item';
      const delBtn = document.createElement('button');
      delBtn.className = 'delete-item';
      delBtn.innerHTML = '×';
      delBtn.title = 'Supprimer cette partie prenante';
      delBtn.addEventListener('click', () => {
        if (!confirm('Supprimer cette partie prenante ?')) return;
        analysis.data.pp.splice(idx, 1);
        saveAnalyses();
        renderPP();
        updateAtelier3Chart();
      });
      el.appendChild(delBtn);
      el.appendChild(createInput('Catégorie', 'text', item.categorie, (v) => {
        item.categorie = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Nom', 'text', item.nom, (v) => {
        item.nom = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Description', 'textarea', item.description, (v) => {
        item.description = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Niveau SSI (0-10)', 'number', item.niveauSSI, (v) => {
        item.niveauSSI = parseFloat(v) || 0;
        saveAnalyses();
        updateAtelier3Chart();
      }));
      el.appendChild(createInput('Indice de menace (0-10)', 'number', item.indiceMenace, (v) => {
        item.indiceMenace = parseFloat(v) || 0;
        saveAnalyses();
        updateAtelier3Chart();
      }));
      listEl.appendChild(el);
    });
  }

  // ----- Atelier 3: Cartographie des parties prenantes
  function renderPPCarto() {
    const tbody = document.getElementById('ppc-body');
    if (!tbody) return;
    tbody.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis.data) analysis.data = {};
    // Array to store cartography entries.  Initialize if missing.
    if (!Array.isArray(analysis.data.ppc)) analysis.data.ppc = [];
    const ppc = analysis.data.ppc;
    // Helper to build options for supports and values
    const supportOptions = [];
    const missionOptions = [];
    // Collect supports across all missions
    (analysis.data.missions || []).forEach(mission => {
      (mission.supports || []).forEach(support => {
        const name = support.name || '';
        if (name && !supportOptions.some(o => o.value === name)) {
          supportOptions.push({ value: name, label: name });
        }
      });
      // Mission options for value selection
      const mName = mission.denom || '';
      if (mName && !missionOptions.some(o => o.value === mission.id)) {
        missionOptions.push({ value: mission.id, label: mName });
      }
    });
    // Colour helpers
    const levelColor = (lvl) => {
      switch (parseInt(lvl, 10)) {
        case 1: return '#2a9d8f'; // green
        case 2: return '#e9c46a'; // yellow
        case 3: return '#f4a261'; // orange
        case 4: return '#e63946'; // red
        default: return '#9aa0a6';
      }
    };
    // Generate table rows
    ppc.forEach((item, idx) => {
      if (!item.id) item.id = uid();
      item.nom = item.nom || '';
      item.categorie = item.categorie || 'prestataire';
      if (!Array.isArray(item.supportIds)) item.supportIds = [];
      if (!Array.isArray(item.valueIds)) item.valueIds = [];
      item.dependance = parseInt(item.dependance, 10) || 1;
      item.penetration = parseInt(item.penetration, 10) || 1;
      item.maturite = parseInt(item.maturite, 10) || 1;
      item.confiance = parseInt(item.confiance, 10) || 1;
      const tr = document.createElement('tr');
      // Nom
      let td = document.createElement('td');
      const nomInput = document.createElement('input');
      nomInput.type = 'text';
      nomInput.value = item.nom;
      nomInput.oninput = (e) => {
        item.nom = e.target.value;
        saveAnalyses();
      };
      td.appendChild(nomInput);
      tr.appendChild(td);
      // Catégorie
      td = document.createElement('td');
      const catSelect = document.createElement('select');
      const catOpts = [
        { value: 'prestataire', label: 'Prestataire' },
        { value: 'partenaire', label: 'Partenaire' },
        { value: 'beneficiaire', label: 'Bénéficiaire' },
        { value: 'interne', label: 'Interne/Autre' }
      ];
      catOpts.forEach(opt => {
        const o = document.createElement('option');
        o.value = opt.value;
        o.textContent = opt.label;
        if (opt.value === item.categorie) o.selected = true;
        catSelect.appendChild(o);
      });
      catSelect.onchange = (e) => {
        item.categorie = e.target.value;
        updateDerivedCells(tr, item);
        saveAnalyses();
        updateAtelier3Chart();
      };
      td.appendChild(catSelect);
      tr.appendChild(td);
      // Biens supports: display selected supports as tags with remove buttons
      td = document.createElement('td');
      const supCell = document.createElement('div');
      supCell.className = 'assoc-cell';
      // Display each selected support
      item.supportIds.forEach((sid) => {
        const opt = supportOptions.find(o => o.value === sid);
        const label = opt ? opt.label : sid;
        const tag = document.createElement('span');
        tag.className = 'assoc-item';
        tag.textContent = label;
        const rmBtn = document.createElement('button');
        rmBtn.textContent = '×';
        rmBtn.title = 'Retirer ce support';
        rmBtn.addEventListener('click', () => {
          const pos = item.supportIds.indexOf(sid);
          if (pos >= 0) item.supportIds.splice(pos, 1);
          saveAnalyses();
          renderPPCarto();
          updateAtelier3Chart();
        });
        tag.appendChild(rmBtn);
        supCell.appendChild(tag);
      });
      // Button to add a new support from the available list
      const addSup = document.createElement('button');
      addSup.className = 'add-assoc-btn';
      addSup.textContent = '+ Ajouter';
      addSup.addEventListener('click', () => {
        const available = supportOptions.filter(opt => !item.supportIds.includes(opt.value));
        if (available.length === 0) {
          alert('Aucun bien support disponible à ajouter.');
          return;
        }
        const msg = 'Sélectionnez un bien support:\n' + available.map((opt, i) => `${i + 1}. ${opt.label}`).join('\n');
        const choice = prompt(msg);
        if (choice === null) return;
        const idxChoice = parseInt(choice, 10) - 1;
        if (!isNaN(idxChoice) && idxChoice >= 0 && idxChoice < available.length) {
          item.supportIds.push(available[idxChoice].value);
          saveAnalyses();
          renderPPCarto();
          updateAtelier3Chart();
        }
      });
      supCell.appendChild(addSup);
      td.appendChild(supCell);
      tr.appendChild(td);
      // Valeurs métier: display selected missions as tags with remove buttons
      td = document.createElement('td');
      const valCell = document.createElement('div');
      valCell.className = 'assoc-cell';
      item.valueIds.forEach((vid) => {
        const opt = missionOptions.find(o => o.value === vid);
        const label2 = opt ? opt.label : vid;
        const tag2 = document.createElement('span');
        tag2.className = 'assoc-item';
        tag2.textContent = label2;
        const rm2 = document.createElement('button');
        rm2.textContent = '×';
        rm2.title = 'Retirer cette valeur';
        rm2.addEventListener('click', () => {
          const pos = item.valueIds.indexOf(vid);
          if (pos >= 0) item.valueIds.splice(pos, 1);
          saveAnalyses();
          renderPPCarto();
          updateAtelier3Chart();
        });
        tag2.appendChild(rm2);
        valCell.appendChild(tag2);
      });
      const addValBtn = document.createElement('button');
      addValBtn.className = 'add-assoc-btn';
      addValBtn.textContent = '+ Ajouter';
      addValBtn.addEventListener('click', () => {
        const availableVals = missionOptions.filter(opt => !item.valueIds.includes(opt.value));
        if (availableVals.length === 0) {
          alert('Aucune valeur métier disponible à ajouter.');
          return;
        }
        const msg2 = 'Sélectionnez une valeur métier:\n' + availableVals.map((opt, i) => `${i + 1}. ${opt.label}`).join('\n');
        const choice2 = prompt(msg2);
        if (choice2 === null) return;
        const idx2 = parseInt(choice2, 10) - 1;
        if (!isNaN(idx2) && idx2 >= 0 && idx2 < availableVals.length) {
          item.valueIds.push(availableVals[idx2].value);
          saveAnalyses();
          renderPPCarto();
          updateAtelier3Chart();
        }
      });
      valCell.appendChild(addValBtn);
      td.appendChild(valCell);
      tr.appendChild(td);
      // Dépendance
      td = document.createElement('td');
      const depSelect = document.createElement('select');
      [1, 2, 3, 4].forEach(val => {
        const o = document.createElement('option');
        o.value = val;
        o.textContent = `${val}`;
        if (val === item.dependance) o.selected = true;
        depSelect.appendChild(o);
      });
      depSelect.style.backgroundColor = levelColor(item.dependance);
      depSelect.onchange = (e) => {
        item.dependance = parseInt(e.target.value, 10);
        e.target.style.backgroundColor = levelColor(item.dependance);
        // update exposition and indice
        updateDerivedCells(tr, item);
        saveAnalyses();
        updateAtelier3Chart();
      };
      td.appendChild(depSelect);
      tr.appendChild(td);
      // Pénétration
      td = document.createElement('td');
      const penSelect = document.createElement('select');
      [1, 2, 3, 4].forEach(val => {
        const o = document.createElement('option');
        o.value = val;
        o.textContent = `${val}`;
        if (val === item.penetration) o.selected = true;
        penSelect.appendChild(o);
      });
      penSelect.style.backgroundColor = levelColor(item.penetration);
      penSelect.onchange = (e) => {
        item.penetration = parseInt(e.target.value, 10);
        e.target.style.backgroundColor = levelColor(item.penetration);
        updateDerivedCells(tr, item);
        saveAnalyses();
        updateAtelier3Chart();
      };
      td.appendChild(penSelect);
      tr.appendChild(td);
      // Maturité SSI
      td = document.createElement('td');
      const matSelect = document.createElement('select');
      [1, 2, 3, 4].forEach(val => {
        const o = document.createElement('option');
        o.value = val;
        o.textContent = `${val}`;
        if (val === item.maturite) o.selected = true;
        matSelect.appendChild(o);
      });
      matSelect.style.backgroundColor = ssiColor(item.maturite);
      matSelect.onchange = (e) => {
        item.maturite = parseInt(e.target.value, 10);
        e.target.style.backgroundColor = ssiColor(item.maturite);
        updateDerivedCells(tr, item);
        saveAnalyses();
        updateAtelier3Chart();
      };
      td.appendChild(matSelect);
      tr.appendChild(td);
      // Confiance
      td = document.createElement('td');
      const confSelect = document.createElement('select');
      [1, 2, 3, 4].forEach(val => {
        const o = document.createElement('option');
        o.value = val;
        o.textContent = `${val}`;
        if (val === item.confiance) o.selected = true;
        confSelect.appendChild(o);
      });
      confSelect.style.backgroundColor = ssiColor(item.confiance);
      confSelect.onchange = (e) => {
        item.confiance = parseInt(e.target.value, 10);
        e.target.style.backgroundColor = ssiColor(item.confiance);
        updateDerivedCells(tr, item);
        saveAnalyses();
        updateAtelier3Chart();
      };
      td.appendChild(confSelect);
      tr.appendChild(td);
      // Exposition (computed)
      td = document.createElement('td');
      td.className = 'expo-cell';
      tr.appendChild(td);
      // Fiabilité cyber (computed)
      td = document.createElement('td');
      td.className = 'fiabilite-cell';
      tr.appendChild(td);
      // Indice de menace (computed)
      td = document.createElement('td');
      td.className = 'indice-cell';
      tr.appendChild(td);
      // Coordonnées (computed)
      td = document.createElement('td');
      td.className = 'coord-cell';
      tr.appendChild(td);
      // Actions
      td = document.createElement('td');
      const delBtn = document.createElement('button');
      delBtn.className = 'delete-item';
      delBtn.textContent = '×';
      delBtn.title = 'Supprimer cette partie prenante';
      delBtn.addEventListener('click', () => {
        if (!confirm('Supprimer cette partie prenante ?')) return;
        ppc.splice(idx, 1);
        saveAnalyses();
        renderPPCarto();
        updateAtelier3Chart();
      });
      td.appendChild(delBtn);
      tr.appendChild(td);
      tbody.appendChild(tr);
      // Compute derived values for this row
      updateDerivedCells(tr, item);
    });
    // Add resizer handles to the cartography table similar to missions and SROV
    addPPCartoTableResizers();

    // Ensure the “Ajouter une partie prenante” button is bound.  When the
    // application is split into multiple pages, the initial binding set up
    // in setupAddButtons() may run before this element exists.  By
    // assigning the handler here, after the table and button have been
    // rendered, we guarantee that clicking the button will append a new
    // stakeholder and refresh the cartography.  Using `onclick` also
    // overwrites any previous listener, preventing duplicate actions.
    const addPPCBtn = document.getElementById('add-ppc-btn');
    if (addPPCBtn) {
      addPPCBtn.onclick = () => {
        const analysis = analyses[currentIndex];
        if (!analysis.data) analysis.data = {};
        if (!Array.isArray(analysis.data.ppc)) analysis.data.ppc = [];
        analysis.data.ppc.push({
          id: uid(),
          nom: '',
          categorie: 'prestataire',
          supportIds: [],
          valueIds: [],
          dependance: 1,
          penetration: 1,
          maturite: 1,
          confiance: 1
        });
        saveAnalyses();
        renderPPCarto();
        updateAtelier3Chart();
      };
    }
    // Helper to update exposition, fiabilité et indice cells
    function updateDerivedCells(row, entry) {
      const expo = (entry.dependance || 1) * (entry.penetration || 1);
      const fiabilite = (entry.maturite || 1) * (entry.confiance || 1);
      const indice = fiabilite ? (expo / fiabilite) : 0;
      entry.exposition = expo;
      entry.fiabilite = fiabilite;
      entry.indiceMenace = indice;
      const expoCell = row.querySelector('.expo-cell');
      const fiabiliteCell = row.querySelector('.fiabilite-cell');
      const indiceCell = row.querySelector('.indice-cell');
      const coordCell = row.querySelector('.coord-cell');
      if (expoCell) {
        expoCell.textContent = `${expo}`;
        const bucket = pertinenceBucketForExpo(expo);
        expoCell.style.backgroundColor = levelColor(bucket);
      }
      if (fiabiliteCell) {
        fiabiliteCell.textContent = `${fiabilite}`;
        const bucket = pertinenceBucketForExpo(fiabilite);
        fiabiliteCell.style.backgroundColor = ssiColor(bucket);
      }
      if (indiceCell) {
        indiceCell.textContent = indice.toFixed(2);
        const zone = threatZoneMeta(indice);
        indiceCell.style.backgroundColor = levelColor(zone.bucket);
      }
      if (coordCell) {
        const zone = threatZoneMeta(indice);
        const typeInfo = stakeholderTypeInfo(entry.categorie);
        const angle = typeInfo.angleDeg;
        coordCell.textContent = `${zone.label} – ${angle}°`;
        entry.rayon = stakeholderRadius(indice);
        entry.angle = angle;
        entry.zoneMenace = zone.key;
      }
    }
    // Map exposition values to a bucket 1–4 similar to pertinence
    function pertinenceBucketForExpo(p) {
      if (p >= 13) return 4;
      if (p >= 9) return 3;
      if (p >= 5) return 2;
      return 1;
    }
  }

  // ----- Atelier 3: Scénarios stratégiques (strategic scenarios)
  function renderStrategies() {
    const tbody = document.getElementById('strategies-body');
    if (!tbody) return;
    tbody.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis.data) analysis.data = {};
    if (!Array.isArray(analysis.data.strategies)) analysis.data.strategies = [];
    const strategies = analysis.data.strategies;
    // Build unique source and objective options from SROV couples
    const sourceOptions = [];
    const objectifOptions = [];
    (analysis.data.srov || []).forEach(couple => {
      const src = (couple.source || '').trim();
      const obj = (couple.objectif || '').trim();
      if (src && !sourceOptions.some(o => o.value === src)) {
        sourceOptions.push({ value: src, label: src });
      }
      if (obj && !objectifOptions.some(o => o.value === obj)) {
        objectifOptions.push({ value: obj, label: obj });
      }
    });
    // Build list of parties prenantes from cartography entries (ppc)
    const ppOptions = (analysis.data.ppc || []).map(pp => ({ value: pp.id, label: pp.nom || 'PP' }));
    // Build list of events (ER) from missions events
    const eventOptions = (analysis.data.events || []).map(ev => {
      const base = ev.evenement || ev.ref || 'Évènement';
      const impact = parseInt(ev.impact, 10);
      const sev = isNaN(impact) ? '' : ` (G${impact})`;
      return { value: ev.id, label: base + sev };
    });
    // Helper to map impact level to color (1:green, 2:yellow, 3:orange, 4:red)
    const levelColor = (lvl) => {
      switch (parseInt(lvl, 10)) {
        case 1: return '#2a9d8f';
        case 2: return '#e9c46a';
        case 3: return '#f4a261';
        case 4: return '#e63946';
        default: return '#9aa0a6';
      }
    };
    strategies.forEach((item, idx) => {
      if (!item.id) item.id = uid();
      // Ensure properties exist
      item.source = item.source || '';
      item.objectif = item.objectif || '';
      if (!Array.isArray(item.chemins)) item.chemins = [];
      if (!Array.isArray(item.intermediaireIds)) item.intermediaireIds = [];
      if (!Array.isArray(item.eventIds)) item.eventIds = [];
      const tr = document.createElement('tr');
      // Source select
      let td = document.createElement('td');
      const srcSelect = document.createElement('select');
      srcSelect.innerHTML = '';
      // Add empty option
      const emptyOpt1 = document.createElement('option');
      emptyOpt1.value = '';
      emptyOpt1.textContent = '—';
      srcSelect.appendChild(emptyOpt1);
      sourceOptions.forEach(opt => {
        const o = document.createElement('option');
        o.value = opt.value;
        o.textContent = opt.label;
        if (opt.value === item.source) o.selected = true;
        srcSelect.appendChild(o);
      });
      srcSelect.onchange = (e) => {
        item.source = e.target.value;
        saveAnalyses();
      };
      td.appendChild(srcSelect);
      tr.appendChild(td);
      // Objectif select
      td = document.createElement('td');
      const objSelect = document.createElement('select');
      objSelect.innerHTML = '';
      const emptyOpt2 = document.createElement('option');
      emptyOpt2.value = '';
      emptyOpt2.textContent = '—';
      objSelect.appendChild(emptyOpt2);
      objectifOptions.forEach(opt => {
        const o = document.createElement('option');
        o.value = opt.value;
        o.textContent = opt.label;
        if (opt.value === item.objectif) o.selected = true;
        objSelect.appendChild(o);
      });
      objSelect.onchange = (e) => {
        item.objectif = e.target.value;
        saveAnalyses();
      };
      td.appendChild(objSelect);
      tr.appendChild(td);
      // Chemins d’attaque (multiple strings)
      td = document.createElement('td');
      const pathCell = document.createElement('div');
      pathCell.className = 'assoc-cell';
      item.chemins.forEach((p) => {
        const tag = document.createElement('span');
        tag.className = 'assoc-item';
        tag.textContent = p;
        const rmBtn = document.createElement('button');
        rmBtn.textContent = '×';
        rmBtn.title = 'Supprimer ce chemin';
        rmBtn.addEventListener('click', () => {
          const pos = item.chemins.indexOf(p);
          if (pos >= 0) item.chemins.splice(pos, 1);
          saveAnalyses();
          renderStrategies();
        });
        tag.appendChild(rmBtn);
        pathCell.appendChild(tag);
      });
      const addPathBtn = document.createElement('button');
      addPathBtn.className = 'add-assoc-btn';
      addPathBtn.textContent = '+ Ajouter';
      addPathBtn.addEventListener('click', () => {
        const input = prompt('Saisissez un chemin d’attaque :');
        if (input && input.trim() !== '') {
          item.chemins.push(input.trim());
          saveAnalyses();
          renderStrategies();
        }
      });
      pathCell.appendChild(addPathBtn);
      td.appendChild(pathCell);
      tr.appendChild(td);
      // Intermédiaires (parties prenantes ids)
      td = document.createElement('td');
      const intermCell = document.createElement('div');
      intermCell.className = 'assoc-cell';
      item.intermediaireIds.forEach((ppId) => {
        const opt = ppOptions.find(o => o.value === ppId);
        const label = opt ? opt.label : ppId;
        const tag = document.createElement('span');
        tag.className = 'assoc-item';
        tag.textContent = label;
        const rmBtn = document.createElement('button');
        rmBtn.textContent = '×';
        rmBtn.title = 'Retirer cet intermédiaire';
        rmBtn.addEventListener('click', () => {
          const pos = item.intermediaireIds.indexOf(ppId);
          if (pos >= 0) item.intermediaireIds.splice(pos, 1);
          saveAnalyses();
          renderStrategies();
        });
        tag.appendChild(rmBtn);
        intermCell.appendChild(tag);
      });
      const addIntermBtn = document.createElement('button');
      addIntermBtn.className = 'add-assoc-btn';
      addIntermBtn.textContent = '+ Ajouter';
      addIntermBtn.addEventListener('click', () => {
        const available = ppOptions.filter(opt => !item.intermediaireIds.includes(opt.value));
        if (available.length === 0) {
          alert('Aucune partie prenante disponible à ajouter.');
          return;
        }
        const msg = 'Sélectionnez une partie prenante :\n' + available.map((opt, i) => `${i + 1}. ${opt.label}`).join('\n');
        const choice = prompt(msg);
        if (choice === null) return;
        const idxChoice = parseInt(choice, 10) - 1;
        if (!isNaN(idxChoice) && idxChoice >= 0 && idxChoice < available.length) {
          item.intermediaireIds.push(available[idxChoice].value);
          saveAnalyses();
          renderStrategies();
        }
      });
      intermCell.appendChild(addIntermBtn);
      td.appendChild(intermCell);
      tr.appendChild(td);
      // Évènements redoutés
      td = document.createElement('td');
      const eventCell = document.createElement('div');
      eventCell.className = 'assoc-cell';
      item.eventIds.forEach((evId) => {
        const opt = eventOptions.find(o => o.value === evId);
        const label = opt ? opt.label : evId;
        const tag = document.createElement('span');
        tag.className = 'assoc-item';
        tag.textContent = label;
        const rmBtn = document.createElement('button');
        rmBtn.textContent = '×';
        rmBtn.title = 'Retirer cet évènement';
        rmBtn.addEventListener('click', () => {
          const pos = item.eventIds.indexOf(evId);
          if (pos >= 0) item.eventIds.splice(pos, 1);
          saveAnalyses();
          renderStrategies();
        });
        tag.appendChild(rmBtn);
        eventCell.appendChild(tag);
      });
      const addEvBtn = document.createElement('button');
      addEvBtn.className = 'add-assoc-btn';
      addEvBtn.textContent = 'Sélectionner';
      addEvBtn.addEventListener('click', () => {
        if (eventOptions.length === 0) {
          alert('Aucun évènement redouté disponible.');
          return;
        }
        strategyEventTarget = item;
        openImportModal('strategyEvents');
      });
      eventCell.appendChild(addEvBtn);
      td.appendChild(eventCell);
      tr.appendChild(td);
      // Gravité (computed as max impact of selected events)
      td = document.createElement('td');
      let maxImpact = 0;
      item.eventIds.forEach((evId) => {
        const ev = (analysis.data.events || []).find(e => e.id === evId);
        const imp = ev ? parseInt(ev.impact, 10) || 0 : 0;
        if (imp > maxImpact) maxImpact = imp;
      });
      if (maxImpact > 0) {
        td.textContent = `${maxImpact}`;
        td.style.backgroundColor = levelColor(maxImpact);
      } else {
        td.textContent = '';
        td.style.backgroundColor = 'transparent';
      }
      tr.appendChild(td);
      // Actions (delete)
      td = document.createElement('td');
      const delBtn = document.createElement('button');
      delBtn.className = 'delete-item';
      delBtn.textContent = '×';
      delBtn.title = 'Supprimer ce scénario';
      delBtn.addEventListener('click', () => {
        if (!confirm('Supprimer ce scénario ?')) return;
        strategies.splice(idx, 1);
        saveAnalyses();
        renderStrategies();
      });
      td.appendChild(delBtn);
      tr.appendChild(td);
      tbody.appendChild(tr);
    });
    // Add column resizers like other tables
    addStrategiesTableResizers();
    // In a multi‑page context the "Ajouter un scénario" button may not
    // have been bound when setupAddButtons() ran (the element might not
    // have existed at that time).  Assign the click handler here to
    // ensure the button always works after rendering.  Using
    // `.onclick` replaces any previous handler, avoiding duplicate
    // invocation when switching analyses or rendering multiple times.
    const addBtn = document.getElementById('add-strategy-btn');
    if (addBtn) {
      addBtn.onclick = () => {
        const analysis = analyses[currentIndex];
        if (!analysis.data) analysis.data = {};
        if (!Array.isArray(analysis.data.strategies)) analysis.data.strategies = [];
        analysis.data.strategies.push({
          id: uid(),
          source: '',
          objectif: '',
          chemins: [],
          intermediaireIds: [],
          eventIds: []
        });
        saveAnalyses();
        renderStrategies();
      };
    }
    updateAtelier3Chart();
  }

  // Add resizer handles to the strategies table
  function addStrategiesTableResizers() {
    const table = document.getElementById('strategies-table');
    if (!table) return;
    const headers = table.querySelectorAll('th');
    headers.forEach((th, i) => {
      let resizer = th.querySelector('.col-resizer');
      if (!resizer) {
        resizer = document.createElement('div');
        resizer.className = 'col-resizer';
        th.appendChild(resizer);
        let startX, startWidth;
        const onMouseDown = (e) => {
          startX = e.clientX;
          startWidth = th.offsetWidth;
          document.addEventListener('mousemove', onMouseMove);
          document.addEventListener('mouseup', onMouseUp);
        };
        const onMouseMove = (e) => {
          const dx = e.clientX - startX;
          const newWidth = Math.max(60, startWidth + dx);
          th.style.minWidth = `${newWidth}px`;
          // Also adjust corresponding td cells
          const index = Array.from(th.parentElement.children).indexOf(th);
          table.querySelectorAll('tr').forEach(row => {
            const cell = row.children[index];
            if (cell) cell.style.minWidth = `${newWidth}px`;
          });
        };
        const onMouseUp = () => {
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
        };
        resizer.addEventListener('mousedown', onMouseDown);
      }
    });
  }

  // Generic helper to add column resizers to a table by id
  function addDataTableResizers(tableId) {
    const table = document.getElementById(tableId);
    if (!table) return;
    const ths = table.querySelectorAll('thead th');
    const rows = table.querySelectorAll('tbody tr');
    ths.forEach((th, index) => {
      th.style.position = 'relative';
      const existing = th.querySelector('.col-resizer');
      if (existing) th.removeChild(existing);
      const resizer = document.createElement('div');
      resizer.className = 'col-resizer';
      th.appendChild(resizer);
      resizer.addEventListener('mousedown', (e) => {
        e.preventDefault();
        const startX = e.clientX;
        const startWidth = th.offsetWidth;
        const computedStyle = window.getComputedStyle(th);
        const minWidth = parseInt(computedStyle.minWidth) || 60;
        function onMouseMove(ev) {
          const delta = ev.clientX - startX;
          let newWidth = startWidth + delta;
          if (newWidth < minWidth) newWidth = minWidth;
          th.style.width = newWidth + 'px';
          th.style.minWidth = newWidth + 'px';
          rows.forEach(row => {
            const cell = row.children[index];
            if (cell) {
              cell.style.width = newWidth + 'px';
              cell.style.minWidth = newWidth + 'px';
            }
          });
        }
        function onMouseUp() {
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
        }
        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
      });
    });
  }

  // ----- Atelier 3: Sources de menace
  function renderSS() {
    const listEl = document.getElementById('ss-list');
    if (!listEl) return;
    listEl.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis.data) analysis.data = {};
    if (!analysis.data.ss) analysis.data.ss = [];
    analysis.data.ss.forEach((item, idx) => {
      if (!item.id) item.id = uid();
      const el = document.createElement('div');
      el.className = 'item';
      const delBtn = document.createElement('button');
      delBtn.className = 'delete-item';
      delBtn.innerHTML = '×';
      delBtn.title = 'Supprimer cette source de menace';
      delBtn.addEventListener('click', () => {
        if (!confirm('Supprimer cette source de menace ?')) return;
        analysis.data.ss.splice(idx, 1);
        saveAnalyses();
        renderSS();
        updateAtelier3Chart();
      });
      el.appendChild(delBtn);
      el.appendChild(createInput('Source de menace', 'text', item.source, (v) => {
        item.source = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Objectif visé', 'text', item.objectif, (v) => {
        item.objectif = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Vraisemblance', 'text', item.vraisemblance, (v) => {
        item.vraisemblance = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Gravité', 'text', item.gravite, (v) => {
        item.gravite = v;
        saveAnalyses();
        updateAtelier3Chart();
      }));
      listEl.appendChild(el);
    });
  }

  // ----- Atelier 4: Scénarios opérationnels
  function renderSO() {
    // This function renders the operational scenarios table if it exists;
    // otherwise it falls back to the previous simple list used in the
    // single‑page version.  The operational scenarios are stored in
    // analysis.data.so.
    const opsBody = document.getElementById('ops-body');
    const listEl = document.getElementById('so-list');
    const analysis = analyses[currentIndex];
    if (!analysis.data) analysis.data = {};
    if (!Array.isArray(analysis.data.so)) analysis.data.so = [];
    // New table layout
    if (opsBody) {
      opsBody.innerHTML = '';
      // Build options for events
      const eventOptions = (analysis.data.events || []).map(ev => ({ value: ev.id, label: ev.evenement || ev.ref || 'ER' }));
      // Build options for strategic paths (from strategies)
      const pathSet = new Set();
      (analysis.data.strategies || []).forEach(st => {
        if (Array.isArray(st.chemins)) {
          st.chemins.forEach(c => {
            const s = (c || '').trim();
            if (s) pathSet.add(s);
          });
        }
      });
      const pathOptions = Array.from(pathSet).map(p => ({ value: p, label: p }));
      // Helper to colour levels
      const levelColor = (lvl) => {
        switch (parseInt(lvl, 10)) {
          case 1: return '#2a9d8f';
          case 2: return '#e9c46a';
          case 3: return '#f4a261';
          case 4: return '#e63946';
          default: return 'transparent';
        }
      };
      analysis.data.so.forEach((item, idx) => {
        // Ensure structure
        if (!item.id) item.id = uid();
        if (!item.eventId) item.eventId = '';
        if (!item.path) item.path = '';
        if (!Array.isArray(item.risks)) item.risks = [];
        const tr = document.createElement('tr');
        // Event select
        let td = document.createElement('td');
        const evSelect = document.createElement('select');
        // empty option
        const emptyOpt = document.createElement('option');
        emptyOpt.value = '';
        emptyOpt.textContent = '—';
        evSelect.appendChild(emptyOpt);
        eventOptions.forEach(opt => {
          const o = document.createElement('option');
          o.value = opt.value;
          o.textContent = opt.label;
          if (opt.value === item.eventId) o.selected = true;
          evSelect.appendChild(o);
        });
        evSelect.onchange = (e) => {
          item.eventId = e.target.value;
          saveAnalyses();
        };
        td.appendChild(evSelect);
        tr.appendChild(td);
        // Path select
        td = document.createElement('td');
        const pathSelect = document.createElement('select');
        const emptyPath = document.createElement('option');
        emptyPath.value = '';
        emptyPath.textContent = '—';
        pathSelect.appendChild(emptyPath);
        pathOptions.forEach(opt => {
          const o = document.createElement('option');
          o.value = opt.value;
          o.textContent = opt.label;
          if (opt.value === item.path) o.selected = true;
          pathSelect.appendChild(o);
        });
        pathSelect.onchange = (e) => {
          item.path = e.target.value;
          saveAnalyses();
        };
        td.appendChild(pathSelect);
        tr.appendChild(td);
        // Risks cell (each risk with its own levels)
        td = document.createElement('td');
        const riskCell = document.createElement('div');
        riskCell.className = 'assoc-cell';
        const formatCellColor = (value) => levelColor(parseLevel(value, 1));
        item.risks.forEach((rk, rIdx) => {
          if (rk.vraisemblance === undefined || rk.vraisemblance === null) rk.vraisemblance = 1;
          if (rk.gravite === undefined || rk.gravite === null) rk.gravite = 1;
          const tag = document.createElement('span');
          tag.className = 'assoc-item';
          const label = document.createElement('span');
          label.textContent = rk.name;
          tag.appendChild(label);
          const vLabel = document.createElement('span');
          vLabel.textContent = 'V:';
          vLabel.title = 'Vraisemblance';
          tag.appendChild(vLabel);
          const vInput = document.createElement('input');
          vInput.type = 'number';
          vInput.step = '0.1';
          vInput.min = '0';
          vInput.max = '4';
          vInput.value = rk.vraisemblance;
          vInput.style.backgroundColor = formatCellColor(rk.vraisemblance);
          vInput.addEventListener('input', (e) => {
            const val = parseLevel(e.target.value, rk.vraisemblance);
            rk.vraisemblance = val;
            e.target.value = val;
            e.target.style.backgroundColor = formatCellColor(val);
            saveAnalyses();
            updateAtelier4Chart();
          });
          tag.appendChild(vInput);
          const gLabel = document.createElement('span');
          gLabel.textContent = 'G:';
          gLabel.title = 'Gravité';
          tag.appendChild(gLabel);
          const gInput = document.createElement('input');
          gInput.type = 'number';
          gInput.step = '0.1';
          gInput.min = '0';
          gInput.max = '4';
          gInput.value = rk.gravite;
          gInput.style.backgroundColor = formatCellColor(rk.gravite);
          gInput.addEventListener('input', (e) => {
            const val = parseLevel(e.target.value, rk.gravite);
            rk.gravite = val;
            e.target.value = val;
            e.target.style.backgroundColor = formatCellColor(val);
            saveAnalyses();
            updateAtelier4Chart();
          });
          tag.appendChild(gInput);
          const rm = document.createElement('button');
          rm.className = 'remove-assoc';
          rm.textContent = '×';
          rm.title = 'Retirer ce risque';
          rm.addEventListener('click', () => {
            if (!confirm('Retirer ce risque ?')) return;
            item.risks.splice(rIdx, 1);
            saveAnalyses();
            renderSO();
          });
          tag.appendChild(rm);
          riskCell.appendChild(tag);
        });
        const addRiskBtn = document.createElement('button');
        addRiskBtn.className = 'add-assoc-btn';
        addRiskBtn.textContent = '+ Ajouter (MITRE)';
        addRiskBtn.addEventListener('click', () => {
          // Open the MITRE/OWASP risk selection modal for this scenario
          openRiskModal(item);
        });
        riskCell.appendChild(addRiskBtn);
        // Manual risk addition
        const addManualBtn = document.createElement('button');
        addManualBtn.className = 'add-assoc-btn';
        addManualBtn.textContent = '+ Ajouter manuel';
        addManualBtn.addEventListener('click', () => {
          const name = prompt('Nom du risque ?');
          if (name && name.trim()) {
            item.risks.push({ name: name.trim(), vraisemblance: 1, gravite: 1 });
            saveAnalyses();
            renderSO();
          }
        });
        riskCell.appendChild(addManualBtn);
        td.appendChild(riskCell);
        tr.appendChild(td);

        // Actions: delete
        td = document.createElement('td');
        const delBtn = document.createElement('button');
        delBtn.className = 'delete-item';
        delBtn.textContent = '×';
        delBtn.title = 'Supprimer ce scénario';
        delBtn.addEventListener('click', () => {
          if (!confirm('Supprimer ce scénario ?')) return;
          analysis.data.so.splice(idx, 1);
          saveAnalyses();
          renderSO();
          updateAtelier4Chart();
        });
        td.appendChild(delBtn);
        tr.appendChild(td);
        opsBody.appendChild(tr);
      });
      // Add resizers to the operations table
      addOpsTableResizers();
      // Bind add button for operations
      const addBtn = document.getElementById('add-op-btn');
      if (addBtn) {
        addBtn.onclick = () => {
          analysis.data.so.push({
            id: uid(),
            eventId: '',
            path: '',
            risks: []
          });
          saveAnalyses();
          renderSO();
          updateAtelier4Chart();
        };
      }
      // Update chart after render
      updateAtelier4Chart();
      return;
    }
    // Fallback: simple list for old single‑page layout
    if (listEl) {
      listEl.innerHTML = '';
      analysis.data.so.forEach((item, idx) => {
        if (!item.id) item.id = uid();
        const el = document.createElement('div');
        el.className = 'item';
        const delBtn = document.createElement('button');
        delBtn.className = 'delete-item';
        delBtn.innerHTML = '×';
        delBtn.title = 'Supprimer ce scénario';
        delBtn.addEventListener('click', () => {
          if (!confirm('Supprimer ce scénario ?')) return;
          analysis.data.so.splice(idx, 1);
          saveAnalyses();
          renderSO();
          updateAtelier4Chart();
        });
        el.appendChild(delBtn);
        el.appendChild(createInput('Chemin d’attaque', 'text', item.chemin, (v) => {
          item.chemin = v;
          saveAnalyses();
        }));
        el.appendChild(createInput('Vraisemblance globale', 'text', item.vraisemblanceGlobale, (v) => {
          item.vraisemblanceGlobale = v;
          saveAnalyses();
          updateAtelier4Chart();
        }));
        listEl.appendChild(el);
      });
    }
  }

  // ----- Atelier 5: Risques
  function renderRisques() {
    const listEl = document.getElementById('risques-list');
    listEl.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis.data) analysis.data = {};
    if (!analysis.data.risques) analysis.data.risques = [];
    analysis.data.risques.forEach((item, idx) => {
      // Ensure risk has unique id and proper arrays
      if (!item.id) item.id = uid();
      if (item.libelle === undefined) item.libelle = '';
      if (!Array.isArray(item.sourceIds)) item.sourceIds = [];
      const el = document.createElement('div');
      el.className = 'item';
      const delBtn = document.createElement('button');
      delBtn.className = 'delete-item';
      delBtn.innerHTML = '×';
      delBtn.title = 'Supprimer ce risque';
      delBtn.addEventListener('click', () => {
        if (!confirm('Supprimer ce risque ?')) return;
        analysis.data.risques.splice(idx, 1);
        saveAnalyses();
        renderRisques();
        updateAtelier4Chart();
        updateAtelier5Chart();
      });
      el.appendChild(delBtn);
      // Select mission
      const missionOptions = (analysis.data.missions || []).map((m, i) => ({ value: m.id, label: m.denom || `Mission ${i + 1}` }));
      el.appendChild(createSelect('Mission', missionOptions, item.missionId || '', (v) => {
        item.missionId = v;
        saveAnalyses();
      }));
      // Select événement
      const eventOptions = (analysis.data.events || []).map((ev, i) => ({ value: ev.id, label: ev.ref || ev.evenement || `ER ${i + 1}` }));
      el.appendChild(createSelect('Évènement', eventOptions, item.eventId || '', (v) => {
        item.eventId = v;
        saveAnalyses();
      }));
      // Select scénario opérationnel
      const scenarioOptions = (analysis.data.so || []).map((so, i) => ({ value: so.id, label: so.chemin || `Scénario ${i + 1}` }));
      el.appendChild(createSelect('Scénario', scenarioOptions, item.scenarioId || '', (v) => {
        item.scenarioId = v;
        saveAnalyses();
      }));
      // Select sources de risque (multi)
      const sourceOptions = (analysis.data.srov || []).map((srov, i) => ({ value: srov.id, label: srov.source || `Source ${i + 1}` }));
      el.appendChild(createMultiSelect('Sources', sourceOptions, item.sourceIds || [], (vals) => {
        item.sourceIds = vals;
        saveAnalyses();
      }));
      // Field for risk label (libellé)
      el.appendChild(createInput('Libellé', 'text', item.libelle, (v) => {
        item.libelle = v;
        saveAnalyses();
        updateAtelier4Chart();
        updateAtelier5Chart();
      }));
      // Fields for titre, description, indice, vraisemblance, gravite, mesures
      el.appendChild(createInput('Titre du risque', 'text', item.titre, (v) => {
        item.titre = v;
        saveAnalyses();
        updateAtelier4Chart();
        updateAtelier5Chart();
      }));
      el.appendChild(createInput('Description', 'textarea', item.description, (v) => {
        item.description = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Indice', 'text', item.indice, (v) => {
        item.indice = v;
        saveAnalyses();
        updateAtelier4Chart();
        updateAtelier5Chart();
      }));
      el.appendChild(createInput('Vraisemblance', 'text', item.vraisemblance, (v) => {
        item.vraisemblance = v;
        saveAnalyses();
        updateAtelier4Chart();
        updateAtelier5Chart();
      }));
      el.appendChild(createInput('Gravité', 'text', item.gravite, (v) => {
        item.gravite = v;
        saveAnalyses();
        updateAtelier4Chart();
        updateAtelier5Chart();
      }));
      el.appendChild(createInput('Mesures de traitement', 'textarea', item.mesures, (v) => {
        item.mesures = v;
        saveAnalyses();
      }));
      listEl.appendChild(el);
    });
  }

  // ----- Atelier 5: Actions et conformité
  // Render actions for GAP requirements (non appliquées ou partiellement appliquées)
  function renderGapActions() {
    const body = document.getElementById('gap-actions-body');
    if (!body) return;
    body.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) return;
    const gap = analysis.data.gap || [];
    if (!Array.isArray(analysis.data.actionsGap)) analysis.data.actionsGap = [];
    analysis.data.actionsGap.forEach(entry => {
      if (!entry) return;
      const tr = document.createElement('tr');
      const req = gap.find(r => r.id === entry.sourceId);
      const tdName = document.createElement('td');
      if (req) {
        tdName.textContent = req.titre || req.domaine || 'Exigence';
      } else {
        const inpReq = document.createElement('input');
        inpReq.value = entry.customTitre || '';
        inpReq.addEventListener('input', (e) => {
          entry.customTitre = e.target.value;
          saveAnalyses();
          renderGapActions();
        });
        tdName.appendChild(inpReq);
      }
      tr.appendChild(tdName);
      const tdActions = document.createElement('td');
      tdActions.className = 'assoc-cell';
      if (!Array.isArray(entry.actions)) entry.actions = [];
      const actTable = document.createElement('table');
      actTable.className = 'nested-table';
      const headerRow = document.createElement('tr');
      headerRow.innerHTML = '<th>Nom</th><th>Description</th><th>Responsable</th><th>Début</th><th>Fin</th><th></th>';
      actTable.appendChild(headerRow);
      entry.actions.forEach((act, aIdx) => {
        const ar = document.createElement('tr');
        let tdA = document.createElement('td');
        const inpName = document.createElement('input');
        inpName.value = act.name || '';
        inpName.addEventListener('input', (e) => {
          act.name = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpName);
        ar.appendChild(tdA);
        tdA = document.createElement('td');
        const inpDesc = document.createElement('textarea');
        inpDesc.value = act.description || '';
        inpDesc.rows = 2;
        inpDesc.addEventListener('input', (e) => {
          act.description = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpDesc);
        ar.appendChild(tdA);
        tdA = document.createElement('td');
        const inpResp = document.createElement('input');
        inpResp.value = act.responsable || '';
        inpResp.addEventListener('input', (e) => {
          act.responsable = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpResp);
        ar.appendChild(tdA);
        tdA = document.createElement('td');
        const inpStart = document.createElement('input');
        inpStart.type = 'date';
        inpStart.value = act.start || '';
        inpStart.addEventListener('change', (e) => {
          act.start = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpStart);
        ar.appendChild(tdA);
        tdA = document.createElement('td');
        const inpEnd = document.createElement('input');
        inpEnd.type = 'date';
        inpEnd.value = act.end || '';
        inpEnd.addEventListener('change', (e) => {
          act.end = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpEnd);
        ar.appendChild(tdA);
        tdA = document.createElement('td');
        const delBtn = document.createElement('button');
        delBtn.className = 'delete-action';
        delBtn.textContent = '×';
        delBtn.addEventListener('click', () => {
          if (!confirm('Supprimer cette action ?')) return;
          entry.actions.splice(aIdx, 1);
          saveAnalyses();
          renderGapActions();
          renderPlanActions();
        });
        tdA.appendChild(delBtn);
        ar.appendChild(tdA);
        actTable.appendChild(ar);
      });
      const addRow = document.createElement('tr');
      const addTd = document.createElement('td');
      addTd.colSpan = 6;
      const addBtn = document.createElement('button');
      addBtn.className = 'add-assoc-btn';
      addBtn.textContent = '+ Action';
      addBtn.addEventListener('click', () => {
        entry.actions.push({ name:'', description:'', responsable:'', start:'', end:'' });
        saveAnalyses();
        renderGapActions();
        renderPlanActions();
      });
      addTd.appendChild(addBtn);
      const addExisting = document.createElement('button');
      addExisting.className = 'add-assoc-btn';
      addExisting.textContent = '+ Action existante';
      addExisting.addEventListener('click', () => {
        const allActs = [];
        analysis.data.actionsGap.forEach(e => {
          if (e !== entry && Array.isArray(e.actions)) {
            e.actions.forEach(a => {
              if (a && a.name) allActs.push(a);
            });
          }
        });
        const currentNames = entry.actions.map(a => (a.name || '').trim());
        const available = allActs.filter(a => !currentNames.includes((a.name || '').trim()));
        if (available.length === 0) {
          alert('Aucune action existante disponible.');
          return;
        }
        const msg = 'Sélectionnez une action existante:\n' + available.map((a,i)=>`${i+1}. ${a.name}`).join('\n');
        const input = prompt(msg);
        if (input === null) return;
        const index = parseInt(input,10) - 1;
        if (!isNaN(index) && index >=0 && index < available.length) {
          entry.actions.push(Object.assign({}, available[index]));
          saveAnalyses();
          renderGapActions();
          renderPlanActions();
        }
      });
      addTd.appendChild(addExisting);
      addRow.appendChild(addTd);
      actTable.appendChild(addRow);
      tdActions.appendChild(actTable);
      tr.appendChild(tdActions);
      const tdDel = document.createElement('td');
      const delRow = document.createElement('button');
      delRow.className = 'delete-item';
      delRow.textContent = '×';
      delRow.addEventListener('click', () => {
        if (!confirm('Supprimer cette exigence ?')) return;
        const idx = analysis.data.actionsGap.indexOf(entry);
        if (idx >= 0) analysis.data.actionsGap.splice(idx, 1);
        saveAnalyses();
        renderGapActions();
        renderPlanActions();
      });
      tdDel.appendChild(delRow);
      tr.appendChild(tdDel);
      body.appendChild(tr);
    });
    addDataTableResizers('gap-actions-table');
  }

  // Render actions for supports: allow user to add rows for each selected support
  function renderSupportActions() {
    const body = document.getElementById('support-actions-body');
    if (!body) return;
    body.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) return;
    if (!Array.isArray(analysis.data.actionsSupports)) analysis.data.actionsSupports = [];
    const supportsSet = new Set();
    (analysis.data.missions || []).forEach(mis => {
      (mis.supports || []).forEach(s => {
        if (s && (s.name || s.denom)) supportsSet.add(s.name || s.denom);
      });
    });
    (analysis.data.supportsQualif || []).forEach(s => {
      if (s && s.name) supportsSet.add(s.name);
    });
    const supportOptions = Array.from(supportsSet);
    analysis.data.actionsSupports.forEach((row, rowIndex) => {
      const tr = document.createElement('tr');
      // Vulnerability select
      let tdV = document.createElement('td');
      const selV = document.createElement('select');
      selV.className = 'form-select';
      const supportObj = (analysis.data.supportsQualif || []).find(s => s.name === row.supportName);
      const vulnOpts = supportObj ? (supportObj.vulnerabilities || []) : [];
      if (row.vulnName && !row.initialLevel) {
        const vObj = vulnOpts.find(v => v.name === row.vulnName);
        if (vObj) {
          row.initialLevel = vObj.level || '';
          if (!row.residualLevel) row.residualLevel = row.initialLevel;
        }
      }
      selV.innerHTML = '<option value="">--Sélectionner--</option>' + vulnOpts.map(v => `<option value="${v.name}" ${row.vulnName===v.name?'selected':''}>${v.name}</option>`).join('');
      selV.addEventListener('change', (e) => {
        row.vulnName = e.target.value;
        const vObj = vulnOpts.find(v => v.name === row.vulnName);
        row.initialLevel = vObj ? (vObj.level || '') : '';
        row.residualLevel = row.initialLevel;
        saveAnalyses();
        renderSupportActions();
        renderPlanActions();
      });
      tdV.appendChild(selV);
      tr.appendChild(tdV);
      // Support select
      const tdSupport = document.createElement('td');
      const sel = document.createElement('select');
      sel.className = 'form-select';
      sel.innerHTML = '<option value="">--Sélectionner--</option>' + supportOptions.map(opt => `<option value="${opt}" ${row.supportName===opt?'selected':''}>${opt}</option>`).join('');
      sel.addEventListener('change', (e) => {
        row.supportName = e.target.value;
        // reset vuln if support changed
        row.vulnName = '';
        row.initialLevel = '';
        row.residualLevel = '';
        saveAnalyses();
        renderSupportActions();
        renderPlanActions();
      });
      tdSupport.appendChild(sel);
      tr.appendChild(tdSupport);
      // Initial level
      const tdInit = document.createElement('td');
      const initSel = document.createElement('select');
      ['info','faible','moderee','forte','critique'].forEach(optVal => {
        const opt = document.createElement('option');
        opt.value = optVal;
        opt.textContent = optVal.charAt(0).toUpperCase() + optVal.slice(1);
        if ((row.initialLevel || '') === optVal) opt.selected = true;
        initSel.appendChild(opt);
      });
      initSel.disabled = true;
      setVulnLevelColor(initSel);
      tdInit.appendChild(initSel);
      tr.appendChild(tdInit);
      // Residual level
      const tdRes = document.createElement('td');
      const resSel = document.createElement('select');
      ['info','faible','moderee','forte','critique'].forEach(optVal => {
        const opt = document.createElement('option');
        opt.value = optVal;
        opt.textContent = optVal.charAt(0).toUpperCase() + optVal.slice(1);
        if ((row.residualLevel || '') === optVal) opt.selected = true;
        resSel.appendChild(opt);
      });
      setVulnLevelColor(resSel);
      resSel.addEventListener('change', (e) => {
        row.residualLevel = e.target.value;
        saveAnalyses();
        setVulnLevelColor(resSel);
        renderSupportLevelChart();
        renderPlanActions();
      });
      tdRes.appendChild(resSel);
      tr.appendChild(tdRes);
      // Actions cell
      const tdActions = document.createElement('td');
      tdActions.className = 'assoc-cell';
      if (!Array.isArray(row.actions)) row.actions = [];
      const actTable = document.createElement('table');
      actTable.className = 'nested-table';
      const headerRow = document.createElement('tr');
      headerRow.innerHTML = '<th>Nom</th><th>Description</th><th>Responsable</th><th>Début</th><th>Fin</th><th></th>';
      actTable.appendChild(headerRow);
      row.actions.forEach((act, aIdx) => {
        const ar = document.createElement('tr');
        let tdA = document.createElement('td');
        const inpName = document.createElement('input');
        inpName.value = act.name || '';
        inpName.addEventListener('input', (e) => {
          act.name = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpName);
        ar.appendChild(tdA);
        // Description
        tdA = document.createElement('td');
        const inpDesc = document.createElement('textarea');
        inpDesc.rows = 2;
        inpDesc.value = act.description || '';
        inpDesc.addEventListener('input', (e) => {
          act.description = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpDesc);
        ar.appendChild(tdA);
        // Responsable
        tdA = document.createElement('td');
        const inpResp = document.createElement('input');
        inpResp.value = act.responsable || '';
        inpResp.addEventListener('input', (e) => {
          act.responsable = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpResp);
        ar.appendChild(tdA);
        // Start
        tdA = document.createElement('td');
        const inpStart = document.createElement('input');
        inpStart.type = 'date';
        inpStart.value = act.start || '';
        inpStart.addEventListener('change', (e) => {
          act.start = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpStart);
        ar.appendChild(tdA);
        // End
        tdA = document.createElement('td');
        const inpEnd = document.createElement('input');
        inpEnd.type = 'date';
        inpEnd.value = act.end || '';
        inpEnd.addEventListener('change', (e) => {
          act.end = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpEnd);
        ar.appendChild(tdA);
        // Delete
        tdA = document.createElement('td');
        const delBtn = document.createElement('button');
        delBtn.className = 'delete-action';
        delBtn.textContent = '×';
        delBtn.addEventListener('click', () => {
          if (!confirm('Supprimer cette action ?')) return;
          row.actions.splice(aIdx, 1);
          saveAnalyses();
          renderSupportActions();
          renderPlanActions();
        });
        tdA.appendChild(delBtn);
        ar.appendChild(tdA);
        actTable.appendChild(ar);
      });
      // Add new or existing action
      const addRow = document.createElement('tr');
      const addTd = document.createElement('td');
      addTd.colSpan = 6;
      const addBtnA = document.createElement('button');
      addBtnA.className = 'add-assoc-btn';
      addBtnA.textContent = '+ Action';
      addBtnA.addEventListener('click', () => {
        row.actions.push({ name:'', description:'', responsable:'', start:'', end:'' });
        saveAnalyses();
        renderSupportActions();
        renderPlanActions();
      });
      addTd.appendChild(addBtnA);
      const addExisting = document.createElement('button');
      addExisting.className = 'add-assoc-btn';
      addExisting.textContent = '+ Action existante';
      addExisting.addEventListener('click', () => {
        const allActs = [];
        analysis.data.actionsSupports.forEach(r => {
          if (r !== row && Array.isArray(r.actions)) {
            r.actions.forEach(a => {
              if (a && a.name) allActs.push(a);
            });
          }
        });
        const currentNames = row.actions.map(a => (a.name || '').trim());
        const available = allActs.filter(a => !currentNames.includes((a.name || '').trim()));
        if (available.length === 0) {
          alert('Aucune action existante disponible.');
          return;
        }
        const msg = 'Sélectionnez une action existante:\n' + available.map((a,i)=>`${i+1}. ${a.name}`).join('\n');
        const input = prompt(msg);
        if (input === null) return;
        const index = parseInt(input,10) - 1;
        if (!isNaN(index) && index >=0 && index < available.length) {
          row.actions.push(Object.assign({}, available[index]));
          saveAnalyses();
          renderSupportActions();
          renderPlanActions();
        }
      });
      addTd.appendChild(addExisting);
      addRow.appendChild(addTd);
      actTable.appendChild(addRow);
      tdActions.appendChild(actTable);
      tr.appendChild(tdActions);
      // Delete row
      const tdDel = document.createElement('td');
      const delRow = document.createElement('button');
      delRow.className = 'delete-item';
      delRow.textContent = '×';
      delRow.addEventListener('click', () => {
        if (!confirm('Supprimer cette ligne ?')) return;
        analysis.data.actionsSupports.splice(rowIndex, 1);
        saveAnalyses();
        renderSupportActions();
        renderPlanActions();
      });
      tdDel.appendChild(delRow);
      tr.appendChild(tdDel);
      body.appendChild(tr);
    });
    addDataTableResizers('support-actions-table');
    renderSupportLevelChart();
    // Add row button is outside in HTML
  }

  function renderSupportLevelChart() {
    const canvas = document.getElementById('support-level-chart');
    if (!canvas) return;
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) {
      clearCanvas(canvas);
      return;
    }
    if (!Array.isArray(analysis.data.actionsSupports)) {
      clearCanvas(canvas);
      return;
    }
    const levels = ['info','faible','moderee','forte','critique'];
    const initCounts = { info:0, faible:0, moderee:0, forte:0, critique:0 };
    const residCounts = { info:0, faible:0, moderee:0, forte:0, critique:0 };
    analysis.data.actionsSupports.forEach(row => {
      const init = (row.initialLevel || '').toLowerCase();
      const res = (row.residualLevel || '').toLowerCase();
      if (initCounts.hasOwnProperty(init)) initCounts[init]++;
      if (residCounts.hasOwnProperty(res)) residCounts[res]++;
    });
    const labels = levels.map(l => l.charAt(0).toUpperCase() + l.slice(1));
    const dataA = levels.map(l => initCounts[l]);
    const dataB = levels.map(l => residCounts[l]);
    const colors = levels.map(l => vulnLevelColor(l));
    if (dataA.some(v=>v>0) || dataB.some(v=>v>0)) drawGroupedBarChart(canvas, labels, dataA, dataB, colors);
    else clearCanvas(canvas);
  }

  // Render actions for parties: allow user to add rows referencing stakeholders
  function renderPartiesActions() {
    const body = document.getElementById('parties-actions-body');
    if (!body) return;
    body.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) return;
    if (!Array.isArray(analysis.data.actionsParties)) analysis.data.actionsParties = [];
    // Build party options from cartography (ppc)
    const ppOptions = (analysis.data.ppc || []).map(pp => ({ id: pp.id, name: pp.nom || pp.name || 'Partie' }));
    // Render each row
    analysis.data.actionsParties.forEach((row, rowIndex) => {
      const tr = document.createElement('tr');
      // Party select
      const tdParty = document.createElement('td');
      const sel = document.createElement('select');
      sel.className = 'form-select';
      sel.innerHTML = '<option value="">--Sélectionner--</option>' + ppOptions.map(opt => `<option value="${opt.id}" ${row.ppId===opt.id?'selected':''}>${opt.name}</option>`).join('');
      sel.addEventListener('change', (e) => {
        row.ppId = e.target.value;
        saveAnalyses();
        renderPartiesActions();
        renderPlanActions();
      });
      tdParty.appendChild(sel);
      tr.appendChild(tdParty);
      // Actions cell: nested table for actions associated with this party
      const tdActions = document.createElement('td');
      tdActions.className = 'assoc-cell';
      if (!Array.isArray(row.actions)) row.actions = [];
      const actTable = document.createElement('table');
      actTable.className = 'nested-table';
      const headerRow = document.createElement('tr');
      headerRow.innerHTML = '<th>Nom</th><th>Description</th><th>Responsable</th><th>Début</th><th>Fin</th><th></th>';
      actTable.appendChild(headerRow);
      row.actions.forEach((act, aIdx) => {
        const ar = document.createElement('tr');
        // Nom
        let tdA = document.createElement('td');
        const inpName = document.createElement('input');
        inpName.value = act.name || '';
        inpName.addEventListener('input', (e) => {
          act.name = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpName);
        ar.appendChild(tdA);
        // Description
        tdA = document.createElement('td');
        const inpDesc = document.createElement('textarea');
        inpDesc.rows = 2;
        inpDesc.value = act.description || '';
        inpDesc.addEventListener('input', (e) => {
          act.description = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpDesc);
        ar.appendChild(tdA);
        // Responsable
        tdA = document.createElement('td');
        const inpResp = document.createElement('input');
        inpResp.value = act.responsable || '';
        inpResp.addEventListener('input', (e) => {
          act.responsable = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpResp);
        ar.appendChild(tdA);
        // Start
        tdA = document.createElement('td');
        const inpStart = document.createElement('input');
        inpStart.type = 'date';
        inpStart.value = act.start || '';
        inpStart.addEventListener('change', (e) => {
          act.start = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpStart);
        ar.appendChild(tdA);
        // End
        tdA = document.createElement('td');
        const inpEnd = document.createElement('input');
        inpEnd.type = 'date';
        inpEnd.value = act.end || '';
        inpEnd.addEventListener('change', (e) => {
          act.end = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpEnd);
        ar.appendChild(tdA);
        // Delete button
        tdA = document.createElement('td');
        const delBtn = document.createElement('button');
        delBtn.className = 'delete-action';
        delBtn.textContent = '×';
        delBtn.addEventListener('click', () => {
          if (!confirm('Supprimer cette action ?')) return;
          row.actions.splice(aIdx, 1);
          saveAnalyses();
          renderPartiesActions();
          renderPlanActions();
        });
        tdA.appendChild(delBtn);
        ar.appendChild(tdA);
        actTable.appendChild(ar);
      });
      // Row to add new or existing action
      const addRow = document.createElement('tr');
      const addTd = document.createElement('td');
      addTd.colSpan = 6;
      const addBtnA = document.createElement('button');
      addBtnA.className = 'add-assoc-btn';
      addBtnA.textContent = '+ Action';
      addBtnA.addEventListener('click', () => {
        row.actions.push({ name:'', description:'', responsable:'', start:'', end:'' });
        saveAnalyses();
        renderPartiesActions();
        renderPlanActions();
      });
      addTd.appendChild(addBtnA);
      const addExisting = document.createElement('button');
      addExisting.className = 'add-assoc-btn';
      addExisting.textContent = '+ Action existante';
      addExisting.addEventListener('click', () => {
        const allActs = [];
        analysis.data.actionsParties.forEach(r => {
          if (r !== row && Array.isArray(r.actions)) {
            r.actions.forEach(a => {
              if (a && a.name) allActs.push(a);
            });
          }
        });
        const currentNames = row.actions.map(a => (a.name || '').trim());
        const available = allActs.filter(a => !currentNames.includes((a.name || '').trim()));
        if (available.length === 0) {
          alert('Aucune action existante disponible.');
          return;
        }
        const msg = 'Sélectionnez une action existante:\n' + available.map((a,i)=>`${i+1}. ${a.name}`).join('\n');
        const input = prompt(msg);
        if (input === null) return;
        const index = parseInt(input,10) - 1;
        if (!isNaN(index) && index >=0 && index < available.length) {
          row.actions.push(Object.assign({}, available[index]));
          saveAnalyses();
          renderPartiesActions();
          renderPlanActions();
        }
      });
      addTd.appendChild(addExisting);
      addRow.appendChild(addTd);
      actTable.appendChild(addRow);
      tdActions.appendChild(actTable);
      tr.appendChild(tdActions);
      // Delete row
      const tdDel = document.createElement('td');
      const delRow = document.createElement('button');
      delRow.className = 'delete-item';
      delRow.textContent = '×';
      delRow.addEventListener('click', () => {
        if (!confirm('Supprimer cette partie prenante ?')) return;
        analysis.data.actionsParties.splice(rowIndex, 1);
        saveAnalyses();
        renderPartiesActions();
        renderPlanActions();
      });
      tdDel.appendChild(delRow);
      tr.appendChild(tdDel);
      body.appendChild(tr);
    });
    addDataTableResizers('parties-actions-table');
  }

  // Render actions for risks: show each risk and allow adding actions and residual levels
  function renderRisquesChart() {
    const container = document.getElementById('risques-chart');
    if (!container) {
      if (risquesChartInstance) {
        risquesChartInstance.dispose();
        risquesChartInstance = null;
      }
      atelier5DragHandlers = null;
      renderRisquesScoreChart([]);
      return;
    }
    if (typeof echarts === 'undefined') return;

    if (risquesChartInstance && risquesChartInstance.getDom() !== container) {
      risquesChartInstance.dispose();
      risquesChartInstance = null;
    }
    if (!risquesChartInstance) {
      risquesChartInstance = echarts.init(container);
    }

    const analysis = analyses[currentIndex];
    const riskMap = new Map();
    if (analysis && analysis.data) {
      (analysis.data.risques || []).forEach(riskObj => {
        const label = riskObj.libelle || riskObj.titre || riskObj.indice || riskObj.id || '';
        if (!label) return;
        const id = riskObj.id || label;
        const current = riskMap.get(id) || {
          id,
          name: label,
          vraisemblance: parseLevel(riskObj.vraisemblance, 1),
          gravite: parseLevel(riskObj.gravite, 1)
        };
        current.vraisemblance = Math.max(current.vraisemblance, parseLevel(riskObj.vraisemblance, 1));
        current.gravite = Math.max(current.gravite, parseLevel(riskObj.gravite, 1));
        riskMap.set(id, current);
      });
      (analysis.data.so || []).forEach(so => {
        (so.risks || []).forEach(rk => {
          const label = rk.name || rk.id;
          if (!label) return;
          const id = rk.id || rk.name;
          const current = riskMap.get(id) || {
            id,
            name: label,
            vraisemblance: parseLevel(rk.vraisemblance, 1),
            gravite: parseLevel(rk.gravite, 1)
          };
          current.vraisemblance = Math.max(current.vraisemblance, parseLevel(rk.vraisemblance, 1));
          current.gravite = Math.max(current.gravite, parseLevel(rk.gravite, 1));
          riskMap.set(id, current);
        });
      });
    }

    const basePoints = [];
    if (analysis && analysis.data) {
      (analysis.data.actionsRisques || []).forEach((row, rowIndex) => {
        if (!row) return;
        if (!row.riskId && row.riskName) {
          for (const [id, obj] of riskMap.entries()) {
            if (obj.name === row.riskName) {
              row.riskId = id;
              break;
            }
          }
          if (!row.riskId) row.riskId = row.riskName ? row.riskName.split(' ')[0] : '';
        }
        const risk = riskMap.get(row.riskId) || {
          id: row.riskId,
          name: row.riskName || row.riskId || 'Risque',
          vraisemblance: parseLevel(row.residualV, 1),
          gravite: parseLevel(row.residualG, 1)
        };
        const initial = [
          parseLevel(risk.gravite, 1),
          parseLevel(risk.vraisemblance, 1)
        ];
        const residual = [
          parseLevel(row.residualG, initial[0]),
          parseLevel(row.residualV, initial[1])
        ];
        const displayId = formatRiskIdentifier(row.riskId || '');
        const label = displayId || row.riskName || risk.name || 'Risque';
        const description = row.riskName || risk.name || '';
        basePoints.push({
          name: label,
          description,
          initial,
          residual,
          meta: { rowIndex, riskId: row.riskId || '' },
          metaKey: `risk::${row.riskId || rowIndex}::${rowIndex}`
        });
      });
    }

    const accent = getCssVar('--accent', '#4da3ff');
    const danger = getCssVar('--danger', '#e74c3c');
    const textPrimary = getCssVar('--text-primary', '#e6e9ef');
    const textSecondary = getCssVar('--text-secondary', '#8aa0c4');

    if (!basePoints.length) {
      renderRisquesScoreChart([]);
      risquesChartInstance.setOption({
        backgroundColor: 'transparent',
        title: {
          text: 'Traitement des risques (Actuel → Résiduel)',
          left: 'center',
          textStyle: { color: textPrimary, fontSize: 18, fontWeight: 'bold' }
        },
        grid: { left: 60, right: 30, top: 80, bottom: 60 },
        xAxis: {
          type: 'value',
          min: 0,
          max: 4,
          splitNumber: 4,
          axisLine: { lineStyle: { color: textSecondary } },
          axisLabel: { color: textSecondary },
          splitLine: { lineStyle: { color: 'rgba(148,163,184,0.2)' } },
          name: 'Gravité',
          nameGap: 30,
          nameTextStyle: { color: textSecondary }
        },
        yAxis: {
          type: 'value',
          min: 0,
          max: 4,
          splitNumber: 4,
          axisLine: { lineStyle: { color: textSecondary } },
          axisLabel: { color: textSecondary },
          splitLine: { lineStyle: { color: 'rgba(148,163,184,0.2)' } },
          name: 'Vraisemblance',
          nameGap: 40,
          nameTextStyle: { color: textSecondary }
        },
        legend: {
          data: ['Actuel', 'Résiduel', 'Traitement'],
          top: 40,
          textStyle: { color: textSecondary }
        },
        series: [],
        graphic: [
          {
            type: 'text',
            left: 'center',
            top: 'middle',
            style: {
              text: 'Aucune action de traitement de risque enregistrée.',
              fill: textSecondary,
              fontSize: 14
            }
          }
        ]
      }, true);
      return;
    }

    const scorePoints = basePoints.map(point => {
      const gravInitial = Number(point.initial && point.initial[0] !== undefined ? point.initial[0] : 0);
      const vraisInitial = Number(point.initial && point.initial[1] !== undefined ? point.initial[1] : 0);
      const gravResidual = Number(point.residual && point.residual[0] !== undefined ? point.residual[0] : 0);
      const vraisResidual = Number(point.residual && point.residual[1] !== undefined ? point.residual[1] : 0);
      const initialScore = Number((gravInitial * vraisInitial * 0.3125).toFixed(5));
      const residualScore = Number((gravResidual * vraisResidual * 0.3125).toFixed(5));
      return {
        name: point.name,
        initial: initialScore,
        residual: residualScore,
        description: point.description
      };
    });

    const initialPoints = basePoints.map(point => ({
      name: point.name,
      value: point.initial.slice(),
      rawValue: point.initial.slice(),
      description: point.description,
      type: 'Actuel',
      meta: point.meta,
      metaKey: point.metaKey
    }));
    const residualPoints = basePoints.map(point => ({
      name: point.name,
      value: point.residual.slice(),
      rawValue: point.residual.slice(),
      description: point.description,
      type: 'Résiduel',
      meta: point.meta,
      metaKey: point.metaKey
    }));

    const initialSpread = spreadInSquare(initialPoints, 0.14);
    const residualSpread = spreadInSquare(residualPoints, 0.14);

    const initialPositions = new Map(initialSpread.map(item => [(item.metaKey || item.name), item.value]));
    const residualPositions = new Map(residualSpread.map(item => [(item.metaKey || item.name), item.value]));

    const lineData = basePoints.map(point => ({
      name: point.name,
      description: point.description,
      from: point.initial.slice(),
      to: point.residual.slice(),
      meta: point.meta,
      metaKey: point.metaKey,
      coords: [
        initialPositions.get(point.metaKey || point.name) || point.initial,
        residualPositions.get(point.metaKey || point.name) || point.residual
      ]
    }));

    const option = {
      backgroundColor: 'transparent',
      title: {
        text: 'Traitement des risques (Actuel → Résiduel)',
        left: 'center',
        textStyle: { color: textPrimary, fontSize: 18, fontWeight: 'bold' }
      },
      legend: {
        data: ['Actuel', 'Résiduel', 'Traitement'],
        top: 40,
        textStyle: { color: textSecondary }
      },
      grid: { left: 60, right: 30, top: 80, bottom: 60 },
      tooltip: {
        trigger: 'item',
        borderWidth: 1,
        backgroundColor: '#0f172a',
        borderColor: accent,
        textStyle: { color: textPrimary },
        formatter: (params) => {
          if (params.seriesType === 'lines') {
            const data = params.data || {};
            const fromRaw = data.from || (data.coords ? data.coords[0] : []);
            const toRaw = data.to || (data.coords ? data.coords[1] : []);
            const desc = data.description ? `<br/><em>${data.description}</em>` : '';
            return `<strong>${data.name}</strong>${desc}<br/>Actuel : G ${fromRaw[0]} / V ${fromRaw[1]}<br/>Résiduel : G ${toRaw[0]} / V ${toRaw[1]}`;
          }
          const desc = params.data && params.data.description ? `<br/><em>${params.data.description}</em>` : '';
          const raw = params.data && params.data.rawValue ? params.data.rawValue : params.value;
          const gravite = raw && raw[0] !== undefined ? raw[0] : params.value[0];
          const vraisemblance = raw && raw[1] !== undefined ? raw[1] : params.value[1];
          return `<strong>${params.name}</strong>${desc}<br/>${params.seriesName} – Gravité : ${gravite}<br/>${params.seriesName} – Vraisemblance : ${vraisemblance}`;
        }
      },
      xAxis: {
        type: 'value',
        min: 0,
        max: 4,
        splitNumber: 4,
        name: 'Gravité',
        nameGap: 30,
        nameTextStyle: { color: textSecondary },
        axisLine: { lineStyle: { color: textSecondary } },
        axisLabel: { color: textSecondary },
        splitLine: { lineStyle: { color: 'rgba(148,163,184,0.2)' } }
      },
      yAxis: {
        type: 'value',
        min: 0,
        max: 4,
        splitNumber: 4,
        name: 'Vraisemblance',
        nameGap: 40,
        nameTextStyle: { color: textSecondary },
        axisLine: { lineStyle: { color: textSecondary } },
        axisLabel: { color: textSecondary },
        splitLine: { lineStyle: { color: 'rgba(148,163,184,0.2)' } }
      },
      series: [
        {
          name: 'Traitement',
          type: 'lines',
          coordinateSystem: 'cartesian2d',
          data: lineData,
          lineStyle: { color: textSecondary, width: 2, type: 'dotted' },
          effect: { show: true, symbol: 'arrow', symbolSize: 9, color: textSecondary, trailLength: 0 },
          z: 1
        },
        {
          name: 'Actuel',
          type: 'scatter',
          data: initialSpread,
          symbolSize: 20,
          itemStyle: { color: accent },
          label: {
            show: true,
            formatter: '{b}',
            position: 'top',
            color: textPrimary,
            fontSize: 12
          },
          z: 2
        },
        {
          name: 'Résiduel',
          type: 'scatter',
          data: residualSpread,
          symbolSize: 16,
          itemStyle: { color: danger },
          label: {
            show: true,
            formatter: '{b}',
            position: 'bottom',
            color: textPrimary,
            fontSize: 12
          },
          z: 3
        }
      ],
      animationDuration: 300,
      graphic: []
    };

    risquesChartInstance.setOption(option, true);
    renderRisquesScoreChart(scorePoints);
    bindAtelier5Drag(residualSpread, lineData);
  }

  function renderRisquesScoreChart(points) {
    const container = document.getElementById('risques-score-chart');
    if (!container) {
      if (risquesScoreChartInstance) {
        risquesScoreChartInstance.dispose();
        risquesScoreChartInstance = null;
      }
      return;
    }
    if (typeof echarts === 'undefined') return;

    if (risquesScoreChartInstance && risquesScoreChartInstance.getDom() !== container) {
      risquesScoreChartInstance.dispose();
      risquesScoreChartInstance = null;
    }
    if (!risquesScoreChartInstance) {
      risquesScoreChartInstance = echarts.init(container);
    }

    const accent = getCssVar('--accent', '#4da3ff');
    const danger = getCssVar('--danger', '#e74c3c');
    const textPrimary = getCssVar('--text-primary', '#e6e9ef');
    const textSecondary = getCssVar('--text-secondary', '#8aa0c4');

    const formatScore = (val) => {
      if (!Number.isFinite(val)) return '-';
      let str = val.toFixed(5);
      str = str.replace(/0+$/, '');
      if (str.endsWith('.')) str = str.slice(0, -1);
      return str;
    };

    if (!points || !points.length) {
      risquesScoreChartInstance.setOption({
        backgroundColor: 'transparent',
        title: {
          text: 'Indice de criticité (Initial vs Résiduel)',
          left: 'center',
          textStyle: { color: textPrimary, fontSize: 18, fontWeight: 'bold' }
        },
        grid: { left: 60, right: 30, top: 80, bottom: 80 },
        xAxis: {
          type: 'category',
          data: [],
          axisLabel: { color: textSecondary },
          axisLine: { lineStyle: { color: textSecondary } },
          splitLine: { show: false }
        },
        yAxis: {
          type: 'value',
          axisLabel: { color: textSecondary },
          axisLine: { lineStyle: { color: textSecondary } },
          splitLine: { lineStyle: { color: 'rgba(148,163,184,0.2)' } },
          name: 'Indice',
          nameTextStyle: { color: textSecondary },
          min: 0
        },
        legend: {
          data: ['Initial', 'Résiduel'],
          top: 40,
          textStyle: { color: textSecondary }
        },
        series: [],
        graphic: [
          {
            type: 'text',
            left: 'center',
            top: 'middle',
            style: {
              text: 'Ajoutez des risques pour visualiser les indices.',
              fill: textSecondary,
              fontSize: 14
            }
          }
        ]
      }, true);
      return;
    }

    const categories = points.map(point => point.name);
    const initialValues = points.map(point => point.initial);
    const residualValues = points.map(point => point.residual);
    const shouldRotate = categories.some(name => (name || '').length > 12);

    risquesScoreChartInstance.setOption({
      backgroundColor: 'transparent',
      title: {
        text: 'Indice de criticité (Initial vs Résiduel)',
        left: 'center',
        textStyle: { color: textPrimary, fontSize: 18, fontWeight: 'bold' }
      },
      tooltip: {
        trigger: 'axis',
        axisPointer: { type: 'shadow' },
        formatter: (params) => {
          if (!Array.isArray(params)) return '';
          const lines = [`<strong>${params[0].axisValue}</strong>`];
          params.forEach(item => {
            lines.push(`${item.marker || ''} ${item.seriesName}: ${formatScore(Number(item.data))}`);
          });
          return lines.join('<br/>');
        }
      },
      legend: {
        data: ['Initial', 'Résiduel'],
        top: 40,
        textStyle: { color: textSecondary }
      },
      grid: { left: 60, right: 30, top: 80, bottom: shouldRotate ? 110 : 80 },
      xAxis: {
        type: 'category',
        data: categories,
        axisLabel: {
          color: textSecondary,
          rotate: shouldRotate ? 20 : 0
        },
        axisLine: { lineStyle: { color: textSecondary } },
        splitLine: { show: false }
      },
      yAxis: {
        type: 'value',
        axisLabel: { color: textSecondary },
        axisLine: { lineStyle: { color: textSecondary } },
        splitLine: { lineStyle: { color: 'rgba(148,163,184,0.2)' } },
        name: 'Indice',
        nameTextStyle: { color: textSecondary },
        min: 0
      },
      series: [
        {
          name: 'Initial',
          type: 'bar',
          data: initialValues,
          itemStyle: { color: danger },
          emphasis: { focus: 'series' }
        },
        {
          name: 'Résiduel',
          type: 'bar',
          data: residualValues,
          itemStyle: { color: accent },
          emphasis: { focus: 'series' }
        }
      ],
      animationDuration: 300
    }, true);
  }

  function bindAtelier5Drag(residualSpread, lineData) {
    if (!risquesChartInstance) return;
    const zr = risquesChartInstance.getZr();
    const dom = risquesChartInstance.getDom();
    if (atelier5DragHandlers) {
      risquesChartInstance.off('mousedown', atelier5DragHandlers.mousedown);
      zr.off('mousemove', atelier5DragHandlers.mousemove);
      zr.off('mouseup', atelier5DragHandlers.mouseup);
      zr.off('globalout', atelier5DragHandlers.mouseup);
    }
    let dragging = null;
    const clamp = (val) => Math.max(0, Math.min(4, val));
    const mousedown = (params) => {
      if (!params || params.seriesIndex !== 2) return;
      if (!params.data || !params.data.meta) return;
      dragging = params.data;
      dom.style.cursor = 'grabbing';
    };
    const mousemove = (event) => {
      if (!dragging) return;
      const pos = [event.offsetX, event.offsetY];
      const dataPos = risquesChartInstance.convertFromPixel({ seriesIndex: 2 }, pos);
      if (!dataPos) return;
      const newValue = [
        Number(clamp(dataPos[0]).toFixed(2)),
        Number(clamp(dataPos[1]).toFixed(2))
      ];
      dragging.value = newValue;
      dragging.rawValue = newValue.slice();
      const metaKey = dragging.metaKey || (dragging.meta ? `risk::${dragging.meta.riskId || dragging.meta.rowIndex}::${dragging.meta.rowIndex}` : null);
      if (metaKey) {
        const line = lineData.find(item => (item.metaKey || item.name) === metaKey);
        if (line) {
          line.coords[1] = newValue.slice();
          line.to = newValue.slice();
        }
      }
      risquesChartInstance.setOption({
        series: [
          { data: lineData },
          {},
          { data: residualSpread }
        ]
      });
    };
    const endDrag = () => {
      if (!dragging) return;
      const meta = dragging.meta || {};
      const analysis = analyses[currentIndex];
      if (analysis && analysis.data && Array.isArray(analysis.data.actionsRisques)) {
        const row = analysis.data.actionsRisques[meta.rowIndex];
        if (row) {
          row.residualG = dragging.value[0];
          row.residualV = dragging.value[1];
          saveAnalyses();
          renderRisquesActions();
          renderRisquesChart();
        }
      }
      dom.style.cursor = '';
      dragging = null;
    };
    risquesChartInstance.on('mousedown', { seriesIndex: 2 }, mousedown);
    zr.on('mousemove', mousemove);
    zr.on('mouseup', endDrag);
    zr.on('globalout', endDrag);
    atelier5DragHandlers = { mousedown, mousemove, mouseup: endDrag };
  }

  function bindAtelier4Drag(scatterData) {
    if (!atelier4ChartInstance) return;
    const zr = atelier4ChartInstance.getZr();
    const dom = atelier4ChartInstance.getDom();
    if (atelier4DragHandlers) {
      atelier4ChartInstance.off('mousedown', atelier4DragHandlers.mousedown);
      zr.off('mousemove', atelier4DragHandlers.mousemove);
      zr.off('mouseup', atelier4DragHandlers.mouseup);
      zr.off('globalout', atelier4DragHandlers.mouseup);
    }
    let dragging = null;
    const clamp = (val) => Math.max(0, Math.min(4, val));
    const mousedown = (params) => {
      if (!params || !params.data || !params.data.meta || params.data.meta.type !== 'scenario') return;
      dragging = params.data;
      dom.style.cursor = 'grabbing';
    };
    const mousemove = (event) => {
      if (!dragging) return;
      const pos = [event.offsetX, event.offsetY];
      const dataPos = atelier4ChartInstance.convertFromPixel({ seriesIndex: 0 }, pos);
      if (!dataPos) return;
      const newValue = [
        Number(clamp(dataPos[0]).toFixed(2)),
        Number(clamp(dataPos[1]).toFixed(2))
      ];
      dragging.value = newValue;
      dragging.rawValue = newValue.slice();
      atelier4ChartInstance.setOption({
        series: [
          { data: scatterData }
        ]
      });
    };
    const endDrag = () => {
      if (!dragging) return;
      const meta = dragging.meta || {};
      if (meta.type === 'scenario') {
        const analysis = analyses[currentIndex];
        if (analysis && analysis.data && Array.isArray(analysis.data.so)) {
          const sIdx = analysis.data.so.findIndex((so, idx) => (so.id || `scenario-${idx}`) === meta.scenarioId);
          const targetScenario = sIdx >= 0 ? analysis.data.so[sIdx] : null;
          if (targetScenario && Array.isArray(targetScenario.risks) && targetScenario.risks[meta.riskIndex]) {
            const rk = targetScenario.risks[meta.riskIndex];
            rk.gravite = dragging.value[0];
            rk.vraisemblance = dragging.value[1];
            saveAnalyses();
            renderSO();
          }
        }
      }
      dom.style.cursor = '';
      dragging = null;
    };
    atelier4ChartInstance.on('mousedown', { seriesIndex: 0 }, mousedown);
    zr.on('mousemove', mousemove);
    zr.on('mouseup', endDrag);
    zr.on('globalout', endDrag);
    atelier4DragHandlers = { mousedown, mousemove, mouseup: endDrag };
  }

  // Render actions for risks: show each risk and allow adding actions and residual levels
  function renderRisquesActions() {
    const body = document.getElementById('risques-actions-body');
    if (!body) return;
    body.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) return;
    if (!Array.isArray(analysis.data.actionsRisques)) analysis.data.actionsRisques = [];
    // Gather risks from analysis list and operational scenarios for reference
    const riskMap = new Map();
    (analysis.data.risques || []).forEach(riskObj => {
      const label = riskObj.libelle || riskObj.titre || riskObj.indice || riskObj.id || '';
      if (!label) return;
      const id = riskObj.id || label;
      const vr = parseLevel(riskObj.vraisemblance, 1);
      const gr = parseLevel(riskObj.gravite, 1);
      const cur = riskMap.get(id) || { id, name: label, vraisemblance: vr, gravite: gr };
      cur.vraisemblance = Math.max(cur.vraisemblance, vr);
      cur.gravite = Math.max(cur.gravite, gr);
      riskMap.set(id, cur);
    });
    (analysis.data.so || []).forEach(so => {
      (so.risks || []).forEach(rk => {
        const label = rk.name || rk.id;
        if (!label) return;
        const id = rk.id || rk.name;
        const vr = parseLevel(rk.vraisemblance, 1);
        const gr = parseLevel(rk.gravite, 1);
        const cur = riskMap.get(id) || { id, name: label, vraisemblance: vr, gravite: gr };
        cur.vraisemblance = Math.max(cur.vraisemblance, vr);
        cur.gravite = Math.max(cur.gravite, gr);
        riskMap.set(id, cur);
      });
    });
    analysis.data.actionsRisques.forEach(row => {
      if (!row.riskId && row.riskName) {
        for (const [id, obj] of riskMap.entries()) {
          if (obj.name === row.riskName) { row.riskId = id; break; }
        }
        if (!row.riskId) row.riskId = row.riskName ? row.riskName.split(' ')[0] : '';
      }
      const risk = riskMap.get(row.riskId) || { id: row.riskId, name: row.riskName, vraisemblance: row.residualV || 1, gravite: row.residualG || 1 };
      const tr = document.createElement('tr');
      const tdId = document.createElement('td');
      if (row.manual) {
        const inpId = document.createElement('input');
        inpId.value = row.riskId || '';
        inpId.addEventListener('input', (e) => {
          row.riskId = e.target.value;
          saveAnalyses();
        });
        tdId.appendChild(inpId);
      } else {
        tdId.textContent = formatRiskIdentifier(row.riskId || '');
      }
      tr.appendChild(tdId);
      const tdName = document.createElement('td');
      if (row.manual) {
        const inpRisk = document.createElement('input');
        inpRisk.value = row.riskName || '';
        inpRisk.addEventListener('input', (e) => {
          row.riskName = e.target.value;
          saveAnalyses();
          renderRisquesActions();
        });
        tdName.appendChild(inpRisk);
      } else {
        tdName.textContent = row.riskName;
      }
      tr.appendChild(tdName);
      const tdVr = document.createElement('td');
      const vrDisplay = parseLevel(risk.vraisemblance, 1);
      tdVr.textContent = vrDisplay;
      tdVr.style.backgroundColor = levelColor(vrDisplay);
      tr.appendChild(tdVr);
      const tdGr = document.createElement('td');
      const grDisplay = parseLevel(risk.gravite, 1);
      tdGr.textContent = grDisplay;
      tdGr.style.backgroundColor = levelColor(grDisplay);
      tr.appendChild(tdGr);
      // Residual vraisemblance input
      const tdResVr = document.createElement('td');
      const inpResidualV = document.createElement('input');
      inpResidualV.type = 'number';
      inpResidualV.step = '0.1';
      inpResidualV.min = '0';
      inpResidualV.max = '4';
      const vrValue = row.residualV !== undefined ? row.residualV : risk.vraisemblance;
      inpResidualV.value = vrValue;
      inpResidualV.style.backgroundColor = levelColor(parseLevel(vrValue, risk.vraisemblance));
      inpResidualV.addEventListener('input', (e) => {
        const val = parseLevel(e.target.value, row.residualV !== undefined ? row.residualV : risk.vraisemblance);
        row.residualV = val;
        e.target.value = val;
        e.target.style.backgroundColor = levelColor(val);
        saveAnalyses();
        renderRisquesChart();
      });
      tdResVr.appendChild(inpResidualV);
      tr.appendChild(tdResVr);
      // Residual gravite input
      const tdResGr = document.createElement('td');
      const inpResidualG = document.createElement('input');
      inpResidualG.type = 'number';
      inpResidualG.step = '0.1';
      inpResidualG.min = '0';
      inpResidualG.max = '4';
      const grValue = row.residualG !== undefined ? row.residualG : risk.gravite;
      inpResidualG.value = grValue;
      inpResidualG.style.backgroundColor = levelColor(parseLevel(grValue, risk.gravite));
      inpResidualG.addEventListener('input', (e) => {
        const val = parseLevel(e.target.value, row.residualG !== undefined ? row.residualG : risk.gravite);
        row.residualG = val;
        e.target.value = val;
        e.target.style.backgroundColor = levelColor(val);
        saveAnalyses();
        renderRisquesChart();
      });
      tdResGr.appendChild(inpResidualG);
      tr.appendChild(tdResGr);
      // Actions cell: nested table for actions associated with this risk
      const tdActions = document.createElement('td');
      tdActions.className = 'assoc-cell';
      if (!Array.isArray(row.actions)) row.actions = [];
      const actTable = document.createElement('table');
      actTable.className = 'nested-table';
      const headerRow = document.createElement('tr');
      headerRow.innerHTML = '<th>Nom</th><th>Description</th><th>Responsable</th><th>Début</th><th>Fin</th><th></th>';
      actTable.appendChild(headerRow);
      row.actions.forEach((act, aIdx) => {
        const ar = document.createElement('tr');
        let tdA = document.createElement('td');
        const inpName = document.createElement('input');
        inpName.value = act.name || '';
        inpName.addEventListener('input', (e) => {
          act.name = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpName);
        ar.appendChild(tdA);
        // Description
        tdA = document.createElement('td');
        const inpDesc = document.createElement('textarea');
        inpDesc.rows = 2;
        inpDesc.value = act.description || '';
        inpDesc.addEventListener('input', (e) => {
          act.description = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpDesc);
        ar.appendChild(tdA);
        // Responsable
        tdA = document.createElement('td');
        const inpResp = document.createElement('input');
        inpResp.value = act.responsable || '';
        inpResp.addEventListener('input', (e) => {
          act.responsable = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpResp);
        ar.appendChild(tdA);
        // Start
        tdA = document.createElement('td');
        const inpStart = document.createElement('input');
        inpStart.type = 'date';
        inpStart.value = act.start || '';
        inpStart.addEventListener('change', (e) => {
          act.start = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpStart);
        ar.appendChild(tdA);
        // End
        tdA = document.createElement('td');
        const inpEnd = document.createElement('input');
        inpEnd.type = 'date';
        inpEnd.value = act.end || '';
        inpEnd.addEventListener('change', (e) => {
          act.end = e.target.value;
          saveAnalyses();
          renderPlanActions();
        });
        tdA.appendChild(inpEnd);
        ar.appendChild(tdA);
        // Delete button
        tdA = document.createElement('td');
        const delBtn = document.createElement('button');
        delBtn.className = 'delete-action';
        delBtn.textContent = '×';
        delBtn.addEventListener('click', () => {
          if (!confirm('Supprimer cette action ?')) return;
          row.actions.splice(aIdx, 1);
          saveAnalyses();
          renderRisquesActions();
          renderPlanActions();
        });
        tdA.appendChild(delBtn);
        ar.appendChild(tdA);
        actTable.appendChild(ar);
      });
      // Row to add new or existing action
      const addRow = document.createElement('tr');
      const addTd = document.createElement('td');
      addTd.colSpan = 6;
      const addBtnA = document.createElement('button');
      addBtnA.className = 'add-assoc-btn';
      addBtnA.textContent = '+ Action';
      addBtnA.addEventListener('click', () => {
        row.actions.push({ name:'', description:'', responsable:'', start:'', end:'' });
        saveAnalyses();
        renderRisquesActions();
        renderPlanActions();
      });
      addTd.appendChild(addBtnA);
      const addExisting = document.createElement('button');
      addExisting.className = 'add-assoc-btn';
      addExisting.textContent = '+ Action existante';
      addExisting.addEventListener('click', () => {
        const allActs = [];
        analysis.data.actionsRisques.forEach(r => {
          if (r !== row && Array.isArray(r.actions)) {
            r.actions.forEach(a => {
              if (a && a.name) allActs.push(a);
            });
          }
        });
        const currentNames = row.actions.map(a => (a.name || '').trim());
        const available = allActs.filter(a => !currentNames.includes((a.name || '').trim()));
        if (available.length === 0) {
          alert('Aucune action existante disponible.');
          return;
        }
        const msg = 'Sélectionnez une action existante:\n' + available.map((a,i)=>`${i+1}. ${a.name}`).join('\n');
        const input = prompt(msg);
        if (input === null) return;
        const index = parseInt(input,10) - 1;
        if (!isNaN(index) && index >=0 && index < available.length) {
          row.actions.push(Object.assign({}, available[index]));
          saveAnalyses();
          renderRisquesActions();
          renderPlanActions();
        }
      });
      addTd.appendChild(addExisting);
      addRow.appendChild(addTd);
      actTable.appendChild(addRow);
      tdActions.appendChild(actTable);
      tr.appendChild(tdActions);
      const tdDel = document.createElement('td');
      const delRow = document.createElement('button');
      delRow.className = 'delete-item';
      delRow.textContent = '×';
      delRow.addEventListener('click', () => {
        if (!confirm('Supprimer ce risque ?')) return;
        const idx = analysis.data.actionsRisques.indexOf(row);
        if (idx >= 0) analysis.data.actionsRisques.splice(idx, 1);
        saveAnalyses();
        renderRisquesActions();
        renderPlanActions();
      });
      tdDel.appendChild(delRow);
      tr.appendChild(tdDel);
      body.appendChild(tr);
    });
    addDataTableResizers('risques-actions-table');
    renderRisquesChart();
  }

  // Aggregate all actions and render plan table + gantt chart
  function renderPlanActions() {
    const body = document.getElementById('plan-actions-body');
    const chartEl = document.getElementById('gantt-chart');
    if (!body || !chartEl) return;
    body.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) return;
    const actions = [];
    // Gap actions
    (analysis.data.actionsGap || []).forEach(entry => {
      const source = (analysis.data.gap || []).find(req => req.id === entry.sourceId);
      const sourceName = source ? (source.titre || source.domaine || 'Exigence') : 'Exigence';
      (entry.actions || []).forEach(act => {
        actions.push({
          name: act.name,
          source: 'GAP: ' + sourceName,
          description: act.description || '',
          responsable: act.responsable || '',
          start: act.start || '',
          end: act.end || ''
        });
      });
    });
    // Support actions
    (analysis.data.actionsSupports || []).forEach(row => {
      const sup = row.supportName || 'Support';
      const vul = row.vulnName ? ` - ${row.vulnName}` : '';
      (row.actions || []).forEach(act => {
        actions.push({
          name: act.name,
          source: 'Support: ' + sup + vul,
          description: act.description || '',
          responsable: act.responsable || '',
          start: act.start || '',
          end: act.end || ''
        });
      });
    });
    // Party actions
    (analysis.data.actionsParties || []).forEach(row => {
      const pp = (analysis.data.ppc || []).find(p => p.id === row.ppId);
      const srcName = pp ? (pp.nom || pp.name || 'Partie') : 'Partie';
      (row.actions || []).forEach(act => {
        actions.push({
          name: act.name,
          source: 'Partie: ' + srcName,
          description: act.description || '',
          responsable: act.responsable || '',
          start: act.start || '',
          end: act.end || ''
        });
      });
    });
    // Risk actions
    (analysis.data.actionsRisques || []).forEach(row => {
      const srcName = row.riskName;
      (row.actions || []).forEach(act => {
        actions.push({
          name: act.name,
          source: 'Risque: ' + srcName,
          description: act.description || '',
          responsable: act.responsable || '',
          start: act.start || '',
          end: act.end || ''
        });
      });
    });
    // Render table rows
    actions.forEach(act => {
      const tr = document.createElement('tr');
      ['name','source','description','responsable','start','end'].forEach(key => {
        const td = document.createElement('td');
        td.textContent = act[key] || '';
        tr.appendChild(td);
      });
      body.appendChild(tr);
    });
    addDataTableResizers('plan-actions-table');
    // Draw gantt chart
    if (typeof drawGanttChartSVG === 'function') {
      drawGanttChartSVG(chartEl, actions);
    }
  }

  // ----- Chart drawing functions (simple bar and radar charts)
  function clearCanvas(canvas) {
    const ctx = canvas.getContext('2d');
    const bg = getComputedStyle(canvas).getPropertyValue('background-color') || '#fff';
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = bg;
    ctx.fillRect(0, 0, canvas.width, canvas.height);
  }

  function drawBarChart(canvas, labels, data, colors) {
    const ctx = canvas.getContext('2d');
    clearCanvas(canvas);
    if (!labels || labels.length === 0) return;
    const width = canvas.width;
    const height = canvas.height;
    const margin = 40;
    const barAreaWidth = width - margin * 2;
    const barWidth = barAreaWidth / labels.length;
    const maxVal = Math.max(...data, 1);
    const style = getComputedStyle(document.documentElement);
    const textPrimary = style.getPropertyValue('--text-primary') || '#e6e9ef';
    const textSecondary = style.getPropertyValue('--text-secondary') || '#8aa0c4';
    const axisColor = textSecondary || 'rgba(200,200,200,0.5)';
    // Draw axes
    ctx.strokeStyle = axisColor;
    ctx.globalAlpha = 0.4;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(margin, margin);
    ctx.lineTo(margin, height - margin);
    ctx.lineTo(width - margin, height - margin);
    ctx.stroke();
    ctx.globalAlpha = 1;
    ctx.font = '12px system-ui';
    ctx.textAlign = 'center';
    // Draw bars
    data.forEach((value, i) => {
      const barHeight = (value / maxVal) * (height - margin * 2);
      const x = margin + i * barWidth + barWidth * 0.2;
      const y = height - margin - barHeight;
      const w = barWidth * 0.6;
      const color = colors && colors[i] ? colors[i] : '#4da3ff';
      ctx.fillStyle = color;
      ctx.fillRect(x, y, w, barHeight);
      // Value label
      ctx.fillStyle = textPrimary;
      ctx.fillText(value, x + w / 2, y - 4);
      // Category label
      ctx.fillStyle = textSecondary;
      ctx.save();
      ctx.translate(x + w / 2, height - margin + 14);
      ctx.rotate(-Math.PI / 4);
      ctx.fillText(labels[i], 0, 0);
      ctx.restore();
    });
  }

  function drawGroupedBarChart(canvas, labels, dataA, dataB, colors) {
    const ctx = canvas.getContext('2d');
    clearCanvas(canvas);
    if (!labels || labels.length === 0) return;
    const width = canvas.width;
    const height = canvas.height;
    const margin = 40;
    const groupWidth = (width - margin * 2) / labels.length;
    const barWidth = groupWidth * 0.35;
    const maxVal = Math.max(...dataA, ...dataB, 1);
    const style = getComputedStyle(document.documentElement);
    const textPrimary = style.getPropertyValue('--text-primary') || '#e6e9ef';
    const textSecondary = style.getPropertyValue('--text-secondary') || '#8aa0c4';
    const axisColor = textSecondary || 'rgba(200,200,200,0.5)';
    // axes
    ctx.strokeStyle = axisColor;
    ctx.globalAlpha = 0.4;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(margin, margin);
    ctx.lineTo(margin, height - margin);
    ctx.lineTo(width - margin, height - margin);
    ctx.stroke();
    ctx.globalAlpha = 1;
    ctx.font = '12px system-ui';
    ctx.textAlign = 'center';
    // legend
    ctx.fillStyle = textSecondary;
    ctx.fillRect(margin, margin - 30, 12, 12);
    ctx.fillText('Initial', margin + 26, margin - 20);
    ctx.fillStyle = textSecondary;
    ctx.globalAlpha = 0.4;
    ctx.fillRect(margin + 80, margin - 30, 12, 12);
    ctx.globalAlpha = 1;
    ctx.fillText('Résiduel', margin + 106, margin - 20);
    // bars
    labels.forEach((lab, i) => {
      const baseX = margin + i * groupWidth;
      const color = colors && colors[i] ? colors[i] : '#888';
      const h1 = (dataA[i] / maxVal) * (height - margin * 2);
      const h2 = (dataB[i] / maxVal) * (height - margin * 2);
      const y1 = height - margin - h1;
      const y2 = height - margin - h2;
      const x1 = baseX + groupWidth * 0.1;
      const x2 = baseX + groupWidth * 0.55;
      ctx.fillStyle = color;
      ctx.fillRect(x1, y1, barWidth, h1);
      ctx.globalAlpha = 0.4;
      ctx.fillRect(x2, y2, barWidth, h2);
      ctx.globalAlpha = 1;
      ctx.fillStyle = textPrimary;
      ctx.fillText(dataA[i], x1 + barWidth / 2, y1 - 4);
      ctx.fillText(dataB[i], x2 + barWidth / 2, y2 - 4);
      ctx.fillStyle = textSecondary;
      ctx.save();
      ctx.translate(baseX + groupWidth / 2, height - margin + 14);
      ctx.rotate(-Math.PI / 4);
      ctx.fillText(lab, 0, 0);
      ctx.restore();
    });
  }

  function drawRadarChart(canvas, labels, dataset) {
    const ctx = canvas.getContext('2d');
    clearCanvas(canvas);
    if (!labels || labels.length === 0) return;
    const width = canvas.width;
    const height = canvas.height;
    const centerX = width / 2;
    const centerY = height / 2;
    const margin = 40;
    const radius = Math.min(width, height) / 2 - margin;
    const n = labels.length;
    // Draw concentric circles and axes
    ctx.strokeStyle = 'rgba(200,200,200,0.2)';
    ctx.lineWidth = 1;
    const steps = 5;
    for (let s = 1; s <= steps; s++) {
      const r = (s / steps) * radius;
      ctx.beginPath();
      for (let i = 0; i < n; i++) {
        const angle = -Math.PI / 2 + (2 * Math.PI * i) / n;
        const x = centerX + r * Math.cos(angle);
        const y = centerY + r * Math.sin(angle);
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      }
      ctx.closePath();
      ctx.stroke();
    }
    // Draw axes lines and labels
    labels.forEach((label, i) => {
      const angle = -Math.PI / 2 + (2 * Math.PI * i) / n;
      const x = centerX + radius * Math.cos(angle);
      const y = centerY + radius * Math.sin(angle);
      ctx.strokeStyle = 'rgba(200,200,200,0.3)';
      ctx.beginPath();
      ctx.moveTo(centerX, centerY);
      ctx.lineTo(x, y);
      ctx.stroke();
      // Label at end
      ctx.fillStyle = 'var(--text-secondary)';
      ctx.font = '12px sans-serif';
      const offsetX = 10 * Math.cos(angle);
      const offsetY = 10 * Math.sin(angle);
      ctx.textAlign = angle > Math.PI / 2 || angle < -Math.PI / 2 ? 'right' : 'left';
      ctx.fillText(label, x + offsetX, y + offsetY);
    });
    // Draw dataset polygon
    if (dataset && dataset.data && dataset.data.length === n) {
      const values = dataset.data;
      const color = dataset.color || 'rgba(77,163,255,0.5)';
      // Determine max for scaling (assuming data in 0-10)
      const maxVal = Math.max(...values, 1);
      ctx.beginPath();
      values.forEach((val, i) => {
        const r = (val / maxVal) * radius;
        const angle = -Math.PI / 2 + (2 * Math.PI * i) / n;
        const x = centerX + r * Math.cos(angle);
        const y = centerY + r * Math.sin(angle);
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      });
      ctx.closePath();
      // Fill polygon
      ctx.fillStyle = color;
      ctx.fill();
      // Outline
      ctx.strokeStyle = 'rgba(77,163,255,0.8)';
      ctx.lineWidth = 2;
      ctx.stroke();
    }
  }

  // ----- Chart updates per atelier
  function updateAtelier1Chart() {
    updateAtelier1Graph();
    renderVulnChart();
  }

  function updateAtelier2Chart() {
    // Draw a custom network diagram representing all SROV couples.  Each
    // source appears on the left, each objective on the right, and edges
    // are coloured according to the pertinence (motivation × ressources).
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) return;
    const srov = analysis.data.srov || [];
    const canvas = document.getElementById('atelier2-chart');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    const w = canvas.width;
    const h = canvas.height;
    // Clear canvas
    ctx.clearRect(0, 0, w, h);
    // Build sets of unique sources and objectives preserving insertion order
    const srcOrder = [];
    const objOrder = [];
    const srcMap = new Map();
    const objMap = new Map();
    srov.forEach(item => {
      const sName = (item.source || '').trim() || 'Source';
      if (!srcMap.has(sName)) {
        srcMap.set(sName, srcOrder.length);
        srcOrder.push(sName);
      }
      const oName = (item.objectif || '').trim() || 'Objectif';
      if (!objMap.has(oName)) {
        objMap.set(oName, objOrder.length);
        objOrder.push(oName);
      }
    });
    const nSrc = srcOrder.length;
    const nObj = objOrder.length;
    // If no entries, nothing to draw
    if (nSrc === 0 && nObj === 0) return;
    // Determine positions
    const srcX = Math.max(60, w * 0.25);
    const objX = Math.min(w - 60, w * 0.75);
    const srcSpacing = nSrc > 0 ? h / (nSrc + 1) : 0;
    const objSpacing = nObj > 0 ? h / (nObj + 1) : 0;
    const positions = {};
    srcOrder.forEach((sName, idx) => {
      positions['src:' + sName] = { x: srcX, y: (idx + 1) * srcSpacing };
    });
    objOrder.forEach((oName, idx) => {
      positions['obj:' + oName] = { x: objX, y: (idx + 1) * objSpacing };
    });
    // Compute max priority per objective
    const objMaxPriority = {};
    srov.forEach(item => {
      const oName = (item.objectif || '').trim() || 'Objectif';
      const prio = parseInt(item.priorite, 10) || 1;
      if (!objMaxPriority[oName] || prio > objMaxPriority[oName]) {
        objMaxPriority[oName] = prio;
      }
    });
    // Helper functions for colours and buckets
    const levelColor = (lvl) => {
      switch (parseInt(lvl, 10)) {
        case 1: return '#2a9d8f';
        case 2: return '#e9c46a';
        case 3: return '#f4a261';
        case 4: return '#e63946';
        default: return '#9aa0a6';
      }
    };
    const pertinenceBucket = (p) => {
      if (p >= 13) return 4;
      if (p >= 9) return 3;
      if (p >= 5) return 2;
      return 1;
    };
    // Draw edges
    srov.forEach(item => {
      const sName = (item.source || '').trim() || 'Source';
      const oName = (item.objectif || '').trim() || 'Objectif';
      const sPos = positions['src:' + sName];
      const oPos = positions['obj:' + oName];
      if (!sPos || !oPos) return;
      const m = parseInt(item.motivation, 10) || 1;
      const r = parseInt(item.ressources, 10) || 1;
      const p = m * r;
      const bucket = pertinenceBucket(p);
      const color = levelColor(bucket);
      const retenue = (typeof item.retenue === 'boolean') ? item.retenue : true;
      // Start and end positions (shortened to avoid overlapping node shapes)
      const rectWidth = 180;
      const diamondWidth = 120;
      const startX = sPos.x + rectWidth / 2;
      const startY = sPos.y;
      const endX = oPos.x - diamondWidth / 2;
      const endY = oPos.y;
      // Draw line
      ctx.strokeStyle = color;
      ctx.lineWidth = 2;
      ctx.setLineDash(retenue ? [] : [6, 4]);
      ctx.beginPath();
      ctx.moveTo(startX, startY);
      ctx.lineTo(endX, endY);
      ctx.stroke();
      // Draw arrow head at end
      const angle = Math.atan2(endY - startY, endX - startX);
      const arrowLen = 8;
      const arrowAngle = Math.PI / 6;
      ctx.beginPath();
      ctx.moveTo(endX, endY);
      ctx.lineTo(endX - arrowLen * Math.cos(angle - arrowAngle), endY - arrowLen * Math.sin(angle - arrowAngle));
      ctx.lineTo(endX - arrowLen * Math.cos(angle + arrowAngle), endY - arrowLen * Math.sin(angle + arrowAngle));
      ctx.closePath();
      ctx.fillStyle = color;
      ctx.fill();
      // Draw label near the middle of the edge
      const midX = (startX + endX) / 2;
      const midY = (startY + endY) / 2;
      ctx.textAlign = 'center';
      ctx.font = '11px sans-serif';
      // Primary label line: motivation, ressources, pertinence
      ctx.fillStyle = '#ffffff';
      const label1 = `M:${m}   R:${r}   P=${p} (niv ${bucket})`;
      ctx.fillText(label1, midX, midY - 4);
      // Secondary label for exclusion
      if (!retenue) {
        const just = (item.justification || '').trim();
        const excl = just ? `EXCLU: ${just}` : 'EXCLU';
        ctx.fillStyle = 'rgba(255,255,255,0.7)';
        ctx.fillText(excl, midX, midY + 10);
      }
    });
    ctx.setLineDash([]);
    // Draw source nodes (rectangles)
    srcOrder.forEach((sName) => {
      const pos = positions['src:' + sName];
      const rectW = 180;
      const rectH = 60;
      const x = pos.x - rectW / 2;
      const y = pos.y - rectH / 2;
      ctx.fillStyle = '#8ecae6';
      ctx.strokeStyle = '#1d4ed8';
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.rect(x, y, rectW, rectH);
      ctx.fill();
      ctx.stroke();
      // Labels
      ctx.fillStyle = '#1f2937';
      ctx.font = 'bold 10px sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText('SOURCE DE RISQUE', pos.x, pos.y - 8);
      ctx.font = '12px sans-serif';
      ctx.fillText(sName, pos.x, pos.y + 10);
    });
    // Draw objective nodes (diamonds)
    objOrder.forEach((oName) => {
      const pos = positions['obj:' + oName];
      const prio = objMaxPriority[oName] || 1;
      const borderColor = levelColor(prio);
      // Diamond dimensions
      const hHalf = 30;
      const wHalf = 50;
      ctx.fillStyle = '#dfe7fd';
      ctx.strokeStyle = borderColor;
      ctx.lineWidth = 3;
      ctx.beginPath();
      ctx.moveTo(pos.x, pos.y - hHalf);
      ctx.lineTo(pos.x + wHalf, pos.y);
      ctx.lineTo(pos.x, pos.y + hHalf);
      ctx.lineTo(pos.x - wHalf, pos.y);
      ctx.closePath();
      ctx.fill();
      ctx.stroke();
      // Labels
      ctx.fillStyle = '#111827';
      ctx.font = 'bold 10px sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(`OBJECTIF VISÉ  P:${prio}`, pos.x, pos.y - 10);
      ctx.font = '12px sans-serif';
      ctx.fillText(oName, pos.x, pos.y + 12);
    });
  } 

  function ensureStakeholderCanvas() {
    const container = document.getElementById('atelier3-stakeholder-chart');
    if (!container) return null;

    if (!stakeholderCanvas || stakeholderCanvas.parentElement !== container) {
      container.innerHTML = '';
      stakeholderCanvas = document.createElement('canvas');
      stakeholderCanvas.id = 'atelier3-stakeholder-canvas';
      stakeholderCanvas.setAttribute('role', 'presentation');
      stakeholderCanvas.setAttribute('aria-hidden', 'true');
      stakeholderCanvas.style.width = '100%';
      stakeholderCanvas.style.height = '100%';
      container.appendChild(stakeholderCanvas);
      stakeholderCtx = stakeholderCanvas.getContext('2d');
      stakeholderPoints = [];
      if (stakeholderTooltipEl) {
        stakeholderTooltipEl.remove();
        stakeholderTooltipEl = null;
      }
    }

    if (!stakeholderTooltipEl) {
      stakeholderTooltipEl = document.createElement('div');
      stakeholderTooltipEl.id = 'stakeholder-tooltip';
      stakeholderTooltipEl.className = 'stakeholder-tooltip';
      stakeholderTooltipEl.setAttribute('aria-hidden', 'true');
      stakeholderTooltipEl.style.opacity = '0';
      container.appendChild(stakeholderTooltipEl);
    }

    if (stakeholderCanvas && !stakeholderCanvas.dataset.eventsBound) {
      stakeholderCanvas.addEventListener('mousemove', handleStakeholderHover);
      stakeholderCanvas.addEventListener('mouseleave', hideStakeholderTooltip);
      stakeholderCanvas.dataset.eventsBound = 'true';
    }

    if (!stakeholderCtx) return null;

    const ratio = window.devicePixelRatio || 1;
    const width = container.clientWidth || 600;
    const height = container.clientHeight || 520;
    const targetWidth = Math.round(width * ratio);
    const targetHeight = Math.round(height * ratio);
    if (stakeholderCanvas.width !== targetWidth || stakeholderCanvas.height !== targetHeight) {
      stakeholderCanvas.width = targetWidth;
      stakeholderCanvas.height = targetHeight;
    }
    stakeholderCanvas.style.width = width + 'px';
    stakeholderCanvas.style.height = height + 'px';
    stakeholderCtx.setTransform(1, 0, 0, 1, 0, 0);
    stakeholderCtx.scale(ratio, ratio);

    if (!stakeholderChartResizeBound) {
      stakeholderChartResizeBound = true;
      window.addEventListener('resize', () => {
        requestAnimationFrame(() => updateAtelier3Chart());
      });
    }

    return { ctx: stakeholderCtx, canvas: stakeholderCanvas, width, height, container };
  }

  function stakeholderSizeByExposition(value) {
    const expo = Number.isFinite(value) ? value : 0;
    if (expo < 3) return 8;
    if (expo < 7) return 14;
    if (expo < 10) return 20;
    return 26;
  }

  function stakeholderColorByFiabilite(value) {
    const fiab = Number.isFinite(value) ? value : 0;
    if (fiab < 4) return '#ef4444';
    if (fiab < 6) return '#f97316';
    if (fiab < 8) return '#facc15';
    return '#22c55e';
  }

  function stakeholderRadius(indice) {
    const val = Number.isFinite(indice) ? Math.max(0, indice) : 0;
    const clamped = Math.min(val, 5);
    return Math.max(0.25, 2.5 - clamped * 0.6);
  }

  function stakeholderTypeInfo(categorie) {
    const raw = (categorie || '').toString().trim().toLowerCase();
    switch (raw) {
      case 'partenaire':
      case 'partenaires':
        return { label: 'Partenaire', angleDeg: 30 };
      case 'beneficiaire':
      case 'bénéficiaire':
      case 'beneficiaires':
      case 'bénéficiaires':
      case 'client':
      case 'clients':
        return { label: 'Bénéficiaire', angleDeg: 150 };
      case 'prestataire':
      case 'prestataires':
      case 'fournisseur':
      case 'fournisseurs':
        return { label: 'Prestataire', angleDeg: 225 };
      case 'interne':
      case 'internes':
      case 'autorite':
      case 'autorité':
      case 'autre':
      case 'autres':
        return { label: 'Interne / Autre', angleDeg: 315 };
      default:
        if (raw) {
          return { label: categorie, angleDeg: 315 };
        }
        return { label: 'Interne / Autre', angleDeg: 315 };
    }
  }

  function stakeholderPolarToXY(radius, angleDeg, jitterSeed) {
    const angleRad = angleDeg * Math.PI / 180;
    const j = (Math.sin((jitterSeed + 1) * 12.9898) * 43758.5453) % 1;
    const jitter = (j - 0.5) * 0.18;
    const r = radius + jitter;
    return [r * Math.cos(angleRad), r * Math.sin(angleRad)];
  }

  function hideStakeholderTooltip() {
    if (stakeholderTooltipEl) {
      stakeholderTooltipEl.style.opacity = '0';
      stakeholderTooltipEl.setAttribute('aria-hidden', 'true');
    }
  }

  function handleStakeholderHover(evt) {
    if (!stakeholderCanvas || !stakeholderTooltipEl) return;
    const rect = stakeholderCanvas.getBoundingClientRect();
    const x = evt.clientX - rect.left;
    const y = evt.clientY - rect.top;
    let found = null;
    for (const point of stakeholderPoints) {
      const dx = x - point.x;
      const dy = y - point.y;
      if (Math.hypot(dx, dy) <= point.hitRadius) {
        found = point;
        break;
      }
    }
    if (!found) {
      hideStakeholderTooltip();
      return;
    }
    stakeholderTooltipEl.innerHTML = [
      `<strong>${found.meta.nom}</strong>`,
      `Type : ${found.meta.type}`,
      `Zone : ${found.meta.zone && found.meta.zone.label ? found.meta.zone.label : found.meta.zone}`,
      `Exposition : ${found.meta.exposition.toFixed(2)}`,
      `Fiabilité : ${found.meta.fiabilite.toFixed(2)}`,
      `Indice de menace : ${found.meta.indice.toFixed(2)}`
    ].join('<br/>');
    stakeholderTooltipEl.style.left = `${x}px`;
    stakeholderTooltipEl.style.top = `${y}px`;
    stakeholderTooltipEl.style.opacity = '1';
    stakeholderTooltipEl.setAttribute('aria-hidden', 'false');
  }

  function renderCartoRadar(ppc) {
    const canvasInfo = ensureStakeholderCanvas();
    const wrapper = document.getElementById('atelier3-radar-wrapper');
    if (!canvasInfo) {
      if (wrapper) wrapper.classList.toggle('empty', !(ppc && ppc.length));
      return;
    }

    const { ctx, width, height } = canvasInfo;
    hideStakeholderTooltip();

    const records = (ppc || []).map((item, index) => {
      const dep = parseInt(item.dependance, 10) || 1;
      const pen = parseInt(item.penetration, 10) || 1;
      const mat = parseInt(item.maturite, 10) || 1;
      const conf = parseInt(item.confiance, 10) || 1;

      const expoCandidate = (item.exposition !== undefined && item.exposition !== null && item.exposition !== '')
        ? Number(item.exposition)
        : NaN;
      const fiabiliteCandidate = (item.fiabilite !== undefined && item.fiabilite !== null && item.fiabilite !== '')
        ? Number(item.fiabilite)
        : NaN;
      const indiceCandidate = (item.indiceMenace !== undefined && item.indiceMenace !== null && item.indiceMenace !== '')
        ? Number(item.indiceMenace)
        : NaN;

      const exposition = Number.isFinite(expoCandidate) ? expoCandidate : (dep * pen);
      const fiabilite = Number.isFinite(fiabiliteCandidate) ? fiabiliteCandidate : ((mat || 1) * (conf || 1));
      const indice = Number.isFinite(indiceCandidate) ? indiceCandidate : (fiabilite ? exposition / fiabilite : 0);

      const zone = threatZoneMeta(indice);
      const typeInfo = stakeholderTypeInfo(item.categorie);
      const radius = stakeholderRadius(indice);
      const coords = stakeholderPolarToXY(radius, typeInfo.angleDeg, index + exposition + fiabilite);
      const displayName = (item.nom && item.nom.trim()) ? item.nom.trim() : 'Partie prenante';

      item.exposition = exposition;
      item.fiabilite = fiabilite;
      item.indiceMenace = indice;
      item.zoneMenace = zone.key;
      item.angle = typeInfo.angleDeg;
      item.rayon = radius;

      return {
        nom: displayName,
        type: typeInfo.label,
        coords,
        exposition,
        fiabilite,
        indice,
        zone,
        color: stakeholderColorByFiabilite(fiabilite),
        size: stakeholderSizeByExposition(exposition)
      };
    });

    ctx.clearRect(0, 0, width, height);
    ctx.fillStyle = '#0b1324';
    ctx.fillRect(0, 0, width, height);

    const centerX = width / 2;
    const centerY = height / 2 + 20;
    const baseRadius = Math.min(width, height) * 0.38;
    const radiusScale = baseRadius / 2.5;

    const zoneCircles = [
      { radius: 2.5, stroke: '#38bdf8', fill: 'rgba(56,189,248,0.06)' },
      { radius: stakeholderRadius(1), stroke: '#22c55e', fill: 'rgba(34,197,94,0.08)' },
      { radius: stakeholderRadius(2), stroke: '#f59e0b', fill: 'rgba(245,158,11,0.12)' },
      { radius: stakeholderRadius(3), stroke: '#ef4444', fill: 'rgba(239,68,68,0.16)' }
    ];

    zoneCircles.forEach(cfg => {
      ctx.beginPath();
      ctx.fillStyle = cfg.fill;
      ctx.strokeStyle = cfg.stroke;
      ctx.lineWidth = 2;
      ctx.arc(centerX, centerY, cfg.radius * radiusScale, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();
    });

    ctx.strokeStyle = '#475569';
    ctx.lineWidth = 1.2;
    ctx.beginPath();
    ctx.moveTo(centerX - baseRadius - 20, centerY);
    ctx.lineTo(centerX + baseRadius + 20, centerY);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(centerX, centerY - baseRadius - 20);
    ctx.lineTo(centerX, centerY + baseRadius + 20);
    ctx.stroke();

    ctx.fillStyle = '#e5e7eb';
    ctx.font = '20px "Segoe UI", system-ui, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Cartographie des parties prenantes', centerX, 36);

    function drawBadge(text, x, y, align) {
      ctx.save();
      ctx.font = '14px "Segoe UI", system-ui, sans-serif';
      const paddingX = 16;
      const paddingY = 10;
      const metrics = ctx.measureText(text);
      const textWidth = metrics.width;
      const boxWidth = textWidth + paddingX * 2;
      const boxHeight = 34;
      let left = x;
      if (align === 'center') {
        left = x - boxWidth / 2;
      } else if (align === 'right') {
        left = x - boxWidth;
      }
      ctx.fillStyle = '#11b3ad';
      ctx.strokeStyle = '#0ea5a4';
      ctx.lineWidth = 1;
      const radius = 16;
      const top = y;
      ctx.beginPath();
      ctx.moveTo(left + radius, top);
      ctx.lineTo(left + boxWidth - radius, top);
      ctx.quadraticCurveTo(left + boxWidth, top, left + boxWidth, top + radius);
      ctx.lineTo(left + boxWidth, top + boxHeight - radius);
      ctx.quadraticCurveTo(left + boxWidth, top + boxHeight, left + boxWidth - radius, top + boxHeight);
      ctx.lineTo(left + radius, top + boxHeight);
      ctx.quadraticCurveTo(left, top + boxHeight, left, top + boxHeight - radius);
      ctx.lineTo(left, top + radius);
      ctx.quadraticCurveTo(left, top, left + radius, top);
      ctx.closePath();
      ctx.fill();
      ctx.stroke();
      ctx.fillStyle = '#e6fffb';
      ctx.textAlign = 'left';
      ctx.textBaseline = 'middle';
      ctx.fillText(text, left + paddingX, top + boxHeight / 2);
      ctx.restore();
    }

    drawBadge('BÉNÉFICIAIRES', centerX - baseRadius - 10, centerY - baseRadius - 70, 'left');
    drawBadge('PARTENAIRES', centerX + baseRadius + 10, centerY - baseRadius - 70, 'right');
    drawBadge('PRESTATAIRES', centerX - baseRadius - 10, centerY + baseRadius + 30, 'left');
    drawBadge('INTERNE / AUTRES', centerX + baseRadius + 10, centerY + baseRadius + 30, 'right');

    const zoneLegend = [
      { color: '#38bdf8', label: 'Zone de confiance' },
      { color: '#22c55e', label: 'Zone de veille' },
      { color: '#f59e0b', label: 'Zone de contrôle' },
      { color: '#ef4444', label: 'Zone de danger' }
    ];
    const legendSpacing = 150;
    const legendStart = centerX - ((zoneLegend.length - 1) * legendSpacing) / 2;
    ctx.font = '13px "Segoe UI", system-ui, sans-serif';
    ctx.textBaseline = 'middle';
    zoneLegend.forEach((entry, idx) => {
      const lx = legendStart + idx * legendSpacing;
      const ly = height - 80;
      ctx.beginPath();
      ctx.strokeStyle = entry.color;
      ctx.fillStyle = '#0b1324';
      ctx.lineWidth = 3;
      ctx.arc(lx, ly, 10, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();
      ctx.fillStyle = '#e5e7eb';
      ctx.textAlign = 'left';
      ctx.fillText(entry.label, lx + 16, ly);
    });

    ctx.font = '12px "Segoe UI", system-ui, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillStyle = '#cbd5e1';
    ctx.fillText('EXPOSITION = Dépendance × Pénétration    |    FIABILITÉ = Maturité × Confiance', centerX, height - 36);

    stakeholderPoints = records.map(rec => {
      const normX = rec.coords[0];
      const normY = rec.coords[1];
      const px = centerX + (normX / 2.5) * baseRadius;
      const py = centerY - (normY / 2.5) * baseRadius;
      const size = rec.size;
      ctx.save();
      ctx.beginPath();
      ctx.fillStyle = rec.color;
      ctx.globalAlpha = 0.92;
      ctx.shadowBlur = 18;
      ctx.shadowColor = 'rgba(8,15,30,0.45)';
      ctx.arc(px, py, size / 2, 0, Math.PI * 2);
      ctx.fill();
      ctx.restore();
      ctx.strokeStyle = 'rgba(8,15,30,0.6)';
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.arc(px, py, size / 2, 0, Math.PI * 2);
      ctx.stroke();

      ctx.fillStyle = '#e5e7eb';
      ctx.font = '12px "Segoe UI", system-ui, sans-serif';
      ctx.textBaseline = 'middle';
      if (px >= centerX) {
        ctx.textAlign = 'left';
        ctx.fillText(rec.nom, px + size / 2 + 8, py);
      } else {
        ctx.textAlign = 'right';
        ctx.fillText(rec.nom, px - size / 2 - 8, py);
      }

      return {
        x: px,
        y: py,
        hitRadius: size / 2 + 10,
        meta: rec
      };
    });

    if (!records.length) {
      ctx.fillStyle = '#94a3b8';
      ctx.font = '14px "Segoe UI", system-ui, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText('Ajoutez des parties prenantes pour alimenter le diagramme.', centerX, centerY);
    }

    if (wrapper) wrapper.classList.toggle('empty', records.length === 0);
  }

  function renderStrategicGraph() {
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) return;
    const svg = document.getElementById('strategic-canvas');
    const gNodes = document.getElementById('strategic-nodes');
    const gLinks = document.getElementById('strategic-links');
    const tip = document.getElementById('strategic-tip');
    if (!svg || !gNodes || !gLinks || !tip) return;
    gNodes.innerHTML = '';
    gLinks.innerHTML = '';

    const strategies = analysis.data.strategies || [];
    const ppList = analysis.data.ppc || [];
    const evList = analysis.data.events || [];

    const nodeMap = new Map();
    const links = [];
    const SVG_NS = 'http://www.w3.org/2000/svg';
    const NODE_WIDTH = 190;
    const HALF_WIDTH = NODE_WIDTH / 2;
    const PADDING_X = 12;
    const PADDING_Y = 14;
    const LINE_HEIGHT = 16;
    const MAX_TEXT_WIDTH = NODE_WIDTH - PADDING_X * 2;
    const MAX_LINES = 4;
    const addNode = (key, type, label, details, severity) => {
      if (!key || nodeMap.has(key)) return nodeMap.get(key);
      const n = { id: key, type, label, details, severity };
      nodeMap.set(key, n);
      return n;
    };

    strategies.forEach(st => {
      if (!st.source || !st.objectif) return;
      const src = addNode('S:' + st.source, 'source', st.source);
      const obj = addNode('O:' + st.objectif, 'objective', st.objectif);
      if (src && obj) links.push({ from: src, to: obj });
      (st.chemins || []).forEach(ch => {
        const atk = addNode('C:' + ch, 'attack', ch);
        if (obj && atk) links.push({ from: obj, to: atk });
        const inters = st.intermediaireIds || [];
        const events = st.eventIds || [];
        if (inters.length > 0) {
          inters.forEach(ppId => {
            const pp = ppList.find(p => p.id === ppId);
            const label = pp ? (pp.nom || 'PP') : ppId;
            const inter = addNode('I:' + ppId, 'inter', label);
            if (atk && inter) links.push({ from: atk, to: inter });
            events.forEach(evId => {
              const ev = evList.find(e => e.id === evId);
              const labelEv = ev ? (ev.evenement || ev.ref || 'Évènement') : evId;
              const sev = ev ? parseInt(ev.impact, 10) || 0 : 0;
              const evNode = addNode('E:' + evId, 'event', labelEv, '', sev);
              if (inter && evNode) links.push({ from: inter, to: evNode });
            });
          });
        } else {
          events.forEach(evId => {
            const ev = evList.find(e => e.id === evId);
            const labelEv = ev ? (ev.evenement || ev.ref || 'Évènement') : evId;
            const sev = ev ? parseInt(ev.impact, 10) || 0 : 0;
            const evNode = addNode('E:' + evId, 'event', labelEv, '', sev);
            if (atk && evNode) links.push({ from: atk, to: evNode });
          });
        }
      });
    });

    const nodes = Array.from(nodeMap.values());
    const laneOrder = ['source','objective','attack','inter','event'];
    const laneX = { source:110, objective:320, attack:540, inter:760, event:1000 };
    const laneTop = 70;
    const laneBottom = 500;
    const laneHeight = laneBottom - laneTop;

    const byType = {};
    nodes.forEach(n => {
      n.x = laneX[n.type] || laneX.attack;
      (byType[n.type] ||= []).push(n);
    });

    const minGapForCount = (count) => {
      if (count <= 1) return laneHeight;
      const theoretical = laneHeight / (count - 1);
      return Math.min(110, Math.max(55, theoretical));
    };

    laneOrder.forEach((type, laneIndex) => {
      const group = byType[type];
      if (!group || !group.length) return;

      const gap = minGapForCount(group.length);
      const fallbackY = (i) => group.length === 1
        ? laneTop + laneHeight / 2
        : Math.min(laneBottom, laneTop + i * gap);
      const desiredPositions = group.map((node, index) => {
        const prevNeighbors = links
          .filter(l => l.to === node && laneOrder.indexOf(l.from.type) < laneIndex)
          .map(l => l.from);
        const nextNeighbors = links
          .filter(l => l.from === node && laneOrder.indexOf(l.to.type) > laneIndex)
          .map(l => l.to);

        const points = [];
        prevNeighbors.forEach(n => { if (typeof n.y === 'number') points.push(n.y); });
        nextNeighbors.forEach(n => { if (typeof n.y === 'number') points.push(n.y); });

        let desired;
        if (points.length) {
          desired = points.reduce((sum, y) => sum + y, 0) / points.length;
        } else {
          desired = fallbackY(index);
        }

        desired = Math.max(laneTop, Math.min(laneBottom, desired));
        return { node, index, desired };
      }).sort((a, b) => a.desired - b.desired || a.index - b.index);

      let lastY = laneTop - gap;
      desiredPositions.forEach((entry, idx) => {
        const remaining = desiredPositions.length - idx - 1;
        const maxAllowed = laneBottom - remaining * gap;
        let y = Math.max(entry.desired, lastY + gap);
        y = Math.min(y, maxAllowed);
        if (!Number.isFinite(y)) y = fallbackY(entry.index);
        entry.node.y = y;
        lastY = y;
      });

      if (group.length === 1) {
        group[0].y = laneTop + laneHeight / 2;
      }
    });

    const colorByType = {
      source: 'var(--source)',
      objective: 'var(--objective)',
      attack: 'var(--attack)',
      inter: 'var(--inter)'
    };
    const sevColor = { 1:'var(--g1)', 2:'var(--g2)', 3:'var(--g3)', 4:'var(--g4)' };

    function clampWithEllipsis(tspan) {
      let current = tspan.textContent || '';
      if (!current) {
        tspan.textContent = '…';
        return;
      }
      if (!current.endsWith('…')) {
        current += '…';
        tspan.textContent = current;
      }
      while (tspan.getComputedTextLength() > MAX_TEXT_WIDTH && current.length) {
        current = current.slice(0, -1);
        tspan.textContent = current + '…';
      }
    }

    function wrapSvgText(textEl, label) {
      const words = (label || '').split(/\s+/).filter(Boolean);
      const createTspan = (lineIndex) => {
        const span = document.createElementNS(SVG_NS, 'tspan');
        span.setAttribute('x', PADDING_X);
        span.setAttribute('dy', lineIndex === 0 ? 0 : LINE_HEIGHT);
        return span;
      };

      if (!words.length) {
        const emptySpan = createTspan(0);
        emptySpan.textContent = '';
        textEl.appendChild(emptySpan);
        return 1;
      }

      let line = '';
      let lineIndex = 0;
      let tspan = createTspan(lineIndex);
      textEl.appendChild(tspan);

      for (let i = 0; i < words.length; i++) {
        const word = words[i];
        const testLine = line ? `${line} ${word}` : word;
        tspan.textContent = testLine;

        if (tspan.getComputedTextLength() > MAX_TEXT_WIDTH && line) {
          // revert to previous line
          tspan.textContent = line;
          lineIndex++;
          if (lineIndex >= MAX_LINES) {
            clampWithEllipsis(tspan);
            return MAX_LINES;
          }
          tspan = createTspan(lineIndex);
          textEl.appendChild(tspan);
          line = word;
          tspan.textContent = word;
          if (tspan.getComputedTextLength() > MAX_TEXT_WIDTH) {
            clampWithEllipsis(tspan);
            return lineIndex + 1;
          }
        } else if (tspan.getComputedTextLength() > MAX_TEXT_WIDTH) {
          // Single long word on a fresh line
          clampWithEllipsis(tspan);
          return lineIndex + 1;
        } else {
          line = testLine;
        }

        if (i === words.length - 1) {
          tspan.textContent = line;
        }
      }

      return lineIndex + 1;
    }

    function makeNode(n){
      const g = document.createElementNS(SVG_NS,'g');
      g.classList.add('node');
      if (n.type === 'attack' || n.type === 'inter') g.classList.add('small');
      gNodes.appendChild(g);
      const txt = document.createElementNS(SVG_NS,'text');
      txt.setAttribute('x', PADDING_X);
      txt.setAttribute('y', PADDING_Y + LINE_HEIGHT - 4);
      g.appendChild(txt);
      const labelText = (n.type === 'event' && n.severity) ? `${n.label} [G${n.severity}]` : n.label;
      const lines = wrapSvgText(txt, labelText);
      const nodeHeight = Math.max(44, lines * LINE_HEIGHT + PADDING_Y * 2);
      g.setAttribute('transform',`translate(${n.x-HALF_WIDTH},${n.y-nodeHeight/2})`);
      const rect = document.createElementNS(SVG_NS,'rect');
      rect.setAttribute('width', NODE_WIDTH);
      rect.setAttribute('height', nodeHeight);
      const fill = (n.type === 'event') ? (sevColor[n.severity] || 'var(--g2)') : (colorByType[n.type] || 'var(--panel)');
      rect.setAttribute('fill', fill);
      g.insertBefore(rect, txt);
      g.addEventListener('mousemove', e => {
        if (!n.details) return;
        tip.style.left = e.clientX + 'px';
        tip.style.top = e.clientY + 'px';
        const extra = (n.type === 'event' && n.severity) ? ` | Gravité: ${n.severity}` : '';
        tip.textContent = n.details + extra;
        tip.style.opacity = 1;
      });
      g.addEventListener('mouseleave', () => tip.style.opacity = 0);
    }

    function linkPath(a,b){
      const bend = Math.max(20, Math.min(140, Math.abs(b.x - a.x)/2.5));
      const p = document.createElementNS(SVG_NS,'path');
      p.setAttribute('class','link');
      const d = `M ${a.x+HALF_WIDTH} ${a.y} C ${a.x+HALF_WIDTH+bend} ${a.y}, ${b.x-HALF_WIDTH-bend} ${b.y}, ${b.x-HALF_WIDTH} ${b.y}`;
      p.setAttribute('d', d);
      gLinks.appendChild(p);
    }

    nodes.forEach(makeNode);
    links.forEach(l => linkPath(l.from, l.to));
  }

  function updateAtelier3Chart() {
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) return;
    const radarWrap = document.getElementById('atelier3-radar-wrapper');
    const graphWrap = document.getElementById('strategic-graph');
    const cartoActive = document.getElementById('atelier3-carto-tab')?.classList.contains('active');
    if (cartoActive) {
      if (radarWrap) radarWrap.style.display = 'block';
      if (graphWrap) graphWrap.style.display = 'none';
      const ppc = analysis.data.ppc || [];
      renderCartoRadar(ppc);
    } else {
      if (radarWrap) radarWrap.style.display = 'none';
      if (graphWrap) graphWrap.style.display = 'block';
      renderStrategicGraph();
    }
  }

  function updateAtelier4Chart() {
    const container = document.getElementById('atelier4-chart');
    if (!container) {
      if (atelier4ChartInstance) {
        atelier4ChartInstance.dispose();
        atelier4ChartInstance = null;
      }
      atelier4DragHandlers = null;
      return;
    }
    if (typeof echarts === 'undefined') return;

    if (atelier4ChartInstance && atelier4ChartInstance.getDom() !== container) {
      atelier4ChartInstance.dispose();
      atelier4ChartInstance = null;
    }
    if (!atelier4ChartInstance) {
      atelier4ChartInstance = echarts.init(container);
    }

    const analysis = analyses[currentIndex];
    const risquesList = [];
    if (analysis && analysis.data) {
      if (Array.isArray(analysis.data.risques)) {
        analysis.data.risques.forEach((risk, idx) => {
          if (!risk) return;
          const gravite = parseLevel(risk.gravite, null);
          const vraisemblance = parseLevel(risk.vraisemblance, null);
          if (gravite === null || vraisemblance === null) return;
          const rawIdentifier = risk.indice || risk.id || risk.libelle || risk.titre || risk.name || '';
          const identifier = formatRiskIdentifier(rawIdentifier) || `R${idx + 1}`;
          const fullName = risk.libelle || risk.titre || risk.name || rawIdentifier || `Risque ${idx + 1}`;
          const description = risk.description || risk.details || risk.detail || '';
          risquesList.push({
            name: identifier,
            identifier,
            fullName,
            value: [gravite, vraisemblance],
            rawValue: [gravite, vraisemblance],
            description,
            meta: { type: 'global', riskId: risk.id || rawIdentifier || identifier || `global-${idx}` },
            metaKey: `global::${risk.id || rawIdentifier || identifier || idx}`
          });
        });
      }
      if (Array.isArray(analysis.data.so)) {
        analysis.data.so.forEach((scenario, sIdx) => {
          if (!scenario || !Array.isArray(scenario.risks)) return;
          const scenarioId = scenario.id || `scenario-${sIdx}`;
          scenario.risks.forEach((risk, rIdx) => {
            if (!risk) return;
            const gravite = parseLevel(risk.gravite, null);
            const vraisemblance = parseLevel(risk.vraisemblance, null);
            if (gravite === null || vraisemblance === null) return;
            const rawIdentifier = risk.id || risk.identifier || risk.code || risk.name || '';
            const identifier = formatRiskIdentifier(rawIdentifier) || `S${sIdx + 1}-${rIdx + 1}`;
            const fullName = risk.name || rawIdentifier || `Risque ${rIdx + 1}`;
            const description = risk.description || risk.details || risk.detail || '';
            risquesList.push({
              name: identifier,
              identifier,
              fullName,
              value: [gravite, vraisemblance],
              rawValue: [gravite, vraisemblance],
              description,
              meta: { type: 'scenario', scenarioId, riskIndex: rIdx, riskId: rawIdentifier || identifier },
              metaKey: `scenario::${scenarioId}::${rIdx}`
            });
          });
        });
      }
    }

    const data = risquesList;

    const scatterData = spreadInSquare(data, 0.14);
    const accent = getCssVar('--accent', '#4da3ff');
    const textPrimary = getCssVar('--text-primary', '#e6e9ef');
    const textSecondary = getCssVar('--text-secondary', '#8aa0c4');

    const option = {
      backgroundColor: 'transparent',
      title: {
        text: 'Matrice des risques (niveau actuel)',
        left: 'center',
        textStyle: { color: textPrimary, fontSize: 18, fontWeight: 'bold' }
      },
      grid: { left: 60, right: 30, top: 70, bottom: 60 },
      tooltip: {
        trigger: 'item',
        borderWidth: 1,
        backgroundColor: '#0f172a',
        borderColor: accent,
        textStyle: { color: textPrimary },
        formatter: (params) => {
          const data = params.data || {};
          const identifier = data.identifier || params.name || 'Risque';
          const fullName = data.fullName && data.fullName !== identifier ? data.fullName : '';
          const desc = data.description ? `<br/><em>${data.description}</em>` : '';
          const raw = data.rawValue ? data.rawValue : params.value;
          const gravite = raw && raw[0] !== undefined ? raw[0] : params.value[0];
          const vraisemblance = raw && raw[1] !== undefined ? raw[1] : params.value[1];
          let html = `<strong>${identifier}</strong>`;
          if (fullName) {
            html += `<br/>Nom : ${fullName}`;
          }
          html += desc;
          html += `<br/>Gravité : ${gravite}<br/>Vraisemblance : ${vraisemblance}`;
          return html;
        }
      },
      xAxis: {
        type: 'value',
        min: 0,
        max: 4,
        splitNumber: 4,
        name: 'Gravité',
        nameGap: 30,
        nameTextStyle: { color: textSecondary },
        axisLine: { lineStyle: { color: textSecondary } },
        axisLabel: { color: textSecondary },
        splitLine: { lineStyle: { color: 'rgba(148,163,184,0.2)' } }
      },
      yAxis: {
        type: 'value',
        min: 0,
        max: 4,
        splitNumber: 4,
        name: 'Vraisemblance',
        nameGap: 40,
        nameTextStyle: { color: textSecondary },
        axisLine: { lineStyle: { color: textSecondary } },
        axisLabel: { color: textSecondary },
        splitLine: { lineStyle: { color: 'rgba(148,163,184,0.2)' } }
      },
      series: [
        {
          name: 'Risque actuel',
          type: 'scatter',
          data: scatterData,
          symbolSize: 20,
          itemStyle: { color: accent },
          label: {
            show: true,
            formatter: (params) => (params.data && params.data.identifier) ? params.data.identifier : params.name,
            color: textPrimary,
            fontSize: 12,
            position: 'top'
          }
        }
      ],
      animationDuration: 300,
      graphic: scatterData.length ? [] : [
        {
          type: 'text',
          left: 'center',
          top: 'middle',
          style: {
            text: 'Ajoutez des risques pour alimenter la matrice.',
            fill: textSecondary,
            fontSize: 14
          }
        }
      ]
    };

    atelier4ChartInstance.setOption(option, true);
    bindAtelier4Drag(scatterData);
  }

  function updateAtelier5Chart() {
    renderRisquesChart();
    renderSupportLevelChart();
  }

  // ----- Event handlers for adding items
  function setupAddButtons() {
    const addMissionBtn = document.getElementById('add-mission-btn');
    if (addMissionBtn) {
      addMissionBtn.addEventListener('click', () => {
        const analysis = analyses[currentIndex];
        if (!analysis.data) analysis.data = {};
        if (!analysis.data.missions) analysis.data.missions = [];
        analysis.data.missions.push({ id: uid(), denom:'', nature:'information', description:'', responsable:'', supports: [] });
        saveAnalyses();
        renderMissionsTable();
        renderSupportsQualifTable();
        updateAtelier1Graph();
      });
    }
    const addSupportQualifBtn = document.getElementById('add-support-qualif-btn');
    if (addSupportQualifBtn) {
      addSupportQualifBtn.addEventListener('click', () => {
        const analysis = analyses[currentIndex];
        if (!analysis.data) analysis.data = {};
        if (!Array.isArray(analysis.data.supportsQualif)) analysis.data.supportsQualif = [];
        analysis.data.supportsQualif.push({ name:'', description:'', responsable:'', vulnerabilities: [] });
        saveAnalyses();
        renderSupportsQualifTable();
        renderSupportActions();
      });
    }
    const refreshSupportsQualifBtn = document.getElementById('refresh-supports-qualif-btn');
    if (refreshSupportsQualifBtn) {
      refreshSupportsQualifBtn.addEventListener('click', () => {
        renderSupportsQualifTable();
        renderSupportActions();
      });
    }
    // GAP analysis: add new requirement
    const addGapBtn = document.getElementById('add-gap-btn');
    if (addGapBtn) {
      addGapBtn.addEventListener('click', () => {
        const analysis = analyses[currentIndex];
        if (!analysis.data) analysis.data = {};
        if (!analysis.data.gap) analysis.data.gap = [];
        analysis.data.gap.push({ id: uid(), domaine:'', titre:'', description:'', application:'Appliqué', justification:'' });
        saveAnalyses();
        renderGapTable();
        updateGapChart();
      });
    }
    // GAP analysis: import requirements from JSON file
    const importGapBtn = document.getElementById('import-gap-btn');
    const gapFileInput = document.getElementById('gap-import-file');
    if (importGapBtn && gapFileInput) {
      importGapBtn.addEventListener('click', () => {
        gapFileInput.value = '';
        gapFileInput.click();
      });
      gapFileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function(evt) {
          try {
            const json = JSON.parse(evt.target.result);
            if (!Array.isArray(json)) {
              alert('Le fichier doit contenir un tableau d\'exigences.');
              return;
            }
            const analysis = analyses[currentIndex];
            if (!analysis.data) analysis.data = {};
            if (!analysis.data.gap) analysis.data.gap = [];
            json.forEach(obj => {
              const req = {
                id: uid(),
                domaine: obj.domaine || obj.domain || '',
                titre: obj.titre || obj.title || '',
                description: obj.description || obj.desc || '',
                application: obj.application || obj.status || '',
                justification: obj.justification || obj.justif || ''
              };
              analysis.data.gap.push(req);
            });
            saveAnalyses();
            renderGapTable();
            updateGapChart();
          } catch (err) {
            alert('Erreur lors de la lecture du fichier : ' + err.message);
          }
        };
        reader.readAsText(file);
      });
    }
    const addSrovBtn = document.getElementById('add-srov-btn');
    if (addSrovBtn) {
      addSrovBtn.addEventListener('click', () => {
        const analysis = analyses[currentIndex];
        if (!analysis.data) analysis.data = {};
        if (!analysis.data.srov) analysis.data.srov = [];
        // Default SROV entry: blank source and objectif, motivation & ressources at 1,
        // priorité 1, retenue true and empty justification
        analysis.data.srov.push({
          id: uid(),
          source: '',
          objectif: '',
          motivation: 1,
          ressources: 1,
          priorite: 1,
          retenue: true,
          justification: ''
        });
        saveAnalyses();
        renderSROV();
        updateAtelier2Chart();
      });
    }

    // Atelier 5: import GAP requirements from Atelier 1
    const addGapRow = document.getElementById('add-gap-action-row');
    if (addGapRow) {
      addGapRow.addEventListener('click', () => {
        openImportModal('gap');
      });
    }

    // Atelier 5: add a new support action row
    const addSupportRow = document.getElementById('add-support-action-row');
    if (addSupportRow) {
      addSupportRow.addEventListener('click', () => {
        const analysis = analyses[currentIndex];
        if (!analysis.data) analysis.data = {};
        if (!Array.isArray(analysis.data.actionsSupports)) analysis.data.actionsSupports = [];
        analysis.data.actionsSupports.push({ supportName: '', vulnName: '', initialLevel: '', residualLevel: '', actions: [] });
        saveAnalyses();
        renderSupportActions();
      });
    }
    const importVulnBtn = document.getElementById('import-vuln-support');
    if (importVulnBtn) {
      importVulnBtn.addEventListener('click', () => openImportModal('vulns'));
    }
    // Atelier 5: add a new party action row
    const addPartieRow = document.getElementById('add-partie-action-row');
    if (addPartieRow) {
      addPartieRow.addEventListener('click', () => {
        const analysis = analyses[currentIndex];
        if (!analysis.data) analysis.data = {};
        if (!Array.isArray(analysis.data.actionsParties)) analysis.data.actionsParties = [];
        analysis.data.actionsParties.push({ ppId: '', actions: [] });
        saveAnalyses();
        renderPartiesActions();
      });
    }

    // Atelier 5: import risks from Atelier 4
    const addRiskRow = document.getElementById('add-risque-action-row');
    if (addRiskRow) {
      addRiskRow.addEventListener('click', () => {
        openImportModal('risques');
      });
    }
    const addPPBtn = document.getElementById('add-pp-btn');
    if (addPPBtn) {
      addPPBtn.addEventListener('click', () => {
        const analysis = analyses[currentIndex];
        if (!analysis.data) analysis.data = {};
        if (!analysis.data.pp) analysis.data.pp = [];
        analysis.data.pp.push({ id: uid(), categorie:'', nom:'', description:'', niveauSSI:0, indiceMenace:0 });
        saveAnalyses();
        renderPP();
      });
    }
    const addSSBtn = document.getElementById('add-ss-btn');
    if (addSSBtn) {
      addSSBtn.addEventListener('click', () => {
        const analysis = analyses[currentIndex];
        if (!analysis.data) analysis.data = {};
        if (!analysis.data.ss) analysis.data.ss = [];
        analysis.data.ss.push({ id: uid(), source:'', objectif:'', vraisemblance:'', gravite:'' });
        saveAnalyses();
        renderSS();
      });
    }
    const addSOBtn = document.getElementById('add-so-btn');
    if (addSOBtn) {
      addSOBtn.addEventListener('click', () => {
        const analysis = analyses[currentIndex];
        if (!analysis.data) analysis.data = {};
        if (!analysis.data.so) analysis.data.so = [];
        analysis.data.so.push({ id: uid(), chemin:'', vraisemblanceGlobale:'' });
        saveAnalyses();
        renderSO();
      });
    }
    const addRisqueBtn = document.getElementById('add-risque-btn');
    if (addRisqueBtn) {
      addRisqueBtn.addEventListener('click', () => {
        const analysis = analyses[currentIndex];
        if (!analysis.data) analysis.data = {};
        if (!analysis.data.risques) analysis.data.risques = [];
        const defaultMission = (analysis.data.missions && analysis.data.missions[0]) ? analysis.data.missions[0].id : '';
        const defaultEvent = (analysis.data.events && analysis.data.events[0]) ? analysis.data.events[0].id : '';
        const defaultScenario = (analysis.data.so && analysis.data.so[0]) ? analysis.data.so[0].id : '';
        analysis.data.risques.push({
          id: uid(),
          libelle:'',
          titre:'',
          description:'',
          missionId: defaultMission,
          eventId: defaultEvent,
          scenarioId: defaultScenario,
          sourceIds: [],
          indice:'',
          vraisemblance:'',
          gravite:'',
          mesures:''
        });
        saveAnalyses();
        renderRisques();
        updateAtelier4Chart();
        updateAtelier5Chart();
      });
    }

    // Cartographie: add new stakeholder row
    const addPPCBtn = document.getElementById('add-ppc-btn');
    if (addPPCBtn) {
      addPPCBtn.addEventListener('click', () => {
        const analysis = analyses[currentIndex];
        if (!analysis.data) analysis.data = {};
        if (!Array.isArray(analysis.data.ppc)) analysis.data.ppc = [];
        analysis.data.ppc.push({
          id: uid(),
          nom: '',
          categorie: 'prestataire',
          supportIds: [],
          valueIds: [],
          dependance: 1,
          penetration: 1,
          maturite: 1,
          confiance: 1
        });
        saveAnalyses();
        renderPPCarto();
        updateAtelier3Chart();
      });
    }

    // Strategic scenarios: add new scenario row
    const addStrategyBtn = document.getElementById('add-strategy-btn');
    if (addStrategyBtn) {
      addStrategyBtn.addEventListener('click', () => {
        const analysis = analyses[currentIndex];
        if (!analysis.data) analysis.data = {};
        if (!Array.isArray(analysis.data.strategies)) analysis.data.strategies = [];
        analysis.data.strategies.push({
          id: uid(),
          source: '',
          objectif: '',
          chemins: [],
          intermediaireIds: [],
          eventIds: []
        });
        saveAnalyses();
        renderStrategies();
      });
    }
  }

  // ----- Navigation and general event handlers
  function setupNavigation() {
    // Top-level tabs in the single-page version are rendered as buttons and
    // toggled via JavaScript.  In the multi-page version, the navigation
    // elements become anchors (<a href="atelierX.html">) and should not be
    // intercepted by JS—clicking them should perform a real page
    // navigation.  Only attach click handlers to tab controls that do
    // **not** have an href attribute.
    const tabs = document.querySelectorAll('.tab-btn');
    tabs.forEach(btn => {
      // Anchors navigate to other pages; store current analysis before leaving
      if (btn.tagName && btn.tagName.toLowerCase() === 'a' && btn.hasAttribute('href')) {
        btn.addEventListener('click', persistCurrentAnalysisId);
        return;
      }
      btn.addEventListener('click', () => {
        tabs.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const target = btn.getAttribute('data-tab');
        document.querySelectorAll('.tab-content').forEach(section => {
          section.classList.toggle('active', section.id === target);
        });
      });
    });

    // Sub‑tabs inside Atelier 1
    const subtabButtons = document.querySelectorAll('#atelier1-subtabs .atelier1-subtab-btn');
    subtabButtons.forEach(btn => {
      btn.addEventListener('click', () => {
        subtabButtons.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const target = btn.getAttribute('data-subtab');
        document.querySelectorAll('.atelier1-subtab-content').forEach(content => {
          content.classList.toggle('active', content.id === 'atelier1-' + target + '-tab');
        });
        // Show/hide graphs depending on sub‑tab
        const graphEl = document.getElementById('atelier1-graph');
        const gapChartEl = document.getElementById('gap-overview-chart');
        if (graphEl && gapChartEl) {
          if (target === 'values') {
            graphEl.style.display = '';
            gapChartEl.style.display = 'none';
            // remove full-width layout when leaving gap analysis
          } else if (target === 'gap') {
            graphEl.style.display = 'none';
            gapChartEl.style.display = '';
            // Draw the GAP chart when switching to this tab
            updateGapChart();
          } else if (target === 'vuln') {
            graphEl.style.display = 'none';
            gapChartEl.style.display = 'none';
            renderSupportsQualifTable();
          }
        }
      });
    });

    // Sub‑tabs inside Atelier 3
    const ppSubBtns = document.querySelectorAll('#atelier3-subtabs .atelier3-subtab-btn');
    ppSubBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        ppSubBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const target = btn.getAttribute('data-subtab');
        document.querySelectorAll('#atelier3 .atelier3-subtab-content').forEach(content => {
          content.classList.toggle('active', content.id === 'atelier3-' + target + '-tab');
        });
        // Update chart when switching
        updateAtelier3Chart();
      });
    });
  }

  function setupSidebarToggle() {
    const btn = document.getElementById('sidebar-toggle');
    if (!btn) return;
    const sidebar = document.getElementById('sidebar');
    try {
      if (localStorage.getItem('ebiosSidebarCollapsed') === '1') {
        sidebar.classList.add('collapsed');
      }
    } catch (e) { /* ignore */ }
    btn.addEventListener('click', () => {
      sidebar.classList.toggle('collapsed');
      try {
        localStorage.setItem('ebiosSidebarCollapsed', sidebar.classList.contains('collapsed') ? '1' : '0');
      } catch (e) { /* ignore */ }
    });
  }

  async function exportPdfReport() {
    if (currentIndex < 0) return;
    const analysis = analyses[currentIndex];
    const reportWindow = window.open('', '_blank');
    if (!reportWindow) return;

    reportWindow.document.write('<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><title>Rapport EBIOS RM</title><link rel="stylesheet" href="styles.css"><style>section{page-break-after:always;margin-bottom:2rem;}h1{text-align:center;margin-top:0;}body{padding:20px;}button{display:none;}*{overflow:visible!important;max-height:none!important;}</style></head><body></body></html>');
    reportWindow.document.close();

    const body = reportWindow.document.body;
    const title = reportWindow.document.createElement('h1');
    title.textContent = (analysis.title || 'Analyse EBIOS RM') + ' – Rapport';
    body.appendChild(title);

    const pages = ['atelier1.html', 'atelier2.html', 'atelier3.html', 'atelier4.html', 'atelier5.html'];
    for (const page of pages) {
      await appendPageToReport(page, body, reportWindow.document, analysis);
    }

    reportWindow.focus();
    reportWindow.print();
  }

  function appendPageToReport(page, targetBody, reportDoc, analysis) {
    return new Promise(resolve => {
      const iframe = document.createElement('iframe');
      iframe.style.position = 'fixed';
      iframe.style.left = '-9999px';
      iframe.src = page;
      iframe.onload = () => {
        setTimeout(() => {
          const doc = iframe.contentDocument;
          const win = iframe.contentWindow;
          if (!doc) {
            document.body.removeChild(iframe);
            return resolve();
          }
          // Ensure that data for all sub‑tabs is rendered before we convert
          // canvases to images and import the section into the report.
          if (win) {
            try {
              if (page === 'atelier1.html') {
                win.updateGapChart && win.updateGapChart();
                win.renderSupportsQualifTable && win.renderSupportsQualifTable();
              } else if (page === 'atelier3.html') {
                win.updateAtelier3Chart && win.updateAtelier3Chart();
              } else if (page === 'atelier5.html') {
                win.renderGapActions && win.renderGapActions();
                win.renderSupportActions && win.renderSupportActions();
                win.renderPartiesActions && win.renderPartiesActions();
                win.renderRisquesActions && win.renderRisquesActions();
                win.renderPlanActions && win.renderPlanActions();
                win.renderRisquesChart && win.renderRisquesChart();
              }
            } catch (e) {
              // Ignore rendering errors; proceed with whatever content is available
            }
          }
          doc.querySelectorAll('canvas').forEach(c => {
            const img = doc.createElement('img');
            try {
              img.src = c.toDataURL('image/png');
            } catch (e) {
              img.src = '';
            }
            img.width = c.width;
            img.height = c.height;
            img.className = c.className;
            img.style.cssText = c.style.cssText;
            c.parentNode.replaceChild(img, c);
          });
          const section = doc.querySelector('section.tab-content');
          if (section) {
            const imported = reportDoc.importNode(section, true);
            imported.querySelectorAll('button').forEach(btn => btn.remove());
            // Remove subtab navigation and display all subtab content
            imported.querySelectorAll('.subtab-nav').forEach(nav => nav.remove());
            imported.querySelectorAll('[class*="subtab-content"]').forEach(div => {
              div.style.display = 'block';
              div.classList.remove('active');
            });
            if (page === 'atelier1.html') {
              const graph = imported.querySelector('#atelier1-graph');
              if (graph) graph.remove();
              const valuesTab = imported.querySelector('#atelier1-values-tab');
              const missionsTable = imported.querySelector('#missions-table');
              if (missionsTable) missionsTable.remove();
              if (valuesTab) {
                // Table: Valeurs métier et biens supports
                const h1 = reportDoc.createElement('h3');
                h1.textContent = 'Valeurs métier et biens supports';
                valuesTab.appendChild(h1);
                const table1 = reportDoc.createElement('table');
                table1.className = 'data-table';
                const thead1 = reportDoc.createElement('thead');
                const hr1 = reportDoc.createElement('tr');
                ['Valeur métier', 'Bien support', 'Description', 'Responsable'].forEach(t => {
                  const th = reportDoc.createElement('th');
                  th.textContent = t;
                  hr1.appendChild(th);
                });
                thead1.appendChild(hr1);
                table1.appendChild(thead1);
                const tbody1 = reportDoc.createElement('tbody');
                (analysis.data.missions || []).forEach(m => {
                  if (Array.isArray(m.supports) && m.supports.length) {
                    m.supports.forEach(s => {
                      const tr = reportDoc.createElement('tr');
                      let td = reportDoc.createElement('td');
                      td.textContent = m.denom || '';
                      tr.appendChild(td);
                      td = reportDoc.createElement('td');
                      td.textContent = s.name || '';
                      tr.appendChild(td);
                      td = reportDoc.createElement('td');
                      td.textContent = s.description || '';
                      tr.appendChild(td);
                      td = reportDoc.createElement('td');
                      td.textContent = s.responsable || '';
                      tr.appendChild(td);
                      tbody1.appendChild(tr);
                    });
                  } else {
                    const tr = reportDoc.createElement('tr');
                    let td = reportDoc.createElement('td');
                    td.textContent = m.denom || '';
                    tr.appendChild(td);
                    td = reportDoc.createElement('td');
                    tr.appendChild(td);
                    td = reportDoc.createElement('td');
                    tr.appendChild(td);
                    td = reportDoc.createElement('td');
                    tr.appendChild(td);
                    tbody1.appendChild(tr);
                  }
                });
                table1.appendChild(tbody1);
                valuesTab.appendChild(table1);

                // Table: Valeurs métier et évènements redoutés
                const h2 = reportDoc.createElement('h3');
                h2.textContent = 'Valeurs métier et évènements redoutés';
                valuesTab.appendChild(h2);
                const table2 = reportDoc.createElement('table');
                table2.className = 'data-table';
                const thead2 = reportDoc.createElement('thead');
                const hr2 = reportDoc.createElement('tr');
                ['Valeur métier', 'Évènement redouté', 'Description des impacts', 'Impact'].forEach(t => {
                  const th = reportDoc.createElement('th');
                  th.textContent = t;
                  hr2.appendChild(th);
                });
                thead2.appendChild(hr2);
                table2.appendChild(thead2);
                const tbody2 = reportDoc.createElement('tbody');
                (analysis.data.missions || []).forEach(m => {
                  const events = (analysis.data.events || []).filter(ev => ev.missionId === m.id);
                  if (events.length) {
                    events.forEach(ev => {
                      const tr = reportDoc.createElement('tr');
                      let td = reportDoc.createElement('td');
                      td.textContent = m.denom || '';
                      tr.appendChild(td);
                      td = reportDoc.createElement('td');
                      td.textContent = ev.evenement || '';
                      tr.appendChild(td);
                      td = reportDoc.createElement('td');
                      td.textContent = ev.impactDescription || '';
                      tr.appendChild(td);
                      td = reportDoc.createElement('td');
                      td.textContent = ev.impact != null ? String(ev.impact) : '';
                      tr.appendChild(td);
                      tbody2.appendChild(tr);
                    });
                  } else {
                    const tr = reportDoc.createElement('tr');
                    let td = reportDoc.createElement('td');
                    td.textContent = m.denom || '';
                    tr.appendChild(td);
                    td = reportDoc.createElement('td');
                    tr.appendChild(td);
                    td = reportDoc.createElement('td');
                    tr.appendChild(td);
                    td = reportDoc.createElement('td');
                    tr.appendChild(td);
                    tbody2.appendChild(tr);
                  }
                });
                table2.appendChild(tbody2);
                valuesTab.appendChild(table2);
              }
            }
            // Replace form fields with plain text so all text is visible and colours are kept
            imported.querySelectorAll('input, textarea, select').forEach(field => {
              const tag = field.tagName.toLowerCase();
              const replacement = reportDoc.createElement(tag === 'textarea' ? 'div' : 'span');
              replacement.className = field.className;
              replacement.style.cssText = field.style.cssText;
              if (tag === 'textarea') {
                replacement.style.whiteSpace = 'pre-wrap';
                replacement.textContent = field.value || field.textContent || '';
              } else if (tag === 'select') {
                const opt = field.options[field.selectedIndex];
                replacement.textContent = opt ? opt.textContent : '';
              } else {
                replacement.textContent = field.value || field.textContent || '';
              }
              field.parentNode.replaceChild(replacement, field);
            });
            // Ensure tables are fully displayed
            imported.querySelectorAll('table').forEach(tbl => {
              tbl.style.width = '100%';
            });
            imported.style.pageBreakAfter = 'always';
            targetBody.appendChild(imported);
          }
          document.body.removeChild(iframe);
          resolve();
        }, 500);
      };
      document.body.appendChild(iframe);
    });
  }

  function setupAnalysisControls() {
    document.getElementById('analysis-title').addEventListener('input', (e) => {
      if (currentIndex < 0) return;
      analyses[currentIndex].title = e.target.value;
      saveAnalyses();
      renderAnalysisList();
    });
    document.getElementById('new-analysis-btn').addEventListener('click', () => {
      const newAnalysis = {
        id: uid(),
        title: 'Nouvelle analyse',
        data: {
          missions: [],
          events: [],
          supportsQualif: [],
          // GAP analysis requirements (domaine, titre, description, application, justification)
          gap: [],
          // Atelier 2 couples source/objectif
          srov: [],
          // Atelier 3: scénarios (ancienne liste) et cartographie (ppc)
          pp: [],
          ppc: [],
          ss: [],
          so: [],
          risques: [],
          // Actions & conformité (Atelier 5)
          actionsGap: [],
          actionsSupports: [],
          actionsParties: [],
          actionsRisques: [],
          // Atelier 3: scénarios stratégiques
          strategies: []
        }
      };
      analyses.push(newAnalysis);
      saveAnalyses();
      currentIndex = analyses.length - 1;
      renderAnalysisList();
      selectAnalysis(currentIndex);
    });
    document.getElementById('delete-btn').addEventListener('click', () => {
      if (currentIndex < 0) return;
      if (!confirm('Supprimer cette analyse ?')) return;
      analyses.splice(currentIndex, 1);
      saveAnalyses();
      currentIndex = analyses.length > 0 ? 0 : -1;
      renderAnalysisList();
      if (currentIndex >= 0) selectAnalysis(currentIndex);
      else {
        // Clear form if no analyses left
        document.getElementById('analysis-title').value = '';
        document.querySelectorAll('.item-list').forEach(list => list.innerHTML = '');
        document.querySelectorAll('canvas').forEach(canvas => clearCanvas(canvas));
      }
    });

    // Sub‑tabs inside Atelier 5
    const sub5Btns = document.querySelectorAll('#atelier5-subtabs .atelier5-subtab-btn');
    sub5Btns.forEach(btn => {
      btn.addEventListener('click', () => {
        sub5Btns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const target = btn.getAttribute('data-subtab');
        document.querySelectorAll('#atelier5 .atelier5-subtab-content').forEach(content => {
          content.classList.toggle('active', content.id === 'atelier5-' + target + '-tab');
        });
        // Refresh charts when switching tabs
        if (target === 'plan') {
          renderPlanActions();
        }
        if (target === 'risques') {
          renderRisquesChart();
        }
      });
    });
    document.getElementById('export-btn').addEventListener('click', () => {
      if (currentIndex < 0) return;
      const analysis = analyses[currentIndex];
      const blob = new Blob([JSON.stringify(analysis, null, 2)], { type: 'application/json' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      const safeTitle = (analysis.title || 'analyse').replace(/[^a-z0-9]/gi, '_').toLowerCase();
      a.download = safeTitle + '.json';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    });
    const pdfBtn = document.getElementById('export-pdf-btn');
    if (pdfBtn) {
      pdfBtn.addEventListener('click', exportPdfReport);
    }
    document.getElementById('export-all-btn').addEventListener('click', () => {
      const blob = new Blob([JSON.stringify(analyses, null, 2)], { type: 'application/json' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'analyses_ebios.json';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    });
    document.getElementById('import-btn').addEventListener('click', () => {
      document.getElementById('import-file').click();
    });
    document.getElementById('import-file').addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = function(ev) {
        try {
          const json = JSON.parse(ev.target.result);
          if (Array.isArray(json)) {
            json.forEach(a => { if (a && !a.id) a.id = uid(); });
            analyses = json;
          } else if (typeof json === 'object' && json !== null) {
            if (!json.id) json.id = uid();
            analyses.push(json);
          } else {
            throw new Error('Invalid format');
          }
          saveAnalyses();
          currentIndex = analyses.length - 1;
          renderAnalysisList();
          selectAnalysis(currentIndex);
        } catch (err) {
          alert('Fichier JSON invalide');
        }
      };
      reader.readAsText(file);
      // Reset the input so the same file can be re‑imported if needed
      e.target.value = '';
    });
  }

  // ----- Initialize
  function init() {
    loadAnalyses();
    // Attempt to restore the previously selected analysis.  We store both
    // the index and the stable ID of the last selected analysis in
    // localStorage.  Prefer the index (avoids an array search) and fall
    // back to the ID if needed.
    let savedIndex = -1;
    try {
      const storedIdx = localStorage.getItem('ebiosCurrentAnalysisIndex');
      if (storedIdx !== null) {
        const idx = parseInt(storedIdx, 10);
        if (!isNaN(idx) && idx >= 0 && idx < analyses.length) {
          savedIndex = idx;
        }
      }
      if (savedIndex === -1) {
        const savedId = localStorage.getItem('ebiosCurrentAnalysisId');
        if (savedId) {
          savedIndex = analyses.findIndex(a => a && a.id === savedId);
        }
      }
    } catch (e) {
      savedIndex = -1;
    }
    renderAnalysisList();
    setupAnalysisControls();
    setupNavigation();
    setupSidebarToggle();
    setupAddButtons();
    setupActionImport();
    // Ensure the current analysis ID is saved even if the user reloads or
    // closes the page without navigating through the provided links.
    window.addEventListener('beforeunload', persistCurrentAnalysisId);
    // Select the previously selected analysis or the first one by default
    if (analyses.length > 0) {
      currentIndex = savedIndex >= 0 ? savedIndex : 0;
      selectAnalysis(currentIndex);
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
