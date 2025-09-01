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

  function loadAnalyses() {
    try {
      const data = localStorage.getItem('ebiosAnalyses');
      analyses = data ? JSON.parse(data) : [];
    } catch (e) {
      console.warn('Failed to parse localStorage data, resetting.', e);
      analyses = [];
    }
  }

  function saveAnalyses() {
    localStorage.setItem('ebiosAnalyses', JSON.stringify(analyses));
  }

  // Persist the ID of the currently selected analysis in localStorage so
  // that navigating between separate workshop pages restores the same
  // analysis automatically.
  function persistCurrentAnalysisId() {
    try {
      const sel = analyses[currentIndex];
      if (sel && sel.id) {
        localStorage.setItem('ebiosCurrentAnalysisId', sel.id);
      }
    } catch (e) {
      // Ignore storage errors (e.g., private browsing)
    }
  }

  // ----- Utility: generate a simple UID
  function uid() {
    return 'id-' + Math.random().toString(36).substring(2, 10) + Date.now().toString(36);
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

    // Ensure Atelier 3 grid layout matches the active sub‑tab
    const grid3 = document.querySelector('#atelier3 .atelier-grid');
    if (grid3) {
      const activeSub = document.querySelector('#atelier3-subtabs .atelier3-subtab-btn.active');
      if (activeSub && activeSub.getAttribute('data-subtab') === 'carto') {
        grid3.classList.add('carto-active');
      } else {
        grid3.classList.remove('carto-active');
      }
    }
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

  // ----- Atelier 1: Mission description
  function renderMissionDescription() {
    const textarea = document.getElementById('mission-description');
    if (!textarea) return;
    const analysis = analyses[currentIndex];
    if (!analysis.data) analysis.data = {};
    if (typeof analysis.data.missionDescription !== 'string') analysis.data.missionDescription = '';
    textarea.value = analysis.data.missionDescription;
    textarea.oninput = (e) => {
      analysis.data.missionDescription = e.target.value;
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
        const sItem = document.createElement('div');
        sItem.className = 'support-item';
        const nameInput = document.createElement('input');
        nameInput.type = 'text';
        nameInput.placeholder = 'Nom';
        nameInput.value = support.name || '';
        nameInput.oninput = (e) => {
          support.name = e.target.value;
          saveAnalyses();
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
        };
        const supRespInput = document.createElement('input');
        supRespInput.type = 'text';
        supRespInput.placeholder = 'Responsable';
        supRespInput.value = support.responsable || '';
        supRespInput.oninput = (e) => {
          support.responsable = e.target.value;
          saveAnalyses();
        };
        const rmBtn = document.createElement('button');
        rmBtn.textContent = '×';
        rmBtn.title = 'Supprimer ce bien support';
        rmBtn.addEventListener('click', () => {
          mission.supports.splice(sIdx, 1);
          saveAnalyses();
          renderMissionsTable();
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
        mission.supports.push({ name: '', description: '', responsable: '' });
        saveAnalyses();
        renderMissionsTable();
        updateAtelier1Graph();
      });
      supportsCell.appendChild(addSupBtn);
      // Button to add an existing support from other missions
      const addExistingBtn = document.createElement('button');
      addExistingBtn.className = 'add-support-btn';
      addExistingBtn.textContent = '+ Support existant';
      addExistingBtn.addEventListener('click', () => {
        // Collect all unique supports across missions
        const allSupports = [];
        analysis.data.missions.forEach(m2 => {
          (m2.supports || []).forEach(s => {
            const name = (s.name || '').trim();
            if (!name) return;
            if (!allSupports.some(ss => ss.name === name)) {
              allSupports.push({ name: name, description: s.description || '', responsable: s.responsable || '' });
            }
          });
        });
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
        analysis.data.missions.splice(idx, 1);
        // Remove associated events when mission deleted
        if (analysis.data.events) {
          analysis.data.events = analysis.data.events.filter(ev => ev.missionId !== mission.id);
        }
        saveAnalyses();
        renderMissionsTable();
        updateAtelier1Graph();
      });
      td.appendChild(delBtn);
      tr.appendChild(td);
      tbody.appendChild(tr);
    });
    // After rendering rows, set up resizable columns on the missions table
    addMissionTableResizers();
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
      // Skip the last column (actions) to avoid resizing it
      if (index === ths.length - 1) return;
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
      // Skip the last column (actions) to avoid resizing it
      if (index === ths.length - 1) return;
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
      // Skip last column (actions) from resizing
      if (index === ths.length - 1) return;
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
      if (index === ths.length - 1) return;
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
      // skip last column (actions)
      if (index === ths.length - 1) return;
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

  // ----- Risk selection modal (MITRE/OWASP) -----
  // A list of common attack techniques inspired by MITRE ATT&CK and OWASP.
  const RISK_LIBRARY = [
    'Phishing',
    'Malware',
    'Ransomware',
    'Injection SQL',
    'Cross‑Site Scripting (XSS)',
    'Escalade de privilèges',
    'Déni de service (DoS)',
    'Vol de données',
    'Fuite d’informations',
    'Compte compromis',
    'Man‑in‑the‑Middle',
    'Attaque par mot de passe',
    'Force brute',
    'Command and Control',
    'Exfiltration de données',
    'Livraison de malware',
    'Injection XML',
    'Inclusion de fichiers',
    'Sécurité insuffisante des API',
    'Configuration non sécurisée'
  ];

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
    const lines = text.split(/\r?\n/);
    if (lines.length < 2) return [];
    const header = lines[0].split(',');
    const idIndex = header.findIndex(h => /technique id/i.test(h));
    const nameIndex = header.findIndex(h => /technique name/i.test(h));
    const descIndex = header.findIndex(h => /technique description/i.test(h));
    const mitIdIndex = header.findIndex(h => /mitigation id/i.test(h));
    const mitNameIndex = header.findIndex(h => /mitigation name/i.test(h));
    const mitDescIndex = header.findIndex(h => /mitigation description/i.test(h));
    const map = new Map();
    for (let i = 1; i < lines.length; i++) {
      const row = lines[i];
      if (!row.trim()) continue;
      // naive CSV split (does not handle quoted commas)
      const cols = row.split(',');
      const tid = cols[idIndex] ? cols[idIndex].trim() : '';
      const tname = cols[nameIndex] ? cols[nameIndex].trim() : '';
      const tdesc = cols[descIndex] ? cols[descIndex].trim() : '';
      const mid = cols[mitIdIndex] ? cols[mitIdIndex].trim() : '';
      const mname = cols[mitNameIndex] ? cols[mitNameIndex].trim() : '';
      const mdesc = cols[mitDescIndex] ? cols[mitDescIndex].trim() : '';
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

  // Handle MITRE CSV import: open a modal to select techniques from the bundled CSV
  function setupMitreImport() {
    const btn = document.getElementById('import-mitre-btn');
    const modal = document.getElementById('mitre-modal');
    const listEl = document.getElementById('mitre-list');
    const searchInput = document.getElementById('mitre-search');
    const applyBtn = document.getElementById('mitre-import-apply');
    const closeBtn = document.getElementById('mitre-close-btn');
    if (!btn || !modal || !listEl || !searchInput || !applyBtn || !closeBtn) return;

    let items = [];
    let selected = new Set();

    function renderList(arr) {
      listEl.innerHTML = '';
      arr.forEach(item => {
        const li = document.createElement('li');
        li.textContent = `${item.id} – ${item.title}`;
        if (selected.has(item.id)) li.classList.add('selected');
        li.addEventListener('click', () => {
          if (selected.has(item.id)) {
            selected.delete(item.id);
            li.classList.remove('selected');
          } else {
            selected.add(item.id);
            li.classList.add('selected');
          }
        });
        listEl.appendChild(li);
      });
    }

    btn.addEventListener('click', () => {
      fetch('mitre_attack.csv')
        .then(res => {
          if (!res.ok) throw new Error('fetch');
          return res.text();
        })
        .then(text => {
          items = parseMitreCsv(text);
          selected = new Set((mitreLibrary || []).map(t => t.id));
          renderList(items);
          searchInput.value = '';
          modal.style.display = 'flex';
        })
        .catch(() => {
          alert('Impossible de charger le fichier MITRE.');
        });
    });

    searchInput.addEventListener('input', () => {
      const term = searchInput.value.toLowerCase();
      const filtered = items.filter(t =>
        (t.id + ' ' + t.title + ' ' + t.description).toLowerCase().includes(term)
      );
      renderList(filtered);
    });

    applyBtn.addEventListener('click', () => {
      mitreLibrary = items.filter(t => selected.has(t.id));
      localStorage.setItem('ebiosMitreLibrary', JSON.stringify(mitreLibrary));
      modal.style.display = 'none';
      alert('Base MITRE importée avec ' + mitreLibrary.length + ' techniques.');
    });

    closeBtn.addEventListener('click', () => {
      modal.style.display = 'none';
    });
  }

  // Setup the kill chain toggle button in Atelier 4.  When clicked,
  // it toggles a CSS class on the operations table to hide/show the
  // columns Connaître/Rester/Trouver/Exploiter, and updates the button
  // label accordingly.
  function setupKillChainToggle() {
    const btn = document.getElementById('toggle-killchain-btn');
    const table = document.getElementById('ops-table');
    if (!btn || !table) return;
    btn.addEventListener('click', () => {
      const hidden = table.classList.toggle('hide-killchain');
      btn.textContent = hidden ? 'Afficher Kill Chain' : 'Masquer Kill Chain';
    });
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

  function setupActionImport() {
    const confirmBtn = document.getElementById('import-confirm');
    if (confirmBtn) confirmBtn.addEventListener('click', () => applyImportSelection());
    const cancelBtn = document.getElementById('import-cancel');
    if (cancelBtn) cancelBtn.addEventListener('click', () => closeImportModal());
    const searchInput = document.getElementById('import-search');
    if (searchInput) searchInput.addEventListener('input', () => filterImportList(searchInput.value));
  }

  // Populate and show the import modal for the given type
  function openImportModal(type) {
    const modal = document.getElementById('import-modal');
    if (!modal) return;
    currentImportType = type;
    importSelections = new Set();
    const titleEl = document.getElementById('import-modal-title');
    const listEl = document.getElementById('import-list');
    const searchInput = document.getElementById('import-search');
    if (!titleEl || !listEl || !searchInput) return;
    // Set modal title based on type
    switch (type) {
      case 'gap':
        titleEl.textContent = 'Importer des exigences';
        break;
      case 'supports':
        titleEl.textContent = 'Importer des supports';
        break;
      case 'parties':
        titleEl.textContent = 'Importer des parties';
        break;
      case 'risques':
        titleEl.textContent = 'Importer des risques';
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
    } else if (type === 'parties') {
      (analysis.data.ppc || []).forEach(pp => {
        const id = pp.id || (pp.id = uid());
        const label = pp.nom || pp.name || 'Partie';
        const desc = pp.categorie || pp.categorie === '' ? (pp.categorie) : '';
        // maybe show dependance / penetration but keep simple
        items.push({ id, label, desc: desc || '', extra: '' });
      });
    } else if (type === 'risques') {
      // gather risks from operational scenarios (analysis.data.so)
      const riskMap = new Map();
      (analysis.data.so || []).forEach(scenario => {
        (scenario.risks || []).forEach(riskObj => {
          const name = riskObj.name || '';
          if (!name) return;
          if (!riskMap.has(name)) {
            // try to find description from mitreLibrary
            let desc = '';
            const tech = mitreLibrary.find(t => t && t['Technique ID'] === name || t && t['Technique Name'] === name);
            if (tech) desc = tech['Technique Description'] || '';
            riskMap.set(name, { name, desc });
          }
        });
      });
      riskMap.forEach((obj) => {
        items.push({ id: obj.name, label: obj.name, desc: obj.desc || '', extra: '' });
      });
    }
    // Render list
    listEl.innerHTML = '';
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
      div.addEventListener('click', () => {
        const id = div.dataset.id;
        if (importSelections.has(id)) {
          importSelections.delete(id);
          div.classList.remove('selected');
        } else {
          importSelections.add(id);
          div.classList.add('selected');
        }
      });
      listEl.appendChild(div);
    });
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

  // Close modal without applying changes
  function closeImportModal() {
    const modal = document.getElementById('import-modal');
    if (modal) modal.style.display = 'none';
    importSelections = new Set();
    currentImportType = null;
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
          analysis.data.actionsSupports.push({ supportName: name, actions: [] });
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
      (analysis.data.so || []).forEach(scenario => {
        const v = scenario.vraisemblance || scenario.vraisemblance;
        const g = scenario.gravite || scenario.gravité || scenario.grav;
        (scenario.risks || []).forEach(riskObj => {
          const name = riskObj.name || '';
          if (!name) return;
          const current = riskMap.get(name) || { name, vraisemblance: riskObj.vraisemblance || v || 1, gravite: riskObj.gravite || g || 1 };
          current.vraisemblance = Math.max(current.vraisemblance, riskObj.vraisemblance || v || 1);
          current.gravite = Math.max(current.gravite, riskObj.gravite || g || 1);
          riskMap.set(name, current);
        });
      });
      importSelections.forEach(name => {
        const risk = riskMap.get(name) || { name, vraisemblance: 1, gravite: 1 };
        if (!analysis.data.actionsRisques.some(row => row.riskName === name)) {
          analysis.data.actionsRisques.push({ riskName: name, residualV: risk.vraisemblance, residualG: risk.gravite, actions: [] });
        }
      });
      saveAnalyses();
      renderRisquesActions();
    }
    // Update plan actions after any import
    renderPlanActions();
    closeImportModal();
  }

  // Open the risk modal for a given scenario item.  This populates the
  // list of risks, sets up search filtering and manual addition, and
  // shows the modal overlay.
  function openRiskModal(targetItem) {
    riskModalTarget = targetItem;
    const modal = document.getElementById('risk-modal');
    if (!modal) return;
    const searchInput = document.getElementById('risk-search');
    const list = document.getElementById('risk-list');
    const manualInput = document.getElementById('risk-manual');
    const manualBtn = document.getElementById('risk-add-manual');
    const closeBtn = document.getElementById('risk-close-btn');
    const applyBtn = document.getElementById('risk-select-apply');
    // Helper to render risk list based on current filter
    // Track currently selected risk names in the modal for multi‑selection
    let selectedRiskNames = [];
    function renderRiskList(filter) {
      list.innerHTML = '';
      const term = (filter || '').toLowerCase();
      const source = (mitreLibrary && mitreLibrary.length > 0) ? mitreLibrary : RISK_LIBRARY;
      const items = Array.isArray(source)
        ? source.map(it => {
            if (typeof it === 'string') {
              return { id: '', title: it, description: '', name: it };
            }
            return {
              id: it.id || '',
              title: it.title || '',
              description: it.description || '',
              name: it.id ? (it.id + ' ' + it.title) : (it.title || it.id)
            };
          })
        : [];
      // Filter by term across id, title and description
      const filtered = items.filter(obj => {
        const full = (obj.name + ' ' + (obj.description || '')).toLowerCase();
        return full.includes(term);
      });
      filtered.forEach(obj => {
        const li = document.createElement('li');
        // Build inner HTML with ID, title and description
        li.innerHTML = `<div><strong>${obj.id}</strong> ${obj.title}</div>` +
          `<div class="risk-desc">${obj.description}</div>`;
        if (selectedRiskNames.includes(obj.name)) {
          li.classList.add('selected');
        }
        li.addEventListener('click', () => {
          const idx = selectedRiskNames.indexOf(obj.name);
          if (idx >= 0) {
            selectedRiskNames.splice(idx, 1);
          } else {
            selectedRiskNames.push(obj.name);
          }
          renderRiskList(searchInput.value);
        });
        list.appendChild(li);
      });
      if (filtered.length === 0) {
        const li = document.createElement('li');
        li.textContent = 'Aucun résultat';
        li.style.fontStyle = 'italic';
        li.style.cursor = 'default';
        list.appendChild(li);
      }
    }
    // Expose helper globally so we can call inside event handlers
    function addRiskToScenario(name) {
      if (!riskModalTarget) return;
      // Use scenario's vraisemblance and gravité levels without prompting
      riskModalTarget.risks.push({ name: name, vraisemblance: riskModalTarget.vraisemblance, gravite: riskModalTarget.gravite });
      saveAnalyses();
      renderSO();
      closeRiskModal();
    }
    // Bind search input
    searchInput.value = '';
    searchInput.oninput = (e) => {
      renderRiskList(e.target.value);
    };
    // Bind manual addition
    manualInput.value = '';
    manualBtn.onclick = () => {
      const name = manualInput.value.trim();
      if (!name) return;
      addRiskToScenario(name);
    };
    // Bind close
    closeBtn.onclick = () => {
      closeRiskModal();
    };
    // Bind apply selection: add all selected risks with same levels
    if (applyBtn) {
      applyBtn.onclick = () => {
        if (!riskModalTarget) return;
        if (!selectedRiskNames || selectedRiskNames.length === 0) {
          closeRiskModal();
          return;
        }
        selectedRiskNames.forEach(name => {
          riskModalTarget.risks.push({ name: name, vraisemblance: riskModalTarget.vraisemblance, gravite: riskModalTarget.gravite });
        });
        saveAnalyses();
        renderSO();
        closeRiskModal();
      };
    }
    // Before showing the modal, ensure the MITRE library is loaded.  If
    // no library is present, attempt to load the bundled CSV on the fly.
    function ensureMitreLoaded(callback) {
      if (mitreLibrary && mitreLibrary.length > 0) {
        callback();
        return;
      }
      // If data is stored in localStorage, reuse it
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
      // Otherwise fetch the CSV from the bundled resource and parse it
      fetch('mitre_attack.csv').then(res => {
        if (!res.ok) throw new Error('Cannot load MITRE CSV');
        return res.text();
      }).then(text => {
        const parsed = parseMitreCsv(text);
        if (parsed && parsed.length > 0) {
          mitreLibrary = parsed;
          localStorage.setItem('ebiosMitreLibrary', JSON.stringify(mitreLibrary));
        }
      }).catch(() => {
        // ignore errors and fall back to default library
      }).finally(() => {
        callback();
      });
    }
    // Show modal and populate list after ensuring the MITRE library is loaded
    ensureMitreLoaded(() => {
      modal.style.display = 'flex';
      renderRiskList(searchInput.value);
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

  // ----- Atelier 1: Graph generation.  A custom SVG graph replaces
  // the previous ECharts-based implementation, so no chart instance
  // is required here.
  let atelier1Chart = null; // unused placeholder, preserved for backward compatibility
  function updateAtelier1Graph() {
    // Custom SVG-based network rendering without external libraries
    const container = document.getElementById('atelier1-graph');
    if (!container) return;
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) return;
    const missions = analysis.data.missions || [];
    const events = analysis.data.events || [];
    // Determine maximum impact per mission
    const impactByMission = new Map();
    events.forEach(ev => {
      const mid = ev.missionId;
      const imp = parseInt(ev.impact, 10) || 0;
      if (!impactByMission.has(mid) || imp > impactByMission.get(mid)) {
        impactByMission.set(mid, imp);
      }
    });
    // Compute support statistics: degree and maximum impact across linked missions
    const supportStats = new Map(); // name -> { degree, maxImpact }
    missions.forEach(m => {
      const mImp = impactByMission.get(m.id) || 0;
      (m.supports || []).forEach(s => {
        const name = (s.name || '').trim();
        if (!name) return;
        if (!supportStats.has(name)) supportStats.set(name, { degree: 0, maxImpact: 0 });
        const stat = supportStats.get(name);
        stat.degree += 1;
        if (mImp > stat.maxImpact) stat.maxImpact = mImp;
        supportStats.set(name, stat);
      });
    });
    // Build mission nodes array
    const missionNodes = [];
    missions.forEach(m => {
      const imp = impactByMission.get(m.id) || 0;
      missionNodes.push({ id: m.id, name: m.denom || 'Valeur', desc: m.description || '', impact: imp });
    });
    // Build support nodes array
    const supportNodes = [];
    supportStats.forEach((stat, name) => {
      supportNodes.push({ id: name, name: name, degree: stat.degree, maxImpact: stat.maxImpact });
    });
    // Clear previous content
    container.innerHTML = '';
    // Determine container dimensions
    const width = container.clientWidth || container.offsetWidth || 600;
    const height = container.clientHeight || container.offsetHeight || 600;
    // Create SVG
    const svgNS = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(svgNS, 'svg');
    svg.setAttribute('width', width);
    svg.setAttribute('height', height);
    svg.style.display = 'block';
    // Helper functions
    function impactColor(lvl) {
      return {1: '#2a9d8f', 2: '#e9c46a', 3: '#f4a261', 4: '#e63946'}[lvl] || '#3c85cc';
    }
    function triangleSize(lvl) {
      return {1: 10, 2: 14, 3: 18, 4: 22}[lvl] || 10;
    }
    function circleRadius(deg) {
      return Math.min(32, 10 + deg * 6);
    }
    // Compute positions: missions on left, supports on right
    const mCount = missionNodes.length;
    const sCount = supportNodes.length;
    const mSpacing = mCount > 0 ? height / (mCount + 1) : 0;
    const sSpacing = sCount > 0 ? height / (sCount + 1) : 0;
    const mX = Math.max(80, width * 0.2);
    const sX = Math.min(width - 80, width * 0.8);
    const missionPos = new Map();
    missionNodes.forEach((node, idx) => {
      const y = (idx + 1) * mSpacing;
      missionPos.set(node.id, { x: mX, y, node });
    });
    const supportPos = new Map();
    supportNodes.forEach((node, idx) => {
      const y = (idx + 1) * sSpacing;
      supportPos.set(node.id, { x: sX, y, node });
    });
    // Draw lines for each link
    missions.forEach(m => {
      (m.supports || []).forEach(s => {
        const name = (s.name || '').trim();
        if (!name) return;
        const mP = missionPos.get(m.id);
        const sP = supportPos.get(name);
        if (!mP || !sP) return;
        const line = document.createElementNS(svgNS, 'line');
        line.setAttribute('x1', mP.x + triangleSize(mP.node.impact));
        line.setAttribute('y1', mP.y);
        line.setAttribute('x2', sP.x - circleRadius(sP.node.degree));
        line.setAttribute('y2', sP.y);
        line.setAttribute('stroke', '#888');
        line.setAttribute('stroke-width', '1');
        svg.appendChild(line);
      });
    });
    // Draw mission nodes
    missionPos.forEach(({ x, y, node }) => {
      const size = triangleSize(node.impact);
      const color = impactColor(node.impact);
      const points = [
        `${x - size},${y - size}`,
        `${x - size},${y + size}`,
        `${x},${y}`
      ].join(' ');
      const poly = document.createElementNS(svgNS, 'polygon');
      poly.setAttribute('points', points);
      poly.setAttribute('fill', color);
      // Tooltip
      const title = document.createElementNS(svgNS, 'title');
      title.textContent = `${node.name}\nImpact: ${node.impact || '0'}${node.desc ? '\n' + node.desc : ''}`;
      poly.appendChild(title);
      svg.appendChild(poly);
      // Label left of triangle
      const text = document.createElementNS(svgNS, 'text');
      text.setAttribute('x', x - size - 4);
      text.setAttribute('y', y);
      text.setAttribute('text-anchor', 'end');
      text.setAttribute('alignment-baseline', 'middle');
      text.setAttribute('fill', 'var(--text-primary)');
      text.setAttribute('font-size', '10');
      text.textContent = node.name;
      svg.appendChild(text);
    });
    // Draw support nodes
    supportPos.forEach(({ x, y, node }) => {
      const r = circleRadius(node.degree);
      const color = impactColor(node.maxImpact);
      const circle = document.createElementNS(svgNS, 'circle');
      circle.setAttribute('cx', x);
      circle.setAttribute('cy', y);
      circle.setAttribute('r', r);
      circle.setAttribute('fill', color || '#9aa0a6');
      const title = document.createElementNS(svgNS, 'title');
      title.textContent = `${node.name}\nLiens: ${node.degree}\nMax impact supporté: ${node.maxImpact || '0'}`;
      circle.appendChild(title);
      svg.appendChild(circle);
      // Label right of circle
      const text = document.createElementNS(svgNS, 'text');
      text.setAttribute('x', x + r + 4);
      text.setAttribute('y', y);
      text.setAttribute('text-anchor', 'start');
      text.setAttribute('alignment-baseline', 'middle');
      text.setAttribute('fill', 'var(--text-primary)');
      text.setAttribute('font-size', '10');
      text.textContent = node.name;
      svg.appendChild(text);
    });
    container.appendChild(svg);
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
    // For SSI metrics (maturité, confiance) we invert colours: 1=red -> 4=green
    const ssiColor = (lvl) => {
      switch (parseInt(lvl, 10)) {
        case 1: return '#e63946'; // red
        case 2: return '#f4a261'; // orange
        case 3: return '#e9c46a'; // yellow
        case 4: return '#2a9d8f'; // green
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
        { value: 'beneficiaire', label: 'Bénéficiaire' }
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
        saveAnalyses();
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
      // Niveau SSI (computed)
      td = document.createElement('td');
      td.className = 'niveau-cell';
      tr.appendChild(td);
      // Indice de menace (computed)
      td = document.createElement('td');
      td.className = 'indice-cell';
      tr.appendChild(td);
      // Actions
      td = document.createElement('td');
      const delBtn = document.createElement('button');
      delBtn.className = 'delete-item';
      delBtn.textContent = '×';
      delBtn.title = 'Supprimer cette partie prenante';
      delBtn.addEventListener('click', () => {
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
    // Helper to update exposition, niveau SSI and indice cells
    function updateDerivedCells(row, entry) {
      const expo = (entry.dependance || 1) * (entry.penetration || 1);
      const niveau = (entry.maturite || 1) * (entry.confiance || 1);
      const indice = niveau ? (expo / niveau) : 0;
      const expoCell = row.querySelector('.expo-cell');
      const niveauCell = row.querySelector('.niveau-cell');
      const indiceCell = row.querySelector('.indice-cell');
      if (expoCell) {
        expoCell.textContent = `${expo}`;
        const bucket = pertinenceBucketForExpo(expo);
        expoCell.style.backgroundColor = levelColor(bucket);
      }
      if (niveauCell) {
        niveauCell.textContent = `${niveau}`;
        const bucket = pertinenceBucketForExpo(niveau);
        niveauCell.style.backgroundColor = ssiColor(bucket);
      }
      if (indiceCell) {
        indiceCell.textContent = indice.toFixed(2);
        // Colour based on indice: low values green, high values red
        let indLvl;
        if (indice >= 4) indLvl = 4;
        else if (indice >= 3) indLvl = 3;
        else if (indice >= 2) indLvl = 2;
        else indLvl = 1;
        indiceCell.style.backgroundColor = levelColor(indLvl);
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
      return { value: ev.id, label: ev.evenement || ev.ref || 'Évènement' };
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
      addEvBtn.textContent = '+ Ajouter';
      addEvBtn.addEventListener('click', () => {
        const available = eventOptions.filter(opt => !item.eventIds.includes(opt.value));
        if (available.length === 0) {
          alert('Aucun évènement redouté disponible à ajouter.');
          return;
        }
        const msg = 'Sélectionnez un évènement :\n' + available.map((opt, i) => `${i + 1}. ${opt.label}`).join('\n');
        const choice = prompt(msg);
        if (choice === null) return;
        const idxChoice = parseInt(choice, 10) - 1;
        if (!isNaN(idxChoice) && idxChoice >= 0 && idxChoice < available.length) {
          item.eventIds.push(available[idxChoice].value);
          saveAnalyses();
          renderStrategies();
        }
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
  }

  // Add resizer handles to the strategies table
  function addStrategiesTableResizers() {
    const table = document.getElementById('strategies-table');
    if (!table) return;
    const headers = table.querySelectorAll('th');
    headers.forEach((th, i) => {
      // Do not add resizer to last column (Actions)
      if (i === headers.length - 1) return;
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
        if (!Array.isArray(item.connaitre)) item.connaitre = [];
        if (!Array.isArray(item.rester)) item.rester = [];
        if (!Array.isArray(item.trouver)) item.trouver = [];
        if (!Array.isArray(item.exploiter)) item.exploiter = [];
        if (!Array.isArray(item.risks)) item.risks = [];
        // vraisemblance and gravité levels
        if (!item.vraisemblance) item.vraisemblance = 1;
        if (!item.gravite) item.gravite = 1;
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
        // Helper to render kill chain stage cell
        function renderStageCell(stageKey) {
          const cell = document.createElement('td');
          const wrapper = document.createElement('div');
          wrapper.className = 'assoc-cell';
          // remove existing
          (item[stageKey] || []).forEach((step, sIdx) => {
            const tag = document.createElement('span');
            tag.className = 'assoc-item';
            tag.textContent = step;
            const rmBtn = document.createElement('button');
            rmBtn.className = 'remove-assoc';
            rmBtn.textContent = '×';
            rmBtn.title = 'Retirer cette étape';
            rmBtn.addEventListener('click', () => {
              item[stageKey].splice(sIdx, 1);
              saveAnalyses();
              renderSO();
            });
            tag.appendChild(rmBtn);
            wrapper.appendChild(tag);
          });
          const addBtn = document.createElement('button');
          addBtn.className = 'add-assoc-btn';
          addBtn.textContent = '+ Ajouter';
          addBtn.addEventListener('click', () => {
            const val = prompt('Décrire une étape pour « ' + stageKey + ' » :');
            if (val) {
              item[stageKey].push(val);
              saveAnalyses();
              renderSO();
            }
          });
          wrapper.appendChild(addBtn);
          cell.appendChild(wrapper);
          return cell;
        }
        // Connaître
        tr.appendChild(renderStageCell('connaitre'));
        // Rester
        tr.appendChild(renderStageCell('rester'));
        // Trouver
        tr.appendChild(renderStageCell('trouver'));
        // Exploiter
        tr.appendChild(renderStageCell('exploiter'));
        // Risks cell
        td = document.createElement('td');
        const riskCell = document.createElement('div');
        riskCell.className = 'assoc-cell';
        item.risks.forEach((rk, rIdx) => {
          const tag = document.createElement('span');
          tag.className = 'assoc-item';
          tag.textContent = rk.name;
          const rm = document.createElement('button');
          rm.className = 'remove-assoc';
          rm.textContent = '×';
          rm.title = 'Retirer ce risque';
          rm.addEventListener('click', () => {
            item.risks.splice(rIdx, 1);
            saveAnalyses();
            renderSO();
          });
          tag.appendChild(rm);
          riskCell.appendChild(tag);
        });
        const addRiskBtn = document.createElement('button');
        addRiskBtn.className = 'add-assoc-btn';
        addRiskBtn.textContent = '+ Ajouter';
        addRiskBtn.addEventListener('click', () => {
          // Open the MITRE/OWASP risk selection modal for this scenario
          openRiskModal(item);
        });
        riskCell.appendChild(addRiskBtn);
        td.appendChild(riskCell);
        tr.appendChild(td);
        // Vraisemblance select
        td = document.createElement('td');
        const vraSelect = document.createElement('select');
        [1,2,3,4].forEach(v => {
          const o = document.createElement('option');
          o.value = v;
          o.textContent = `${v}`;
          if (v === item.vraisemblance) o.selected = true;
          vraSelect.appendChild(o);
        });
        vraSelect.style.backgroundColor = levelColor(item.vraisemblance);
        vraSelect.onchange = (e) => {
          item.vraisemblance = parseInt(e.target.value, 10);
          e.target.style.backgroundColor = levelColor(item.vraisemblance);
          saveAnalyses();
          updateAtelier4Chart();
        };
        td.appendChild(vraSelect);
        tr.appendChild(td);
        // Gravité select
        td = document.createElement('td');
        const grSelect = document.createElement('select');
        [1,2,3,4].forEach(v => {
          const o = document.createElement('option');
          o.value = v;
          o.textContent = `${v}`;
          if (v === item.gravite) o.selected = true;
          grSelect.appendChild(o);
        });
        grSelect.style.backgroundColor = levelColor(item.gravite);
        grSelect.onchange = (e) => {
          item.gravite = parseInt(e.target.value, 10);
          e.target.style.backgroundColor = levelColor(item.gravite);
          saveAnalyses();
          updateAtelier4Chart();
        };
        td.appendChild(grSelect);
        tr.appendChild(td);
        // Actions: delete
        td = document.createElement('td');
        const delBtn = document.createElement('button');
        delBtn.className = 'delete-item';
        delBtn.textContent = '×';
        delBtn.title = 'Supprimer ce scénario';
        delBtn.addEventListener('click', () => {
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
            connaitre: [],
            rester: [],
            trouver: [],
            exploiter: [],
            risks: [],
            vraisemblance: 1,
            gravite: 1
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
      if (!Array.isArray(item.sourceIds)) item.sourceIds = [];
      const el = document.createElement('div');
      el.className = 'item';
      const delBtn = document.createElement('button');
      delBtn.className = 'delete-item';
      delBtn.innerHTML = '×';
      delBtn.title = 'Supprimer ce risque';
      delBtn.addEventListener('click', () => {
        analysis.data.risques.splice(idx, 1);
        saveAnalyses();
        renderRisques();
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
      // Fields for titre, description, indice, vraisemblance, gravite, mesures
      el.appendChild(createInput('Titre du risque', 'text', item.titre, (v) => {
        item.titre = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Description', 'textarea', item.description, (v) => {
        item.description = v;
        saveAnalyses();
      }));
      el.appendChild(createInput('Indice', 'text', item.indice, (v) => {
        item.indice = v;
        saveAnalyses();
        updateAtelier5Chart();
      }));
      el.appendChild(createInput('Vraisemblance', 'text', item.vraisemblance, (v) => {
        item.vraisemblance = v;
        saveAnalyses();
        updateAtelier5Chart();
      }));
      el.appendChild(createInput('Gravité', 'text', item.gravite, (v) => {
        item.gravite = v;
        saveAnalyses();
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
      addRow.appendChild(addTd);
      actTable.appendChild(addRow);
      tdActions.appendChild(actTable);
      tr.appendChild(tdActions);
      body.appendChild(tr);
    });
  }

  // Render actions for supports: allow user to add rows for each selected support
  function renderSupportActions() {
    const body = document.getElementById('support-actions-body');
    if (!body) return;
    body.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) return;
    // Ensure data array
    if (!Array.isArray(analysis.data.actionsSupports)) analysis.data.actionsSupports = [];
    // Build list of available supports from missions
    const supportsSet = new Set();
    (analysis.data.missions || []).forEach(mis => {
      if (Array.isArray(mis.supports)) {
        mis.supports.forEach(s => {
          if (s && (s.name || s.denom)) supportsSet.add(s.name || s.denom);
        });
      }
    });
    const supportOptions = Array.from(supportsSet);
    // Render each row from actionsSupports
    analysis.data.actionsSupports.forEach((row, rowIndex) => {
      const tr = document.createElement('tr');
      // Support select
      const tdSupport = document.createElement('td');
      const sel = document.createElement('select');
      sel.className = 'form-select';
      sel.innerHTML = '<option value="">--Sélectionner--</option>' + supportOptions.map(opt => `<option value="${opt}" ${row.supportName===opt?'selected':''}>${opt}</option>`).join('');
      sel.addEventListener('change', (e) => {
        row.supportName = e.target.value;
        saveAnalyses();
        renderSupportActions();
        renderPlanActions();
      });
      tdSupport.appendChild(sel);
      tr.appendChild(tdSupport);
      // Actions cell: nested table
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
          row.actions.splice(aIdx, 1);
          saveAnalyses();
          renderSupportActions();
          renderPlanActions();
        });
        tdA.appendChild(delBtn);
        ar.appendChild(tdA);
        actTable.appendChild(ar);
      });
      // Add new action row
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
        analysis.data.actionsSupports.splice(rowIndex, 1);
        saveAnalyses();
        renderSupportActions();
        renderPlanActions();
      });
      tdDel.appendChild(delRow);
      tr.appendChild(tdDel);
      body.appendChild(tr);
    });
    // Add row button is outside in HTML
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
          row.actions.splice(aIdx, 1);
          saveAnalyses();
          renderPartiesActions();
          renderPlanActions();
        });
        tdA.appendChild(delBtn);
        ar.appendChild(tdA);
        actTable.appendChild(ar);
      });
      // Row to add new action
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
        analysis.data.actionsParties.splice(rowIndex, 1);
        saveAnalyses();
        renderPartiesActions();
        renderPlanActions();
      });
      tdDel.appendChild(delRow);
      tr.appendChild(tdDel);
      body.appendChild(tr);
    });
  }

  // Render actions for risks: show each risk and allow adding actions and residual levels
  function renderRisquesActions() {
    const body = document.getElementById('risques-actions-body');
    if (!body) return;
    body.innerHTML = '';
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) return;
    if (!Array.isArray(analysis.data.actionsRisques)) analysis.data.actionsRisques = [];
    // Gather risks from operational scenarios (atelier 4) for reference
    const riskMap = new Map();
    (analysis.data.so || []).forEach(scenario => {
      const vs = scenario.vraisemblance || scenario.vraisemblance;
      const gr = scenario.gravite || scenario.gravité || scenario.grav;
      (scenario.risks || []).forEach(riskObj => {
        const name = riskObj.name || '';
        if (!name) return;
        const cur = riskMap.get(name) || { name, vraisemblance: riskObj.vraisemblance || vs || 1, gravite: riskObj.gravite || gr || 1 };
        cur.vraisemblance = Math.max(cur.vraisemblance, riskObj.vraisemblance || vs || 1);
        cur.gravite = Math.max(cur.gravite, riskObj.gravite || gr || 1);
        riskMap.set(name, cur);
      });
    });
    analysis.data.actionsRisques.forEach(row => {
      const risk = riskMap.get(row.riskName) || { name: row.riskName, vraisemblance: row.residualV || 1, gravite: row.residualG || 1 };
      const tr = document.createElement('tr');
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
      tdVr.textContent = risk.vraisemblance;
      tr.appendChild(tdVr);
      const tdGr = document.createElement('td');
      tdGr.textContent = risk.gravite;
      tr.appendChild(tdGr);
      // Residual vraisemblance select
      const tdResVr = document.createElement('td');
      const selVr = document.createElement('select');
      selVr.className = 'form-select';
      for (let i = 1; i <= 4; i++) {
        const opt = document.createElement('option');
        opt.value = i;
        opt.textContent = i;
        if (row.residualV == i) opt.selected = true;
        selVr.appendChild(opt);
      }
      selVr.addEventListener('change', (e) => {
        row.residualV = parseInt(e.target.value, 10);
        saveAnalyses();
      });
      tdResVr.appendChild(selVr);
      tr.appendChild(tdResVr);
      // Residual gravite select
      const tdResGr = document.createElement('td');
      const selGr = document.createElement('select');
      selGr.className = 'form-select';
      for (let i = 1; i <= 4; i++) {
        const opt = document.createElement('option');
        opt.value = i;
        opt.textContent = i;
        if (row.residualG == i) opt.selected = true;
        selGr.appendChild(opt);
      }
      selGr.addEventListener('change', (e) => {
        row.residualG = parseInt(e.target.value, 10);
        saveAnalyses();
      });
      tdResGr.appendChild(selGr);
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
          row.actions.splice(aIdx, 1);
          saveAnalyses();
          renderRisquesActions();
          renderPlanActions();
        });
        tdA.appendChild(delBtn);
        ar.appendChild(tdA);
        actTable.appendChild(ar);
      });
      // Row to add new action
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
      addRow.appendChild(addTd);
      actTable.appendChild(addRow);
      tdActions.appendChild(actTable);
      tr.appendChild(tdActions);
      body.appendChild(tr);
    });
  }

  // Aggregate all actions and render plan table + gantt chart
  function renderPlanActions() {
    const body = document.getElementById('plan-actions-body');
    const canvas = document.getElementById('gantt-chart');
    if (!body || !canvas) return;
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
      const srcName = row.supportName || 'Support';
      (row.actions || []).forEach(act => {
        actions.push({
          name: act.name,
          source: 'Support: ' + srcName,
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
    // Draw gantt chart
    drawGanttChart(canvas, actions);
  }

  // Draw a simple Gantt chart on a canvas from a list of actions with start/end dates
  function drawGanttChart(canvas, actions) {
    const ctx = canvas.getContext('2d');
    // Clear
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    if (!actions || actions.length === 0) return;
    // Parse dates
    const parsed = actions.map(act => {
      const start = act.start ? new Date(act.start) : null;
      const end = act.end ? new Date(act.end) : null;
      return { act, start, end };
    }).filter(item => item.start && item.end && !isNaN(item.start.getTime()) && !isNaN(item.end.getTime()));
    if (parsed.length === 0) return;
    // Determine min and max dates
    let minDate = parsed[0].start;
    let maxDate = parsed[0].end;
    parsed.forEach(it => {
      if (it.start < minDate) minDate = it.start;
      if (it.end > maxDate) maxDate = it.end;
    });
    const range = maxDate - minDate;
    const rowHeight = 18;
    const yMargin = 20;
    const xMargin = 60;
    canvas.height = yMargin * 2 + rowHeight * parsed.length;
    // Draw axes
    ctx.strokeStyle = 'rgba(255,255,255,0.2)';
    ctx.beginPath();
    ctx.moveTo(xMargin, yMargin);
    ctx.lineTo(xMargin, canvas.height - yMargin);
    ctx.lineTo(canvas.width - 10, canvas.height - yMargin);
    ctx.stroke();
    parsed.forEach((item, index) => {
      const x = xMargin + ((item.start - minDate) / range) * (canvas.width - xMargin - 20);
      const w = ((item.end - item.start) / range) * (canvas.width - xMargin - 20);
      const y = yMargin + index * rowHeight + 4;
      ctx.fillStyle = 'rgba(77,163,255,0.6)';
      ctx.fillRect(x, y, w, rowHeight - 8);
      ctx.fillStyle = 'var(--text-primary)';
      ctx.font = '12px sans-serif';
      ctx.fillText(item.act.name, 2, y + rowHeight - 12);
    });
    // Draw date labels on x-axis
    const tickCount = Math.min(6, parsed.length > 0 ? parsed.length + 1 : 1);
    for (let i = 0; i < tickCount; i++) {
      const t = minDate.getTime() + (range * i) / (tickCount - 1);
      const date = new Date(t);
      const x = xMargin + (i / (tickCount - 1)) * (canvas.width - xMargin - 20);
      const label = date.toISOString().slice(0, 10);
      ctx.fillStyle = 'var(--text-secondary)';
      ctx.font = '10px sans-serif';
      ctx.fillText(label, x - 20, canvas.height - 5);
    }
  }

  // ----- Chart drawing functions (simple bar and radar charts)
  function clearCanvas(canvas) {
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
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
    // Draw axes
    ctx.strokeStyle = 'rgba(200,200,200,0.5)';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(margin, margin);
    ctx.lineTo(margin, height - margin);
    ctx.lineTo(width - margin, height - margin);
    ctx.stroke();
    ctx.font = '12px sans-serif';
    ctx.fillStyle = 'var(--text-secondary)';
    // Draw bars
    data.forEach((value, i) => {
      const barHeight = (value / maxVal) * (height - margin * 2);
      const x = margin + i * barWidth + barWidth * 0.2;
      const y = height - margin - barHeight;
      const w = barWidth * 0.6;
      const color = colors && colors[i] ? colors[i] : 'rgba(77,163,255,0.8)';
      ctx.fillStyle = color;
      ctx.fillRect(x, y, w, barHeight);
      // Value label
      ctx.fillStyle = 'var(--text-primary)';
      ctx.textAlign = 'center';
      ctx.fillText(value, x + w / 2, y - 4);
      // Category label
      ctx.fillStyle = 'var(--text-secondary)';
      ctx.save();
      ctx.translate(x + w / 2, height - margin + 14);
      ctx.rotate(-Math.PI / 4);
      ctx.fillText(labels[i], 0, 0);
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
    // The bar chart for Atelier 1 has been replaced by a network graph.
    updateAtelier1Graph();
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

  function updateAtelier3Chart() {
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) return;
    const canvas = document.getElementById('atelier3-chart');
    // Determine which subtab is active (carto or strategies)
    const cartoActive = document.getElementById('atelier3-carto-tab')?.classList.contains('active');
    if (cartoActive) {
      const ppc = analysis.data.ppc || [];
      let sumExpo = 0;
      let sumNiveau = 0;
      let sumIndice = 0;
      let n = 0;
      ppc.forEach(item => {
        const dep = parseInt(item.dependance, 10) || 1;
        const pen = parseInt(item.penetration, 10) || 1;
        const mat = parseInt(item.maturite, 10) || 1;
        const conf = parseInt(item.confiance, 10) || 1;
        const expo = dep * pen;
        const niveau = mat * conf;
        const indice = niveau ? expo / niveau : 0;
        sumExpo += expo;
        sumNiveau += niveau;
        sumIndice += indice;
        n++;
      });
      const avgExpo = n ? sumExpo / n : 0;
      const avgNiveau = n ? sumNiveau / n : 0;
      const avgIndice = n ? sumIndice / n : 0;
      if (n > 0) {
        drawRadarChart(canvas, ['Exposition', 'Niveau SSI', 'Indice'], {
          data: [avgExpo, avgNiveau, avgIndice],
          color: 'rgba(77,163,255,0.4)'
        });
      } else {
        clearCanvas(canvas);
      }
    } else {
      // Scenario subtab: use old PP list (niveauSSI & indiceMenace)
      const ppList = analysis.data.pp || [];
      let sumSSI = 0;
      let sumMenace = 0;
      let count = 0;
      ppList.forEach(item => {
        if (typeof item.niveauSSI === 'number') {
          sumSSI += item.niveauSSI;
          count++;
        }
        if (typeof item.indiceMenace === 'number') {
          sumMenace += item.indiceMenace;
        }
      });
      const avgSSI = count ? sumSSI / count : 0;
      const avgMenace = count ? sumMenace / count : 0;
      if (count > 0) {
        drawRadarChart(canvas, ['Niveau SSI', 'Indice de menace'], {
          data: [avgSSI, avgMenace],
          color: 'rgba(77,163,255,0.4)'
        });
      } else {
        clearCanvas(canvas);
      }
    }
  }

  function updateAtelier4Chart() {
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) return;
    const ops = analysis.data.so || [];
    // Count scenarios by their vraisemblance level (1..4).  If no scenarios
    // exist, the chart will be empty.
    const counts = { 1: 0, 2: 0, 3: 0, 4: 0 };
    ops.forEach(item => {
      const lvl = parseInt(item.vraisemblance, 10);
      if (lvl >= 1 && lvl <= 4) counts[lvl] = (counts[lvl] || 0) + 1;
    });
    const labels = ['1','2','3','4'];
    const data = labels.map(l => counts[parseInt(l,10)]);
    // Colours based on level for each bar
    const levelColor = (lvl) => {
      switch (parseInt(lvl, 10)) {
        case 1: return '#2a9d8f';
        case 2: return '#e9c46a';
        case 3: return '#f4a261';
        case 4: return '#e63946';
        default: return 'rgba(77,163,255,0.8)';
      }
    };
    const colors = labels.map(l => levelColor(l));
    const canvas = document.getElementById('atelier4-chart');
    drawBarChart(canvas, labels, data, colors);
  }

  function updateAtelier5Chart() {
    const analysis = analyses[currentIndex];
    if (!analysis || !analysis.data) return;
    const risques = analysis.data.risques || [];
    const counts = {};
    risques.forEach(item => {
      const key = (item.gravite || '').trim() || 'Non précisé';
      counts[key] = (counts[key] || 0) + 1;
    });
    const labels = Object.keys(counts);
    const data = labels.map(l => counts[l]);
    const canvas = document.getElementById('atelier5-chart');
    drawBarChart(canvas, labels, data);
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
        updateAtelier1Graph();
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
        analysis.data.actionsSupports.push({ supportName: '', actions: [] });
        saveAnalyses();
        renderSupportActions();
      });
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
    document.getElementById('add-pp-btn').addEventListener('click', () => {
      const analysis = analyses[currentIndex];
      if (!analysis.data) analysis.data = {};
      if (!analysis.data.pp) analysis.data.pp = [];
      analysis.data.pp.push({ id: uid(), categorie:'', nom:'', description:'', niveauSSI:0, indiceMenace:0 });
      saveAnalyses();
      renderPP();
    });
    document.getElementById('add-ss-btn').addEventListener('click', () => {
      const analysis = analyses[currentIndex];
      if (!analysis.data) analysis.data = {};
      if (!analysis.data.ss) analysis.data.ss = [];
      analysis.data.ss.push({ id: uid(), source:'', objectif:'', vraisemblance:'', gravite:'' });
      saveAnalyses();
      renderSS();
    });
    document.getElementById('add-so-btn').addEventListener('click', () => {
      const analysis = analyses[currentIndex];
      if (!analysis.data) analysis.data = {};
      if (!analysis.data.so) analysis.data.so = [];
      analysis.data.so.push({ id: uid(), chemin:'', vraisemblanceGlobale:'' });
      saveAnalyses();
      renderSO();
    });
    document.getElementById('add-risque-btn').addEventListener('click', () => {
      const analysis = analyses[currentIndex];
      if (!analysis.data) analysis.data = {};
      if (!analysis.data.risques) analysis.data.risques = [];
      const defaultMission = (analysis.data.missions && analysis.data.missions[0]) ? analysis.data.missions[0].id : '';
      const defaultEvent = (analysis.data.events && analysis.data.events[0]) ? analysis.data.events[0].id : '';
      const defaultScenario = (analysis.data.so && analysis.data.so[0]) ? analysis.data.so[0].id : '';
      analysis.data.risques.push({
        id: uid(),
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
    });

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
            const grid = document.querySelector('#atelier1 .atelier-grid');
            if (grid) grid.classList.remove('gap-active');
          } else if (target === 'gap') {
            graphEl.style.display = 'none';
            gapChartEl.style.display = '';
            // Draw the GAP chart when switching to this tab
            updateGapChart();
            // make the grid full width for gap analysis
            const grid = document.querySelector('#atelier1 .atelier-grid');
            if (grid) grid.classList.add('gap-active');
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

        // When switching between cartography and strategies, toggle a class
        // on the Atelier 3 grid to collapse the layout to a single column
        // for the cartography view.  This ensures the chart appears below
        // the table rather than beside it.
        const grid3 = document.querySelector('#atelier3 .atelier-grid');
        if (grid3) {
          if (target === 'carto') {
            grid3.classList.add('carto-active');
          } else {
            grid3.classList.remove('carto-active');
          }
        }
      });
    });
  }

  function setupSidebarToggle() {
    const btn = document.getElementById('sidebar-toggle');
    if (!btn) return;
    const sidebar = document.getElementById('sidebar');
    btn.addEventListener('click', () => {
      sidebar.classList.toggle('collapsed');
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
        // When switching to the plan tab, refresh the plan view and gantt chart
        if (target === 'plan') {
          renderPlanActions();
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
            analyses = json;
          } else if (typeof json === 'object') {
            analyses.push(json);
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
    // Attempt to restore the previously selected analysis.  The ID of
    // the last selected analysis is persisted in localStorage under
    // `ebiosCurrentAnalysisId`.  If it exists and matches one of
    // the loaded analyses, use that index; otherwise default to the
    // first analysis.
    let savedIndex = -1;
    try {
      const savedId = localStorage.getItem('ebiosCurrentAnalysisId');
      if (savedId) {
        savedIndex = analyses.findIndex(a => a && a.id === savedId);
      }
    } catch (e) {
      savedIndex = -1;
    }
    renderAnalysisList();
    setupAnalysisControls();
    setupNavigation();
    setupSidebarToggle();
    setupAddButtons();
    setupMitreImport();
    setupKillChainToggle();
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
