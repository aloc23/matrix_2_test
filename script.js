document.addEventListener('DOMContentLoaded', function() {
  // -------------------- State --------------------
  let rawData = [];
  let mappedData = [];
  let mappingConfigured = false;
  let config = {
    weekLabelRow: 0,
    weekColStart: 0,
    weekColEnd: 0,
    firstDataRow: 1,
    lastDataRow: 1
  };
  let weekLabels = [];
  let weekCheckboxStates = [];
  let repaymentRows = [];
  let openingBalance = 0;
  let loanOutstanding = 0;
  let roiInvestment = 120000;
  let roiInterest = 0.0;

  // ROI week/date mapping
  let weekStartDates = [];
  let investmentWeekIndex = 0;

  // --- ROI SUGGESTION STATE ---
  let showSuggestions = false;
  let suggestedRepayments = null;
  let achievedSuggestedIRR = null;
  
  // --- TARGET IRR/NPV SETTINGS ---
  let targetIRR = 0.20; // Default 20%
  let installmentCount = 12; // Default 12 installments
  let liveUpdateEnabled = true;
  
  // --- BUFFER/GAP SETTINGS ---
  let bufferSettings = {
    type: 'none', // none, 2weeks, 1month, 2months, quarter, custom
    customWeeks: 1
  };

  // --- Chart.js chart instances for destroy ---
  let mainChart = null;
  let roiPieChart = null;
  let roiLineChart = null;
  window.tornadoChartObj = null; // global for tornado chart
  let summaryChart = null;

  // -------------------- Tabs & UI Interactions --------------------
  function setupTabs() {
    document.querySelectorAll('.tabs button').forEach(btn => {
      btn.addEventListener('click', function() {
        document.querySelectorAll('.tabs button').forEach(b => b.classList.remove('active'));
        this.classList.add('active');
        document.querySelectorAll('.tab-content').forEach(sec => sec.classList.remove('active'));
        var tabId = btn.getAttribute('data-tab');
        var panel = document.getElementById(tabId);
        if (panel) panel.classList.add('active');
        setTimeout(() => {
          updateAllTabs();
        }, 50);
      });
    });
    document.querySelectorAll('.subtabs button').forEach(btn => {
      btn.addEventListener('click', function() {
        document.querySelectorAll('.subtabs button').forEach(b => b.classList.remove('active'));
        this.classList.add('active');
        document.querySelectorAll('.subtab-panel').forEach(sec => sec.classList.remove('active'));
        var subtabId = 'subtab-' + btn.getAttribute('data-subtab');
        var subpanel = document.getElementById(subtabId);
        if (subpanel) subpanel.classList.add('active');
        setTimeout(updateAllTabs, 50);
      });
    });
    document.querySelectorAll('.collapsible-header').forEach(btn => {
      btn.addEventListener('click', function() {
        var content = btn.nextElementSibling;
        var caret = btn.querySelector('.caret');
        if (content && content.classList.contains('active')) {
          content.classList.remove('active');
          if (caret) caret.style.transform = 'rotate(-90deg)';
        } else if (content) {
          content.classList.add('active');
          if (caret) caret.style.transform = 'none';
        }
      });
    });
  }
  setupTabs();
  setupTargetIrrModal();
  setupBufferModal();

  // -------------------- Spreadsheet Upload & Mapping --------------------
  function setupSpreadsheetUpload() {
    var spreadsheetUpload = document.getElementById('spreadsheetUpload');
    if (spreadsheetUpload) {
      spreadsheetUpload.addEventListener('change', function(event) {
        const reader = new FileReader();
        reader.onload = function (e) {
          const dataArr = new Uint8Array(e.target.result);
          const workbook = XLSX.read(dataArr, { type: 'array' });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
          if (!json.length) return;
          rawData = json;
          mappedData = json;
          autoDetectMapping(mappedData);
          mappingConfigured = false;
          renderMappingPanel(mappedData);
          updateWeekLabels();
          updateAllTabs();
        };
        reader.readAsArrayBuffer(event.target.files[0]);
      });
    }
  }
  setupSpreadsheetUpload();

  function renderMappingPanel(allRows) {
    const panel = document.getElementById('mappingPanel');
    if (!panel) return;
    panel.innerHTML = '';

    function drop(label, id, max, sel, onChange, items) {
      let lab = document.createElement('label');
      lab.textContent = label;
      let selElem = document.createElement('select');
      selElem.className = 'mapping-dropdown';
      for (let i = 0; i < max; i++) {
        let opt = document.createElement('option');
        opt.value = i;
        let textVal = items && items[i] ? items[i] : (allRows[i] ? allRows[i].slice(0,8).join(',').slice(0,32) : '');
        opt.textContent = `${id==='row'?'Row':'Col'} ${i+1}: ${textVal}`;
        selElem.appendChild(opt);
      }
      selElem.value = sel;
      selElem.onchange = function() { onChange(parseInt(this.value,10)); };
      lab.appendChild(selElem);
      panel.appendChild(lab);
    }

    drop('Which row contains week labels? ', 'row', Math.min(allRows.length, 30), config.weekLabelRow, v => { config.weekLabelRow = v; updateWeekLabels(); renderMappingPanel(allRows); updateAllTabs(); });
    panel.appendChild(document.createElement('br'));

    let weekRow = allRows[config.weekLabelRow] || [];
    drop('First week column: ', 'col', weekRow.length, config.weekColStart, v => { config.weekColStart = v; updateWeekLabels(); renderMappingPanel(allRows); updateAllTabs(); }, weekRow);
    drop('Last week column: ', 'col', weekRow.length, config.weekColEnd, v => { config.weekColEnd = v; updateWeekLabels(); renderMappingPanel(allRows); updateAllTabs(); }, weekRow);
    panel.appendChild(document.createElement('br'));

    drop('First data row: ', 'row', allRows.length, config.firstDataRow, v => { config.firstDataRow = v; renderMappingPanel(allRows); updateAllTabs(); });
    drop('Last data row: ', 'row', allRows.length, config.lastDataRow, v => { config.lastDataRow = v; renderMappingPanel(allRows); updateAllTabs(); });
    panel.appendChild(document.createElement('br'));

    // Opening balance input
    let obDiv = document.createElement('div');
    obDiv.innerHTML = `Opening Balance: <input type="number" id="openingBalanceInput" value="${openingBalance}" style="width:120px;">`;
    panel.appendChild(obDiv);
    setTimeout(() => {
      let obInput = document.getElementById('openingBalanceInput');
      if (obInput) obInput.oninput = function() {
        openingBalance = parseFloat(obInput.value) || 0;
        updateAllTabs();
        renderMappingPanel(allRows);
      };
    }, 0);

    // Reset button for mapping
    const resetBtn = document.createElement('button');
    resetBtn.textContent = "Reset Mapping";
    resetBtn.style.marginLeft = '10px';
    resetBtn.onclick = function() {
      autoDetectMapping(allRows);
      weekCheckboxStates = weekLabels.map(()=>true);
      openingBalance = 0;
      renderMappingPanel(allRows);
      updateWeekLabels();
      updateAllTabs();
    };
    panel.appendChild(resetBtn);

    // Collapsible Week filter UI
    if (weekLabels.length) {
      const weekFilterDiv = document.createElement('div');
      weekFilterDiv.className = "collapsible-week-filter";

      // Collapsible header
      const collapseBtn = document.createElement('button');
      collapseBtn.type = 'button';
      collapseBtn.className = 'collapse-toggle';
      collapseBtn.innerHTML = `<span class="caret" style="display:inline-block;transition:transform 0.2s;margin-right:6px;">&#9654;</span>Filter week columns to include:`;
      collapseBtn.style.marginBottom = '10px';
      collapseBtn.style.background = 'none';
      collapseBtn.style.color = '#1976d2';
      collapseBtn.style.fontWeight = 'bold';
      collapseBtn.style.fontSize = '1.06em';
      collapseBtn.style.border = 'none';
      collapseBtn.style.cursor = 'pointer';
      collapseBtn.style.outline = 'none';
      collapseBtn.style.padding = '4px 0';

      // Collapsible content
      const collapsibleContent = document.createElement('div');
      collapsibleContent.className = "week-checkbox-collapsible-content";
      collapsibleContent.style.display = 'none';
      collapsibleContent.style.margin = '14px 0 4px 0';

      // Buttons
      const selectAllBtn = document.createElement('button');
      selectAllBtn.textContent = "Select All";
      selectAllBtn.type = 'button';
      selectAllBtn.style.marginRight = '8px';
      selectAllBtn.onclick = function() {
        weekCheckboxStates = weekCheckboxStates.map(()=>true);
        updateAllTabs();
        renderMappingPanel(allRows);
      };
      const deselectAllBtn = document.createElement('button');
      deselectAllBtn.textContent = "Deselect All";
      deselectAllBtn.type = 'button';
      deselectAllBtn.onclick = function() {
        weekCheckboxStates = weekCheckboxStates.map(()=>false);
        updateAllTabs();
        renderMappingPanel(allRows);
      };
      collapsibleContent.appendChild(selectAllBtn);
      collapsibleContent.appendChild(deselectAllBtn);

      // Checkbox group
      const groupDiv = document.createElement('div');
      groupDiv.className = 'week-checkbox-group';
      groupDiv.style.marginTop = '8px';
      weekLabels.forEach((label, idx) => {
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = weekCheckboxStates[idx] !== false;
        cb.id = 'weekcol_cb_' + idx;
        cb.onchange = function() {
          weekCheckboxStates[idx] = cb.checked;
          updateAllTabs();
          renderMappingPanel(allRows);
        };
        const lab = document.createElement('label');
        lab.htmlFor = cb.id;
        lab.textContent = label;
        lab.style.marginRight = '13px';
        groupDiv.appendChild(cb);
        groupDiv.appendChild(lab);
      });
      collapsibleContent.appendChild(groupDiv);

      // Collapsible logic
      collapseBtn.addEventListener('click', function() {
        const isOpen = collapsibleContent.style.display !== 'none';
        collapsibleContent.style.display = isOpen ? 'none' : 'block';
        const caret = collapseBtn.querySelector('.caret');
        caret.style.transform = isOpen ? 'rotate(0)' : 'rotate(90deg)';
      });

      weekFilterDiv.appendChild(collapseBtn);
      weekFilterDiv.appendChild(collapsibleContent);
      panel.appendChild(weekFilterDiv);
    }

    // Save Mapping Button
    const saveBtn = document.createElement('button');
    saveBtn.textContent = "Save Mapping";
    saveBtn.style.margin = "10px 0";
    saveBtn.onclick = function() {
      mappingConfigured = true;
      updateWeekLabels();
      updateAllTabs();
      renderMappingPanel(allRows);
    };
    panel.appendChild(saveBtn);

    // Compact preview
    if (weekLabels.length && mappingConfigured) {
      const previewWrap = document.createElement('div');
      const compactTable = document.createElement('table');
      compactTable.className = "compact-preview-table";
      const tr1 = document.createElement('tr');
      tr1.appendChild(document.createElement('th'));
      getFilteredWeekIndices().forEach(fi => {
        const th = document.createElement('th');
        th.textContent = weekLabels[fi];
        tr1.appendChild(th);
      });
      compactTable.appendChild(tr1);
      const tr2 = document.createElement('tr');
      const lbl = document.createElement('td');
      lbl.textContent = "Bank Balance (rolling)";
      tr2.appendChild(lbl);
      let rolling = getRollingBankBalanceArr();
      getFilteredWeekIndices().forEach((fi, i) => {
        let bal = rolling[i];
        let td = document.createElement('td');
        td.textContent = isNaN(bal) ? '' : `€${Math.round(bal)}`;
        if (bal < 0) td.style.background = "#ffeaea";
        tr2.appendChild(td);
      });
      compactTable.appendChild(tr2);
      previewWrap.style.overflowX = "auto";
      previewWrap.appendChild(compactTable);
      panel.appendChild(previewWrap);
    }
  }

  function autoDetectMapping(sheet) {
    for (let r = 0; r < Math.min(sheet.length, 10); r++) {
      for (let c = 0; c < Math.min(sheet[r].length, 30); c++) {
        const val = (sheet[r][c] || '').toString().toLowerCase();
        if (/week\s*\d+/.test(val) || /week\s*\d+\/\d+/.test(val)) {
          config.weekLabelRow = r;
          config.weekColStart = c;
          let lastCol = c;
          while (
            lastCol < sheet[r].length &&
            ((sheet[r][lastCol] || '').toLowerCase().indexOf('week') >= 0 ||
            /^\d{1,2}\/\d{1,2}/.test(sheet[r][lastCol] || ''))
          ) {
            lastCol++;
          }
          config.weekColEnd = lastCol - 1;
          config.firstDataRow = r + 1;
          config.lastDataRow = sheet.length-1;
          return;
        }
      }
    }
    config.weekLabelRow = 4;
    config.weekColStart = 5;
    config.weekColEnd = Math.max(5, (sheet[4]||[]).length-1);
    config.firstDataRow = 6;
    config.lastDataRow = sheet.length-1;
  }

  function extractWeekStartDates(weekLabels, baseYear) {
    let currentYear = baseYear;
    let lastMonthIdx = -1;
    const months = [
      "jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"
    ];
    return weekLabels.map(label => {
      let match = label.match(/(\d{1,2})\s*([A-Za-z]{3,})/);
      if (!match) return null;
      let [_, day, monthStr] = match;
      let monthIdx = months.findIndex(m =>
        monthStr.toLowerCase().startsWith(m)
      );
      if (monthIdx === -1) return null;
      if (lastMonthIdx !== -1 && monthIdx < lastMonthIdx) currentYear++;
      lastMonthIdx = monthIdx;
      let date = new Date(currentYear, monthIdx, parseInt(day, 10));
      return date;
    });
  }

  function populateInvestmentWeekDropdown() {
    const dropdown = document.getElementById('investmentWeek');
    if (!dropdown) return;
    dropdown.innerHTML = '';
    weekLabels.forEach((label, i) => {
      const opt = document.createElement('option');
      let dateStr = weekStartDates[i]
        ? weekStartDates[i].toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })
        : 'N/A';
      opt.value = i;
      opt.textContent = `${label} (${dateStr})`;
      dropdown.appendChild(opt);
    });
    dropdown.value = investmentWeekIndex;
  }

  function updateWeekLabels() {
    let weekRow = mappedData[config.weekLabelRow] || [];
    weekLabels = weekRow.slice(config.weekColStart, config.weekColEnd+1).map(x => x || '');
    window.weekLabels = weekLabels; // make global for charts
    if (!weekCheckboxStates || weekCheckboxStates.length !== weekLabels.length) {
      weekCheckboxStates = weekLabels.map(() => true);
    }
    populateWeekDropdown(weekLabels);

    // ROI week start date integration. Use a default base year (2025) or prompt user for year.
    weekStartDates = extractWeekStartDates(weekLabels, 2025);
    populateInvestmentWeekDropdown();
  }

  function getFilteredWeekIndices() {
    if (weekCheckboxStates && weekCheckboxStates.length > 0) {
      return weekCheckboxStates.map((checked, idx) => checked ? idx : null).filter(idx => idx !== null);
    } else {
      // When no mapping is configured, return all week indices up to 52 weeks
      return Array.from({length: 52}, (_, i) => i);
    }
  }

  // -------------------- Calculation Helpers --------------------
  function getIncomeArr() {
    if (!mappedData || !mappingConfigured) return [];
    let arr = [];
    for (let w = 0; w < weekLabels.length; w++) {
      if (!weekCheckboxStates[w]) continue;
      let absCol = config.weekColStart + w;
      let sum = 0;
      for (let r = config.firstDataRow; r <= config.lastDataRow; r++) {
        let val = mappedData[r][absCol];
        if (typeof val === "string") val = val.replace(/,/g, '').replace(/€|\s/g,'');
        let num = parseFloat(val);
        if (!isNaN(num) && num > 0) sum += num;
      }
      arr[w] = sum;
    }
    return arr;
  }
  function getExpenditureArr() {
    if (!mappedData || !mappingConfigured) return [];
    let arr = [];
    for (let w = 0; w < weekLabels.length; w++) {
      if (!weekCheckboxStates[w]) continue;
      let absCol = config.weekColStart + w;
      let sum = 0;
      for (let r = config.firstDataRow; r <= config.lastDataRow; r++) {
        let val = mappedData[r][absCol];
        if (typeof val === "string") val = val.replace(/,/g, '').replace(/€|\s/g,'');
        let num = parseFloat(val);
        if (!isNaN(num) && num < 0) sum += Math.abs(num);
      }
      arr[w] = sum;
    }
    return arr;
  }
  function getRepaymentArr() {
    // If no mapping is configured, use default week labels for repayment calculations
    let actualWeekLabels = weekLabels && weekLabels.length > 0 ? weekLabels : Array.from({length: 52}, (_, i) => `Week ${i + 1}`);
    let actualWeekStartDates = weekStartDates && weekStartDates.length > 0 ? weekStartDates : Array.from({length: 52}, (_, i) => new Date(2025, 0, 1 + i * 7));
    let arr = Array(actualWeekLabels.length).fill(0);
    
    // Process all repayments using their explicit dates
    repaymentRows.forEach(r => {
      let repaymentDate = null;
      
      // Get the explicit date for this repayment
      if (r.explicitDate) {
        repaymentDate = new Date(r.explicitDate);
      } else if (r.type === "week" && r.week) {
        // Fallback: calculate date from week if explicitDate is missing
        let weekIdx = actualWeekLabels.indexOf(r.week);
        if (weekIdx === -1) weekIdx = 0;
        repaymentDate = actualWeekStartDates[weekIdx] || new Date(2025, 0, 1 + weekIdx * 7);
      } else if (r.type === "frequency") {
        // For frequency-based, use current approach for backward compatibility
        if (r.frequency === "monthly") {
          let perMonth = Math.ceil(arr.length/12);
          for (let m=0; m<12; m++) {
            for (let w=m*perMonth; w<(m+1)*perMonth && w<arr.length; w++) arr[w] += r.amount;
          }
        } else if (r.frequency === "quarterly") {
          let perQuarter = Math.ceil(arr.length/4);
          for (let q=0;q<4;q++) {
            for (let w=q*perQuarter; w<(q+1)*perQuarter && w<arr.length; w++) arr[w] += r.amount;
          }
        } else if (r.frequency === "one-off") { 
          arr[0] += r.amount; 
        }
        return; // Skip the date-based logic for frequency repayments
      }
      
      if (repaymentDate) {
        // Find the closest week to this explicit date
        let closestWeekIdx = 0;
        let closestDiff = Math.abs(actualWeekStartDates[0] - repaymentDate);
        
        for (let i = 1; i < actualWeekStartDates.length; i++) {
          let diff = Math.abs(actualWeekStartDates[i] - repaymentDate);
          if (diff < closestDiff) {
            closestDiff = diff;
            closestWeekIdx = i;
          }
        }
        
        // Add the repayment to the closest week
        arr[closestWeekIdx] += r.amount;
      }
    });
    
    // If mapping is configured, return filtered results. Otherwise, return all results.
    if (mappingConfigured && weekLabels.length > 0) {
      return getFilteredWeekIndices().map(idx => arr[idx]);
    } else {
      // Return all weeks when no mapping is configured
      return arr;
    }
  }
  
  // Function to get explicit repayment dates and amounts for NPV/IRR calculations
  function getExplicitRepaymentSchedule() {
    let schedule = [];
    let actualWeekLabels = weekLabels && weekLabels.length > 0 ? weekLabels : Array.from({length: 52}, (_, i) => `Week ${i + 1}`);
    let actualWeekStartDates = weekStartDates && weekStartDates.length > 0 ? weekStartDates : Array.from({length: 52}, (_, i) => new Date(2025, 0, 1 + i * 7));
    
    repaymentRows.forEach(r => {
      let repaymentDate = null;
      
      // Get the explicit date for this repayment
      if (r.explicitDate) {
        repaymentDate = new Date(r.explicitDate);
      } else if (r.type === "week" && r.week) {
        // Fallback: calculate date from week if explicitDate is missing
        let weekIdx = actualWeekLabels.indexOf(r.week);
        if (weekIdx === -1) weekIdx = 0;
        repaymentDate = actualWeekStartDates[weekIdx] || new Date(2025, 0, 1 + weekIdx * 7);
      }
      
      if (repaymentDate) {
        schedule.push({ date: repaymentDate, amount: r.amount });
      } else if (r.type === "frequency") {
        // Handle frequency-based repayments with calculated dates
        if (r.frequency === "monthly") {
          let perMonth = Math.ceil(actualWeekLabels.length/12);
          for (let m=0; m<12; m++) {
            let weekIdx = m * perMonth;
            if (weekIdx < actualWeekStartDates.length) {
              let date = actualWeekStartDates[weekIdx] || new Date(2025, 0, 1 + weekIdx * 7);
              schedule.push({ date, amount: r.amount });
            }
          }
        } else if (r.frequency === "quarterly") {
          let perQuarter = Math.ceil(actualWeekLabels.length/4);
          for (let q=0;q<4;q++) {
            let weekIdx = q * perQuarter;
            if (weekIdx < actualWeekStartDates.length) {
              let date = actualWeekStartDates[weekIdx] || new Date(2025, 0, 1 + weekIdx * 7);
              schedule.push({ date, amount: r.amount });
            }
          }
        } else if (r.frequency === "one-off") { 
          let date = actualWeekStartDates[0] || new Date(2025, 0, 1);
          schedule.push({ date, amount: r.amount });
        }
      }
    });
    
    // Sort by date to handle back-dating properly
    schedule.sort((a, b) => a.date - b.date);
    
    return schedule;
  }
  function getNetProfitArr(incomeArr, expenditureArr, repaymentArr) {
    return incomeArr.map((inc, i) => (inc || 0) - (expenditureArr[i] || 0) - (repaymentArr[i] || 0));
  }
  function getRollingBankBalanceArr() {
    let incomeArr = getIncomeArr();
    let expenditureArr = getExpenditureArr();
    let repaymentArr = getRepaymentArr();
    let rolling = [];
    let ob = openingBalance;
    getFilteredWeekIndices().forEach((fi, i) => {
      let income = incomeArr[fi] || 0;
      let out = expenditureArr[fi] || 0;
      let repay = repaymentArr[i] || 0;
      let prev = (i === 0 ? ob : rolling[i-1]);
      rolling[i] = prev + income - out - repay;
    });
    return rolling;
  }
  function getMonthAgg(arr, months=12) {
    let filtered = arr.filter((_,i)=>getFilteredWeekIndices().includes(i));
    let perMonth = Math.ceil(filtered.length/months);
    let out = [];
    for(let m=0;m<months;m++) {
      let sum=0;
      for(let w=m*perMonth;w<(m+1)*perMonth && w<filtered.length;w++) sum += filtered[w];
      out.push(sum);
    }
    return out;
  }

  // -------------------- Repayments UI --------------------
  const weekSelect = document.getElementById('weekSelect');
  const repaymentFrequency = document.getElementById('repaymentFrequency');
  
  function getDateForWeek(weekLabel) {
    const actualWeekLabels = weekLabels && weekLabels.length > 0 ? weekLabels : Array.from({length: 52}, (_, i) => `Week ${i + 1}`);
    const actualWeekStartDates = weekStartDates && weekStartDates.length > 0 ? weekStartDates : Array.from({length: 52}, (_, i) => new Date(2025, 0, 1 + i * 7));
    
    const weekIdx = actualWeekLabels.indexOf(weekLabel);
    if (weekIdx !== -1 && actualWeekStartDates[weekIdx]) {
      return actualWeekStartDates[weekIdx];
    }
    // Fallback for default weeks
    const weekNum = parseInt(weekLabel.replace(/\D/g, '')) || 1;
    return new Date(2025, 0, 1 + (weekNum - 1) * 7);
  }
  
  function populateWeekDropdown(labels) {
    if (!weekSelect) return;
    weekSelect.innerHTML = '';
    (labels && labels.length ? labels : Array.from({length: 52}, (_, i) => `Week ${i+1}`)).forEach(label => {
      const opt = document.createElement('option');
      opt.value = label;
      opt.textContent = label;
      weekSelect.appendChild(opt);
    });
    
    // Update the date field when weeks are populated
    const repaymentDateInput = document.getElementById('repaymentDate');
    if (repaymentDateInput && weekSelect.value) {
      const weekDate = getDateForWeek(weekSelect.value);
      repaymentDateInput.value = weekDate.toISOString().split('T')[0];
    }
  }

  function setupRepaymentForm() {
    if (!weekSelect || !repaymentFrequency) return;
    const repaymentDateInput = document.getElementById('repaymentDate');
    
    // Function to update date when week changes
    function updateDateFromWeek() {
      if (repaymentDateInput && weekSelect.value) {
        const weekDate = getDateForWeek(weekSelect.value);
        repaymentDateInput.value = weekDate.toISOString().split('T')[0];
      }
    }
    
    // Auto-populate date when week selection changes
    if (weekSelect) {
      weekSelect.addEventListener('change', updateDateFromWeek);
    }
    
    document.querySelectorAll('input[name="repaymentType"]').forEach(radio => {
      radio.addEventListener('change', function() {
        if (this.value === "week") {
          weekSelect.disabled = false;
          repaymentFrequency.disabled = true;
          updateDateFromWeek(); // Auto-populate date for selected week
        } else if (this.value === "date") {
          weekSelect.disabled = true;
          repaymentFrequency.disabled = true;
          // Keep date picker enabled for manual selection
        } else {
          weekSelect.disabled = true;
          repaymentFrequency.disabled = false;
          // For frequency mode, set a default date or current date
          if (repaymentDateInput) {
            repaymentDateInput.value = new Date().toISOString().split('T')[0];
          }
        }
      });
    });

    let addRepaymentForm = document.getElementById('addRepaymentForm');
    if (addRepaymentForm) {
      addRepaymentForm.onsubmit = function(e) {
        e.preventDefault();
        const type = document.querySelector('input[name="repaymentType"]:checked').value;
        let week = null, frequency = null, explicitDate = null;
        
        // Always get the explicit date from the date input
        explicitDate = repaymentDateInput ? repaymentDateInput.value : null;
        if (!explicitDate) {
          alert('Please select a date for the repayment.');
          return;
        }
        
        if (type === "week") {
          week = weekSelect.value;
        } else if (type === "frequency") {
          frequency = repaymentFrequency.value;
        }
        
        const amount = document.getElementById('repaymentAmount').value;
        if (!amount) return;
        
        repaymentRows.push({ 
          type, 
          week, 
          frequency, 
          explicitDate, 
          amount: parseFloat(amount), 
          editing: false 
        });
        renderRepaymentRows();
        this.reset();
        populateWeekDropdown(weekLabels);
        document.getElementById('weekSelect').selectedIndex = 0;
        document.getElementById('repaymentFrequency').selectedIndex = 0;
        if (repaymentDateInput) repaymentDateInput.value = '';
        document.querySelector('input[name="repaymentType"][value="week"]').checked = true;
        weekSelect.disabled = false;
        repaymentFrequency.disabled = true;
        updateAllTabs();
      };
    }
    
    // Initialize with first week's date when weeks are available
    setTimeout(() => {
      updateDateFromWeek();
    }, 100);
  }
  setupRepaymentForm();
  
  // Initialize week dropdown with default weeks if no mapping is configured
  if (!mappingConfigured || !weekLabels || weekLabels.length === 0) {
    populateWeekDropdown([]);
  }

  function renderRepaymentRows() {
    const container = document.getElementById('repaymentRows');
    if (!container) return;
    container.innerHTML = "";
    repaymentRows.forEach((row, i) => {
      const div = document.createElement('div');
      div.className = 'repayment-row';
      div.style.display = 'flex';
      div.style.alignItems = 'center';
      div.style.gap = '10px';
      div.style.marginBottom = '8px';
      
      // Week selector
      const weekSelectElem = document.createElement('select');
      (weekLabels.length ? weekLabels : Array.from({length:52}, (_,i)=>`Week ${i+1}`)).forEach(label => {
        const opt = document.createElement('option');
        opt.value = label;
        opt.textContent = label;
        weekSelectElem.appendChild(opt);
      });
      weekSelectElem.value = row.week || "";
      weekSelectElem.disabled = !row.editing || row.type !== "week";
      
      // Update date when week changes in edit mode
      weekSelectElem.addEventListener('change', function() {
        if (row.editing && row.type === "week") {
          const weekDate = getDateForWeek(this.value);
          dateInput.value = weekDate.toISOString().split('T')[0];
        }
      });

      // Date selector - always visible and shows the explicit date
      const dateInput = document.createElement('input');
      dateInput.type = 'date';
      dateInput.value = row.explicitDate || "";
      dateInput.disabled = !row.editing;
      dateInput.style.width = '140px';

      // Frequency selector
      const freqSelect = document.createElement('select');
      ["monthly", "quarterly", "one-off"].forEach(f => {
        const opt = document.createElement('option');
        opt.value = f;
        opt.textContent = f.charAt(0).toUpperCase() + f.slice(1);
        freqSelect.appendChild(opt);
      });
      freqSelect.value = row.frequency || "monthly";
      freqSelect.disabled = !row.editing || row.type !== "frequency";

      // Amount input
      const amountInput = document.createElement('input');
      amountInput.type = 'number';
      amountInput.value = row.amount;
      amountInput.placeholder = 'Repayment €';
      amountInput.disabled = !row.editing;
      amountInput.style.width = '120px';

      // Edit button
      const editBtn = document.createElement('button');
      editBtn.textContent = row.editing ? 'Save' : 'Edit';
      editBtn.onclick = function() {
        if (row.editing) {
          // Always save the explicit date
          row.explicitDate = dateInput.value;
          if (row.type === "week") {
            row.week = weekSelectElem.value;
          } else if (row.type === "frequency") {
            row.frequency = freqSelect.value;
          }
          row.amount = parseFloat(amountInput.value);
        }
        row.editing = !row.editing;
        renderRepaymentRows();
        updateAllTabs();
      };

      // Remove button
      const removeBtn = document.createElement('button');
      removeBtn.textContent = 'Remove';
      removeBtn.onclick = function() {
        repaymentRows.splice(i, 1);
        renderRepaymentRows();
        updateAllTabs();
      };

      // Display mode and controls
      const modeLabel = document.createElement('span');
      modeLabel.style.marginRight = "10px";
      modeLabel.style.fontWeight = "bold";
      
      if (row.type === "week") {
        modeLabel.textContent = "Week:";
        div.appendChild(modeLabel);
        div.appendChild(weekSelectElem);
      } else if (row.type === "date") {
        modeLabel.textContent = "Explicit:";
        div.appendChild(modeLabel);
      } else {
        modeLabel.textContent = "Frequency:";
        div.appendChild(modeLabel);
        div.appendChild(freqSelect);
      }
      
      // Always show the date
      const dateLabel = document.createElement('span');
      dateLabel.textContent = "Date:";
      dateLabel.style.marginLeft = row.type === "date" ? "0" : "15px";
      dateLabel.style.fontWeight = "bold";
      div.appendChild(dateLabel);
      div.appendChild(dateInput);
      
      const amountLabel = document.createElement('span');
      amountLabel.textContent = "Amount:";
      amountLabel.style.fontWeight = "bold";
      div.appendChild(amountLabel);
      div.appendChild(amountInput);
      div.appendChild(editBtn);
      div.appendChild(removeBtn);

      container.appendChild(div);
    });
  }

  function updateLoanSummary() {
    const totalRepaid = getRepaymentArr().reduce((a,b)=>a+b,0);
    let totalRepaidBox = document.getElementById('totalRepaidBox');
    let remainingBox = document.getElementById('remainingBox');
    if (totalRepaidBox) totalRepaidBox.textContent = "Total Repaid: €" + totalRepaid.toLocaleString();
    if (remainingBox) remainingBox.textContent = "Remaining: €" + (loanOutstanding-totalRepaid).toLocaleString();
  }
  let loanOutstandingInput = document.getElementById('loanOutstandingInput');
  if (loanOutstandingInput) {
    loanOutstandingInput.oninput = function() {
      loanOutstanding = parseFloat(this.value) || 0;
      updateLoanSummary();
    };
  }

  // -------------------- Main Chart & Summary --------------------
  function updateChartAndSummary() {
    let mainChartElem = document.getElementById('mainChart');
    let mainChartSummaryElem = document.getElementById('mainChartSummary');
    let mainChartNoDataElem = document.getElementById('mainChartNoData');
    if (!mainChartElem || !mainChartSummaryElem || !mainChartNoDataElem) return;

    if (!mappingConfigured || !weekLabels.length || getFilteredWeekIndices().length === 0) {
      if (mainChartNoDataElem) mainChartNoDataElem.style.display = "";
      if (mainChartSummaryElem) mainChartSummaryElem.innerHTML = "";
      if (mainChart && typeof mainChart.destroy === "function") mainChart.destroy();
      return;
    } else {
      if (mainChartNoDataElem) mainChartNoDataElem.style.display = "none";
    }

    const filteredWeeks = getFilteredWeekIndices();
    const labels = filteredWeeks.map(idx => weekLabels[idx]);
    const incomeArr = getIncomeArr();
    const expenditureArr = getExpenditureArr();
    const repaymentArr = getRepaymentArr();
    const rollingArr = getRollingBankBalanceArr();
    const netProfitArr = getNetProfitArr(incomeArr, expenditureArr, repaymentArr);

    const data = {
      labels: labels,
      datasets: [
        {
          label: "Income",
          data: filteredWeeks.map(idx => incomeArr[idx] || 0),
          backgroundColor: "rgba(76,175,80,0.6)",
          borderColor: "#388e3c",
          fill: false,
          type: "bar"
        },
        {
          label: "Expenditure",
          data: filteredWeeks.map(idx => expenditureArr[idx] || 0),
          backgroundColor: "rgba(244,67,54,0.6)",
          borderColor: "#c62828",
          fill: false,
          type: "bar"
        },
        {
          label: "Repayment",
          data: filteredWeeks.map((_, i) => repaymentArr[i] || 0),
          backgroundColor: "rgba(255,193,7,0.6)",
          borderColor: "#ff9800",
          fill: false,
          type: "bar"
        },
        {
          label: "Net Profit",
          data: filteredWeeks.map((_, i) => netProfitArr[filteredWeeks[i]] || 0),
          backgroundColor: "rgba(33,150,243,0.3)",
          borderColor: "#1976d2",
          type: "line",
          fill: false,
          yAxisID: "y"
        },
        {
          label: "Rolling Bank Balance",
          data: rollingArr,
          backgroundColor: "rgba(156,39,176,0.2)",
          borderColor: "#8e24aa",
          type: "line",
          fill: true,
          yAxisID: "y"
        }
      ]
    };

    if (mainChart && typeof mainChart.destroy === "function") mainChart.destroy();

    mainChart = new Chart(mainChartElem.getContext('2d'), {
      type: 'bar',
      data: data,
      options: {
        responsive: true,
        plugins: {
          legend: { display: true },
          tooltip: { mode: "index", intersect: false }
        },
        scales: {
          x: { stacked: true },
          y: {
            beginAtZero: true,
            title: { display: true, text: "€" }
          }
        }
      }
    });

    let totalIncome = incomeArr.reduce((a,b)=>a+(b||0), 0);
    let totalExpenditure = expenditureArr.reduce((a,b)=>a+(b||0), 0);
    let totalRepayment = repaymentArr.reduce((a,b)=>a+(b||0), 0);
    let finalBalance = rollingArr[rollingArr.length - 1] || 0;
    let lowestBalance = Math.min(...rollingArr);

    mainChartSummaryElem.innerHTML = `
      <b>Total Income:</b> €${Math.round(totalIncome).toLocaleString()}<br>
      <b>Total Expenditure:</b> €${Math.round(totalExpenditure).toLocaleString()}<br>
      <b>Total Repayments:</b> €${Math.round(totalRepayment).toLocaleString()}<br>
      <b>Final Bank Balance:</b> <span style="color:${finalBalance<0?'#c00':'#388e3c'}">€${Math.round(finalBalance).toLocaleString()}</span><br>
      <b>Lowest Bank Balance:</b> <span style="color:${lowestBalance<0?'#c00':'#388e3c'}">€${Math.round(lowestBalance).toLocaleString()}</span>
    `;
  }

  // ---------- P&L Tab Functions ----------
  function renderSectionSummary(headerId, text, arr) {
    const headerElem = document.getElementById(headerId);
    if (!headerElem) return;
    headerElem.innerHTML = text;
  }
  function renderPnlTables() {
  // Weekly Breakdown
  const weeklyTable = document.getElementById('pnlWeeklyBreakdown');
  const monthlyTable = document.getElementById('pnlMonthlyBreakdown');
  const cashFlowTable = document.getElementById('pnlCashFlow');
  const pnlSummary = document.getElementById('pnlSummary');
  if (!weeklyTable || !monthlyTable || !cashFlowTable) return;

  // ---- Weekly table ----
  let tbody = weeklyTable.querySelector('tbody');
  if (tbody) tbody.innerHTML = '';
  let incomeArr = getIncomeArr();
  let expenditureArr = getExpenditureArr();
  let repaymentArr = getRepaymentArr();
  let rollingArr = getRollingBankBalanceArr();
  let netArr = getNetProfitArr(incomeArr, expenditureArr, repaymentArr);
  let weekIdxs = getFilteredWeekIndices();
  let rows = '';
  let minBal = null, minBalWeek = null;

  weekIdxs.forEach((idx, i) => {
    const net = (incomeArr[idx] || 0) - (expenditureArr[idx] || 0) - (repaymentArr[i] || 0);
    const netTooltip = `Income - Expenditure - Repayment\n${incomeArr[idx]||0} - ${expenditureArr[idx]||0} - ${repaymentArr[i]||0} = ${net}`;
    const balTooltip = `Prev Bal + Income - Expenditure - Repayment\n${i===0?openingBalance:rollingArr[i-1]} + ${incomeArr[idx]||0} - ${expenditureArr[idx]||0} - ${repaymentArr[i]||0} = ${rollingArr[i]||0}`;
    let row = `<tr${rollingArr[i]<0?' class="negative-balance-row"':''}>` +
      `<td>${weekLabels[idx]}</td>` +
      `<td${incomeArr[idx]<0?' class="negative-number"':''}>€${Math.round(incomeArr[idx]||0).toLocaleString()}</td>` +
      `<td${expenditureArr[idx]<0?' class="negative-number"':''}>€${Math.round(expenditureArr[idx]||0).toLocaleString()}</td>` +
      `<td${repaymentArr[i]<0?' class="negative-number"':''}>€${Math.round(repaymentArr[i]||0).toLocaleString()}</td>` +
      `<td class="${net<0?'negative-number':''}" data-tooltip="${netTooltip}">€${Math.round(net||0).toLocaleString()}</td>` +
      `<td${rollingArr[i]<0?' class="negative-number"':''} data-tooltip="${balTooltip}">€${Math.round(rollingArr[i]||0).toLocaleString()}</td></tr>`;
    rows += row;
    if (minBal===null||rollingArr[i]<minBal) {minBal=rollingArr[i];minBalWeek=weekLabels[idx];}
  });
  if (tbody) tbody.innerHTML = rows;
  renderSectionSummary('weekly-breakdown-header', `Total Net: €${netArr.reduce((a,b)=>a+(b||0),0).toLocaleString()}`, netArr);

  // ---- Monthly Breakdown ----
  let months = 12;
  let incomeMonth = getMonthAgg(incomeArr, months);
  let expMonth = getMonthAgg(expenditureArr, months);
  let repayMonth = getMonthAgg(repaymentArr, months);
  let netMonth = incomeMonth.map((inc, i) => inc - (expMonth[i]||0) - (repayMonth[i]||0));
  let mtbody = monthlyTable.querySelector('tbody');
  if (mtbody) {
    mtbody.innerHTML = '';
    for (let m=0; m<months; m++) {
      const netTooltip = `Income - Expenditure - Repayment\n${incomeMonth[m]||0} - ${expMonth[m]||0} - ${repayMonth[m]||0} = ${netMonth[m]||0}`;
      mtbody.innerHTML += `<tr>
        <td>Month ${m+1}</td>
        <td${incomeMonth[m]<0?' class="negative-number"':''}>€${Math.round(incomeMonth[m]||0).toLocaleString()}</td>
        <td${expMonth[m]<0?' class="negative-number"':''}>€${Math.round(expMonth[m]||0).toLocaleString()}</td>
        <td class="${netMonth[m]<0?'negative-number':''}" data-tooltip="${netTooltip}">€${Math.round(netMonth[m]||0).toLocaleString()}</td>
        <td${repayMonth[m]<0?' class="negative-number"':''}>€${Math.round(repayMonth[m]||0).toLocaleString()}</td>
      </tr>`;
    }
  }
  renderSectionSummary('monthly-breakdown-header', `Total Net: €${netMonth.reduce((a,b)=>a+(b||0),0).toLocaleString()}`, netMonth);

  // ---- Cash Flow Table ----
  let ctbody = cashFlowTable.querySelector('tbody');
  let closingArr = [];
  if (ctbody) {
    ctbody.innerHTML = '';
    let closing = opening = openingBalance;
    for (let m=0; m<months; m++) {
      let inflow = incomeMonth[m] || 0;
      let outflow = (expMonth[m] || 0) + (repayMonth[m] || 0);
      closing = opening + inflow - outflow;
      closingArr.push(closing);
      const closingTooltip = `Opening + Inflow - Outflow\n${opening} + ${inflow} - ${outflow} = ${closing}`;
      ctbody.innerHTML += `<tr>
        <td>Month ${m+1}</td>
        <td>€${Math.round(opening).toLocaleString()}</td>
        <td>€${Math.round(inflow).toLocaleString()}</td>
        <td>€${Math.round(outflow).toLocaleString()}</td>
        <td${closing<0?' class="negative-number"':''} data-tooltip="${closingTooltip}">€${Math.round(closing).toLocaleString()}</td>
      </tr>`;
      opening = closing;
    }
  }
  renderSectionSummary('cashflow-header', `Closing Bal: €${Math.round(closingArr[closingArr.length-1]||0).toLocaleString()}`, closingArr);

  // ---- P&L Summary ----
  if (pnlSummary) {
    pnlSummary.innerHTML = `
      <b>Total Income:</b> €${Math.round(incomeArr.reduce((a,b)=>a+(b||0),0)).toLocaleString()}<br>
      <b>Total Expenditure:</b> €${Math.round(expenditureArr.reduce((a,b)=>a+(b||0),0)).toLocaleString()}<br>
      <b>Total Repayments:</b> €${Math.round(repaymentArr.reduce((a,b)=>a+(b||0),0)).toLocaleString()}<br>
      <b>Final Bank Balance:</b> <span style="color:${rollingArr[rollingArr.length-1]<0?'#c00':'#388e3c'}">€${Math.round(rollingArr[rollingArr.length-1]||0).toLocaleString()}</span><br>
      <b>Lowest Bank Balance:</b> <span style="color:${minBal<0?'#c00':'#388e3c'}">${minBalWeek?minBalWeek+': ':''}€${Math.round(minBal||0).toLocaleString()}</span>
    `;
  }
}
  // ---------- Summary Tab Functions ----------
  function renderSummaryTab() {
    // Key Financials
    let incomeArr = getIncomeArr();
    let expenditureArr = getExpenditureArr();
    let repaymentArr = getRepaymentArr();
    let rollingArr = getRollingBankBalanceArr();
    let netArr = getNetProfitArr(incomeArr, expenditureArr, repaymentArr);
    let totalIncome = incomeArr.reduce((a,b)=>a+(b||0),0);
    let totalExpenditure = expenditureArr.reduce((a,b)=>a+(b||0),0);
    let totalRepayment = repaymentArr.reduce((a,b)=>a+(b||0),0);
    let finalBal = rollingArr[rollingArr.length-1]||0;
    let minBal = Math.min(...rollingArr);

    // Update KPI cards if present
    if (document.getElementById('kpiTotalIncome')) {
      document.getElementById('kpiTotalIncome').textContent = '€' + totalIncome.toLocaleString();
      document.getElementById('kpiTotalExpenditure').textContent = '€' + totalExpenditure.toLocaleString();
      document.getElementById('kpiTotalRepayments').textContent = '€' + totalRepayment.toLocaleString();
      document.getElementById('kpiFinalBank').textContent = '€' + Math.round(finalBal).toLocaleString();
      document.getElementById('kpiLowestBank').textContent = '€' + Math.round(minBal).toLocaleString();
    }

    let summaryElem = document.getElementById('summaryKeyFinancials');
    if (summaryElem) {
      summaryElem.innerHTML = `
        <b>Total Income:</b> €${Math.round(totalIncome).toLocaleString()}<br>
        <b>Total Expenditure:</b> €${Math.round(totalExpenditure).toLocaleString()}<br>
        <b>Total Repayments:</b> €${Math.round(totalRepayment).toLocaleString()}<br>
        <b>Final Bank Balance:</b> <span style="color:${finalBal<0?'#c00':'#388e3c'}">€${Math.round(finalBal).toLocaleString()}</span><br>
        <b>Lowest Bank Balance:</b> <span style="color:${minBal<0?'#c00':'#388e3c'}">€${Math.round(minBal).toLocaleString()}</span>
      `;
    }
    // Summary Chart
    let summaryChartElem = document.getElementById('summaryChart');
    if (summaryChart && typeof summaryChart.destroy === "function") summaryChart.destroy();
    if (summaryChartElem) {
      summaryChart = new Chart(summaryChartElem.getContext('2d'), {
        type: 'bar',
        data: {
          labels: ["Income", "Expenditure", "Repayment", "Final Bank", "Lowest Bank"],
          datasets: [{
            label: "Totals",
            data: [
              Math.round(totalIncome),
              -Math.round(totalExpenditure),
              -Math.round(totalRepayment),
              Math.round(finalBal),
              Math.round(minBal)
            ],
            backgroundColor: [
              "#4caf50","#f44336","#ffc107","#2196f3","#9c27b0"
            ]
          }]
        },
        options: {
          responsive:true,
          plugins:{legend:{display:false}},
          scales: { y: { beginAtZero: true } }
        }
      });
    }

    // Tornado Chart logic
    function renderTornadoChart() {
      // Calculate row impact by "sum of absolute values" for each data row
      let impact = [];
      if (!mappedData || !mappingConfigured) return;
      for (let r = config.firstDataRow; r <= config.lastDataRow; r++) {
        let label = mappedData[r][0] || `Row ${r + 1}`;
        let vals = [];
        for (let w = 0; w < weekLabels.length; w++) {
          if (!weekCheckboxStates[w]) continue;
          let absCol = config.weekColStart + w;
          let val = mappedData[r][absCol];
          if (typeof val === "string") val = val.replace(/,/g,'').replace(/€|\s/g,'');
          let num = parseFloat(val);
          if (!isNaN(num)) vals.push(num);
        }
        let total = vals.reduce((a,b)=>a+Math.abs(b),0);
        if (total > 0) impact.push({label, total});
      }
      impact.sort((a,b)=>b.total-a.total);
      impact = impact.slice(0, 10);

      let ctx = document.getElementById('tornadoChart').getContext('2d');
      if (window.tornadoChartObj && typeof window.tornadoChartObj.destroy === "function") window.tornadoChartObj.destroy();
      window.tornadoChartObj = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: impact.map(x=>x.label),
          datasets: [{ label: "Total Impact (€)", data: impact.map(x=>x.total), backgroundColor: '#1976d2' }]
        },
        options: { indexAxis: 'y', responsive: true, plugins: { legend: { display: false } } }
      });
    }
    renderTornadoChart();
  }

  // -------------------- TARGET IRR/NPV MODAL CONTROLS --------------------
  function setupTargetIrrModal() {
    const modal = document.getElementById('targetIrrModal');
    const editBtn = document.getElementById('editTargetIrrBtn');
    const closeBtn = document.getElementById('closeIrrModal');
    const applyBtn = document.getElementById('applyIrrSettings');
    const cancelBtn = document.getElementById('cancelIrrSettings');
    const slider = document.getElementById('targetIrrSlider');
    const sliderValue = document.getElementById('targetIrrValue');
    const installmentInput = document.getElementById('installmentCountInput');
    const liveUpdateCheckbox = document.getElementById('liveUpdateCheckbox');
    const suggestionDisplay = document.getElementById('suggestedIrrDisplay');
    const npvDisplay = document.getElementById('equivalentNpvDisplay');
    const firstRepaymentWeekSelect = document.getElementById('firstRepaymentWeekSelect');
    
    if (!modal || !editBtn) return;
    
    // Function to populate first repayment week dropdown
    function populateFirstRepaymentWeekDropdown() {
      if (!firstRepaymentWeekSelect) return;
      
      // Clear existing options except the first one
      firstRepaymentWeekSelect.innerHTML = '<option value="">Select week...</option>';
      
      // Use mapped week labels if available, otherwise generate default weeks
      const availableWeekLabels = weekLabels && weekLabels.length > 0 ? weekLabels : 
        Array.from({length: 52}, (_, i) => `Week ${i + 1}`);
      
      // Populate dropdown with weeks after investment week
      for (let i = investmentWeekIndex + 1; i < availableWeekLabels.length; i++) {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = availableWeekLabels[i] || `Week ${i + 1}`;
        firstRepaymentWeekSelect.appendChild(option);
      }
    }
    
    // Calculate NPV for given IRR
    function calculateNPVForIRR(irrRate) {
      const investment = parseFloat(document.getElementById('roiInvestmentInput').value) || 0;
      if (investment <= 0) return 0;
      
      // Get current repayments or use default installment calculation
      const repaymentsFull = getRepaymentArr ? getRepaymentArr() : [];
      const repayments = repaymentsFull.slice(investmentWeekIndex);
      
      // If no repayments exist, calculate equivalent repayments for target IRR
      if (repayments.length === 0 || repayments.every(r => r === 0)) {
        // Calculate what total repayments would be needed for this IRR
        const targetReturn = investment * (1 + irrRate);
        return targetReturn - investment; // This is the NPV equivalent
      }
      
      // Use actual repayments with date-based discounting
      const cashflows = [-investment, ...repayments];
      const cashflowDates = [weekStartDates[investmentWeekIndex] || new Date()];
      
      for (let i = 1; i < cashflows.length; i++) {
        let idx = investmentWeekIndex + i;
        cashflowDates[i] = weekStartDates[idx] || new Date();
      }
      
      // Use the global npv_date function for consistent calculation
      // This implements: NPV = sum(CF_i / (1 + r)^(t_i/365.25)) - Investment
      if (typeof npv_date === 'function') {
        return npv_date(irrRate, cashflows, cashflowDates);
      } else {
        // Simple NPV calculation without dates (fallback)
        return cashflows.reduce((acc, val, i) => acc + val / Math.pow(1 + irrRate, i), 0);
      }
    }
    
    // Update NPV display
    function updateNPVDisplay() {
      if (!npvDisplay) return;
      const irrRate = parseFloat(slider.value) / 100;
      const npvValue = calculateNPVForIRR(irrRate);
      npvDisplay.textContent = `€${npvValue.toLocaleString(undefined, {maximumFractionDigits: 2})}`;
    }
    
    // Update display
    function updateDisplay() {
      if (suggestionDisplay) {
        suggestionDisplay.textContent = `Target IRR: ${Math.round(targetIRR * 100)}%`;
      }
      updateNPVDisplay();
    }
    
    // Open modal
    editBtn.addEventListener('click', function() {
      slider.value = Math.round(targetIRR * 100);
      sliderValue.textContent = Math.round(targetIRR * 100) + '%';
      installmentInput.value = installmentCount;
      liveUpdateCheckbox.checked = liveUpdateEnabled;
      populateFirstRepaymentWeekDropdown();
      updateNPVDisplay(); // Initialize NPV display
      modal.style.display = 'flex';
    });
    
    // Close modal
    function closeModal() {
      modal.style.display = 'none';
    }
    
    closeBtn.addEventListener('click', closeModal);
    cancelBtn.addEventListener('click', closeModal);
    
    // Click outside modal to close
    modal.addEventListener('click', function(e) {
      if (e.target === modal) closeModal();
    });
    
    // Slider input
    slider.addEventListener('input', function() {
      const value = this.value;
      sliderValue.textContent = value + '%';
      updateNPVDisplay(); // Update NPV live as slider changes
      if (liveUpdateCheckbox.checked) {
        targetIRR = parseFloat(value) / 100;
        updateDisplay();
        if (showSuggestions) {
          generateAndUpdateSuggestions();
        }
      }
    });
    
    // Installment count input
    installmentInput.addEventListener('input', function() {
      if (liveUpdateCheckbox.checked) {
        installmentCount = parseInt(this.value) || 12;
        if (showSuggestions) {
          generateAndUpdateSuggestions();
        }
      }
    });
    
    // Apply settings
    applyBtn.addEventListener('click', function() {
      targetIRR = parseFloat(slider.value) / 100;
      installmentCount = parseInt(installmentInput.value) || 12;
      liveUpdateEnabled = liveUpdateCheckbox.checked;
      updateDisplay();
      if (showSuggestions) {
        generateAndUpdateSuggestions();
      }
      closeModal();
    });
    
    // Initialize display
    updateDisplay();
  }

  // -------------------- BUFFER SELECTION MODAL CONTROLS --------------------
  function setupBufferModal() {
    const modal = document.getElementById('bufferModal');
    const bufferBtn = document.getElementById('bufferSelectionBtn');
    const closeBtn = document.getElementById('closeBufferModal');
    const applyBtn = document.getElementById('applyBufferSettings');
    const cancelBtn = document.getElementById('cancelBufferSettings');
    const bufferDisplay = document.getElementById('bufferSelectionDisplay');
    const customBufferSettings = document.getElementById('customBufferSettings');
    const customBufferWeeks = document.getElementById('customBufferWeeks');
    
    if (!modal || !bufferBtn) return;
    
    // Function to update buffer display
    function updateBufferDisplay() {
      if (!bufferDisplay) return;
      
      let displayText = '';
      switch (bufferSettings.type) {
        case 'none':
          displayText = 'None (fastest schedule)';
          break;
        case '2weeks':
          displayText = '2 weeks gap';
          break;
        case '1month':
          displayText = '1 month gap';
          break;
        case '2months':
          displayText = '2 months gap';
          break;
        case 'quarter':
          displayText = 'Quarter (3 months) gap';
          break;
        case 'custom':
          displayText = `Custom: ${bufferSettings.customWeeks} weeks gap`;
          break;
        default:
          displayText = 'None selected';
      }
      bufferDisplay.textContent = displayText;
    }
    
    // Function to get buffer weeks based on settings
    function getBufferWeeks() {
      switch (bufferSettings.type) {
        case 'none':
          return 0;
        case '2weeks':
          return 2;
        case '1month':
          return 4; // Approximate 4 weeks per month
        case '2months':
          return 8;
        case 'quarter':
          return 12; // Approximate 12 weeks per quarter
        case 'custom':
          return bufferSettings.customWeeks;
        default:
          return 0;
      }
    }
    
    // Make getBufferWeeks available globally for suggestion algorithm
    window.getBufferWeeks = getBufferWeeks;
    
    // Open modal
    bufferBtn.addEventListener('click', function() {
      // Set current selection
      const radioButtons = modal.querySelectorAll('input[name="bufferOption"]');
      radioButtons.forEach(radio => {
        radio.checked = radio.value === bufferSettings.type;
      });
      
      if (customBufferWeeks) {
        customBufferWeeks.value = bufferSettings.customWeeks;
      }
      
      // Show/hide custom settings
      if (customBufferSettings) {
        customBufferSettings.style.display = bufferSettings.type === 'custom' ? 'block' : 'none';
      }
      
      modal.style.display = 'flex';
    });
    
    // Close modal
    function closeModal() {
      modal.style.display = 'none';
    }
    
    closeBtn.addEventListener('click', closeModal);
    cancelBtn.addEventListener('click', closeModal);
    
    // Click outside modal to close
    modal.addEventListener('click', function(e) {
      if (e.target === modal) closeModal();
    });
    
    // Handle radio button changes
    modal.addEventListener('change', function(e) {
      if (e.target.name === 'bufferOption') {
        if (customBufferSettings) {
          customBufferSettings.style.display = e.target.value === 'custom' ? 'block' : 'none';
        }
      }
    });
    
    // Apply settings
    applyBtn.addEventListener('click', function() {
      const selectedRadio = modal.querySelector('input[name="bufferOption"]:checked');
      if (selectedRadio) {
        bufferSettings.type = selectedRadio.value;
        
        if (bufferSettings.type === 'custom' && customBufferWeeks) {
          bufferSettings.customWeeks = parseInt(customBufferWeeks.value) || 1;
        }
        
        updateBufferDisplay();
        
        // If suggestions are currently shown, regenerate them with new buffer
        if (showSuggestions) {
          generateAndUpdateSuggestions();
        }
      }
      closeModal();
    });
    
    // Initialize display
    updateBufferDisplay();
  }

  // -------------------- ENHANCED SUGGESTION ALGORITHM --------------------
  function computeEnhancedSuggestedRepayments({investment, targetIRR, installmentCount, filteredWeeks, investmentWeekIndex, openingBalance, cashflow, weekStartDates}) {
    if (!filteredWeeks || !filteredWeeks.length || targetIRR <= 0 || installmentCount <= 0) {
      return { suggestedRepayments: [], achievedIRR: null, warnings: [] };
    }

    let warnings = [];
    
    // Calculate total amount that needs to be recouped (investment + target return)
    const targetReturn = investment * (1 + targetIRR);
    
    // Calculate suggested installment amount based on target installment count
    const suggestedInstallmentAmount = targetReturn / installmentCount;
    
    // Find investment index in filtered weeks
    const investmentIndex = filteredWeeks.indexOf(investmentWeekIndex);
    const startIndex = investmentIndex === -1 ? 0 : investmentIndex + 1;
    
    // Get first repayment week if specified
    const firstRepaymentWeekSelect = document.getElementById('firstRepaymentWeekSelect');
    let firstRepaymentWeek = startIndex;
    if (firstRepaymentWeekSelect && firstRepaymentWeekSelect.value) {
      firstRepaymentWeek = parseInt(firstRepaymentWeekSelect.value);
    }
    
    // Get buffer settings
    const bufferWeeks = window.getBufferWeeks ? window.getBufferWeeks() : 0;
    
    // Initialize variables for the new logic
    let outstanding = targetReturn;
    let repayments = [];
    let currentWeekIndex = Math.max(startIndex, firstRepaymentWeek);
    let extendedWeeks = [...filteredWeeks];
    let extendedWeekStartDates = [...weekStartDates];
    let lastRepaymentWeek = -1; // Track last repayment for buffer logic
    
    // Helper function to check if a week has sufficient bank balance for a repayment
    function hasValidBankBalance(weekIndex, repaymentAmount) {
      if (!cashflow || !cashflow.income || !cashflow.expenditure) {
        return true; // No cashflow data, assume sufficient
      }
      
      // Calculate rolling balance up to this week
      let rolling = openingBalance;
      for (let i = 0; i <= weekIndex && i < Math.max(cashflow.income.length, cashflow.expenditure.length); i++) {
        const income = cashflow.income[i] || 0;
        const expenditure = cashflow.expenditure[i] || 0;
        rolling = rolling + income - expenditure;
        
        // If this is the week we're checking, subtract the proposed repayment
        if (i === weekIndex) {
          rolling -= repaymentAmount;
        }
      }
      
      return rolling >= 0; // Bank balance should not go negative
    }
    
    // Helper function to add a week with proper date calculation
    function addWeekToSchedule() {
      const newWeekIndex = extendedWeeks.length;
      extendedWeeks.push(newWeekIndex);
      
      // Calculate date for the new week (7 days after the last week)
      const lastDate = extendedWeekStartDates[extendedWeekStartDates.length - 1] || new Date(2025, 0, 1);
      const newDate = new Date(lastDate);
      newDate.setDate(lastDate.getDate() + 7);
      extendedWeekStartDates.push(newDate);
      
      return newWeekIndex;
    }
    
    // Helper function to check if buffer requirement is met
    function isBufferSatisfied(weekIndex) {
      if (bufferWeeks === 0 || lastRepaymentWeek === -1) return true;
      return (weekIndex - lastRepaymentWeek) >= bufferWeeks;
    }
    
    let maxAttempts = 500; // Prevent infinite loops
    let attempts = 0;
    
    // Continue adding repayments until outstanding is fully covered
    while (outstanding > 0.01 && attempts < maxAttempts) { // Use small threshold to handle floating point precision
      attempts++;
      
      // Ensure we have enough weeks in the schedule
      if (currentWeekIndex >= extendedWeeks.length) {
        addWeekToSchedule();
      }
      
      // Check if buffer requirement is satisfied
      if (!isBufferSatisfied(currentWeekIndex)) {
        currentWeekIndex++;
        continue;
      }
      
      // Calculate payment amount (either full installment or remaining outstanding)
      let payment = Math.min(suggestedInstallmentAmount, outstanding);
      
      // Check if bank balance is sufficient for this payment
      if (!hasValidBankBalance(currentWeekIndex, payment)) {
        // Try smaller payment amounts
        let maxAffordablePayment = 0;
        for (let testPayment = payment * 0.1; testPayment <= payment; testPayment += payment * 0.1) {
          if (hasValidBankBalance(currentWeekIndex, testPayment)) {
            maxAffordablePayment = testPayment;
          }
        }
        
        if (maxAffordablePayment > 0.01) {
          payment = maxAffordablePayment;
        } else {
          // Skip this week if no affordable payment found
          currentWeekIndex++;
          continue;
        }
      }
      
      repayments.push({
        weekIndex: currentWeekIndex,
        amount: payment,
        date: extendedWeekStartDates[currentWeekIndex] || new Date(2025, 0, 1 + currentWeekIndex * 7)
      });
      
      outstanding -= payment;
      lastRepaymentWeek = currentWeekIndex;
      currentWeekIndex += (bufferWeeks + 1); // Move to next allowed week based on buffer
    }
    
    // Check if plan is achievable
    if (outstanding > 0.01) {
      warnings.push(`Unable to achieve target IRR with current settings. Remaining amount: €${outstanding.toLocaleString(undefined, {maximumFractionDigits: 2})}. Consider reducing buffer, increasing installment count, or extending the schedule.`);
    }
    
    // Ensure the last payment covers any remaining amount due to rounding
    if (repayments.length > 0 && outstanding > 0.01) {
      repayments[repayments.length - 1].amount += outstanding;
      outstanding = 0;
    }
    
    // Create the suggested array covering all weeks (original + extended if needed)
    const totalExtendedWeeks = Math.max(extendedWeeks.length, currentWeekIndex);
    const suggestedArray = new Array(totalExtendedWeeks).fill(0);
    
    // Fill in the repayments
    repayments.forEach(repayment => {
      if (repayment.weekIndex < suggestedArray.length) {
        suggestedArray[repayment.weekIndex] = repayment.amount;
      }
    });
    
    // Validate that total repayments equal target return (within small margin)
    const totalRepayments = suggestedArray.reduce((sum, amount) => sum + amount, 0);
    const shortfall = targetReturn - totalRepayments;
    
    if (Math.abs(shortfall) > 0.01) {
      if (shortfall > 0 && repayments.length > 0) {
        // Add shortfall to last payment
        const lastRepaymentIndex = repayments[repayments.length - 1].weekIndex;
        if (lastRepaymentIndex < suggestedArray.length) {
          suggestedArray[lastRepaymentIndex] += shortfall;
        }
      }
    }
    
    // Calculate achieved IRR using XIRR with actual dates for accurate annualized return
    const cashflows = [-investment];
    const cashflowDates = [extendedWeekStartDates[investmentWeekIndex] || new Date(2025, 0, 1)];
    
    // Add repayment cashflows starting from the investment week
    for (let i = startIndex; i < suggestedArray.length; i++) {
      if (suggestedArray[i] > 0) {
        cashflows.push(suggestedArray[i]);
        cashflowDates.push(extendedWeekStartDates[i] || new Date(2025, 0, 1 + i * 7));
      }
    }
    
    const achievedIRR = calculateIRR(cashflows, cashflowDates);
    
    // Check if achieved IRR is significantly different from target
    if (isFinite(achievedIRR) && Math.abs(achievedIRR - targetIRR) > 0.01) {
      warnings.push(`Achieved IRR (${(achievedIRR * 100).toFixed(2)}%) differs from target IRR (${(targetIRR * 100).toFixed(2)}%). Consider adjusting settings.`);
    }
    
    // Trim the suggested array to only include weeks up to the last repayment
    const lastRepaymentIndex = suggestedArray.findLastIndex(amount => amount > 0);
    const trimmedArray = lastRepaymentIndex >= 0 ? suggestedArray.slice(0, lastRepaymentIndex + 1) : suggestedArray;
    
    return {
      suggestedRepayments: trimmedArray,
      achievedIRR: achievedIRR,
      extendedWeeks: extendedWeeks.slice(0, trimmedArray.length),
      extendedWeekStartDates: extendedWeekStartDates.slice(0, trimmedArray.length),
      warnings: warnings
    };
  }
  
  function validateBankBalanceConstraint(suggestedRepayments, cashflow, openingBalance, filteredWeeks, investmentIndex) {
    const validatedRepayments = [...suggestedRepayments];
    let rolling = openingBalance;
    
    for (let i = 0; i < filteredWeeks.length; i++) {
      const weekIndex = filteredWeeks[i];
      const income = cashflow.income[weekIndex] || 0;
      const expenditure = cashflow.expenditure[weekIndex] || 0;
      const suggestedRepayment = i > investmentIndex ? validatedRepayments[i] : 0;
      
      const projectedBalance = rolling + income - expenditure - suggestedRepayment;
      
      // If balance would go negative, reduce the repayment
      if (projectedBalance < 0 && suggestedRepayment > 0) {
        const maxRepayment = rolling + income - expenditure;
        validatedRepayments[i] = Math.max(0, maxRepayment);
      }
      
      // Update rolling balance
      rolling = rolling + income - expenditure - (i > investmentIndex ? validatedRepayments[i] : 0);
    }
    
    return validatedRepayments;
  }
  
  function generateAndUpdateSuggestions() {
    const investment = parseFloat(document.getElementById('roiInvestmentInput').value) || 0;
    const filteredWeeks = getFilteredWeekIndices ? getFilteredWeekIndices() : Array.from({length: 52}, (_, i) => i);
    
    // If no filtered weeks (no data loaded), use default 52-week timeline
    const actualFilteredWeeks = filteredWeeks.length > 0 ? filteredWeeks : Array.from({length: 52}, (_, i) => i);
    
    const incomeArr = getIncomeArr ? getIncomeArr() : Array(52).fill(0);
    const expenditureArr = getExpenditureArr ? getExpenditureArr() : Array(52).fill(0);
    const cashflow = {income: incomeArr, expenditure: expenditureArr};
    
    const result = computeEnhancedSuggestedRepayments({
      investment,
      targetIRR,
      installmentCount,
      filteredWeeks: actualFilteredWeeks,
      investmentWeekIndex,
      openingBalance,
      cashflow,
      weekStartDates
    });
    
    suggestedRepayments = result.suggestedRepayments;
    achievedSuggestedIRR = result.achievedIRR;
    
    // Display warnings if any
    if (result.warnings && result.warnings.length > 0) {
      displaySuggestionWarnings(result.warnings);
    } else {
      clearSuggestionWarnings();
    }
    
    // Store extended weeks and dates for rendering if schedule was extended
    if (result.extendedWeeks && result.extendedWeekStartDates) {
      // Update global variables to include extended weeks for rendering
      window.extendedWeekLabels = result.extendedWeeks.map((_, i) => 
        i < weekLabels.length ? weekLabels[i] : `Extended Week ${i + 1}`
      );
      window.extendedWeekStartDates = result.extendedWeekStartDates;
    }
    
    // Re-render the ROI section to show updated suggestions
    renderRoiSection();
  }
  
  // -------------------- WARNING DISPLAY FUNCTIONS --------------------
  function displaySuggestionWarnings(warnings) {
    const warningElement = document.getElementById('roiWarningAlert');
    if (!warningElement || !warnings || warnings.length === 0) return;
    
    let warningHtml = '<div class="alert alert-warning" style="margin-bottom: 1em;">';
    warningHtml += '<strong>Planning Warnings:</strong><br>';
    warnings.forEach(warning => {
      warningHtml += `• ${warning}<br>`;
    });
    warningHtml += '</div>';
    
    warningElement.innerHTML = warningHtml;
  }
  
  function clearSuggestionWarnings() {
    const warningElement = document.getElementById('roiWarningAlert');
    if (warningElement) {
      warningElement.innerHTML = '';
    }
  }
  function computeSuggestedRepayments({investment, targetIRR, filteredWeeks, investmentWeekIndex, openingBalance, cashflow, weekStartDates}) {
    if (!filteredWeeks || !filteredWeeks.length || targetIRR <= 0) {
      return { suggestedRepayments: [], achievedIRR: null };
    }

    // Create array for suggested repayments covering all filtered weeks from investment point
    const totalWeeks = filteredWeeks.length;
    const suggestedArray = new Array(totalWeeks).fill(0);
    
    // Simple suggestion algorithm: distribute repayments to achieve target IRR
    // Start from the investment week and spread repayments across remaining weeks
    const investmentIndex = filteredWeeks.indexOf(investmentWeekIndex);
    if (investmentIndex === -1) return { suggestedRepayments: suggestedArray, achievedIRR: null };
    
    const remainingWeeks = totalWeeks - investmentIndex - 1;
    if (remainingWeeks <= 0) return { suggestedRepayments: suggestedArray, achievedIRR: null };
    
    // Calculate equal repayments to achieve target IRR
    // Using simple approximation: Total return = investment * (1 + targetIRR)
    const targetReturn = investment * (1 + targetIRR);
    const weeklyRepayment = targetReturn / remainingWeeks;
    
    // Fill suggested repayments from investment week onwards
    for (let i = investmentIndex + 1; i < totalWeeks; i++) {
      suggestedArray[i] = weeklyRepayment;
    }
    
    // Calculate achieved IRR for the suggested repayments using XIRR with actual dates
    const cashflows = [-investment, ...suggestedArray.slice(investmentIndex + 1)];
    const cashflowDates = [weekStartDates[investmentWeekIndex] || new Date()];
    
    // Build dates for cash flows from investment week onwards
    for (let i = 1; i < cashflows.length; i++) {
      let weekIdx = investmentWeekIndex + i;
      cashflowDates[i] = weekStartDates[weekIdx] || new Date(2025, 0, 1 + weekIdx * 7);
    }
    
    const achievedIRR = calculateIRR(cashflows, cashflowDates);
    
    return {
      suggestedRepayments: suggestedArray,
      achievedIRR: achievedIRR
    };
  }

  /**
   * XIRR - Extended Internal Rate of Return calculation for irregular cash flow dates
   * This function calculates annualized IRR based on actual cash flow dates rather than 
   * assuming evenly spaced periods. Essential for accurate ROI calculations with irregular
   * repayment schedules.
   * 
   * @param {Array} cashflows - Array of cash flow values (negative for outflows, positive for inflows)
   * @param {Array} dates - Array of Date objects corresponding to each cash flow
   * @param {number} guess - Initial guess for the rate (default: 0.1 or 10%)
   * @returns {number} Annualized IRR or NaN if calculation fails
   */
  function xirr(cashflows, dates, guess = 0.1) {
    if (!cashflows || !dates || cashflows.length !== dates.length || cashflows.length < 2) {
      return NaN;
    }
    
    // Helper function to calculate NPV using actual dates
    function xnpv(rate, cashflows, dates) {
      const msPerDay = 24 * 3600 * 1000;
      const baseDate = dates[0];
      return cashflows.reduce((acc, val, i) => {
        if (!dates[i]) return acc;
        let days = (dates[i] - baseDate) / msPerDay;
        let years = days / 365.25; // Use 365.25 for more accurate annualization
        return acc + val / Math.pow(1 + rate, years);
      }, 0);
    }
    
    // Newton-Raphson method to find the rate where XNPV = 0
    let rate = guess;
    const epsilon = 1e-6;
    const maxIter = 100;
    
    for (let iter = 0; iter < maxIter; iter++) {
      let npv0 = xnpv(rate, cashflows, dates);
      let npv1 = xnpv(rate + epsilon, cashflows, dates);
      let derivative = (npv1 - npv0) / epsilon;
      
      if (Math.abs(derivative) < 1e-10) break; // Avoid division by very small numbers
      
      let newRate = rate - npv0 / derivative;
      
      if (!isFinite(newRate)) break;
      if (Math.abs(newRate - rate) < 1e-7) return newRate; // Convergence achieved
      
      rate = newRate;
    }
    
    return NaN; // Failed to converge
  }

  /**
   * Calculate IRR using XIRR logic for irregular cash flow schedules
   * This function replaces the previous calculateIRR to ensure accurate 
   * annualized returns based on actual cash flow dates.
   * 
   * @param {Array} cashflows - Array of cash flow values
   * @param {Array} dates - Optional array of dates for XIRR calculation
   * @returns {number} Annualized IRR or NaN if calculation fails
   */
  function calculateIRR(cashflows, dates = null) {
    if (!cashflows || cashflows.length < 2) return NaN;
    
    // If dates are provided, use XIRR for accurate date-based calculation
    if (dates && dates.length === cashflows.length) {
      return xirr(cashflows, dates);
    }
    
    // Fallback to standard IRR for evenly spaced periods (legacy compatibility)
    function npv(rate, cashflows) {
      if (!cashflows.length) return 0;
      return cashflows.reduce((acc, val, i) => acc + val/Math.pow(1+rate, i), 0);
    }
    
    let rate = 0.1, epsilon = 1e-6, maxIter = 100;
    for (let iter=0; iter<maxIter; iter++) {
      let npv0 = npv(rate, cashflows);
      let npv1 = npv(rate+epsilon, cashflows);
      let deriv = (npv1-npv0)/epsilon;
      if (Math.abs(deriv) < 1e-10) break;
      let newRate = rate - npv0/deriv;
      if (!isFinite(newRate)) break;
      if (Math.abs(newRate-rate) < 1e-7) return newRate;
      rate = newRate;
    }
    return NaN;
  }

  // --- ROI TABLE RENDERING (SINGLE TABLE, OVERLAY SUGGESTIONS) ---
  function renderRoiPaybackTable({actualRepayments, suggestedRepayments, filteredWeeks, weekLabels, weekStartDates, investmentWeekIndex, explicitSchedule}) {
    if (!weekLabels) return '';
    
    // Use extended week data if available (for suggestions that go beyond original schedule)
    const effectiveWeekLabels = window.extendedWeekLabels || weekLabels;
    const effectiveWeekStartDates = window.extendedWeekStartDates || weekStartDates;
    
    // Always show the table - actual repayments are always displayed even if zero
    const hasActualRepayments = actualRepayments && actualRepayments.length > 0;
    const hasSuggestedRepayments = suggestedRepayments && suggestedRepayments.length > 0 && suggestedRepayments.some(r => r > 0);
    const hasExplicitDates = explicitSchedule && explicitSchedule.length > 0;
    
    // Initialize actual repayments array if not provided
    if (!hasActualRepayments) {
      actualRepayments = Array(effectiveWeekLabels.length).fill(0);
    }
    
    const discountRate = parseFloat(document.getElementById('roiInterestInput').value) || 0;
    
    let tableHtml = `
      <table class="table table-sm">
        <thead>
          <tr>
            <th>Period</th>
            <th>Date</th>
            <th>Actual Repayment</th>
            ${suggestedRepayments ? '<th>Suggested Repayment</th>' : ''}
            <th>Cumulative Actual</th>
            ${suggestedRepayments ? '<th>Cumulative Suggested</th>' : ''}
            <th>Discounted Cumulative</th>
          </tr>
        </thead>
        <tbody>
    `;
    
    let cum = 0, discCum = 0;
    let sugCum = 0, sugDiscCum = 0;
    
    // Determine the maximum range to display
    const maxWeeks = Math.max(
      actualRepayments.length,
      hasSuggestedRepayments ? suggestedRepayments.length : 0,
      effectiveWeekLabels.length
    );
    
    // Create a combined index set of all weeks that need to be shown
    const weeksToShow = new Set();
    
    // Add all weeks with actual repayments (even if zero) from investment week onwards
    for (let i = investmentWeekIndex + 1; i < actualRepayments.length; i++) {
      weeksToShow.add(i);
    }
    
    // Add weeks with suggested repayments
    if (hasSuggestedRepayments) {
      for (let i = 0; i < suggestedRepayments.length; i++) {
        if (suggestedRepayments[i] > 0) {
          weeksToShow.add(i);
        }
      }
    }
    
    // Convert to sorted array and ensure we include at least some weeks for display
    const sortedWeeks = Array.from(weeksToShow).sort((a, b) => a - b);
    
    // If no weeks to show, show at least a few weeks from investment onwards for context
    if (sortedWeeks.length === 0) {
      for (let i = investmentWeekIndex + 1; i < Math.min(investmentWeekIndex + 5, effectiveWeekLabels.length); i++) {
        sortedWeeks.push(i);
      }
    }
    
    // Handle explicit dates vs week-based display
    if (hasExplicitDates) {
      // For explicit dates, show entries sorted by date
      explicitSchedule.forEach((scheduleItem, index) => {
        const actualRepayment = scheduleItem.amount;
        const suggestedRepayment = (suggestedRepayments && suggestedRepayments[index]) || 0;
        
        cum += actualRepayment;
        if (actualRepayment > 0) {
          // Use actual days from investment date for accurate discounting
          const investmentDate = effectiveWeekStartDates[investmentWeekIndex] || new Date(2025, 0, 1);
          const repaymentDate = scheduleItem.date;
          
          const days = (repaymentDate - investmentDate) / (24 * 3600 * 1000);
          const years = days / 365.25; // Use 365.25 for consistency with XIRR
          discCum += actualRepayment / Math.pow(1 + discountRate / 100, years);
        }
        
        if (suggestedRepayments) {
          sugCum += suggestedRepayment;
          if (suggestedRepayment > 0) {
            const investmentDate = effectiveWeekStartDates[investmentWeekIndex] || new Date(2025, 0, 1);
            const repaymentDate = scheduleItem.date;
            const days = (repaymentDate - investmentDate) / (24 * 3600 * 1000);
            const years = days / 365.25;
            sugDiscCum += suggestedRepayment / Math.pow(1 + discountRate / 100, years);
          }
        }
        
        // Show explicit date and a generated label
        const weekLabel = `Repayment ${index + 1}`;
        const weekDate = scheduleItem.date.toLocaleDateString('en-GB');
        
        // Highlight rows with suggested repayments
        const rowStyle = suggestedRepayments && suggestedRepayment > 0 ? 'style="background-color: rgba(33, 150, 243, 0.05);"' : '';
        
        tableHtml += `
          <tr ${rowStyle}>
            <td>${weekLabel}</td>
            <td style="font-weight: bold; color: #1976d2;">${weekDate}</td>
            <td>€${actualRepayment.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>
            ${suggestedRepayments ? `<td style="color: #2196f3; font-weight: bold;">€${suggestedRepayment.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>` : ''}
            <td>€${cum.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>
            ${suggestedRepayments ? `<td style="color: #2196f3;">€${sugCum.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>` : ''}
            <td>€${discCum.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>
          </tr>
        `;
      });
    } else {
      // Use traditional week-based display
      for (const weekIndex of sortedWeeks) {
        const actualRepayment = (actualRepayments && actualRepayments[weekIndex]) || 0;
        const suggestedRepayment = (suggestedRepayments && suggestedRepayments[weekIndex]) || 0;
        
        cum += actualRepayment;
        if (actualRepayment > 0) {
          // Use actual days from investment date for accurate discounting
          const investmentDate = effectiveWeekStartDates[investmentWeekIndex];
          const repaymentDate = effectiveWeekStartDates[weekIndex];
          if (investmentDate && repaymentDate) {
            const days = (repaymentDate - investmentDate) / (24 * 3600 * 1000);
            const years = days / 365.25; // Use 365.25 for consistency with XIRR
            discCum += actualRepayment / Math.pow(1 + discountRate / 100, years);
          } else {
            // Fallback to period-based calculation if dates unavailable
            const periodIndex = weekIndex - investmentWeekIndex;
            discCum += actualRepayment / Math.pow(1 + discountRate / 100, periodIndex);
          }
        }
        
        if (suggestedRepayments) {
          sugCum += suggestedRepayment;
          if (suggestedRepayment > 0) {
            // Use actual days from investment date for accurate discounting  
            const investmentDate = effectiveWeekStartDates[investmentWeekIndex];
            const repaymentDate = effectiveWeekStartDates[weekIndex];
            if (investmentDate && repaymentDate) {
              const days = (repaymentDate - investmentDate) / (24 * 3600 * 1000);
              const years = days / 365.25; // Use 365.25 for consistency with XIRR
              sugDiscCum += suggestedRepayment / Math.pow(1 + discountRate / 100, years);
            } else {
              // Fallback to period-based calculation if dates unavailable
              const periodIndex = weekIndex - investmentWeekIndex;
              sugDiscCum += suggestedRepayment / Math.pow(1 + discountRate / 100, periodIndex);
            }
          }
        }
        
        const weekLabel = effectiveWeekLabels[weekIndex] || `Week ${weekIndex + 1}`;
        const weekDate = effectiveWeekStartDates[weekIndex] ? effectiveWeekStartDates[weekIndex].toLocaleDateString('en-GB') : '-';
        
        // Highlight rows with suggested repayments
        const rowStyle = suggestedRepayments && suggestedRepayment > 0 ? 'style="background-color: rgba(33, 150, 243, 0.05);"' : '';
        
        tableHtml += `
          <tr ${rowStyle}>
            <td>${weekLabel}</td>
            <td>${weekDate}</td>
            <td>€${actualRepayment.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>
            ${suggestedRepayments ? `<td style="color: #2196f3; font-weight: bold;">€${suggestedRepayment.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>` : ''}
            <td>€${cum.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>
            ${suggestedRepayments ? `<td style="color: #2196f3;">€${sugCum.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>` : ''}
            <td>€${discCum.toLocaleString(undefined, {maximumFractionDigits: 2})}</td>
          </tr>
        `;
      }
    }
    
    tableHtml += `</tbody></table>`;
    
    // Add suggested summary if suggestions are shown
    if (suggestedRepayments && achievedSuggestedIRR !== null) {
      tableHtml += `
        <div style="margin-top: 10px; padding: 10px; background-color: rgba(33, 150, 243, 0.1); border-radius: 4px;">
          <strong>Suggested Repayments Summary:</strong><br>
          Total Suggested: €${sugCum.toLocaleString(undefined, {maximumFractionDigits: 2})}<br>
          Suggested Discounted Total: €${sugDiscCum.toLocaleString(undefined, {maximumFractionDigits: 2})}<br>
          Achieved IRR: ${isFinite(achievedSuggestedIRR) && !isNaN(achievedSuggestedIRR) ? (achievedSuggestedIRR * 100).toFixed(2) + '%' : 'n/a'}
        </div>
      `;
    }
    
    return tableHtml;
  }

  // -------------------- ROI/Payback Section --------------------
function renderRoiSection() {
  const dropdown = document.getElementById('investmentWeek');
  if (!dropdown) return;
  
  investmentWeekIndex = parseInt(dropdown.value, 10) || 0;
  const investment = parseFloat(document.getElementById('roiInvestmentInput').value) || 0;
  const discountRate = parseFloat(document.getElementById('roiInterestInput').value) || 0;
  const investmentWeek = investmentWeekIndex;
  const investmentDate = weekStartDates[investmentWeek] || null;

  // Handle case when no week mapping is available - use default weeks
  let actualWeekLabels = weekLabels && weekLabels.length > 0 ? weekLabels : Array.from({length: 52}, (_, i) => `Week ${i + 1}`);
  let actualWeekStartDates = weekStartDates && weekStartDates.length > 0 ? weekStartDates : Array.from({length: 52}, (_, i) => new Date(2025, 0, 1 + i * 7));
  
  // Get explicit repayment schedule for accurate date-based calculations
  const explicitSchedule = getExplicitRepaymentSchedule();
  
  // Check if we have any explicit date repayments
  const hasExplicitDates = explicitSchedule.some(item => 
    repaymentRows.some(r => r.type === "date" && r.explicitDate)
  );
  
  let cashflows, cashflowDates;
  
  if (hasExplicitDates && explicitSchedule.length > 0) {
    // Use explicit repayment schedule for NPV/IRR calculations
    const investmentScheduleDate = investmentDate || actualWeekStartDates[investmentWeek] || new Date(2025, 0, 1);
    
    // Filter schedule to only include repayments after investment date
    const futureRepayments = explicitSchedule.filter(item => item.date >= investmentScheduleDate);
    
    cashflows = [-investment, ...futureRepayments.map(item => item.amount)];
    cashflowDates = [investmentScheduleDate, ...futureRepayments.map(item => item.date)];
  } else {
    // Fall back to traditional week-based approach
    const repaymentsFull = getRepaymentArr ? getRepaymentArr() : [];
    const repayments = repaymentsFull.slice(investmentWeek);
    
    cashflows = [-investment, ...repayments];
    cashflowDates = [investmentDate];
    for (let i = 1; i < cashflows.length; i++) {
      let idx = investmentWeek + i;
      cashflowDates[i] = actualWeekStartDates[idx] || null;
    }
  }

  function npv(rate, cashflows) {
    if (!cashflows.length) return 0;
    return cashflows.reduce((acc, val, i) => acc + val/Math.pow(1+rate, i), 0);
  }
  
  // Legacy IRR function - kept for compatibility but replaced by XIRR for irregular schedules
  function irr(cashflows, guess=0.1) {
    let rate = guess, epsilon = 1e-6, maxIter = 100;
    for (let iter=0; iter<maxIter; iter++) {
      let npv0 = npv(rate, cashflows);
      let npv1 = npv(rate+epsilon, cashflows);
      let deriv = (npv1-npv0)/epsilon;
      let newRate = rate - npv0/deriv;
      if (!isFinite(newRate)) break;
      if (Math.abs(newRate-rate) < 1e-7) return newRate;
      rate = newRate;
    }
    return NaN;
  }

  function npv_date(rate, cashflows, dateArr) {
    const msPerDay = 24 * 3600 * 1000;
    const baseDate = dateArr[0];
    return cashflows.reduce((acc, val, i) => {
      if (!dateArr[i]) return acc;
      let days = (dateArr[i] - baseDate) / msPerDay;
      let years = days / 365.25; // Use 365.25 for consistency with XIRR and more accurate annualization
      return acc + val / Math.pow(1 + rate, years);
    }, 0);
  }

  // Calculate NPV using correct date-based formula: NPV = sum(CF_i / (1 + r)^(t_i/365.25)) - Investment
  // where r = annual discount rate from ROI input, t_i = days since investment
  let npvVal = (discountRate && cashflows.length > 1 && cashflowDates[0]) ?
    npv_date(discountRate / 100, cashflows, cashflowDates) : null;
  
  // Use XIRR for accurate annualized IRR calculation with actual cash flow dates
  // XIRR replaces standard IRR to handle irregular repayment schedules properly.
  // This ensures accurate annualized returns regardless of payment timing irregularities.
  let irrVal = (cashflows.length > 1 && cashflowDates.length > 1 && cashflowDates[0]) ? 
    xirr(cashflows, cashflowDates) : NaN;

  // Calculate discounted payback using actual repayment dates
  let discCum = 0, payback = null;
  for (let i = 1; i < cashflows.length; i++) {
    let discounted;
    
    if (cashflowDates[0] && cashflowDates[i]) {
      const days = (cashflowDates[i] - cashflowDates[0]) / (24 * 3600 * 1000);
      const years = days / 365.25; // Use 365.25 for consistency with XIRR
      discounted = cashflows[i] / Math.pow(1 + discountRate / 100, years);
    } else {
      // Fallback to period-based calculation if dates unavailable
      discounted = cashflows[i] / Math.pow(1 + discountRate / 100, i);
    }
    
    discCum += discounted;
    if (payback === null && discCum >= investment) payback = i;
  }

  // Prepare data for table rendering
  let tableRepayments;
  if (hasExplicitDates && explicitSchedule.length > 0) {
    // For explicit dates, create a combined schedule for display
    tableRepayments = [];
    const totalScheduleItems = Math.max(explicitSchedule.length, actualWeekLabels.length);
    
    // Create expanded table data that includes both week-based and explicit date repayments
    for (let i = 0; i < totalScheduleItems; i++) {
      const explicitItem = explicitSchedule[i];
      const weekRepayment = i < getRepaymentArr().length ? getRepaymentArr()[i] : 0;
      
      if (explicitItem) {
        tableRepayments.push(explicitItem.amount);
      } else {
        tableRepayments.push(weekRepayment);
      }
    }
  } else {
    // Use traditional week-based repayments
    const repaymentsFull = getRepaymentArr ? getRepaymentArr() : [];
    tableRepayments = repaymentsFull.slice(investmentWeek);
  }

  // Instead of inline table generation, use renderRoiPaybackTable
  const filteredWeeks = getFilteredWeekIndices ? getFilteredWeekIndices() : Array.from({length: actualWeekLabels.length}, (_, i) => i);
  const tableHtml = renderRoiPaybackTable({
    actualRepayments: tableRepayments,
    suggestedRepayments: showSuggestions ? suggestedRepayments : null,
    filteredWeeks,
    weekLabels: actualWeekLabels,
    weekStartDates: actualWeekStartDates,
    investmentWeekIndex: investmentWeek,
    explicitSchedule: hasExplicitDates ? explicitSchedule : null
  });

  // Calculate total repayments for summary
  const totalRepayments = hasExplicitDates && explicitSchedule.length > 0 
    ? explicitSchedule.reduce((sum, item) => sum + item.amount, 0)
    : (tableRepayments ? tableRepayments.reduce((a, b) => a + b, 0) : 0);

  let summary = `<b>Total Investment:</b> €${investment.toLocaleString()}<br>
    <b>Total Repayments:</b> €${totalRepayments.toLocaleString()}<br>
    <b>NPV (${discountRate}%):</b> ${typeof npvVal === "number" ? "€" + npvVal.toLocaleString(undefined, { maximumFractionDigits: 2 }) : "n/a"}<br>
    <b>IRR:</b> ${isFinite(irrVal) && !isNaN(irrVal) ? (irrVal * 100).toFixed(2) + '%' : 'n/a'}<br>
    <b>Discounted Payback (periods):</b> ${payback ?? 'n/a'}`;

  // Add note about explicit dates if used
  if (hasExplicitDates) {
    summary += `<br><div style="margin-top:8px; padding:4px; background-color:#e3f2fd; border-radius:4px; font-size:0.9em;">
      <strong>Note:</strong> Calculations use explicit dates provided for repayments.
    </div>`;
  }

  // Show achievedSuggestedIRR if present
  if (showSuggestions && achievedSuggestedIRR !== null && isFinite(achievedSuggestedIRR)) {
    summary += `<br><b>Suggested IRR:</b> ${(achievedSuggestedIRR * 100).toFixed(2)}%`;
  }

  let badge = '';
  if (irrVal > 0.15) badge = '<div class="alert alert-success" style="margin-bottom: 1em;"><strong>Attractive ROI</strong> - This investment shows excellent returns</div>';
  else if (irrVal > 0.08) badge = '<div class="alert alert-warning" style="margin-bottom: 1em;"><strong>Moderate ROI</strong> - This investment shows reasonable returns</div>';
  else if (!isNaN(irrVal)) badge = '<div class="alert alert-danger" style="margin-bottom: 1em;"><strong>Low ROI</strong> - This investment shows poor returns</div>';
  else badge = '';

  // Update the prominent warning display
  const warningElement = document.getElementById('roiWarningAlert');
  if (warningElement) {
    warningElement.innerHTML = badge;
  }

  document.getElementById('roiSummary').innerHTML = summary;
  document.getElementById('roiPaybackTableWrap').innerHTML = tableHtml;

  // Charts
  renderRoiCharts(investment, repayments);

  if (!repayments.length || repayments.reduce((a, b) => a + b, 0) === 0) {
    document.getElementById('roiSummary').innerHTML += '<div class="alert alert-warning">No repayments scheduled. ROI cannot be calculated.</div>';
  }
}

// ROI Performance Chart (line) + Pie chart
function renderRoiCharts(investment, repayments) {
  if (!Array.isArray(repayments) || repayments.length === 0) return;

  // Build cumulative and discounted cumulative arrays
  let cumArr = [];
  let discCumArr = [];
  let cum = 0, discCum = 0;
  const discountRate = parseFloat(document.getElementById('roiInterestInput').value) || 0;
  for (let i = 0; i < repayments.length; i++) {
    cum += repayments[i] || 0;
    cumArr.push(cum);

    // Discounted only if repayment > 0
    if (repayments[i] > 0) {
      discCum += repayments[i] / Math.pow(1 + discountRate / 100, i + 1);
    }
    discCumArr.push(discCum);
  }

  // Build X labels
  const weekLabels = window.weekLabels || repayments.map((_, i) => `Week ${i + 1}`);

  // ROI Performance Chart (Line)
  let roiLineElem = document.getElementById('roiLineChart');
  if (roiLineElem) {
    const roiLineCtx = roiLineElem.getContext('2d');
    if (window.roiLineChart && typeof window.roiLineChart.destroy === "function") window.roiLineChart.destroy();
    window.roiLineChart = new Chart(roiLineCtx, {
      type: 'line',
      data: {
        labels: weekLabels.slice(0, repayments.length),
        datasets: [
          {
            label: "Cumulative Repayments",
            data: cumArr,
            borderColor: "#4caf50",
            backgroundColor: "#4caf5040",
            fill: false,
            tension: 0.15
          },
          {
            label: "Discounted Cumulative",
            data: discCumArr,
            borderColor: "#1976d2",
            backgroundColor: "#1976d240",
            borderDash: [6,4],
            fill: false,
            tension: 0.15
          },
          {
            label: "Initial Investment",
            data: Array(repayments.length).fill(investment),
            borderColor: "#f44336",
            borderDash: [3,3],
            borderWidth: 1,
            pointRadius: 0,
            fill: false
          }
        ]
      },
      options: {
        responsive: true,
        plugins: { legend: { display: true } },
        scales: {
          y: { beginAtZero: true, title: { display: true, text: "€" } }
        }
      }
    });
  }

  // Pie chart (optional)
  let roiPieElem = document.getElementById('roiPieChart');
  if (roiPieElem) {
    const roiPieCtx = roiPieElem.getContext('2d');
    if (window.roiPieChart && typeof window.roiPieChart.destroy === "function") window.roiPieChart.destroy();
    window.roiPieChart = new Chart(roiPieCtx, {
      type: 'pie',
      data: {
        labels: ["Total Repayments", "Unrecouped"],
        datasets: [{
          data: [
            cumArr[cumArr.length - 1] || 0,
            Math.max(investment - (cumArr[cumArr.length - 1] || 0), 0)
          ],
          backgroundColor: ["#4caf50", "#f3b200"]
        }]
      },
      options: { responsive: true, maintainAspectRatio: false }
    });
  }
}

// --- ROI input events ---
document.getElementById('roiInvestmentInput').addEventListener('input', function() {
  clearRoiSuggestions();
  renderRoiSection();
  // Update NPV display in modal if open
  const modal = document.getElementById('targetIrrModal');
  const npvDisplay = document.getElementById('equivalentNpvDisplay');
  if (modal && modal.style.display !== 'none' && npvDisplay) {
    const slider = document.getElementById('targetIrrSlider');
    if (slider) {
      const irrRate = parseFloat(slider.value) / 100;
      const investment = parseFloat(this.value) || 0;
      if (investment <= 0) {
        npvDisplay.textContent = '€0';
      } else {
        // Update NPV display
        updateNPVDisplayInModal();
      }
    }
  }
});
document.getElementById('roiInterestInput').addEventListener('input', function() {
  clearRoiSuggestions();
  renderRoiSection();
  // Update NPV display in modal if open
  updateNPVDisplayInModal();
});
document.getElementById('refreshRoiBtn').addEventListener('click', function() {
  clearRoiSuggestions();
  renderRoiSection();
  updateNPVDisplayInModal();
});
document.getElementById('investmentWeek').addEventListener('change', function() {
  clearRoiSuggestions();
  renderRoiSection();
  updateNPVDisplayInModal();
});

// Helper function to update NPV display when modal is open
function updateNPVDisplayInModal() {
  const modal = document.getElementById('targetIrrModal');
  const npvDisplay = document.getElementById('equivalentNpvDisplay');
  const slider = document.getElementById('targetIrrSlider');
  
  if (modal && modal.style.display !== 'none' && npvDisplay && slider) {
    const irrRate = parseFloat(slider.value) / 100;
    const investment = parseFloat(document.getElementById('roiInvestmentInput').value) || 0;
    
    if (investment <= 0) {
      npvDisplay.textContent = '€0';
      return;
    }
    
    // Calculate NPV for current IRR
    const repaymentsFull = getRepaymentArr ? getRepaymentArr() : [];
    const repayments = repaymentsFull.slice(investmentWeekIndex);
    
    if (repayments.length === 0 || repayments.every(r => r === 0)) {
      // Calculate what total repayments would be needed for this IRR
      const targetReturn = investment * (1 + irrRate);
      const npvValue = targetReturn - investment;
      npvDisplay.textContent = `€${npvValue.toLocaleString(undefined, {maximumFractionDigits: 2})}`;
    } else {
      // Use actual repayments with date-based discounting
      const cashflows = [-investment, ...repayments];
      const cashflowDates = [weekStartDates[investmentWeekIndex] || new Date()];
      
      for (let i = 1; i < cashflows.length; i++) {
        let idx = investmentWeekIndex + i;
        cashflowDates[i] = weekStartDates[idx] || new Date();
      }
      
      // Use date-based NPV calculation with correct formula: 
      // NPV = sum(CF_i / (1 + r)^(t_i/365.25)) - Investment
      const npvValue = npv_date(irrRate, cashflows, cashflowDates);
      npvDisplay.textContent = `€${npvValue.toLocaleString(undefined, {maximumFractionDigits: 2})}`;
    }
  }
}

// --- SUGGESTION BUTTON EVENT ---
document.getElementById('showSuggestedRepaymentsBtn').addEventListener('click', function() {
  if (!showSuggestions) {
    // Generate suggestions using current target IRR and installment count
    showSuggestions = true;
    this.textContent = 'Hide Suggested Repayments';
    generateAndUpdateSuggestions();
  } else {
    // Hide suggestions
    clearRoiSuggestions();
    this.textContent = 'Show Suggested Repayments';
    renderRoiSection();
  }
});

// --- CLEAR SUGGESTIONS WHEN DATA CHANGES ---
function clearRoiSuggestions() {
  showSuggestions = false;
  suggestedRepayments = null;
  achievedSuggestedIRR = null;
  clearSuggestionWarnings();
  const btn = document.getElementById('showSuggestedRepaymentsBtn');
  if (btn) btn.textContent = 'Show Suggested Repayments';
}

// -------------------- EXCEL EXPORT FUNCTIONALITY --------------------
function setupExcelExport() {
  const exportBtn = document.getElementById('exportToExcelBtn');
  if (!exportBtn) return;
  
  exportBtn.addEventListener('click', function() {
    try {
      // Create workbook
      const workbook = XLSX.utils.book_new();
      
      // Get current repayments data
      const actualRepayments = getRepaymentArr ? getRepaymentArr() : [];
      const investment = parseFloat(document.getElementById('roiInvestmentInput').value) || 0;
      const discountRate = parseFloat(document.getElementById('roiInterestInput').value) || 0;
      
      // Use mapped week labels if available
      const actualWeekLabels = weekLabels && weekLabels.length > 0 ? weekLabels : 
        Array.from({length: 52}, (_, i) => `Week ${i + 1}`);
      const actualWeekStartDates = weekStartDates && weekStartDates.length > 0 ? weekStartDates : 
        Array.from({length: 52}, (_, i) => new Date(2025, 0, 1 + i * 7));
      
      // Sheet 1: Repayments Inputted (Actual)
      const actualData = [];
      actualData.push(['Week', 'Date', 'Repayment Amount', 'Cumulative Total', 'Discounted Cumulative']);
      
      let cumulative = 0;
      let discountedCumulative = 0;
      
      for (let i = investmentWeekIndex + 1; i < actualRepayments.length; i++) {
        const repayment = actualRepayments[i] || 0;
        if (repayment > 0) {
          cumulative += repayment;
          const periodIndex = i - investmentWeekIndex;
          discountedCumulative += repayment / Math.pow(1 + discountRate / 100, periodIndex);
          
          actualData.push([
            actualWeekLabels[i] || `Week ${i + 1}`,
            actualWeekStartDates[i] ? actualWeekStartDates[i].toLocaleDateString('en-GB') : '-',
            repayment,
            cumulative,
            discountedCumulative
          ]);
        }
      }
      
      // Add summary row for actual repayments
      if (actualData.length > 1) {
        actualData.push(['', '', '', '', '']);
        actualData.push(['SUMMARY', '', '', '', '']);
        actualData.push(['Total Investment', '', investment, '', '']);
        actualData.push(['Total Repayments', '', cumulative, '', '']);
        actualData.push(['Net Return', '', cumulative - investment, '', '']);
      }
      
      const actualSheet = XLSX.utils.aoa_to_sheet(actualData);
      XLSX.utils.book_append_sheet(workbook, actualSheet, 'Repayments Inputted');
      
      // Sheet 2: Adjusted IRR Suggestions (if available)
      if (showSuggestions && suggestedRepayments) {
        const suggestedData = [];
        suggestedData.push(['Week', 'Date', 'Suggested Amount', 'Cumulative Total', 'Discounted Cumulative']);
        
        let sugCumulative = 0;
        let sugDiscountedCumulative = 0;
        
        // Use extended week data if available
        const effectiveWeekLabels = window.extendedWeekLabels || actualWeekLabels;
        const effectiveWeekStartDates = window.extendedWeekStartDates || actualWeekStartDates;
        
        for (let i = 0; i < suggestedRepayments.length; i++) {
          const suggestedAmount = suggestedRepayments[i] || 0;
          if (suggestedAmount > 0) {
            sugCumulative += suggestedAmount;
            const periodIndex = i - investmentWeekIndex;
            sugDiscountedCumulative += suggestedAmount / Math.pow(1 + discountRate / 100, periodIndex);
            
            suggestedData.push([
              effectiveWeekLabels[i] || `Week ${i + 1}`,
              effectiveWeekStartDates[i] ? effectiveWeekStartDates[i].toLocaleDateString('en-GB') : '-',
              suggestedAmount,
              sugCumulative,
              sugDiscountedCumulative
            ]);
          }
        }
        
        // Add summary row for suggested repayments
        if (suggestedData.length > 1) {
          suggestedData.push(['', '', '', '', '']);
          suggestedData.push(['SUMMARY', '', '', '', '']);
          suggestedData.push(['Total Investment', '', investment, '', '']);
          suggestedData.push(['Total Suggested Repayments', '', sugCumulative, '', '']);
          suggestedData.push(['Net Return', '', sugCumulative - investment, '', '']);
          if (achievedSuggestedIRR !== null && isFinite(achievedSuggestedIRR)) {
            suggestedData.push(['Achieved IRR', '', (achievedSuggestedIRR * 100).toFixed(2) + '%', '', '']);
          }
        }
        
        const suggestedSheet = XLSX.utils.aoa_to_sheet(suggestedData);
        XLSX.utils.book_append_sheet(workbook, suggestedSheet, 'Adjusted IRR Suggestions');
      }
      
      // Generate and download the file
      const fileName = `MATRIX_Repayments_${new Date().toISOString().split('T')[0]}.xlsx`;
      XLSX.writeFile(workbook, fileName);
      
    } catch (error) {
      console.error('Error exporting to Excel:', error);
      alert('Error exporting to Excel. Please ensure you have a modern browser that supports file downloads.');
    }
  });
}

// Initialize Excel export functionality
setupExcelExport();

  // -------------------- Update All Tabs --------------------
  function updateAllTabs() {
    clearRoiSuggestions(); // Clear suggestions when data changes
    renderRepaymentRows();
    updateLoanSummary();
    updateChartAndSummary();
    renderPnlTables();
    renderSummaryTab();
    renderRoiSection();
    renderTornadoChart();
    
    // Update NPV display in modal if open
    updateNPVDisplayInModal();
  }
  updateAllTabs();
});
