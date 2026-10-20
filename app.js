/**
 * ============================================================================
 * SCBT316 eNose 3D Interactive Analytics Dashboard
 * 100% Dynamic Multi-Sheet .xlsx Schema & Sample Parser Engine
 * - Zero hardcoded sample class names
 * - Dynamic Regex Sheet Scanner: /^(Raw|Sampling)_(.+)$/i
 * - Auto-classification: "ข้อมูลตัวอย่าง (Samples)" vs "ข้อมูลอ้างอิง (Blanks)"
 * - Golden-Angle HSL Dynamic Color Palette Generator
 * - Dynamic Group Filter from BlankAssignment / Date columns
 * - Dynamic Sensor Columns (G1..Gn), Cycle Loop, and Mode filters
 * - Paired scatter3d + mesh3d (3D Convex Hull / Ellipsoid) per class
 * - Locked 3D Camera State (uirevision: 'constant')
 * ============================================================================
 */

(function () {
  'use strict';

  // Camera Presets for 3D Orbit / Turntable
  const CAMERA_PRESETS = {
    isometric: { eye: { x: 1.55, y: 1.55, z: 1.15 }, up: { x: 0, y: 0, z: 1 }, center: { x: 0, y: 0, z: 0 } },
    top:       { eye: { x: 0.01, y: 0.01, z: 2.45 }, up: { x: 0, y: 1, z: 0 }, center: { x: 0, y: 0, z: 0 } },
    side:      { eye: { x: 0.01, y: -2.4, z: 0.05 }, up: { x: 0, y: 0, z: 1 }, center: { x: 0, y: 0, z: 0 } },
    front:     { eye: { x: 2.4,  y: 0.01, z: 0.05 }, up: { x: 0, y: 0, z: 1 }, center: { x: 0, y: 0, z: 0 } }
  };

  // Ignored metadata / non-sensor sheet names
  const IGNORED_SHEET_REGEX = /^(readme|read_me|info|metadata|summary|instruction|instructions|notes?|config)$/i;
  const BLANK_ASSIGN_SHEET_REGEX = /^blank[\s_\-]*assignment$/i;

  // --------------------------------------------------------------------------
  // 1. DYNAMIC APPLICATION STATE
  // --------------------------------------------------------------------------
  const state = {
    fileName: '',
    isDemoData: true,
    sheetSummary: [],

    // Protocol Metadata extracted from ReadMe sheet (or fallback)
    protocolInfo: {
      code: 'R10–S120–P60–F200',
      desc: 'Resting 10 s • Sampling 120 s • Purging 60 s • Flow 200',
      rawText: 'Protocol: R10–S120–P60–F200 (Resting 10 s, Sampling 120 s, Purging 60 s, Flow Rate setpoint 200)'
    },

    // Dynamically discovered schema from the active workbook
    sampleClasses: [],      // [{ id, label, isBlank, color, groupLabel, assignedBlank }]
    sampleMap: {},          // { [sampleId]: sampleObj }
    sensorCols: [],         // e.g. ['G1', 'G2', ..., 'G15'] discovered dynamically
    extraNumericCols: [],   // e.g. ['Temp', 'Flow Rate'] discovered dynamically
    availableCycles: [],    // e.g. [1, 2, 3, 4, 5, 6, 7] discovered dynamically
    availableModes: [],     // e.g. ['Resting', 'Sampling', 'Purging'] discovered dynamically
    dynamicGroups: [],      // [{ id, label, sampleIds: [...] }] discovered from BlankAssignment / Date

    // Datasets keyed by 'Sampling' and 'Raw', each mapping sampleId -> rows[]
    datasets: {
      Sampling: {},
      Raw: {}
    },

    // BlankAssignment mapping: { [sampleId]: blankSampleId }
    blankAssignment: {},

    // Active Pipeline Controls
    activeDataset: 'Raw',      // 'Sampling' | 'Raw'
    baselineCorrection: false,
    perCycleBlank: true,       // Subtract Blank from the same Cycle Loop k

    // Axis & 3D PCA Controls
    axisMode: 'pca',          // 'custom' | 'pca'
    customAxes: {
      x: '',
      y: '',
      z: ''
    },
    defaultCustomAxes: {
      x: '',
      y: '',
      z: ''
    },
    pcaStandardize: true,     // Z-Score Standardization ((X - μ) / σ)
    pcaGranularity: 'cycle',  // 'all' (All Time-Points Mode) | 'cycle' (Cycle Summary Mode: 1 point/cycle)
    pcaSummaryMethod: 'mean', // 'mean' | 'max' | 'delta_mean' | 'delta_max'
    pcaFitScope: 'active',    // 'global' (Global PCA all classes) | 'active' (Dynamic Re-fit on Active Samples)
    excludeBlanksPca: true,   // Exclude Blank classes when computing PCA eigenvectors
    pcaResult: null,

    // Multi-View 7-Tab Suite Controls (toggled via 'แสดงผลการวิเคราะห์รูปแบบอื่นๆ')
    analysisSuiteOpen: false, // Whether the 7-tab analysis bar & suite view is open
    activeViewTab: 'pca3d',   // 'pca3d' | 'biplot' | 'scree' | 'radar' | 'heatmap' | 'timeseries' | 'qc'
    lastSecondaryTab: 'biplot',
    scoreDimension: '3d',     // '3d' | '2d'
    show3dBiplot: false,      // Overlay 3D sensor loading vectors on 3D PCA plot
    showCentroids: true,      // Show group centroids on 3D / 2D score plots
    showBiplotVectors: true,  // Master switch: show/hide all sensor loading vectors on Tab 2 Biplot
    activeBiplotSensors: new Set(), // Individual sensor vectors enabled on Tab 2 Biplot (G1..Gn)
    radarScale: 'zscore',     // 'zscore' | 'norm' | 'raw'
    heatmapSort: 'pc1',       // 'pc1' | 'name'
    tsSource: 'Raw',          // 'Raw' | 'Sampling'
    tsMode: 'compare_samples',// 'compare_samples' | 'multi_sensor'
    tsSensor: '',
    tsSample: '',
    qcSubmode: 'stability',   // 'stability' | 'distance'
    distanceSpace: 'pca',     // 'pca' | 'zscore'

    // Interactive Filtering Controls
    activeSamples: new Set(),
    activeCycles: new Set(),
    activeModes: new Set(),
    sampleSearchQuery: '',

    // 3D Area Hull & Visual Controls
    showAreaHull: true,
    hullType: 'convex', // 'convex' | 'ellipsoid'
    hullOpacity: 0.20,
    markerSize: 4.5,
    focusedSample: null,

    // Plotly trace metadata for Focus Highlight restyling
    traceMeta: [],
    currentCamera: CAMERA_PRESETS.isometric,
    autoRotateTimer: null,
    autoRotateAngle: Math.PI / 4
  };

  // --------------------------------------------------------------------------
  // 2. DYNAMIC COLOR PALETTE GENERATOR (GOLDEN ANGLE HSL -> HEX)
  // --------------------------------------------------------------------------

  /**
   * Convert HSL (h: 0..360, s: 0..100, l: 0..100) to #RRGGBB Hex string
   * (Hex format ensures 100% compatibility with both Plotly scatter3d & mesh3d)
   */
  function hslToHex(h, s, l) {
    const sNorm = s / 100;
    const lNorm = l / 100;
    const c = (1 - Math.abs(2 * lNorm - 1)) * sNorm;
    const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
    const m = lNorm - c / 2;
    let rPrime = 0, gPrime = 0, bPrime = 0;

    if (h >= 0 && h < 60) { rPrime = c; gPrime = x; bPrime = 0; }
    else if (h >= 60 && h < 120) { rPrime = x; gPrime = c; bPrime = 0; }
    else if (h >= 120 && h < 180) { rPrime = 0; gPrime = c; bPrime = x; }
    else if (h >= 180 && h < 240) { rPrime = 0; gPrime = x; bPrime = c; }
    else if (h >= 240 && h < 300) { rPrime = x; gPrime = 0; bPrime = c; }
    else { rPrime = c; gPrime = 0; bPrime = x; }

    const toHex = n => Math.round((n + m) * 255).toString(16).padStart(2, '0');
    return `#${toHex(rPrime)}${toHex(gPrime)}${toHex(bPrime)}`.toUpperCase();
  }

  /**
   * Generate a distinct, high-visibility color on #0B0F19 dark background
   * using Golden Angle (137.50776°) distribution so any number of samples are clearly separated.
   */
  function generateDynamicColor(index, isBlank = false) {
    if (isBlank) {
      // Distinct cool silver / slate / ice hues for Blanks
      const blankLightness = [70, 86, 58, 78][index % 4];
      const blankHue = (205 + index * 32) % 360;
      return hslToHex(blankHue, 22, blankLightness);
    }
    const goldenAngle = 137.50776405;
    const hue = (195 + index * goldenAngle) % 360;
    const saturation = 86 - (index % 3) * 6; // 86%, 80%, 74%
    const lightness = 60 + ((index % 2) * 6); // 60%, 66%
    return hslToHex(hue, saturation, lightness);
  }

  /**
   * Simple 3D Geometric Marker Shapes supported natively by Plotly WebGL scatter3d
   * Cycled automatically across sample classes so each sample has a distinct shape + color.
   */
  const MARKER_SHAPES_3D = [
    { symbol: 'circle',       icon: '●', name: 'วงกลมทึบ (Circle)' },
    { symbol: 'square',       icon: '■', name: 'สี่เหลี่ยมทึบ (Square)' },
    { symbol: 'diamond',      icon: '◆', name: 'ข้าวหลามตัดทึบ (Diamond)' },
    { symbol: 'cross',        icon: '✚', name: 'เครื่องหมายบวก (Cross)' },
    { symbol: 'x',            icon: '✖', name: 'กากบาท (X)' },
    { symbol: 'circle-open',  icon: '○', name: 'วงกลมโปร่ง (Open Circle)' },
    { symbol: 'square-open',  icon: '□', name: 'สี่เหลี่ยมโปร่ง (Open Square)' },
    { symbol: 'diamond-open', icon: '◇', name: 'ข้าวหลามตัดโปร่ง (Open Diamond)' }
  ];

  function escapeHtml(str) {
    return String(str ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // --------------------------------------------------------------------------
  // 3. DYNAMIC SHEET & SCHEMA PARSER (ZERO HARDCODING)
  // --------------------------------------------------------------------------

  /**
   * Reset all workbook-dependent state before ingesting a newly uploaded .xlsx file
   */
  function resetWorkbookState() {
    state.sampleClasses = [];
    state.sampleMap = {};
    state.sensorCols = [];
    state.extraNumericCols = [];
    state.availableCycles = [];
    state.availableModes = [];
    state.dynamicGroups = [];
    state.datasets = { Sampling: {}, Raw: {} };
    state.blankAssignment = {};
    state.activeSamples = new Set();
    state.activeCycles = new Set();
    state.activeModes = new Set();
    state.sampleSearchQuery = '';
    state.focusedSample = null;
    state.pcaResult = null;
    state.sheetSummary = [];
    state.protocolInfo = {
      code: 'R10–S120–P60–F200',
      desc: 'Resting 10 s • Sampling 120 s • Purging 60 s • Flow 200',
      rawText: 'Protocol: R10–S120–P60–F200 (Resting 10 s, Sampling 120 s, Purging 60 s, Flow Rate setpoint 200)'
    };

    const searchInput = document.getElementById('sampleSearchInput');
    const clearBtn = document.getElementById('btnClearSampleSearch');
    if (searchInput) searchInput.value = '';
    if (clearBtn) clearBtn.style.display = 'none';
  }

  /**
   * Parse ReadMe worksheet to extract protocol metadata automatically
   */
  function parseReadMeProtocol(worksheet) {
    if (!worksheet) return;
    const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
    if (!aoa || aoa.length === 0) return;

    const allCells = [];
    for (const row of aoa) {
      if (!Array.isArray(row)) continue;
      for (const cell of row) {
        if (cell !== null && cell !== undefined && String(cell).trim() !== '') {
          allCells.push(String(cell).trim());
        }
      }
    }

    const joined = allCells.join(' | ');
    const codeMatch = joined.match(/R\d+[\s–\-—]*S\d+[\s–\-—]*P\d+(?:[\s–\-—]*F\d+)?/i);
    if (codeMatch) {
      state.protocolInfo.code = codeMatch[0].replace(/-/g, '–').toUpperCase();
    }

    const protoLine = allCells.find(line => /protocol|resting|sampling|purging/i.test(line));
    if (protoLine) {
      state.protocolInfo.rawText = protoLine;
      const parenMatch = protoLine.match(/\(([^)]+)\)/);
      if (parenMatch) {
        state.protocolInfo.desc = parenMatch[1].replace(/,\s*/g, ' • ');
      } else {
        state.protocolInfo.desc = protoLine.replace(/^protocol\s*:\s*/i, '');
      }
    }
  }

  /**
   * Extract dataset type ('Sampling' | 'Raw' | 'BlankAssignment' | 'Ignore' | 'Unprefixed')
   * and dynamic sample name from any sheet name using Regex /^(Raw|Sampling)_(.+)$/i
   */
  function parseSheetIdentity(rawSheetName) {
    const sheetName = String(rawSheetName || '').trim();
    if (!sheetName) return { role: 'Ignore', sampleId: null };

    if (BLANK_ASSIGN_SHEET_REGEX.test(sheetName) || /blank[\s_\-]*assign/i.test(sheetName)) {
      return { role: 'BlankAssignment', sampleId: null };
    }

    if (IGNORED_SHEET_REGEX.test(sheetName)) {
      return { role: 'Ignore', sampleId: null };
    }

    // Match /^(Raw|Sampling)_(.+)$/i (also tolerating space or hyphen after prefix)
    const prefixMatch = sheetName.match(/^(Raw|Sampling)[\s_\-]+(.+)$/i);
    if (prefixMatch) {
      const prefix = prefixMatch[1].toLowerCase() === 'raw' ? 'Raw' : 'Sampling';
      const extractedName = prefixMatch[2].trim();
      if (!extractedName || IGNORED_SHEET_REGEX.test(extractedName)) {
        return { role: 'Ignore', sampleId: null };
      }
      return { role: prefix, sampleId: extractedName };
    }

    return { role: 'Unprefixed', sampleId: sheetName };
  }

  /**
   * Check whether a sample name represents a Reference Blank (contains 'blank')
   */
  function isBlankSampleName(sampleId) {
    return /blank/i.test(String(sampleId || ''));
  }

  /**
   * Clean & Trim Mode string
   */
  function cleanModeValue(rawVal, fallbackMode = 'Sampling') {
    if (rawVal === undefined || rawVal === null || String(rawVal).trim() === '') {
      return fallbackMode;
    }
    const trimmed = String(rawVal).trim();
    const lower = trimmed.toLowerCase();
    if (lower.startsWith('rest')) return 'Resting';
    if (lower.startsWith('samp')) return 'Sampling';
    if (lower.startsWith('purg') || lower.startsWith('clean') || lower.startsWith('flush')) return 'Purging';
    return trimmed;
  }

  /**
   * Format Excel Time value into human-readable string
   */
  function formatTimeValue(val, rowIndex) {
    if (val === undefined || val === null || val === '') return `${rowIndex}s`;
    if (typeof val === 'number') {
      if (val >= 0 && val < 1) {
        const totalSeconds = Math.round(val * 86400);
        const hrs = Math.floor(totalSeconds / 3600);
        const mins = Math.floor((totalSeconds % 3600) / 60);
        const secs = totalSeconds % 60;
        return [hrs, mins, secs].map(v => String(v).padStart(2, '0')).join(':');
      }
      return Number.isInteger(val) ? `${val}s` : val.toFixed(1);
    }
    return String(val).trim();
  }

  function toFloat(val, fallback = 0) {
    if (typeof val === 'number') return Number.isFinite(val) ? val : fallback;
    if (val === undefined || val === null) return fallback;
    const cleaned = String(val).trim().replace(/,/g, '');
    const num = parseFloat(cleaned);
    return Number.isFinite(num) ? num : fallback;
  }

  function toInt(val, fallback = 1) {
    if (typeof val === 'number' && Number.isFinite(val)) return Math.max(1, Math.round(val));
    if (val === undefined || val === null) return fallback;
    const match = String(val).trim().match(/\d+/);
    if (match) {
      const parsed = parseInt(match[0], 10);
      if (Number.isFinite(parsed) && parsed > 0) return parsed;
    }
    return fallback;
  }

  /**
   * Parse raw Cycle Loop integer WITHOUT clamping <= 0 to 1,
   * so rows with Cycle Loop <= 0 (e.g. initial 0-loop row in Sampling_B) can be filtered out.
   */
  function parseRawCycleLoop(val, fallback = 1) {
    if (typeof val === 'number' && Number.isFinite(val)) return Math.round(val);
    if (val === undefined || val === null || String(val).trim() === '') return fallback;
    const match = String(val).trim().match(/-?\d+/);
    if (match) {
      const parsed = parseInt(match[0], 10);
      if (Number.isFinite(parsed)) return parsed;
    }
    return fallback;
  }

  /**
   * Dynamically scan a worksheet, detect sensor columns (G1..Gn), extra numeric columns,
   * Mode, Cycle Loop, Time, Date, and return { rows, detectedSensors, detectedExtras, detectedDates }
   */
  function parseDynamicWorksheet(worksheet, sampleId, defaultMode = 'Sampling') {
    const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: true });
    if (!aoa || aoa.length < 2) {
      return { rows: [], detectedSensors: [], detectedExtras: [], detectedDates: [] };
    }

    // Locate header row within first 15 rows by looking for G1, G2.., Mode, Cycle, Temp, Flow
    let headerRowIdx = 0;
    let bestScore = 0;
    const maxScan = Math.min(15, aoa.length);

    for (let r = 0; r < maxScan; r++) {
      const row = aoa[r];
      if (!Array.isArray(row)) continue;
      let score = 0;
      for (const cell of row) {
        if (cell === null || cell === undefined) continue;
        const token = String(cell).trim().toUpperCase();
        if (/^G\d+$/.test(token)) score += 3;
        if (token === 'MODE' || token.includes('CYCLE') || token.includes('TEMP') || token.includes('FLOW')) score += 2;
      }
      if (score > bestScore) {
        bestScore = score;
        headerRowIdx = r;
      }
    }

    // If no sensor-like headers were found at all, this sheet is not a valid eNose data sheet
    if (bestScore < 3) {
      return { rows: [], detectedSensors: [], detectedExtras: [], detectedDates: [] };
    }

    const rawHeaders = (aoa[headerRowIdx] || []).map(h => (h === null || h === undefined ? '' : String(h).trim()));

    const sensorColIndices = {}; // { G1: idx, G2: idx, ... }
    const extraColIndices = {};  // { 'Temp': idx, 'Flow Rate': idx, ... }
    let cycleColIdx = -1;
    let modeColIdx = -1;
    let timeColIdx = -1;
    let dateColIdx = -1;

    rawHeaders.forEach((h, idx) => {
      if (!h) return;
      const cleanH = h.trim();
      const upper = cleanH.toUpperCase().replace(/[\s_\-]+/g, '');

      // Match Sensor columns: G1, G2, ..., G15, G16...
      const sensorMatch = upper.match(/^(?:SENSOR)?(G\d+)$/);
      if (sensorMatch) {
        sensorColIndices[sensorMatch[1]] = idx;
        return;
      }

      if (upper === 'MODE' || upper === 'STATE' || upper === 'PHASE') {
        modeColIdx = idx;
      } else if (upper.includes('CYCLE') || upper.includes('LOOP') || upper === 'REP') {
        cycleColIdx = idx;
      } else if (upper === 'DATE' || upper === 'EXPERIMENTDATE' || upper === 'GROUPDATE') {
        dateColIdx = idx;
      } else if (upper.includes('TIME') || upper === 'T' || upper === 'SEC' || upper === 'ELAPSED') {
        if (timeColIdx === -1) timeColIdx = idx;
      } else if (upper.startsWith('TEMP')) {
        extraColIndices['Temp'] = idx;
      } else if (upper.includes('FLOW')) {
        extraColIndices['Flow Rate'] = idx;
      } else if (upper.includes('HUMID') || upper === 'RH') {
        extraColIndices['Humidity'] = idx;
      }
    });

    const detectedSensors = Object.keys(sensorColIndices);
    const detectedExtras = Object.keys(extraColIndices);
    const detectedDatesSet = new Set();

    if (detectedSensors.length === 0) {
      return { rows: [], detectedSensors: [], detectedExtras: [], detectedDates: [] };
    }

    const cleanedRows = [];
    for (let r = headerRowIdx + 1; r < aoa.length; r++) {
      const row = aoa[r];
      if (!Array.isArray(row)) continue;

      // Filter out invalid / pre-start rows where Cycle Loop <= 0 (e.g. Cycle Loop = 0 in Sampling_B)
      let parsedCycle = 1;
      if (cycleColIdx !== -1) {
        parsedCycle = parseRawCycleLoop(row[cycleColIdx], 1);
        if (parsedCycle <= 0) continue;
      } else {
        parsedCycle = (Math.floor(cleanedRows.length / 30) % 7) + 1;
      }

      let validSensorCount = 0;
      const record = {
        sample: sampleId,
        rowIdx: cleanedRows.length + 1
      };

      for (const sCol of detectedSensors) {
        const cIdx = sensorColIndices[sCol];
        const rawCell = row[cIdx];
        if (rawCell !== null && rawCell !== undefined && String(rawCell).trim() !== '') {
          const val = toFloat(rawCell, NaN);
          if (Number.isFinite(val)) {
            record[sCol] = val;
            validSensorCount++;
          } else {
            record[sCol] = 0;
          }
        } else {
          record[sCol] = 0;
        }
      }

      if (validSensorCount < Math.min(2, detectedSensors.length)) continue;

      for (const eCol of detectedExtras) {
        const cIdx = extraColIndices[eCol];
        record[eCol] = toFloat(row[cIdx], 0);
      }

      record['Cycle Loop'] = parsedCycle;

      record['Mode'] = modeColIdx !== -1
        ? cleanModeValue(row[modeColIdx], defaultMode)
        : defaultMode;

      record['Time'] = timeColIdx !== -1
        ? formatTimeValue(row[timeColIdx], cleanedRows.length + 1)
        : `${cleanedRows.length + 1}s`;

      if (dateColIdx !== -1 && row[dateColIdx] !== null && row[dateColIdx] !== undefined) {
        const dStr = String(row[dateColIdx]).trim();
        if (dStr) {
          record['Date'] = dStr;
          detectedDatesSet.add(dStr);
        }
      }

      cleanedRows.push(record);
    }

    return {
      rows: cleanedRows,
      detectedSensors,
      detectedExtras,
      detectedDates: Array.from(detectedDatesSet)
    };
  }

  /**
   * Dynamically parse BlankAssignment worksheet to extract:
   * - blankMap: { [sampleId]: assignedBlankId }
   * - dateOrGroupMap: { [sampleId]: groupName }
   */
  function parseDynamicBlankAssignment(worksheet, discoveredSampleIds, discoveredBlankIds) {
    const blankMap = {};
    const groupMap = {};

    if (!worksheet) return { blankMap, groupMap };

    const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
    if (!aoa || aoa.length === 0) return { blankMap, groupMap };

    // Build case-insensitive lookup for discovered sample IDs and blank IDs
    const sampleLookup = new Map();
    discoveredSampleIds.forEach(id => {
      sampleLookup.set(id.toLowerCase(), id);
      sampleLookup.set(`sampling_${id}`.toLowerCase(), id);
      sampleLookup.set(`raw_${id}`.toLowerCase(), id);
    });

    const blankLookup = new Map();
    discoveredBlankIds.forEach(bId => {
      blankLookup.set(bId.toLowerCase(), bId);
      blankLookup.set(`sampling_${bId}`.toLowerCase(), bId);
      blankLookup.set(`raw_${bId}`.toLowerCase(), bId);
    });

    for (let r = 0; r < aoa.length; r++) {
      const row = aoa[r];
      if (!Array.isArray(row)) continue;

      let rowSampleId = null;
      let rowBlankId = null;
      let rowGroupLabel = null;

      for (const cell of row) {
        if (cell === null || cell === undefined) continue;
        const token = String(cell).trim();
        if (!token) continue;
        const lower = token.toLowerCase();

        // Check if token matches a discovered Blank
        if (blankLookup.has(lower)) {
          rowBlankId = blankLookup.get(lower);
          continue;
        }
        // Also check partial match against discovered Blank IDs (e.g., cell says "12Oct" and blank is "Blank_12Oct")
        const matchedBlank = discoveredBlankIds.find(bId => {
          const suffix = bId.replace(/^blank[\s_\-]*/i, '').toLowerCase();
          return suffix && (lower === suffix || lower.includes(bId.toLowerCase()));
        });
        if (matchedBlank && isBlankSampleName(token)) {
          rowBlankId = matchedBlank;
          continue;
        }

        // Check if token matches a discovered Sample
        if (sampleLookup.has(lower) && !isBlankSampleName(sampleLookup.get(lower))) {
          rowSampleId = sampleLookup.get(lower);
          continue;
        }

        // Otherwise, if it's not a header word, it may be a Date / Group label (e.g., "12-Oct", "14-Oct")
        if (!/^(sample|tea|blank|assigned|group|date|sheet|no|id|#)/i.test(token)) {
          rowGroupLabel = token;
          if (!rowBlankId && matchedBlank) {
            rowBlankId = matchedBlank;
          }
        }
      }

      if (rowSampleId && rowBlankId) {
        blankMap[rowSampleId] = rowBlankId;
      }
      if (rowSampleId && rowGroupLabel) {
        groupMap[rowSampleId] = rowGroupLabel;
      }
    }

    return { blankMap, groupMap };
  }

  /**
   * Natural sort helper for sensor names ('G1', 'G2', ..., 'G10', 'G15') and sample names
   */
  function naturalCompare(a, b) {
    return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
  }

  /**
   * Master Workbook Ingestion Pipeline (100% Dynamic)
   */
  function ingestWorkbook(workbook, fileName, isDemo = false) {
    resetWorkbookState();

    let blankAssignWorksheet = null;
    const rawSamplingMap = new Map(); // sampleId -> rows[]
    const rawFullMap = new Map();     // sampleId -> rows[]
    const unprefixedMap = new Map();  // sampleId -> rows[]
    const sampleDateMap = new Map();  // sampleId -> dateStr

    const sensorSet = new Set();
    const extraSet = new Set();
    const sheetSummary = [];

    // 1. Scan all sheets in workbook.SheetNames
    for (const sheetName of workbook.SheetNames) {
      const ws = workbook.Sheets[sheetName];
      if (/^(readme|read_me|info|protocol)$/i.test(String(sheetName).trim())) {
        parseReadMeProtocol(ws);
      }
      const identity = parseSheetIdentity(sheetName);

      if (identity.role === 'BlankAssignment') {
        blankAssignWorksheet = ws;
        sheetSummary.push({ sheetName, role: 'BlankAssignment', sampleId: '-', rowCount: '-' });
        continue;
      }

      if (identity.role === 'Ignore' || !identity.sampleId) {
        sheetSummary.push({ sheetName, role: 'Skipped (Non-Sensor)', sampleId: '-', rowCount: '-' });
        continue;
      }

      const defaultMode = identity.role === 'Sampling' ? 'Sampling' : 'Sampling';
      const parsed = parseDynamicWorksheet(ws, identity.sampleId, defaultMode);

      if (parsed.rows.length === 0) {
        sheetSummary.push({ sheetName, role: 'Skipped (No Sensor Data)', sampleId: '-', rowCount: 0 });
        continue;
      }

      parsed.detectedSensors.forEach(s => sensorSet.add(s));
      parsed.detectedExtras.forEach(e => extraSet.add(e));
      if (parsed.detectedDates.length > 0) {
        sampleDateMap.set(identity.sampleId, parsed.detectedDates[0]);
      }

      if (identity.role === 'Sampling') {
        rawSamplingMap.set(identity.sampleId, parsed.rows);
        sheetSummary.push({ sheetName, role: 'Sampling_*', sampleId: identity.sampleId, rowCount: parsed.rows.length });
      } else if (identity.role === 'Raw') {
        rawFullMap.set(identity.sampleId, parsed.rows);
        sheetSummary.push({ sheetName, role: 'Raw_*', sampleId: identity.sampleId, rowCount: parsed.rows.length });
      } else {
        unprefixedMap.set(identity.sampleId, parsed.rows);
        sheetSummary.push({ sheetName, role: 'Auto Sensor Sheet', sampleId: identity.sampleId, rowCount: parsed.rows.length });
      }
    }

    // 2. Consolidate all unique discovered sample IDs in order of appearance
    const discoveredIdsSet = new Set([
      ...rawSamplingMap.keys(),
      ...rawFullMap.keys(),
      ...unprefixedMap.keys()
    ]);

    const allDiscoveredIds = Array.from(discoveredIdsSet);

    // Separate into "ข้อมูลตัวอย่าง (Samples)" and "ข้อมูลอ้างอิง (Blanks)"
    const nonBlankIds = allDiscoveredIds.filter(id => !isBlankSampleName(id)).sort(naturalCompare);
    const blankIds = allDiscoveredIds.filter(id => isBlankSampleName(id)).sort(naturalCompare);
    const orderedAllIds = [...nonBlankIds, ...blankIds];

    // 3. Populate state.datasets.Sampling and state.datasets.Raw with cross-fallback
    orderedAllIds.forEach(id => {
      let sampRows = rawSamplingMap.get(id) || [];
      let rawRows = rawFullMap.get(id) || [];
      const unprefRows = unprefixedMap.get(id) || [];

      if (sampRows.length === 0 && rawRows.length === 0 && unprefRows.length > 0) {
        rawRows = unprefRows;
        const onlySamp = unprefRows.filter(r => r.Mode === 'Sampling');
        sampRows = onlySamp.length > 0 ? onlySamp : [...unprefRows];
      } else if (sampRows.length === 0 && rawRows.length > 0) {
        const onlySamp = rawRows.filter(r => r.Mode === 'Sampling');
        sampRows = onlySamp.length > 0 ? onlySamp : [...rawRows];
      } else if (rawRows.length === 0 && sampRows.length > 0) {
        rawRows = [...sampRows];
      }

      state.datasets.Sampling[id] = sampRows;
      state.datasets.Raw[id] = rawRows;
    });

    // 4. Parse BlankAssignment & Date grouping dynamically
    const { blankMap, groupMap } = parseDynamicBlankAssignment(
      blankAssignWorksheet,
      orderedAllIds,
      blankIds
    );

    const defaultBlankFallback = blankIds[0] || null;
    orderedAllIds.forEach(id => {
      if (isBlankSampleName(id)) {
        state.blankAssignment[id] = id;
      } else if (blankMap[id]) {
        state.blankAssignment[id] = blankMap[id];
      } else {
        // Check if sample name or date shares a suffix with any blank (e.g. 12Oct / 14Oct)
        const dateHint = sampleDateMap.get(id) || groupMap[id] || '';
        const matchedBlank = blankIds.find(bId => {
          const cleanSuffix = bId.replace(/^blank[\s_\-]*/i, '').toLowerCase();
          return cleanSuffix && (id.toLowerCase().includes(cleanSuffix) || dateHint.toLowerCase().includes(cleanSuffix));
        });
        state.blankAssignment[id] = matchedBlank || defaultBlankFallback;
      }
    });

    // 5. Build state.sampleClasses with Golden-Angle Dynamic Color Palette + 3D Marker Shapes
    let sampleColorIdx = 0;
    let blankColorIdx = 0;
    let shapeIdx = 0;

    state.sampleClasses = orderedAllIds.map(id => {
      const isBlank = isBlankSampleName(id);
      const color = generateDynamicColor(isBlank ? blankColorIdx++ : sampleColorIdx++, isBlank);
      const shapeSpec = MARKER_SHAPES_3D[shapeIdx++ % MARKER_SHAPES_3D.length];
      const assignedBlank = state.blankAssignment[id] || null;
      const dateOrGroup = groupMap[id] || sampleDateMap.get(id) || (assignedBlank && !isBlank ? assignedBlank : null);

      const sampleObj = {
        id,
        label: id,
        isBlank,
        color,
        shapeSymbol: shapeSpec.symbol,
        shapeIcon: shapeSpec.icon,
        shapeName: shapeSpec.name,
        assignedBlank,
        groupLabel: isBlank
          ? 'Reference Blank'
          : (dateOrGroup ? `${dateOrGroup}` : 'Sample')
      };
      state.sampleMap[id] = sampleObj;
      return sampleObj;
    });

    state.activeSamples = new Set(orderedAllIds);

    // 6. Build Dynamic Group Filter buttons from BlankAssignment / Date / Blanks
    const groupsByKey = new Map();
    nonBlankIds.forEach(id => {
      const gKey = groupMap[id] || sampleDateMap.get(id) || state.blankAssignment[id];
      if (gKey) {
        if (!groupsByKey.has(gKey)) {
          groupsByKey.set(gKey, []);
        }
        groupsByKey.get(gKey).push(id);
      }
    });

    state.dynamicGroups = [];
    groupsByKey.forEach((sampleIds, gKey) => {
      if (sampleIds.length > 0 && sampleIds.length < orderedAllIds.length) {
        const cleanKey = String(gKey).replace(/^Blank[\s_\-]*/i, '');
        state.dynamicGroups.push({
          id: gKey,
          cleanKey,
          sampleIds
        });
      }
    });

    // 7. Finalize Dynamic Sensor Columns, Extra Columns, Cycle Loops, and Modes
    state.sensorCols = Array.from(sensorSet).sort(naturalCompare);
    state.activeBiplotSensors = new Set(state.sensorCols);
    state.extraNumericCols = Array.from(extraSet);

    // Ensure all rows have numeric values for all discovered sensors/extras
    for (const dsKey of ['Sampling', 'Raw']) {
      for (const id of orderedAllIds) {
        const rows = state.datasets[dsKey][id] || [];
        for (const r of rows) {
          for (const sCol of state.sensorCols) {
            if (r[sCol] === undefined) r[sCol] = 0;
          }
          for (const eCol of state.extraNumericCols) {
            if (r[eCol] === undefined) r[eCol] = 0;
          }
        }
      }
    }

    // Set default custom X, Y, Z axes dynamically (prefer G13, G14, G10 if present, else first 3 sensors)
    const pickSensor = (preferred, fallbackIdx) => {
      if (state.sensorCols.includes(preferred)) return preferred;
      return state.sensorCols[fallbackIdx % Math.max(1, state.sensorCols.length)] || '';
    };
    state.defaultCustomAxes = {
      x: pickSensor('G13', 0),
      y: pickSensor('G14', 1),
      z: pickSensor('G10', 2)
    };
    state.customAxes = { ...state.defaultCustomAxes };

    // Discover all Cycle Loops and Modes present in the file
    const cycleSet = new Set();
    const modeSet = new Set();
    for (const dsKey of ['Raw', 'Sampling']) {
      for (const id of orderedAllIds) {
        for (const r of state.datasets[dsKey][id] || []) {
          if (Number.isFinite(r['Cycle Loop'])) cycleSet.add(r['Cycle Loop']);
          if (r['Mode']) modeSet.add(r['Mode']);
        }
      }
    }

    state.availableCycles = Array.from(cycleSet).sort((a, b) => a - b);

    // Order Modes canonically if Resting/Sampling/Purging exist, plus any custom modes
    const canonicalModes = ['Resting', 'Sampling', 'Purging'];
    const discoveredModes = Array.from(modeSet);
    state.availableModes = [
      ...canonicalModes.filter(m => discoveredModes.includes(m)),
      ...discoveredModes.filter(m => !canonicalModes.includes(m))
    ];
    state.activeModes = new Set(state.availableModes);

    state.fileName = fileName;
    state.isDemoData = isDemo;
    state.sheetSummary = sheetSummary;

    // Automatically configure Interface Filter presets upon file upload/ingest:
    // - Section 1: Dataset Group = 'Raw' (Raw_*)
    // - Section 2: Axis Mode = 'pca' (3D PCA Mode), Granularity = 'cycle' (Cycle Summary),
    //              Fit Scope = 'active' (Active Re-fit PCA), Exclude Blanks = true
    // - Section 4: Active Cycles = Cycles 3–5 ('รอบที่ 3-5')
    state.activeDataset = 'Raw';
    state.axisMode = 'pca';
    state.pcaGranularity = 'cycle';
    state.pcaFitScope = 'active';
    state.excludeBlanksPca = true;

    const targetCycles3to5 = [3, 4, 5].filter(c => state.availableCycles.includes(c));
    if (targetCycles3to5.length > 0) {
      state.activeCycles = new Set(targetCycles3to5);
    } else if (state.availableCycles.length >= 3) {
      state.activeCycles = new Set(state.availableCycles.slice(2, 5));
    } else {
      state.activeCycles = new Set(state.availableCycles);
    }

    // Sync DOM elements in Interface Filter with automatic preset
    document.querySelectorAll('#datasetSegmented .seg-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.dataset === state.activeDataset);
    });
    document.querySelectorAll('#axisModeSegmented .seg-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.axismode === state.axisMode);
    });
    const customAxesPanelEl = document.getElementById('customAxesPanel');
    const pcaAxesPanelEl = document.getElementById('pcaAxesPanel');
    if (customAxesPanelEl) customAxesPanelEl.style.display = state.axisMode === 'custom' ? 'flex' : 'none';
    if (pcaAxesPanelEl) pcaAxesPanelEl.style.display = state.axisMode === 'pca' ? 'block' : 'none';

    document.querySelectorAll('#pcaGranularitySegmented .seg-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.granularity === state.pcaGranularity);
    });
    const pcaMethodRowEl = document.getElementById('pcaCycleSummaryMethodRow');
    if (pcaMethodRowEl) {
      pcaMethodRowEl.style.display = state.pcaGranularity === 'cycle' ? 'block' : 'none';
    }

    document.querySelectorAll('#pcaFitScopeSegmented .seg-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.fitscope === state.pcaFitScope);
    });
    const toggleExcludeBlanksEl = document.getElementById('toggleExcludeBlanksPca');
    if (toggleExcludeBlanksEl) {
      toggleExcludeBlanksEl.checked = state.excludeBlanksPca;
    }

    // 8. Update UI controls dynamically
    updateHeaderAndPipelineUI();
    populateAxisDropdowns();
    populateTimeSeriesControls();
    renderDynamicGroupButtons();
    renderSampleTogglePills();
    renderCyclePills();
    renderModePills();
    refreshDashboard(true);
  }

  // --------------------------------------------------------------------------
  // 4. BASELINE CORRECTION, DRIFT CORRECTION & ACTIVE DATA PIPELINE
  // --------------------------------------------------------------------------

  /**
   * Compute mean baseline vector for every discovered Blank in the active dataset,
   * supporting Overall, ByMode, ByCycle (Per-Cycle Blank Subtraction), and ByCycleAndMode.
   */
  function computeBlankBaselines(datasetKey) {
    const ds = state.datasets[datasetKey] || {};
    const baselines = {};
    const blankClasses = state.sampleClasses.filter(s => s.isBlank);

    blankClasses.forEach(blankObj => {
      const blankId = blankObj.id;
      const rows = ds[blankId] || [];
      baselines[blankId] = {
        overall: {},
        byMode: {},
        byCycle: {},
        byCycleAndMode: {}
      };

      state.sensorCols.forEach(col => {
        baselines[blankId].overall[col] = 0;
      });
      if (rows.length === 0) return;

      state.sensorCols.forEach(col => {
        let sum = 0;
        for (let i = 0; i < rows.length; i++) sum += rows[i][col];
        baselines[blankId].overall[col] = sum / rows.length;
      });

      state.availableModes.forEach(mode => {
        const modeRows = rows.filter(r => r.Mode === mode);
        if (modeRows.length > 0) {
          baselines[blankId].byMode[mode] = {};
          state.sensorCols.forEach(col => {
            let sum = 0;
            for (let i = 0; i < modeRows.length; i++) sum += modeRows[i][col];
            baselines[blankId].byMode[mode][col] = sum / modeRows.length;
          });
        }
      });

      // Group by Cycle Loop for Per-Cycle Blank Subtraction (Loop k - Blank Loop k)
      state.availableCycles.forEach(cycleNum => {
        const cycRows = rows.filter(r => r['Cycle Loop'] === cycleNum);
        if (cycRows.length > 0) {
          baselines[blankId].byCycle[cycleNum] = {};
          state.sensorCols.forEach(col => {
            let sum = 0;
            for (let i = 0; i < cycRows.length; i++) sum += cycRows[i][col];
            baselines[blankId].byCycle[cycleNum][col] = sum / cycRows.length;
          });

          baselines[blankId].byCycleAndMode[cycleNum] = {};
          state.availableModes.forEach(mode => {
            const cmRows = cycRows.filter(r => r.Mode === mode);
            if (cmRows.length > 0) {
              baselines[blankId].byCycleAndMode[cycleNum][mode] = {};
              state.sensorCols.forEach(col => {
                let sum = 0;
                for (let i = 0; i < cmRows.length; i++) sum += cmRows[i][col];
                baselines[blankId].byCycleAndMode[cycleNum][mode][col] = sum / cmRows.length;
              });
            }
          });
        }
      });
    });

    return baselines;
  }

  /**
   * Compute per-cycle Resting baseline from Raw_* dataset for each sample
   * so ΔMean (Samp - Rest) and ΔMax (Max Samp - Rest) can eliminate inter-loop baseline drift.
   */
  function computeRestingBaselinesForSample(sampleId, rawBlankRef) {
    const rawRows = (state.datasets.Raw && state.datasets.Raw[sampleId]) || [];
    const byCycleRest = {};
    if (rawRows.length === 0) return byCycleRest;

    const grouped = new Map();
    for (let i = 0; i < rawRows.length; i++) {
      const r = rawRows[i];
      if (r.Mode !== 'Resting') continue;
      const c = r['Cycle Loop'];
      if (!grouped.has(c)) grouped.set(c, []);
      grouped.get(c).push(r);
    }

    grouped.forEach((rList, cycleNum) => {
      if (rList.length === 0) return;
      byCycleRest[cycleNum] = {};
      const blankVec = (state.baselineCorrection && rawBlankRef)
        ? ((state.perCycleBlank &&
            rawBlankRef.byCycleAndMode &&
            rawBlankRef.byCycleAndMode[cycleNum] &&
            rawBlankRef.byCycleAndMode[cycleNum]['Resting']) ||
           (state.perCycleBlank && rawBlankRef.byCycle && rawBlankRef.byCycle[cycleNum]) ||
           (rawBlankRef.byMode && rawBlankRef.byMode['Resting']) ||
           rawBlankRef.overall)
        : null;

      for (const col of state.sensorCols) {
        let sum = 0;
        for (let i = 0; i < rList.length; i++) {
          const rawVal = rList[i][col] || 0;
          const bVal = blankVec ? (blankVec[col] || 0) : 0;
          sum += (rawVal - bVal);
        }
        byCycleRest[cycleNum][col] = sum / rList.length;
      }
    });

    return byCycleRest;
  }

  /**
   * Feature Extraction for Cycle Summary Mode (4 Methods):
   * - 'mean': Mean during Sampling
   * - 'max': Max during Sampling
   * - 'delta_mean': Mean(Sampling) - Mean(Resting of same Cycle Loop from Raw_*)
   * - 'delta_max': Max(Sampling) - Mean(Resting of same Cycle Loop from Raw_*)
   */
  function aggregateRowsByCycle(rows, sampleId, method = 'mean', restingByCycle = null) {
    if (!rows || rows.length === 0) return [];
    const byCycle = new Map();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const c = r['Cycle Loop'];
      if (!byCycle.has(c)) byCycle.set(c, []);
      byCycle.get(c).push(r);
    }

    const numericCols = [...state.sensorCols, ...state.extraNumericCols];
    const summaryRows = [];
    const methodLabelMap = {
      mean: 'Mean',
      max: 'Max',
      delta_mean: 'ΔMean(Samp−Rest)',
      delta_max: 'ΔMax(Max−Rest)'
    };

    Array.from(byCycle.keys())
      .sort((a, b) => a - b)
      .forEach(cycleNum => {
        const cRows = byCycle.get(cycleNum);
        const n = cRows.length;
        if (n === 0) return;

        // If ΔMean or ΔMax is selected and multiple modes exist in cRows, prefer Sampling rows for the response
        const sampOnlyRows = (method === 'delta_mean' || method === 'delta_max')
          ? cRows.filter(r => r.Mode === 'Sampling')
          : [];
        const targetRows = sampOnlyRows.length > 0 ? sampOnlyRows : cRows;
        const tn = targetRows.length;

        const restVec = (restingByCycle && restingByCycle[cycleNum]) ? restingByCycle[cycleNum] : null;
        const modesInCycle = Array.from(new Set(targetRows.map(r => r.Mode))).join('/');

        const summaryRecord = {
          sample: sampleId,
          rowIdx: cycleNum,
          'Cycle Loop': cycleNum,
          Mode: modesInCycle || 'Sampling',
          Time: `Cycle #${cycleNum} (${methodLabelMap[method] || method}, n=${tn})`
        };

        for (const col of numericCols) {
          const isSensor = state.sensorCols.includes(col);
          const restOffset = (isSensor && (method === 'delta_mean' || method === 'delta_max') && restVec)
            ? (restVec[col] || 0)
            : 0;

          if (method === 'max' || method === 'delta_max') {
            let maxVal = -Infinity;
            for (let i = 0; i < tn; i++) {
              const v = targetRows[i][col];
              if (Number.isFinite(v) && v > maxVal) maxVal = v;
            }
            const baseMax = Number.isFinite(maxVal) ? maxVal : 0;
            summaryRecord[col] = baseMax - restOffset;
          } else {
            let sumVal = 0;
            for (let i = 0; i < tn; i++) {
              sumVal += targetRows[i][col] || 0;
            }
            const baseMean = sumVal / tn;
            summaryRecord[col] = baseMean - restOffset;
          }
        }

        summaryRows.push(summaryRecord);
      });

    return summaryRows;
  }

  /**
   * Filter rows by active Cycles & Modes, apply Baseline Correction (with Per-Cycle Blank support),
   * and apply Cycle Summary aggregation (Mean / Max / ΔMean / ΔMax) when PCA Granularity Mode is 'cycle'.
   */
  function getProcessedDataByClass() {
    const datasetKey = state.activeDataset;
    const rawDs = state.datasets[datasetKey] || {};
    const baselines = state.baselineCorrection ? computeBlankBaselines(datasetKey) : null;
    const rawBaselines = (state.baselineCorrection && (state.pcaSummaryMethod === 'delta_mean' || state.pcaSummaryMethod === 'delta_max'))
      ? computeBlankBaselines('Raw')
      : null;
    const useCycleSummary = state.axisMode === 'pca' && state.pcaGranularity === 'cycle';

    const resultByClass = {};
    const allFilteredRows = [];
    const allGlobalRows = [];

    for (const s of state.sampleClasses) {
      const sampleId = s.id;
      const sourceRows = rawDs[sampleId] || [];
      const assignedBlankId = state.blankAssignment[sampleId];
      const blankRef = (baselines && assignedBlankId) ? baselines[assignedBlankId] : null;
      const rawBlankRef = (rawBaselines && assignedBlankId) ? rawBaselines[assignedBlankId] : null;

      const classFiltered = [];
      for (let i = 0; i < sourceRows.length; i++) {
        const row = sourceRows[i];
        const cyc = row['Cycle Loop'];

        if (!state.activeCycles.has(cyc)) continue;
        if (!state.activeModes.has(row['Mode'])) continue;

        if (!state.baselineCorrection || !blankRef) {
          classFiltered.push(row);
        } else {
          // Per-Cycle Blank Subtraction: prefer matching Cycle Loop + Mode, then matching Cycle Loop, then Mode, then Overall
          const blankBaseline = (
            state.perCycleBlank &&
            blankRef.byCycleAndMode &&
            blankRef.byCycleAndMode[cyc] &&
            blankRef.byCycleAndMode[cyc][row.Mode]
          ) || (
            state.perCycleBlank &&
            blankRef.byCycle &&
            blankRef.byCycle[cyc]
          ) || (
            blankRef.byMode &&
            blankRef.byMode[row.Mode]
          ) || blankRef.overall;

          const correctedRow = { ...row, _assignedBlank: assignedBlankId };
          for (const col of state.sensorCols) {
            correctedRow[col] = row[col] - (blankBaseline[col] || 0);
          }
          classFiltered.push(correctedRow);
        }
      }

      let restingByCycle = null;
      if (useCycleSummary && (state.pcaSummaryMethod === 'delta_mean' || state.pcaSummaryMethod === 'delta_max')) {
        restingByCycle = computeRestingBaselinesForSample(sampleId, rawBlankRef);
      }

      const finalClassRows = useCycleSummary
        ? aggregateRowsByCycle(classFiltered, sampleId, state.pcaSummaryMethod, restingByCycle)
        : classFiltered;

      resultByClass[sampleId] = finalClassRows;

      for (let i = 0; i < finalClassRows.length; i++) {
        allGlobalRows.push(finalClassRows[i]);
        if (state.activeSamples.has(sampleId)) {
          allFilteredRows.push(finalClassRows[i]);
        }
      }
    }

    return { resultByClass, allFilteredRows, allGlobalRows };
  }

  // --------------------------------------------------------------------------
  // 5. 3D PCA ENGINE (AUTO-CLEAN VARIANCE=0, Z-SCORE STANDARDIZATION & JACOBI EIGEN SOLVER)
  // --------------------------------------------------------------------------

  function jacobiEigenSymmetric(matrix, maxSweeps = 60) {
    const n = matrix.length;
    const A = matrix.map(row => [...row]);
    const V = Array.from({ length: n }, (_, i) =>
      Array.from({ length: n }, (__, j) => (i === j ? 1 : 0))
    );

    for (let sweep = 0; sweep < maxSweeps; sweep++) {
      let maxOffDiag = 0;
      for (let p = 0; p < n - 1; p++) {
        for (let q = p + 1; q < n; q++) {
          const absVal = Math.abs(A[p][q]);
          if (absVal > maxOffDiag) maxOffDiag = absVal;
        }
      }
      if (maxOffDiag < 1e-11) break;

      for (let p = 0; p < n - 1; p++) {
        for (let q = p + 1; q < n; q++) {
          const apq = A[p][q];
          if (Math.abs(apq) < 1e-13) continue;

          const app = A[p][p];
          const aqq = A[q][q];
          const tau = (aqq - app) / (2 * apq);
          const t = tau >= 0
            ? 1 / (tau + Math.sqrt(1 + tau * tau))
            : -1 / (-tau + Math.sqrt(1 + tau * tau));
          const c = 1 / Math.sqrt(1 + t * t);
          const s = t * c;

          A[p][p] = app - t * apq;
          A[q][q] = aqq + t * apq;
          A[p][q] = 0;
          A[q][p] = 0;

          for (let k = 0; k < n; k++) {
            if (k !== p && k !== q) {
              const akp = A[k][p];
              const akq = A[k][q];
              A[k][p] = c * akp - s * akq;
              A[p][k] = A[k][p];
              A[k][q] = s * akp + c * akq;
              A[q][k] = A[k][q];
            }
          }

          for (let k = 0; k < n; k++) {
            const vkp = V[k][p];
            const vkq = V[k][q];
            V[k][p] = c * vkp - s * vkq;
            V[k][q] = s * vkp + c * vkq;
          }
        }
      }
    }

    const eigenPairs = [];
    for (let i = 0; i < n; i++) {
      const vec = V.map(row => row[i]);
      let maxIdx = 0;
      for (let j = 1; j < n; j++) {
        if (Math.abs(vec[j]) > Math.abs(vec[maxIdx])) maxIdx = j;
      }
      if (vec[maxIdx] < 0) {
        for (let j = 0; j < n; j++) vec[j] = -vec[j];
      }
      eigenPairs.push({
        value: Math.max(0, A[i][i]),
        vector: vec
      });
    }

    eigenPairs.sort((a, b) => b.value - a.value);
    return eigenPairs;
  }

  /**
   * Compute Principal Component Analysis (PC1..PCn):
   * 1. Auto-Clean: Automatically detects and removes sensor columns with Variance = 0 (constant/zero columns).
   * 2. Z-Score Standardization: Z = (X - μ) / σ so all active sensors share equal unit variance.
   * 3. Eigenvalue Decomposition: Projects onto orthogonal principal components and returns full loadings & scree data.
   */
  function computePCA3D(rows, standardize = true) {
    const allSensorCols = state.sensorCols;
    const n = rows ? rows.length : 0;
    if (n < 3 || allSensorCols.length === 0) {
      return {
        explainedVariance: [0, 0, 0],
        cumulativeVariance: 0,
        cumulativeVariance2D: 0,
        allExplainedVariance: [],
        allCumulativeVariance: [],
        allEigenvalues: [],
        vectors: [[], [], []],
        allVectors: [],
        means: [],
        stds: [],
        topLoadings: [[], [], []],
        validSensors: [],
        droppedSensors: [...allSensorCols],
        projectRow: () => [0, 0, 0],
        getZScoreRow: () => []
      };
    }

    // Step 1: Compute mean (μ) and standard deviation (σ) for each sensor column
    // and Auto-Clean columns where Variance = 0 (σ <= 1e-9)
    const validSensors = [];
    const droppedSensors = [];
    const means = [];
    const stds = [];

    for (let j = 0; j < allSensorCols.length; j++) {
      const col = allSensorCols[j];
      let sum = 0;
      for (let i = 0; i < n; i++) {
        sum += rows[i][col] || 0;
      }
      const mean = sum / n;

      let sumSq = 0;
      for (let i = 0; i < n; i++) {
        const diff = (rows[i][col] || 0) - mean;
        sumSq += diff * diff;
      }
      const variance = sumSq / Math.max(1, n - 1);
      const sd = Math.sqrt(variance);

      if (sd <= 1e-9) {
        droppedSensors.push(col);
      } else {
        validSensors.push(col);
        means.push(mean);
        stds.push(standardize ? sd : 1);
      }
    }

    const p = validSensors.length;
    if (p === 0) {
      return {
        explainedVariance: [0, 0, 0],
        cumulativeVariance: 0,
        cumulativeVariance2D: 0,
        allExplainedVariance: [],
        allCumulativeVariance: [],
        allEigenvalues: [],
        vectors: [[], [], []],
        allVectors: [],
        means: [],
        stds: [],
        topLoadings: [[], [], []],
        validSensors: [],
        droppedSensors,
        projectRow: () => [0, 0, 0],
        getZScoreRow: () => []
      };
    }

    // Step 2: Build Covariance / Correlation Matrix on Z-score standardized valid sensors
    const cov = Array.from({ length: p }, () => new Array(p).fill(0));
    const centered = new Array(p);

    for (let i = 0; i < n; i++) {
      const r = rows[i];
      for (let j = 0; j < p; j++) {
        centered[j] = ((r[validSensors[j]] || 0) - means[j]) / stds[j];
      }
      for (let j = 0; j < p; j++) {
        const cj = centered[j];
        for (let k = j; k < p; k++) {
          cov[j][k] += cj * centered[k];
        }
      }
    }

    const denom = Math.max(1, n - 1);
    for (let j = 0; j < p; j++) {
      for (let k = j; k < p; k++) {
        cov[j][k] /= denom;
        cov[k][j] = cov[j][k];
      }
    }

    // Step 3: Solve Eigenvalues & Eigenvectors
    const eigenPairs = jacobiEigenSymmetric(cov);
    const totalVar = eigenPairs.reduce((acc, ep) => acc + ep.value, 0) || 1;

    const allEigenvalues = eigenPairs.map(ep => ep.value);
    const allExplainedVariance = eigenPairs.map(ep => (ep.value / totalVar) * 100);
    const allCumulativeVariance = [];
    let runningCum = 0;
    for (let i = 0; i < allExplainedVariance.length; i++) {
      runningCum += allExplainedVariance[i];
      allCumulativeVariance.push(runningCum);
    }

    const ev1 = allExplainedVariance[0] || 0;
    const ev2 = allExplainedVariance[1] || 0;
    const ev3 = allExplainedVariance[2] || 0;
    const explainedVariance = [ev1, ev2, ev3];
    const cumulativeVariance2D = ev1 + ev2;
    const cumulativeVariance = ev1 + ev2 + ev3;

    const zeroVec = new Array(p).fill(0);
    const v1 = eigenPairs[0] ? eigenPairs[0].vector : zeroVec;
    const v2 = eigenPairs[1] ? eigenPairs[1].vector : zeroVec;
    const v3 = eigenPairs[2] ? eigenPairs[2].vector : zeroVec;

    const getTopSensors = vec =>
      vec
        .map((weight, idx) => ({ sensor: validSensors[idx], weight }))
        .sort((a, b) => Math.abs(b.weight) - Math.abs(a.weight))
        .slice(0, 3)
        .map(item => `${item.sensor} (${item.weight >= 0 ? '+' : ''}${item.weight.toFixed(2)})`);

    const topLoadings = [getTopSensors(v1), getTopSensors(v2), getTopSensors(v3)];

    const projectRow = row => {
      let pc1 = 0, pc2 = 0, pc3 = 0;
      for (let j = 0; j < p; j++) {
        const z = ((row[validSensors[j]] || 0) - means[j]) / stds[j];
        pc1 += z * v1[j];
        pc2 += z * v2[j];
        pc3 += z * v3[j];
      }
      return [pc1, pc2, pc3];
    };

    const getZScoreRow = row => {
      const zArr = new Array(p);
      for (let j = 0; j < p; j++) {
        zArr[j] = ((row[validSensors[j]] || 0) - means[j]) / stds[j];
      }
      return zArr;
    };

    return {
      explainedVariance,
      cumulativeVariance,
      cumulativeVariance2D,
      allExplainedVariance,
      allCumulativeVariance,
      allEigenvalues,
      vectors: [v1, v2, v3],
      allVectors: eigenPairs.map(ep => ep.vector),
      means,
      stds,
      topLoadings,
      validSensors,
      droppedSensors,
      projectRow,
      getZScoreRow
    };
  }

  // --------------------------------------------------------------------------
  // 6. 3D AREA HULL ENGINE (3D CONVEX HULL & 3D CONFIDENCE ELLIPSOID FOR MESH3D)
  // --------------------------------------------------------------------------

  function computeConvexHullMesh3D(pts) {
    if (!pts || pts.length < 4) return null;

    let workPts = pts;
    if (pts.length > 90) {
      const selectedIndices = new Set();
      const numDirs = 64;
      const goldenRatio = (1 + Math.sqrt(5)) / 2;
      for (let d = 0; d < numDirs; d++) {
        const theta = (2 * Math.PI * d) / goldenRatio;
        const phi = Math.acos(1 - (2 * (d + 0.5)) / numDirs);
        const dx = Math.cos(theta) * Math.sin(phi);
        const dy = Math.sin(theta) * Math.sin(phi);
        const dz = Math.cos(phi);

        let minDot = Infinity, maxDot = -Infinity;
        let minIdx = 0, maxIdx = 0;
        for (let idx = 0; idx < pts.length; idx++) {
          const p = pts[idx];
          const dot = p[0] * dx + p[1] * dy + p[2] * dz;
          if (dot < minDot) { minDot = dot; minIdx = idx; }
          if (dot > maxDot) { maxDot = dot; maxIdx = idx; }
        }
        selectedIndices.add(minIdx);
        selectedIndices.add(maxIdx);
      }
      workPts = Array.from(selectedIndices).map(idx => pts[idx]);
    }

    const n = workPts.length;
    if (n < 4) return null;

    let minX = Infinity, maxX = -Infinity;
    let minY = Infinity, maxY = -Infinity;
    let minZ = Infinity, maxZ = -Infinity;
    for (const p of workPts) {
      if (p[0] < minX) minX = p[0]; if (p[0] > maxX) maxX = p[0];
      if (p[1] < minY) minY = p[1]; if (p[1] > maxY) maxY = p[1];
      if (p[2] < minZ) minZ = p[2]; if (p[2] > maxZ) maxZ = p[2];
    }
    const scale = Math.max(maxX - minX, maxY - minY, maxZ - minZ, 1e-4);
    const eps = scale * 1e-5;

    const V = workPts.map((p, idx) => [
      p[0] + eps * Math.sin(idx * 12.9898 + 1),
      p[1] + eps * Math.cos(idx * 78.233 + 2),
      p[2] + eps * Math.sin(idx * 37.719 + 3)
    ]);

    function faceNormal(a, b, c) {
      const ux = b[0] - a[0], uy = b[1] - a[1], uz = b[2] - a[2];
      const vx = c[0] - a[0], vy = c[1] - a[1], vz = c[2] - a[2];
      return [
        uy * vz - uz * vy,
        uz * vx - ux * vz,
        ux * vy - uy * vx
      ];
    }

    function signedDist(face, pt) {
      const a = V[face[0]];
      const nrm = face.normal;
      return nrm[0] * (pt[0] - a[0]) + nrm[1] * (pt[1] - a[1]) + nrm[2] * (pt[2] - a[2]);
    }

    function makeFace(i0, i1, i2, interiorPt) {
      let f = [i0, i1, i2];
      let nrm = faceNormal(V[i0], V[i1], V[i2]);
      const d = nrm[0] * (interiorPt[0] - V[i0][0]) +
                nrm[1] * (interiorPt[1] - V[i0][1]) +
                nrm[2] * (interiorPt[2] - V[i0][2]);
      if (d > 0) {
        f = [i0, i2, i1];
        nrm = [-nrm[0], -nrm[1], -nrm[2]];
      }
      f.normal = nrm;
      return f;
    }

    let i0 = 0, i1 = 1, i2 = 2, i3 = 3;
    const interior = [
      (V[i0][0] + V[i1][0] + V[i2][0] + V[i3][0]) / 4,
      (V[i0][1] + V[i1][1] + V[i2][1] + V[i3][1]) / 4,
      (V[i0][2] + V[i1][2] + V[i2][2] + V[i3][2]) / 4
    ];

    let faces = [
      makeFace(i0, i1, i2, V[i3]),
      makeFace(i0, i1, i3, V[i2]),
      makeFace(i0, i2, i3, V[i1]),
      makeFace(i1, i2, i3, V[i0])
    ];

    for (let idx = 4; idx < n; idx++) {
      const pt = V[idx];
      const visibleFaces = [];
      const keepFaces = [];

      for (let fIdx = 0; fIdx < faces.length; fIdx++) {
        const f = faces[fIdx];
        if (signedDist(f, pt) > 1e-11 * scale) {
          visibleFaces.push(f);
        } else {
          keepFaces.push(f);
        }
      }

      if (visibleFaces.length === 0) continue;

      const edgeCounts = new Map();
      const directedEdges = [];

      for (const f of visibleFaces) {
        const edges = [
          [f[0], f[1]],
          [f[1], f[2]],
          [f[2], f[0]]
        ];
        for (const [u, v] of edges) {
          const key = u < v ? `${u}_${v}` : `${v}_${u}`;
          edgeCounts.set(key, (edgeCounts.get(key) || 0) + 1);
          directedEdges.push([u, v, key]);
        }
      }

      for (const [u, v, key] of directedEdges) {
        if (edgeCounts.get(key) === 1) {
          keepFaces.push(makeFace(u, v, idx, interior));
        }
      }

      faces = keepFaces;
    }

    if (faces.length < 4) return computeEllipsoidMesh3D(pts);

    return {
      x: workPts.map(p => p[0]),
      y: workPts.map(p => p[1]),
      z: workPts.map(p => p[2]),
      i: faces.map(f => f[0]),
      j: faces.map(f => f[1]),
      k: faces.map(f => f[2])
    };
  }

  function computeEllipsoidMesh3D(pts, sdMultiplier = 2.15) {
    if (!pts || pts.length < 3) return null;
    const n = pts.length;
    let mx = 0, my = 0, mz = 0;
    for (const p of pts) {
      mx += p[0]; my += p[1]; mz += p[2];
    }
    mx /= n; my /= n; mz /= n;

    const cov = [
      [0, 0, 0],
      [0, 0, 0],
      [0, 0, 0]
    ];
    for (const p of pts) {
      const dx = p[0] - mx;
      const dy = p[1] - my;
      const dz = p[2] - mz;
      cov[0][0] += dx * dx; cov[0][1] += dx * dy; cov[0][2] += dx * dz;
      cov[1][1] += dy * dy; cov[1][2] += dy * dz;
      cov[2][2] += dz * dz;
    }
    const denom = Math.max(1, n - 1);
    cov[0][0] /= denom; cov[0][1] /= denom; cov[0][2] /= denom;
    cov[1][0] = cov[0][1]; cov[1][1] /= denom; cov[1][2] /= denom;
    cov[2][0] = cov[0][2]; cov[2][1] = cov[1][2]; cov[2][2] /= denom;

    const eig = jacobiEigenSymmetric(cov);
    const r0 = Math.max(1e-3, Math.sqrt(Math.max(1e-6, eig[0].value)) * sdMultiplier);
    const r1 = Math.max(1e-3, Math.sqrt(Math.max(1e-6, eig[1].value)) * sdMultiplier);
    const r2 = Math.max(1e-3, Math.sqrt(Math.max(1e-6, eig[2].value)) * sdMultiplier);

    const v0 = eig[0].vector;
    const v1 = eig[1].vector;
    const v2 = eig[2].vector;

    const latSteps = 12;
    const lonSteps = 18;
    const vx = [], vy = [], vz = [];

    for (let a = 0; a <= latSteps; a++) {
      const theta = (a * Math.PI) / latSteps;
      const sinT = Math.sin(theta);
      const cosT = Math.cos(theta);

      for (let b = 0; b < lonSteps; b++) {
        const phi = (b * 2 * Math.PI) / lonSteps;
        const ux = r0 * sinT * Math.cos(phi);
        const uy = r1 * sinT * Math.sin(phi);
        const uz = r2 * cosT;

        vx.push(mx + ux * v0[0] + uy * v1[0] + uz * v2[0]);
        vy.push(my + ux * v0[1] + uy * v1[1] + uz * v2[1]);
        vz.push(mz + ux * v0[2] + uy * v1[2] + uz * v2[2]);
      }
    }

    const fi = [], fj = [], fk = [];
    for (let a = 0; a < latSteps; a++) {
      for (let b = 0; b < lonSteps; b++) {
        const nextB = (b + 1) % lonSteps;
        const p1 = a * lonSteps + b;
        const p2 = a * lonSteps + nextB;
        const p3 = (a + 1) * lonSteps + b;
        const p4 = (a + 1) * lonSteps + nextB;

        fi.push(p1); fj.push(p2); fk.push(p3);
        fi.push(p2); fj.push(p4); fk.push(p3);
      }
    }

    return { x: vx, y: vy, z: vz, i: fi, j: fj, k: fk };
  }

  // --------------------------------------------------------------------------
  // 7. STATISTICAL HELPERS, KPI STRIP & MULTI-VIEW 8-TAB RENDERING ENGINE
  // --------------------------------------------------------------------------

  /**
   * Compute Count (n), Mean, SD, Median, Min, Max for an array of numbers
   */
  function computeDescriptiveStats(values) {
    if (!values || values.length === 0) {
      return { count: 0, mean: 0, sd: 0, median: 0, min: 0, max: 0 };
    }
    const n = values.length;
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;
    for (let i = 0; i < n; i++) {
      const v = Number(values[i]) || 0;
      sum += v;
      if (v < min) min = v;
      if (v > max) max = v;
    }
    const mean = sum / n;
    let sumSq = 0;
    for (let i = 0; i < n; i++) {
      const d = (Number(values[i]) || 0) - mean;
      sumSq += d * d;
    }
    const sd = n > 1 ? Math.sqrt(sumSq / (n - 1)) : 0;

    const sorted = [...values].map(v => Number(v) || 0).sort((a, b) => a - b);
    const mid = Math.floor(n / 2);
    const median = n % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;

    return {
      count: n,
      mean,
      sd,
      median,
      min: Number.isFinite(min) ? min : 0,
      max: Number.isFinite(max) ? max : 0
    };
  }

  /**
   * Compute 2D Confidence Ellipse (default 1.96 SD ~ 95%) for 2D PCA / Sensor projection
   */
  function computeEllipse2DPoints(xs, ys, sdMultiplier = 1.96, steps = 48) {
    const n = Math.min(xs ? xs.length : 0, ys ? ys.length : 0);
    if (n < 3) return null;

    let mx = 0, my = 0;
    for (let i = 0; i < n; i++) {
      mx += xs[i];
      my += ys[i];
    }
    mx /= n;
    my /= n;

    let sxx = 0, syy = 0, sxy = 0;
    for (let i = 0; i < n; i++) {
      const dx = xs[i] - mx;
      const dy = ys[i] - my;
      sxx += dx * dx;
      syy += dy * dy;
      sxy += dx * dy;
    }
    const denom = Math.max(1, n - 1);
    sxx /= denom;
    syy /= denom;
    sxy /= denom;

    const trace = sxx + syy;
    const det = sxx * syy - sxy * sxy;
    const disc = Math.sqrt(Math.max(0, (trace * trace) / 4 - det));
    const l1 = Math.max(1e-6, trace / 2 + disc);
    const l2 = Math.max(1e-6, trace / 2 - disc);

    const r1 = Math.sqrt(l1) * sdMultiplier;
    const r2 = Math.sqrt(l2) * sdMultiplier;
    const angle = Math.abs(sxy) > 1e-12 ? Math.atan2(l1 - sxx, sxy) : (sxx >= syy ? 0 : Math.PI / 2);
    const cosA = Math.cos(angle);
    const sinA = Math.sin(angle);

    const ex = [];
    const ey = [];
    for (let k = 0; k <= steps; k++) {
      const t = (k * 2 * Math.PI) / steps;
      const ux = r1 * Math.cos(t);
      const uy = r2 * Math.sin(t);
      ex.push(mx + ux * cosA - uy * sinA);
      ey.push(my + ux * sinA + uy * cosA);
    }
    return { x: ex, y: ey, cx: mx, cy: my };
  }

  /**
   * Update Part 1: Top 4-Card KPI Summary Strip
   */
  function updateKpiSummaryStrip(resultByClass, allFilteredRows, pca) {
    // Card 1: Protocol & Active Cycles
    const protoCodeEl = document.getElementById('kpiProtocolCode');
    const protoDescEl = document.getElementById('kpiProtocolDesc');
    const cycleBadgeEl = document.getElementById('kpiCycleBadge');

    if (protoCodeEl) protoCodeEl.textContent = state.protocolInfo.code || 'R10–S120–P60–F200';
    if (protoDescEl) {
      protoDescEl.textContent = state.protocolInfo.desc || 'Resting 10s • Sampling 120s • Purging 60s • Flow 200';
      protoDescEl.title = state.protocolInfo.rawText || '';
    }
    if (cycleBadgeEl) {
      const totalCyc = state.availableCycles.length;
      const actCyc = state.activeCycles.size;
      const warmupCut = totalCyc > 1 && !state.activeCycles.has(state.availableCycles[0]) && actCyc === totalCyc - 1;
      cycleBadgeEl.textContent = warmupCut
        ? `Loops: ${actCyc}/${totalCyc} (Warm-up Cut)`
        : `Loops: ${actCyc}/${totalCyc}`;
    }

    // Card 2: Active Sensor Array Status (Auto-Clean Var=0)
    const activeSensorRatioEl = document.getElementById('kpiActiveSensorRatio');
    const droppedBadgeEl = document.getElementById('kpiDroppedSensorsBadge');
    const droppedListEl = document.getElementById('kpiDroppedSensorsList');

    const totalSensors = state.sensorCols.length;
    const validSensors = pca && pca.validSensors ? pca.validSensors : state.sensorCols;
    const droppedSensors = pca && pca.droppedSensors ? pca.droppedSensors : [];

    if (activeSensorRatioEl) {
      activeSensorRatioEl.textContent = `${validSensors.length} / ${totalSensors} Active Sensors`;
    }
    if (droppedBadgeEl) {
      if (droppedSensors.length > 0) {
        droppedBadgeEl.textContent = `${droppedSensors.length} Excluded (Var=0)`;
        droppedBadgeEl.className = 'badge badge-amber';
      } else {
        droppedBadgeEl.textContent = '100% Active';
        droppedBadgeEl.className = 'badge badge-emerald';
      }
    }
    if (droppedListEl) {
      droppedListEl.textContent = droppedSensors.length > 0
        ? `ตัดออกอัตโนมัติ (Var=0): ${droppedSensors.join(', ')}`
        : `เซนเซอร์ทำงานครบทุกตัว (${validSensors[0] || 'G1'}–${validSensors[validSensors.length - 1] || 'Gn'})`;
      droppedListEl.title = droppedListEl.textContent;
    }

    // Card 3: PCA Discriminant Power (2D & 3D Cumulative Variance)
    const pca3dEl = document.getElementById('kpiPca3dVariance');
    const pca2dEl = document.getElementById('kpiPca2dVariance');
    const pcaBadgeEl = document.getElementById('kpiPcaPowerBadge');

    if (pca && pca.explainedVariance) {
      const [v1, v2, v3] = pca.explainedVariance;
      const cum2d = v1 + v2;
      const cum3d = v1 + v2 + v3;
      if (pca3dEl) pca3dEl.textContent = `3D (PC1–PC3): ${cum3d.toFixed(1)}%`;
      if (pca2dEl) pca2dEl.textContent = `2D (PC1+PC2): ${cum2d.toFixed(1)}% • PC1: ${v1.toFixed(1)}%`;
      if (pcaBadgeEl) {
        if (cum3d >= 85) {
          pcaBadgeEl.textContent = `✓ High Power (≥85%)`;
          pcaBadgeEl.className = 'badge badge-emerald';
        } else {
          pcaBadgeEl.textContent = `Moderate (${cum3d.toFixed(0)}%)`;
          pcaBadgeEl.className = 'badge badge-amber';
        }
      }
    }

    // Card 4: Environment & QC Stability (Temp & Flow Rate)
    const tempEl = document.getElementById('kpiTempSummary');
    const flowEl = document.getElementById('kpiFlowSummary');
    const qcBadgeEl = document.getElementById('kpiQcStatusBadge');

    const tempVals = [];
    const flowVals = [];
    for (let i = 0; i < allFilteredRows.length; i++) {
      const r = allFilteredRows[i];
      if (r['Temp'] !== undefined && Number.isFinite(r['Temp'])) tempVals.push(r['Temp']);
      if (r['Flow Rate'] !== undefined && Number.isFinite(r['Flow Rate'])) flowVals.push(r['Flow Rate']);
    }

    const tStats = computeDescriptiveStats(tempVals);
    const fStats = computeDescriptiveStats(flowVals);

    if (tempEl) {
      tempEl.textContent = tStats.count > 0
        ? `Temp: ${tStats.mean.toFixed(2)} ± ${tStats.sd.toFixed(2)} °C [${tStats.min.toFixed(1)}–${tStats.max.toFixed(1)}]`
        : 'Temp: N/A';
    }
    if (flowEl) {
      flowEl.textContent = fStats.count > 0
        ? `Flow: ${fStats.mean.toFixed(1)} ± ${fStats.sd.toFixed(1)} mL/min [${fStats.min.toFixed(0)}–${fStats.max.toFixed(0)}]`
        : 'Flow Rate: N/A';
    }
    if (qcBadgeEl) {
      const isStable = tStats.sd < 1.0 && fStats.sd < 5.0;
      qcBadgeEl.textContent = isStable ? '✓ QC Stable' : '⚠ Check Drift';
      qcBadgeEl.className = isStable ? 'badge badge-emerald' : 'badge badge-amber';
    }
  }

  /**
   * Master Dashboard Refresh (updates 3D/2D Score Plot, KPI Strip, Active Secondary Tab, and Bottom Summary Cards)
   */
  function refreshDashboard(preserveCamera = true) {
    const { resultByClass, allFilteredRows, allGlobalRows } = getProcessedDataByClass();

    // Update Sample Pill counts (n = ...) & active states in sidebar
    updateSampleTogglesUI(resultByClass);

    // Always compute PCA so KPI Strip, Biplot, Scree Plot, Heatmap, Distance Matrix & Excel Export are live
    const baseScopeRows = (state.pcaFitScope === 'active' && allFilteredRows.length >= 3)
      ? allFilteredRows
      : (allGlobalRows.length >= 3 ? allGlobalRows : allFilteredRows);

    const nonBlankScopeRows = state.excludeBlanksPca
      ? baseScopeRows.filter(r => {
          const sObj = state.sampleMap[r.sample];
          return sObj && !sObj.isBlank;
        })
      : baseScopeRows;

    const pcaBasisRows = nonBlankScopeRows.length >= 3 ? nonBlankScopeRows : baseScopeRows;
    const pca = computePCA3D(pcaBasisRows, state.pcaStandardize);
    state.pcaResult = pca;

    // Update PCA Sidebar Info Panel
    const [v1, v2, v3] = pca.explainedVariance;
    const cumVar = pca.cumulativeVariance;
    document.getElementById('pcaVar1').textContent = `${v1.toFixed(1)}%`;
    document.getElementById('pcaVar2').textContent = `${v2.toFixed(1)}%`;
    document.getElementById('pcaVar3').textContent = `${v3.toFixed(1)}%`;
    document.getElementById('pcaTotalVarianceBadge').textContent = `${cumVar.toFixed(1)}%`;

    const autoCleanEl = document.getElementById('pcaAutoCleanNote');
    if (autoCleanEl) {
      if (pca.droppedSensors && pca.droppedSensors.length > 0) {
        autoCleanEl.innerHTML =
          `<strong>Auto-Clean:</strong> ใช้ ${pca.validSensors.length}/${state.sensorCols.length} เซนเซอร์ ` +
          `<span style="color:#F59E0B">(ตัด Variance=0: ${pca.droppedSensors.join(', ')})</span>`;
      } else {
        autoCleanEl.innerHTML =
          `<strong>Auto-Clean:</strong> เซนเซอร์มีความแปรปรวนครบทั้ง ${pca.validSensors.length}/${state.sensorCols.length} คอลัมน์ (Variance &gt; 0)`;
      }
    }

    document.getElementById('pcaLoadingsNote').innerHTML =
      `<div><strong>PC1 Top:</strong> ${pca.topLoadings[0].join(', ') || '-'}</div>` +
      `<div><strong>PC2 Top:</strong> ${pca.topLoadings[1].join(', ') || '-'}</div>` +
      `<div><strong>PC3 Top:</strong> ${pca.topLoadings[2].join(', ') || '-'}</div>`;

    // Determine active X, Y, Z axes
    let axisTitleX = state.customAxes.x || 'X';
    let axisTitleY = state.customAxes.y || 'Y';
    let axisTitleZ = state.customAxes.z || 'Z';
    let getXYZ = row => [
      row[state.customAxes.x] || 0,
      row[state.customAxes.y] || 0,
      row[state.customAxes.z] || 0
    ];

    if (state.axisMode === 'pca') {
      axisTitleX = `PC1 (${v1.toFixed(1)}%)`;
      axisTitleY = `PC2 (${v2.toFixed(1)}%)`;
      axisTitleZ = `PC3 (${v3.toFixed(1)}%)`;
      getXYZ = row => pca.projectRow(row);
    } else if (state.baselineCorrection) {
      if (state.sensorCols.includes(axisTitleX)) axisTitleX = `${axisTitleX} (ΔBlank)`;
      if (state.sensorCols.includes(axisTitleY)) axisTitleY = `${axisTitleY} (ΔBlank)`;
      if (state.sensorCols.includes(axisTitleZ)) axisTitleZ = `${axisTitleZ} (ΔBlank)`;
    }

    // Update Top 4-Card KPI Summary Strip
    updateKpiSummaryStrip(resultByClass, allFilteredRows, pca);

    // Build Paired 3D Traces per Active Sample Class: (1) scatter3d + (2) mesh3d Area Hull
    const traces = [];
    const traceMeta = [];
    let totalVisiblePoints = 0;
    let maxCoordRadius = 1.0;

    for (const s of state.sampleClasses) {
      if (!state.activeSamples.has(s.id)) continue;
      const rows = resultByClass[s.id] || [];
      if (rows.length === 0) continue;

      totalVisiblePoints += rows.length;

      const xs = new Array(rows.length);
      const ys = new Array(rows.length);
      const zs = new Array(rows.length);
      const customdata = new Array(rows.length);
      let sumX = 0, sumY = 0, sumZ = 0;

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const [x, y, z] = getXYZ(r);
        xs[i] = x;
        ys[i] = y;
        zs[i] = z;
        sumX += x;
        sumY += y;
        sumZ += z;
        const rad = Math.max(Math.abs(x), Math.abs(y), Math.abs(z));
        if (rad > maxCoordRadius) maxCoordRadius = rad;

        customdata[i] = [
          s.label,
          r['Cycle Loop'],
          r['Mode'],
          r['Time'],
          s.groupLabel,
          s.shapeName || 'วงกลมทึบ (Circle)'
        ];
      }

      const isOpenShape = Boolean(s.shapeSymbol && s.shapeSymbol.endsWith('-open'));

      // Layer 1: 3D Scatter Points (scatter3d) with Distinct Simple 3D Shape per Sample Class
      traces.push({
        type: 'scatter3d',
        mode: 'markers',
        name: s.label,
        legendgroup: s.id,
        showlegend: false,
        x: xs,
        y: ys,
        z: zs,
        customdata: customdata,
        marker: {
          size: state.markerSize,
          color: s.color,
          opacity: 0.92,
          symbol: s.shapeSymbol || 'circle',
          line: isOpenShape
            ? { width: 2.2, color: s.color }
            : { width: 0.6, color: 'rgba(11, 15, 25, 0.75)' }
        },
        hovertemplate:
          `<b>ตัวอย่าง (Sample): ${s.shapeIcon || '●'} %{customdata[0]}</b> <span style="color:#94A3B8">(%{customdata[4]})</span><br>` +
          `รูปทรง (Shape): <b>%{customdata[5]}</b><br>` +
          `รอบการวัด (Cycle Loop): <b>#%{customdata[1]}</b><br>` +
          `โหมด (Mode): <b>%{customdata[2]}</b><br>` +
          `เวลา (Time): <b>%{customdata[3]}</b><br>` +
          `<span style="color:#38BDF8">X (${axisTitleX}):</span> <b>%{x:.4f}</b><br>` +
          `<span style="color:#34D399">Y (${axisTitleY}):</span> <b>%{y:.4f}</b><br>` +
          `<span style="color:#F472B6">Z (${axisTitleZ}):</span> <b>%{z:.4f}</b>` +
          `<extra></extra>`
      });
      traceMeta.push({ sampleId: s.id, traceType: 'scatter', color: s.color });

      // Optional Group Centroid Marker & Label in 3D Space
      if (state.showCentroids && rows.length >= 2) {
        const cx = sumX / rows.length;
        const cy = sumY / rows.length;
        const cz = sumZ / rows.length;
        traces.push({
          type: 'scatter3d',
          mode: 'markers+text',
          name: `${s.label} Centroid`,
          legendgroup: s.id,
          showlegend: false,
          x: [cx],
          y: [cy],
          z: [cz],
          text: [s.label],
          textposition: 'top center',
          textfont: {
            family: "'JetBrains Mono', 'Noto Sans Thai', monospace",
            size: 10,
            color: s.color
          },
          marker: {
            size: Math.max(6, state.markerSize + 2),
            color: s.color,
            opacity: 1,
            symbol: 'diamond',
            line: { width: 1.5, color: '#FFFFFF' }
          },
          hovertemplate:
            `<b>🎯 Centroid: ${escapeHtml(s.label)}</b> (n=${rows.length})<br>` +
            `Mean X: <b>%{x:.4f}</b><br>Mean Y: <b>%{y:.4f}</b><br>Mean Z: <b>%{z:.4f}</b><extra></extra>`
        });
        traceMeta.push({ sampleId: s.id, traceType: 'centroid', color: s.color });
      }

      // Layer 2: Paired 3D Area Highlight (mesh3d Convex Hull / Confidence Ellipsoid)
      if (state.showAreaHull && rows.length >= 3) {
        const pts3D = xs.map((x, idx) => [x, ys[idx], zs[idx]]);
        const meshGeom = (state.hullType === 'ellipsoid' || rows.length < 4)
          ? computeEllipsoidMesh3D(pts3D)
          : computeConvexHullMesh3D(pts3D);

        if (meshGeom) {
          traces.push({
            type: 'mesh3d',
            name: `${s.label} Area Hull`,
            legendgroup: s.id,
            showlegend: false,
            x: meshGeom.x,
            y: meshGeom.y,
            z: meshGeom.z,
            i: meshGeom.i,
            j: meshGeom.j,
            k: meshGeom.k,
            color: s.color,
            opacity: state.hullOpacity,
            flatshading: state.hullType === 'convex',
            lighting: {
              ambient: 0.78,
              diffuse: 0.65,
              specular: 0.25,
              roughness: 0.5,
              fresnel: 0.2
            },
            hoverinfo: 'skip'
          });
          traceMeta.push({ sampleId: s.id, traceType: 'hull', color: s.color });
        }
      }
    }

    // Optional 3D Biplot Loading Vectors Overlay inside 3D Plot
    if (state.show3dBiplot && pca && pca.validSensors.length > 0) {
      const [v1Vec, v2Vec, v3Vec] = pca.vectors;
      let maxLoad = 0;
      for (let j = 0; j < pca.validSensors.length; j++) {
        const mag = Math.sqrt(v1Vec[j] * v1Vec[j] + v2Vec[j] * v2Vec[j] + v3Vec[j] * v3Vec[j]);
        if (mag > maxLoad) maxLoad = mag;
      }
      const arrowScale = (maxCoordRadius * 0.85) / (maxLoad || 1);

      const lineX = [];
      const lineY = [];
      const lineZ = [];
      const tipX = [];
      const tipY = [];
      const tipZ = [];
      const tipLabels = [];

      for (let j = 0; j < pca.validSensors.length; j++) {
        const tx = v1Vec[j] * arrowScale;
        const ty = v2Vec[j] * arrowScale;
        const tz = v3Vec[j] * arrowScale;
        lineX.push(0, tx, null);
        lineY.push(0, ty, null);
        lineZ.push(0, tz, null);
        tipX.push(tx);
        tipY.push(ty);
        tipZ.push(tz);
        tipLabels.push(pca.validSensors[j]);
      }

      traces.push({
        type: 'scatter3d',
        mode: 'lines',
        name: '3D Sensor Loadings',
        showlegend: false,
        x: lineX,
        y: lineY,
        z: lineZ,
        line: { width: 3.5, color: 'rgba(251, 191, 36, 0.85)' },
        hoverinfo: 'skip'
      });
      traceMeta.push({ sampleId: '__biplot__', traceType: 'biplot', color: '#FBBF24' });

      traces.push({
        type: 'scatter3d',
        mode: 'markers+text',
        name: 'Sensor Loading Tips',
        showlegend: false,
        x: tipX,
        y: tipY,
        z: tipZ,
        text: tipLabels,
        textposition: 'top center',
        textfont: { family: "'JetBrains Mono', monospace", size: 11, color: '#FDE68A' },
        marker: { size: 4.5, color: '#F59E0B', symbol: 'diamond' },
        hovertemplate: `<b>🧭 Sensor Vector: %{text}</b><br>X: %{x:.3f}<br>Y: %{y:.3f}<br>Z: %{z:.3f}<extra></extra>`
      });
      traceMeta.push({ sampleId: '__biplot__', traceType: 'biplot', color: '#F59E0B' });
    }

    state.traceMeta = traceMeta;

    // Toggle visibility between 3D Plot and 2D Projection inside Tab 1
    const plot3dEl = document.getElementById('plot3d');
    const plot2dEl = document.getElementById('plot2dProjection');
    const floating3dToolbar = document.getElementById('floating3dToolbar');
    const rotVertWrap = document.getElementById('rotVerticalBarWrap');
    const rotHorizWrap = document.getElementById('rotHorizontalBarWrap');

    if (state.scoreDimension === '2d') {
      if (plot3dEl) plot3dEl.style.display = 'none';
      if (plot2dEl) plot2dEl.style.display = 'block';
      if (floating3dToolbar) floating3dToolbar.style.display = 'none';
      if (rotVertWrap) rotVertWrap.style.display = 'none';
      if (rotHorizWrap) rotHorizWrap.style.display = 'none';
      render2DScoreProjection(resultByClass, getXYZ, axisTitleX, axisTitleY, pca);
    } else {
      if (plot3dEl) plot3dEl.style.display = 'block';
      if (plot2dEl) plot2dEl.style.display = 'none';
      if (floating3dToolbar) floating3dToolbar.style.display = 'flex';
      if (rotVertWrap) rotVertWrap.style.display = 'flex';
      if (rotHorizWrap) rotHorizWrap.style.display = 'flex';

      // Lock current 3D camera state before updating Plotly
      if (preserveCamera && plot3dEl && plot3dEl._fullLayout && plot3dEl._fullLayout.scene && plot3dEl._fullLayout.scene._scene) {
        try {
          const liveCam = plot3dEl._fullLayout.scene._scene.getCamera();
          if (liveCam) state.currentCamera = liveCam;
        } catch (_) {}
      }

      const layout = {
        uirevision: 'constant',
        paper_bgcolor: '#0B0F19',
        plot_bgcolor: '#0B0F19',
        margin: { l: 0, r: 0, t: 0, b: 0 },
        font: {
          family: "'Plus Jakarta Sans', 'Noto Sans Thai', sans-serif",
          color: '#F8FAFC'
        },
        hoverlabel: {
          bgcolor: 'rgba(15, 23, 42, 0.96)',
          bordercolor: 'rgba(56, 189, 248, 0.55)',
          font: {
            family: "'JetBrains Mono', 'Noto Sans Thai', monospace",
            size: 12,
            color: '#F8FAFC'
          }
        },
        scene: {
          uirevision: 'constant',
          bgcolor: '#0B0F19',
          dragmode: 'orbit',
          camera: state.currentCamera,
          aspectmode: 'cube',
          xaxis: {
            title: { text: `X: ${axisTitleX}`, font: { size: 12, color: '#38BDF8' } },
            gridcolor: 'rgba(148, 163, 184, 0.14)',
            zerolinecolor: 'rgba(56, 189, 248, 0.35)',
            showbackground: true,
            backgroundcolor: 'rgba(17, 24, 39, 0.55)',
            tickfont: { family: "'JetBrains Mono', monospace", size: 10, color: '#94A3B8' }
          },
          yaxis: {
            title: { text: `Y: ${axisTitleY}`, font: { size: 12, color: '#34D399' } },
            gridcolor: 'rgba(148, 163, 184, 0.14)',
            zerolinecolor: 'rgba(52, 211, 153, 0.35)',
            showbackground: true,
            backgroundcolor: 'rgba(17, 24, 39, 0.55)',
            tickfont: { family: "'JetBrains Mono', monospace", size: 10, color: '#94A3B8' }
          },
          zaxis: {
            title: { text: `Z: ${axisTitleZ}`, font: { size: 12, color: '#F472B6' } },
            gridcolor: 'rgba(148, 163, 184, 0.14)',
            zerolinecolor: 'rgba(244, 114, 182, 0.35)',
            showbackground: true,
            backgroundcolor: 'rgba(17, 24, 39, 0.55)',
            tickfont: { family: "'JetBrains Mono', monospace", size: 10, color: '#94A3B8' }
          }
        }
      };

      const config = {
        responsive: true,
        displaylogo: false,
        scrollZoom: true,
        modeBarButtonsToRemove: ['resetCameraLastSave3d']
      };

      Plotly.react(plot3dEl, traces, layout, config);
      syncSlidersFromCamera(state.currentCamera);

      if (!plot3dEl._hasRelayoutSync) {
        plot3dEl._hasRelayoutSync = true;
        plot3dEl.on('plotly_relayout', ev => {
          if (ev && ev['scene.camera']) {
            state.currentCamera = ev['scene.camera'];
            syncSlidersFromCamera(state.currentCamera);
          }
        });
      }
    }

    // Render active secondary tab if user is viewing Tabs 2–8
    renderActiveSecondaryTab(resultByClass, pca, getXYZ, axisTitleX, axisTitleY, axisTitleZ);

    updateTopStatusHeader(axisTitleX, axisTitleY, axisTitleZ, totalVisiblePoints);
    renderBottomSummaryCards(resultByClass, getXYZ, axisTitleX, axisTitleY, axisTitleZ, totalVisiblePoints);
  }

  /**
   * Tab 1 Sub-Mode: 2D Score Projection (PC1 vs PC2 or X vs Y) with 95% Confidence Ellipse & Centroid Labels
   */
  function render2DScoreProjection(resultByClass, getXYZ, axisTitleX, axisTitleY, pca) {
    const plot2dEl = document.getElementById('plot2dProjection');
    if (!plot2dEl) return;

    const traces = [];
    for (const s of state.sampleClasses) {
      if (!state.activeSamples.has(s.id)) continue;
      const rows = resultByClass[s.id] || [];
      if (rows.length === 0) continue;

      const xs = [];
      const ys = [];
      const custom = [];
      for (let i = 0; i < rows.length; i++) {
        const [x, y] = getXYZ(rows[i]);
        xs.push(x);
        ys.push(y);
        custom.push([s.label, rows[i]['Cycle Loop'], rows[i]['Mode'], rows[i]['Time']]);
      }

      // 2D Confidence Ellipse (95%)
      if (state.showAreaHull && rows.length >= 3) {
        const ell = computeEllipse2DPoints(xs, ys, 1.96, 48);
        if (ell) {
          traces.push({
            type: 'scatter',
            mode: 'lines',
            name: `${s.label} (95% Ellipse)`,
            legendgroup: s.id,
            showlegend: false,
            x: ell.x,
            y: ell.y,
            fill: 'toself',
            fillcolor: `${s.color}28`,
            line: { color: s.color, width: 1.5, dash: 'dot' },
            hoverinfo: 'skip'
          });
        }
      }

      // 2D Scatter Points
      traces.push({
        type: 'scatter',
        mode: 'markers',
        name: s.label,
        legendgroup: s.id,
        x: xs,
        y: ys,
        customdata: custom,
        marker: {
          size: Math.max(6, state.markerSize + 2),
          color: s.color,
          opacity: 0.9,
          line: { width: 0.8, color: '#0B0F19' }
        },
        hovertemplate:
          `<b>${s.shapeIcon || '●'} %{customdata[0]}</b> (Loop #%{customdata[1]} • %{customdata[2]})<br>` +
          `X (${axisTitleX}): <b>%{x:.4f}</b><br>` +
          `Y (${axisTitleY}): <b>%{y:.4f}</b><extra></extra>`
      });

      // 2D Centroid Label
      if (state.showCentroids && rows.length >= 1) {
        const cx = xs.reduce((a, b) => a + b, 0) / xs.length;
        const cy = ys.reduce((a, b) => a + b, 0) / ys.length;
        traces.push({
          type: 'scatter',
          mode: 'markers+text',
          name: `${s.label} Centroid`,
          legendgroup: s.id,
          showlegend: false,
          x: [cx],
          y: [cy],
          text: [s.label],
          textposition: 'top center',
          textfont: { family: "'JetBrains Mono', 'Noto Sans Thai', monospace", size: 11, color: s.color },
          marker: { size: 10, color: s.color, symbol: 'diamond', line: { width: 2, color: '#FFFFFF' } },
          hovertemplate: `<b>🎯 Centroid: ${escapeHtml(s.label)}</b><br>X: %{x:.4f}<br>Y: %{y:.4f}<extra></extra>`
        });
      }
    }

    const layout = {
      paper_bgcolor: '#0B0F19',
      plot_bgcolor: '#0B0F19',
      margin: { l: 64, r: 28, t: 56, b: 54 },
      font: { family: "'Plus Jakarta Sans', 'Noto Sans Thai', sans-serif", color: '#F8FAFC' },
      legend: { orientation: 'h', y: 1.08, x: 0.5, xanchor: 'center', font: { size: 11 } },
      xaxis: {
        title: { text: `X: ${axisTitleX}`, font: { size: 12, color: '#38BDF8' } },
        gridcolor: 'rgba(148, 163, 184, 0.14)',
        zerolinecolor: 'rgba(56, 189, 248, 0.35)'
      },
      yaxis: {
        title: { text: `Y: ${axisTitleY}`, font: { size: 12, color: '#34D399' } },
        gridcolor: 'rgba(148, 163, 184, 0.14)',
        zerolinecolor: 'rgba(52, 211, 153, 0.35)'
      }
    };

    Plotly.react(plot2dEl, traces, layout, { responsive: true, displaylogo: false });
  }

  /**
   * Dispatch rendering for the active secondary tab (Tabs 2–7)
   */
  function renderActiveSecondaryTab(resultByClass, pca, getXYZ, axisTitleX, axisTitleY, axisTitleZ) {
    switch (state.activeViewTab) {
      case 'biplot':
        renderBiplotTab(resultByClass, pca);
        break;
      case 'scree':
        renderScreeTab(pca);
        break;
      case 'radar':
        renderRadarTab(resultByClass, pca);
        break;
      case 'heatmap':
        renderHeatmapTab(resultByClass, pca);
        break;
      case 'timeseries':
        renderTimeSeriesTab();
        break;
      case 'qc':
        if (state.qcSubmode === 'distance') {
          renderDistanceTab(resultByClass, pca, getXYZ, axisTitleX, axisTitleY, axisTitleZ);
        } else {
          renderQcTab(resultByClass);
        }
        break;
      case 'distance':
        renderDistanceTab(resultByClass, pca, getXYZ, axisTitleX, axisTitleY, axisTitleZ);
        break;
      default:
        break;
    }
  }

  /**
   * Render individual sensor vector toggle pills (G1–Gn) in Tab 2 PCA Biplot toolbar
   */
  function renderBiplotSensorPills(validSensors = [], droppedSensors = []) {
    const container = document.getElementById('biplotSensorPillsContainer');
    const masterCheckbox = document.getElementById('toggleBiplotAllVectors');
    if (masterCheckbox) {
      masterCheckbox.checked = Boolean(state.showBiplotVectors);
    }
    if (!container) return;

    container.innerHTML = '';
    state.sensorCols.forEach(sCol => {
      const isExcluded = droppedSensors.includes(sCol) || (validSensors.length > 0 && !validSensors.includes(sCol));
      const isChecked = state.activeBiplotSensors.has(sCol);
      const isEffectivelyActive = Boolean(state.showBiplotVectors) && isChecked && !isExcluded;

      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = `biplot-sensor-pill ${isEffectivelyActive ? 'active' : 'inactive'} ${isExcluded ? 'excluded-sensor' : ''}`;
      btn.dataset.sensor = sCol;
      btn.title = isExcluded
        ? `${sCol}: ความแปรปรวนเป็นศูนย์ (Excluded Variance = 0)`
        : `คลิกเพื่อเปิด/ปิดเส้นเวกเตอร์ของเซนเซอร์ ${sCol} บนกราฟ Biplot`;
      btn.innerHTML = `<span class="sensor-dot"></span><span>${escapeHtml(sCol)}</span>`;

      if (!isExcluded) {
        btn.addEventListener('click', () => {
          if (!state.showBiplotVectors) {
            state.showBiplotVectors = true;
            state.activeBiplotSensors.add(sCol);
            if (masterCheckbox) masterCheckbox.checked = true;
          } else if (state.activeBiplotSensors.has(sCol)) {
            state.activeBiplotSensors.delete(sCol);
          } else {
            state.activeBiplotSensors.add(sCol);
          }
          refreshDashboard(true);
        });
      }

      container.appendChild(btn);
    });
  }

  /**
   * TAB 2: PCA Biplot & Sensor Loading Vectors (G1–G15 Correlation Circle + Loadings Matrix Table)
   */
  function renderBiplotTab(resultByClass, pca) {
    const plotEl = document.getElementById('plotBiplot2d');
    const tableWrap = document.getElementById('biplotLoadingsTableWrap');
    const badgeEl = document.getElementById('biplotActiveSensorBadge');
    if (!plotEl || !pca) return;

    const [v1Pct, v2Pct] = pca.explainedVariance;
    const [vec1, vec2, vec3] = pca.vectors;
    const validSensors = pca.validSensors || [];
    const droppedSensors = pca.droppedSensors || [];

    renderBiplotSensorPills(validSensors, droppedSensors);

    const visibleSensorCount = state.showBiplotVectors
      ? validSensors.filter(s => state.activeBiplotSensors.has(s)).length
      : 0;

    if (badgeEl) {
      badgeEl.textContent = `${visibleSensorCount} / ${validSensors.length} Vectors`;
    }

    const traces = [];
    let maxScore = 1.0;

    // 1. Sample Score Points & Centroids in PC1 vs PC2 space
    for (const s of state.sampleClasses) {
      if (!state.activeSamples.has(s.id)) continue;
      const rows = resultByClass[s.id] || [];
      if (rows.length === 0) continue;

      const pc1s = [];
      const pc2s = [];
      for (let i = 0; i < rows.length; i++) {
        const [p1, p2] = pca.projectRow(rows[i]);
        pc1s.push(p1);
        pc2s.push(p2);
        const m = Math.max(Math.abs(p1), Math.abs(p2));
        if (m > maxScore) maxScore = m;
      }

      traces.push({
        type: 'scatter',
        mode: 'markers',
        name: s.label,
        legendgroup: s.id,
        x: pc1s,
        y: pc2s,
        marker: { size: 6.5, color: s.color, opacity: 0.55 },
        hovertemplate: `<b>${escapeHtml(s.label)}</b><br>PC1: %{x:.3f}<br>PC2: %{y:.3f}<extra></extra>`
      });

      const cx = pc1s.reduce((a, b) => a + b, 0) / pc1s.length;
      const cy = pc2s.reduce((a, b) => a + b, 0) / pc2s.length;
      traces.push({
        type: 'scatter',
        mode: 'markers+text',
        name: `${s.label} Centroid`,
        legendgroup: s.id,
        showlegend: false,
        x: [cx],
        y: [cy],
        text: [s.label],
        textposition: 'top center',
        textfont: { family: "'JetBrains Mono', monospace", size: 11, color: s.color },
        marker: { size: 10, color: s.color, symbol: 'diamond', line: { width: 1.8, color: '#FFFFFF' } },
        hovertemplate: `<b>🎯 Centroid: ${escapeHtml(s.label)}</b><br>PC1: %{x:.3f}<br>PC2: %{y:.3f}<extra></extra>`
      });
    }

    // 2. Unit Correlation Circle & Sensor Loading Vectors
    let maxLoading = 0;
    for (let j = 0; j < validSensors.length; j++) {
      const m = Math.sqrt((vec1[j] || 0) ** 2 + (vec2[j] || 0) ** 2);
      if (m > maxLoading) maxLoading = m;
    }
    const scaleFactor = (maxScore * 0.85) / (maxLoading || 1);

    // Correlation Circle
    const circleX = [];
    const circleY = [];
    const circleRad = maxLoading * scaleFactor;
    for (let k = 0; k <= 64; k++) {
      const t = (k * 2 * Math.PI) / 64;
      circleX.push(circleRad * Math.cos(t));
      circleY.push(circleRad * Math.sin(t));
    }
    traces.push({
      type: 'scatter',
      mode: 'lines',
      name: 'Correlation Circle',
      x: circleX,
      y: circleY,
      line: { color: 'rgba(148, 163, 184, 0.28)', width: 1.2, dash: 'dash' },
      hoverinfo: 'skip'
    });

    // Loading Vectors & Annotations (filtered by state.showBiplotVectors & state.activeBiplotSensors)
    const annotations = [];
    const tipX = [];
    const tipY = [];
    const tipNames = [];

    if (state.showBiplotVectors) {
      for (let j = 0; j < validSensors.length; j++) {
        const sName = validSensors[j];
        if (!state.activeBiplotSensors.has(sName)) continue;

        const lx = (vec1[j] || 0) * scaleFactor;
        const ly = (vec2[j] || 0) * scaleFactor;
        tipX.push(lx);
        tipY.push(ly);
        tipNames.push(sName);

        annotations.push({
          x: lx,
          y: ly,
          ax: 0,
          ay: 0,
          xref: 'x',
          yref: 'y',
          axref: 'x',
          ayref: 'y',
          showarrow: true,
          arrowhead: 2,
          arrowsize: 1.1,
          arrowwidth: 2,
          arrowcolor: '#FBBF24'
        });
      }
    }

    if (tipX.length > 0) {
      traces.push({
        type: 'scatter',
        mode: 'markers+text',
        name: 'Sensor Loadings (G1–Gn)',
        x: tipX,
        y: tipY,
        text: tipNames,
        textposition: 'top center',
        textfont: { family: "'JetBrains Mono', monospace", size: 11.5, color: '#FDE68A' },
        marker: { size: 7, color: '#F59E0B', symbol: 'diamond' },
        hovertemplate: `<b>🧭 เซนเซอร์ %{text}</b><br>Scaled X: %{x:.3f}<br>Scaled Y: %{y:.3f}<extra></extra>`
      });
    }

    const layout = {
      title: {
        text: `🧭 2D PCA Biplot — Score Centroids & Sensor Loading Vectors (PC1 ${v1Pct.toFixed(1)}% vs PC2 ${v2Pct.toFixed(1)}%)`,
        font: { size: 13.5, color: '#F8FAFC' }
      },
      paper_bgcolor: '#0B0F19',
      plot_bgcolor: '#0B0F19',
      margin: { l: 60, r: 28, t: 50, b: 50 },
      font: { family: "'Plus Jakarta Sans', 'Noto Sans Thai', sans-serif", color: '#F8FAFC' },
      legend: { orientation: 'h', y: -0.14, x: 0.5, xanchor: 'center', font: { size: 10.5 } },
      annotations,
      xaxis: {
        title: { text: `PC1 (${v1Pct.toFixed(1)}%)`, font: { size: 12, color: '#38BDF8' } },
        gridcolor: 'rgba(148, 163, 184, 0.14)',
        zerolinecolor: 'rgba(56, 189, 248, 0.4)'
      },
      yaxis: {
        title: { text: `PC2 (${v2Pct.toFixed(1)}%)`, font: { size: 12, color: '#34D399' } },
        gridcolor: 'rgba(148, 163, 184, 0.14)',
        zerolinecolor: 'rgba(52, 211, 153, 0.4)'
      }
    };

    Plotly.react(plotEl, traces, layout, { responsive: true, displaylogo: false });

    // Populate Right-Side Sensor Loadings Table
    if (tableWrap) {
      const rowsHtml = state.sensorCols
        .map(col => {
          const idx = validSensors.indexOf(col);
          if (idx === -1 || droppedSensors.includes(col)) {
            return `
              <tr style="opacity:0.55;">
                <td><strong>${escapeHtml(col)}</strong></td>
                <td colspan="3" style="color:#F59E0B;">Excluded (Variance = 0)</td>
                <td><span class="badge badge-amber">Var=0</span></td>
              </tr>
            `;
          }
          const isVectorVisible = Boolean(state.showBiplotVectors) && state.activeBiplotSensors.has(col);
          const w1 = vec1[idx] || 0;
          const w2 = vec2[idx] || 0;
          const w3 = vec3[idx] || 0;
          const power2d = Math.sqrt(w1 * w1 + w2 * w2);
          return `
            <tr data-biplotsensor="${escapeHtml(col)}" style="cursor:pointer; opacity:${isVectorVisible ? '1' : '0.45'};" title="คลิกเพื่อเปิด/ปิดเส้นเวกเตอร์ ${escapeHtml(col)} บนกราฟ Biplot">
              <td><strong style="color:${isVectorVisible ? '#38BDF8' : '#64748B'};">${isVectorVisible ? '● ' : '○ '}${escapeHtml(col)}</strong></td>
              <td style="color:${Math.abs(w1) >= 0.3 ? '#6EE7B7' : '#F8FAFC'}">${w1 >= 0 ? '+' : ''}${w1.toFixed(3)}</td>
              <td style="color:${Math.abs(w2) >= 0.3 ? '#6EE7B7' : '#F8FAFC'}">${w2 >= 0 ? '+' : ''}${w2.toFixed(3)}</td>
              <td style="color:${Math.abs(w3) >= 0.3 ? '#F472B6' : '#94A3B8'}">${w3 >= 0 ? '+' : ''}${w3.toFixed(3)}</td>
              <td><span class="badge badge-emerald">${power2d.toFixed(3)}</span></td>
            </tr>
          `;
        })
        .join('');

      tableWrap.innerHTML = `
        <table class="data-table">
          <thead>
            <tr>
              <th>Sensor</th>
              <th>PC1</th>
              <th>PC2</th>
              <th>PC3</th>
              <th>|L₁₂|</th>
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      `;

      tableWrap.querySelectorAll('tr[data-biplotsensor]').forEach(rowEl => {
        rowEl.addEventListener('click', () => {
          const sCol = rowEl.dataset.biplotsensor;
          if (!sCol) return;
          if (!state.showBiplotVectors) {
            state.showBiplotVectors = true;
            state.activeBiplotSensors.add(sCol);
          } else if (state.activeBiplotSensors.has(sCol)) {
            state.activeBiplotSensors.delete(sCol);
          } else {
            state.activeBiplotSensors.add(sCol);
          }
          refreshDashboard(true);
        });
      });
    }
  }

  /**
   * TAB 3: Scree Plot & Cumulative Explained Variance (PC1..PCn with 85% Threshold Line)
   */
  function renderScreeTab(pca) {
    const plotEl = document.getElementById('plotScree');
    if (!plotEl || !pca) return;

    const expVars = pca.allExplainedVariance || [];
    const cumVars = pca.allCumulativeVariance || [];
    const pcLabels = expVars.map((_, i) => `PC${i + 1}`);

    const traces = [
      {
        type: 'bar',
        name: 'Individual Explained Variance (%)',
        x: pcLabels,
        y: expVars,
        text: expVars.map(v => `${v.toFixed(1)}%`),
        textposition: 'auto',
        marker: {
          color: expVars.map((_, i) => (i < 3 ? '#38BDF8' : 'rgba(56, 189, 248, 0.38)')),
          line: { color: '#0284C7', width: 1 }
        },
        hovertemplate: `<b>%{x}</b><br>Individual Variance: <b>%{y:.2f}%</b><extra></extra>`
      },
      {
        type: 'scatter',
        mode: 'lines+markers+text',
        name: 'Cumulative Explained Variance (%)',
        x: pcLabels,
        y: cumVars,
        text: cumVars.map(v => `${v.toFixed(1)}%`),
        textposition: 'top center',
        textfont: { family: "'JetBrains Mono', monospace", size: 10.5, color: '#34D399' },
        line: { color: '#10B981', width: 3 },
        marker: { size: 8, color: '#34D399', symbol: 'circle' },
        hovertemplate: `<b>%{x}</b><br>Cumulative Variance: <b>%{y:.2f}%</b><extra></extra>`
      }
    ];

    const layout = {
      title: {
        text: `📊 Scree Plot & Cumulative Explained Variance — 2D (PC1+PC2 = ${(pca.cumulativeVariance2D || 0).toFixed(1)}%) • 3D (PC1–PC3 = ${(pca.cumulativeVariance || 0).toFixed(1)}%)`,
        font: { size: 13.5, color: '#F8FAFC' }
      },
      paper_bgcolor: '#0B0F19',
      plot_bgcolor: '#0B0F19',
      margin: { l: 64, r: 36, t: 54, b: 54 },
      font: { family: "'Plus Jakarta Sans', 'Noto Sans Thai', sans-serif", color: '#F8FAFC' },
      legend: { orientation: 'h', y: 1.08, x: 0.5, xanchor: 'center' },
      shapes: [
        {
          type: 'line',
          xref: 'paper',
          x0: 0,
          x1: 1,
          yref: 'y',
          y0: 85,
          y1: 85,
          line: { color: '#10B981', width: 2, dash: 'dash' }
        }
      ],
      annotations: [
        {
          xref: 'paper',
          x: 0.98,
          yref: 'y',
          y: 87.5,
          text: '✓ 85% Chemometrics Threshold',
          showarrow: false,
          font: { family: "'JetBrains Mono', monospace", size: 11, color: '#34D399' }
        }
      ],
      xaxis: {
        title: { text: 'Principal Components (PC1 – PCn)', font: { size: 12, color: '#94A3B8' } },
        gridcolor: 'rgba(148, 163, 184, 0.12)'
      },
      yaxis: {
        title: { text: 'Explained Variance (%)', font: { size: 12, color: '#38BDF8' } },
        range: [0, 106],
        gridcolor: 'rgba(148, 163, 184, 0.14)'
      }
    };

    Plotly.react(plotEl, traces, layout, { responsive: true, displaylogo: false });
  }

  /**
   * TAB 4: Odor Fingerprint Radar / Spider Chart (scatterpolar across Active Sensors)
   */
  function renderRadarTab(resultByClass, pca) {
    const plotEl = document.getElementById('plotRadar');
    if (!plotEl) return;

    const activeSensors = (pca && pca.validSensors && pca.validSensors.length > 0)
      ? pca.validSensors
      : state.sensorCols;
    if (activeSensors.length === 0) return;

    // Compute per-sample mean across active sensors
    const sampleMeans = [];
    for (const s of state.sampleClasses) {
      if (!state.activeSamples.has(s.id)) continue;
      const rows = resultByClass[s.id] || [];
      if (rows.length === 0) continue;

      const meansVec = activeSensors.map(col => {
        let sum = 0;
        for (let i = 0; i < rows.length; i++) sum += rows[i][col] || 0;
        return sum / rows.length;
      });
      sampleMeans.push({ sampleObj: s, meansVec, count: rows.length });
    }

    // Compute per-sensor min/max/mean/sd across all active samples for scaling
    const sensorMin = new Array(activeSensors.length).fill(Infinity);
    const sensorMax = new Array(activeSensors.length).fill(-Infinity);
    for (const item of sampleMeans) {
      item.meansVec.forEach((v, idx) => {
        if (v < sensorMin[idx]) sensorMin[idx] = v;
        if (v > sensorMax[idx]) sensorMax[idx] = v;
      });
    }

    const traces = sampleMeans.map(item => {
      const s = item.sampleObj;
      const rVals = item.meansVec.map((val, j) => {
        if (state.radarScale === 'zscore' && pca && pca.means && pca.stds) {
          const mu = pca.means[j] || 0;
          const sd = pca.stds[j] || 1;
          return (val - mu) / sd;
        }
        if (state.radarScale === 'norm') {
          const span = sensorMax[j] - sensorMin[j];
          return span > 1e-9 ? ((val - sensorMin[j]) / span) * 100 : 50;
        }
        return val;
      });

      // Close loop for polar chart
      const thetaClosed = [...activeSensors, activeSensors[0]];
      const rClosed = [...rVals, rVals[0]];

      return {
        type: 'scatterpolar',
        mode: 'lines+markers',
        name: s.label,
        r: rClosed,
        theta: thetaClosed,
        fill: 'toself',
        fillcolor: `${s.color}1F`,
        line: { color: s.color, width: 2.2 },
        marker: { size: 6, color: s.color },
        hovertemplate: `<b>${s.shapeIcon || '●'} ${escapeHtml(s.label)}</b><br>Sensor %{theta}: <b>%{r:.3f}</b><extra></extra>`
      };
    });

    const scaleTitleMap = {
      zscore: 'Z-Score Standardized ((X − μ) / σ)',
      norm: 'Normalized Relative Response (0 – 100%)',
      raw: 'Mean Sensor Response'
    };

    const layout = {
      title: {
        text: `🕸️ Odor Fingerprint Radar Chart — ${scaleTitleMap[state.radarScale] || ''} (${activeSensors.length} Active Sensors)`,
        font: { size: 13.5, color: '#F8FAFC' }
      },
      paper_bgcolor: '#0B0F19',
      plot_bgcolor: '#0B0F19',
      margin: { l: 60, r: 60, t: 54, b: 44 },
      font: { family: "'Plus Jakarta Sans', 'Noto Sans Thai', sans-serif", color: '#F8FAFC' },
      legend: { orientation: 'v', x: 1.02, y: 0.95, font: { size: 11 } },
      polar: {
        bgcolor: 'rgba(17, 24, 39, 0.65)',
        angularaxis: {
          gridcolor: 'rgba(148, 163, 184, 0.22)',
          linecolor: 'rgba(56, 189, 248, 0.4)',
          tickfont: { family: "'JetBrains Mono', monospace", size: 11.5, color: '#38BDF8' }
        },
        radialaxis: {
          gridcolor: 'rgba(148, 163, 184, 0.18)',
          linecolor: 'rgba(148, 163, 184, 0.3)',
          tickfont: { family: "'JetBrains Mono', monospace", size: 10, color: '#94A3B8' }
        }
      }
    };

    Plotly.react(plotEl, traces, layout, { responsive: true, displaylogo: false });
  }

  /**
   * TAB 5: Sensor Fingerprint Heatmap (Sample × Active Sensor Z-Score Matrix, sorted by PC1 or Name)
   */
  function renderHeatmapTab(resultByClass, pca) {
    const plotEl = document.getElementById('plotHeatmap');
    if (!plotEl || !pca) return;

    const activeSensors = pca.validSensors && pca.validSensors.length > 0 ? pca.validSensors : state.sensorCols;
    const sampleRowsData = [];

    for (const s of state.sampleClasses) {
      if (!state.activeSamples.has(s.id)) continue;
      const rows = resultByClass[s.id] || [];
      if (rows.length === 0) continue;

      let pc1Sum = 0;
      const zSum = new Array(activeSensors.length).fill(0);

      for (let i = 0; i < rows.length; i++) {
        const [pc1] = pca.projectRow(rows[i]);
        pc1Sum += pc1;
        const zVec = pca.getZScoreRow(rows[i]);
        for (let j = 0; j < activeSensors.length; j++) {
          zSum[j] += zVec[j] || 0;
        }
      }

      sampleRowsData.push({
        sampleObj: s,
        label: `${s.shapeIcon || '●'} ${s.label}`,
        meanPc1: pc1Sum / rows.length,
        meanZVec: zSum.map(v => Number((v / rows.length).toFixed(2)))
      });
    }

    if (state.heatmapSort === 'pc1') {
      sampleRowsData.sort((a, b) => b.meanPc1 - a.meanPc1);
    } else {
      sampleRowsData.sort((a, b) => naturalCompare(a.sampleObj.label, b.sampleObj.label));
    }

    const yLabels = sampleRowsData.map(item => item.label);
    const zMatrix = sampleRowsData.map(item => item.meanZVec);

    const trace = {
      type: 'heatmap',
      x: activeSensors,
      y: yLabels,
      z: zMatrix,
      colorscale: 'RdBu',
      reversescale: true,
      zmid: 0,
      text: zMatrix,
      texttemplate: '%{z:.2f}',
      textfont: { family: "'JetBrains Mono', monospace", size: 11 },
      colorbar: {
        title: { text: 'Mean Z-Score', font: { size: 11, color: '#94A3B8' } },
        tickfont: { family: "'JetBrains Mono', monospace", size: 10, color: '#94A3B8' }
      },
      hovertemplate: `Sample: <b>%{y}</b><br>Sensor: <b>%{x}</b><br>Mean Z-Score: <b>%{z:.3f}</b><extra></extra>`
    };

    const layout = {
      title: {
        text: `🔥 Sensor Fingerprint Heatmap — Mean Z-Score Matrix (${yLabels.length} Samples × ${activeSensors.length} Active Sensors)`,
        font: { size: 13.5, color: '#F8FAFC' }
      },
      paper_bgcolor: '#0B0F19',
      plot_bgcolor: '#0B0F19',
      margin: { l: 130, r: 40, t: 52, b: 56 },
      font: { family: "'Plus Jakarta Sans', 'Noto Sans Thai', sans-serif", color: '#F8FAFC' },
      xaxis: {
        title: { text: 'Active Gas Sensors (Auto-Cleaned)', font: { size: 12, color: '#38BDF8' } },
        tickfont: { family: "'JetBrains Mono', monospace", size: 11, color: '#38BDF8' }
      },
      yaxis: {
        autorange: 'reversed',
        tickfont: { family: "'JetBrains Mono', 'Noto Sans Thai', monospace", size: 11, color: '#F8FAFC' }
      }
    };

    Plotly.react(plotEl, [trace], layout, { responsive: true, displaylogo: false });
  }

  /**
   * Populate Time-Series Tab Dropdowns dynamically
   */
  function populateTimeSeriesControls() {
    const sensorSel = document.getElementById('tsSensorSelect');
    const sampleSel = document.getElementById('tsSampleSelect');
    if (sensorSel) {
      sensorSel.innerHTML = '';
      state.sensorCols.forEach(col => {
        const opt = document.createElement('option');
        opt.value = col;
        opt.textContent = col;
        sensorSel.appendChild(opt);
      });
      if (!state.tsSensor || !state.sensorCols.includes(state.tsSensor)) {
        state.tsSensor = state.defaultCustomAxes.x || state.sensorCols[0] || '';
      }
      sensorSel.value = state.tsSensor;
    }

    if (sampleSel) {
      sampleSel.innerHTML = '';
      state.sampleClasses.forEach(s => {
        const opt = document.createElement('option');
        opt.value = s.id;
        opt.textContent = `${s.label} (${s.groupLabel})`;
        sampleSel.appendChild(opt);
      });
      if (!state.tsSample || !state.sampleMap[state.tsSample]) {
        state.tsSample = (state.sampleClasses[0] && state.sampleClasses[0].id) || '';
      }
      sampleSel.value = state.tsSample;
    }
  }

  /**
   * TAB 6: Dynamic Sensor Time-Series Curve with Automatic Phase Bands (Resting / Sampling / Purging)
   */
  function renderTimeSeriesTab() {
    const plotEl = document.getElementById('plotTimeSeries');
    if (!plotEl) return;

    const dsKey = state.tsSource || 'Raw';
    const ds = state.datasets[dsKey] || {};
    const baselines = state.baselineCorrection ? computeBlankBaselines(dsKey) : null;

    const traces = [];
    const shapes = [];
    let referenceRowsForPhases = [];

    if (state.tsMode === 'compare_samples') {
      const targetSensor = state.tsSensor || state.defaultCustomAxes.x || state.sensorCols[0];
      for (const s of state.sampleClasses) {
        if (!state.activeSamples.has(s.id)) continue;
        const rawRows = (ds[s.id] || []).filter(r => state.activeCycles.has(r['Cycle Loop']));
        if (rawRows.length === 0) continue;
        if (referenceRowsForPhases.length === 0) referenceRowsForPhases = rawRows;

        const assignedBlank = state.blankAssignment[s.id];
        const blankRef = (baselines && assignedBlank) ? baselines[assignedBlank] : null;

        const xIdx = [];
        const yVal = [];
        const custom = [];

        for (let i = 0; i < rawRows.length; i++) {
          const r = rawRows[i];
          const cyc = r['Cycle Loop'];
          const bVec = blankRef
            ? ((state.perCycleBlank && blankRef.byCycleAndMode?.[cyc]?.[r.Mode]) ||
               (state.perCycleBlank && blankRef.byCycle?.[cyc]) ||
               blankRef.byMode?.[r.Mode] ||
               blankRef.overall)
            : null;
          const v = (r[targetSensor] || 0) - (bVec ? (bVec[targetSensor] || 0) : 0);
          xIdx.push(i + 1);
          yVal.push(v);
          custom.push([r.Time, cyc, r.Mode]);
        }

        traces.push({
          type: 'scatter',
          mode: 'lines',
          name: s.label,
          x: xIdx,
          y: yVal,
          customdata: custom,
          line: { color: s.color, width: 2 },
          hovertemplate:
            `<b>${escapeHtml(s.label)}</b> (${targetSensor})<br>` +
            `Point #%{x} • Time: %{customdata[0]}<br>` +
            `Cycle Loop: #%{customdata[1]} (%{customdata[2]})<br>` +
            `Signal: <b>%{y:.4f}</b><extra></extra>`
        });
      }
    } else {
      // Multi-Sensor view for 1 selected sample
      const sampleId = state.tsSample || (state.sampleClasses[0] && state.sampleClasses[0].id);
      const sObj = state.sampleMap[sampleId];
      const rawRows = (ds[sampleId] || []).filter(r => state.activeCycles.has(r['Cycle Loop']));
      referenceRowsForPhases = rawRows;

      const assignedBlank = state.blankAssignment[sampleId];
      const blankRef = (baselines && assignedBlank) ? baselines[assignedBlank] : null;
      const activeSensors = (state.pcaResult && state.pcaResult.validSensors.length > 0)
        ? state.pcaResult.validSensors
        : state.sensorCols;

      activeSensors.forEach((col, sIdx) => {
        const xIdx = [];
        const yVal = [];
        const custom = [];
        const color = generateDynamicColor(sIdx, false);

        for (let i = 0; i < rawRows.length; i++) {
          const r = rawRows[i];
          const cyc = r['Cycle Loop'];
          const bVec = blankRef
            ? ((state.perCycleBlank && blankRef.byCycleAndMode?.[cyc]?.[r.Mode]) ||
               (state.perCycleBlank && blankRef.byCycle?.[cyc]) ||
               blankRef.byMode?.[r.Mode] ||
               blankRef.overall)
            : null;
          const v = (r[col] || 0) - (bVec ? (bVec[col] || 0) : 0);
          xIdx.push(i + 1);
          yVal.push(v);
          custom.push([r.Time, cyc, r.Mode]);
        }

        traces.push({
          type: 'scatter',
          mode: 'lines',
          name: col,
          x: xIdx,
          y: yVal,
          customdata: custom,
          line: { color, width: 2 },
          hovertemplate:
            `<b>Sensor ${col}</b> (${escapeHtml(sObj ? sObj.label : sampleId)})<br>` +
            `Time: %{customdata[0]} • Loop #%{customdata[1]} (%{customdata[2]})<br>` +
            `Signal: <b>%{y:.4f}</b><extra></extra>`
        });
      });
    }

    // Build automatic background phase bands (Resting / Sampling / Purging) from referenceRowsForPhases
    if (referenceRowsForPhases.length > 0) {
      const modeColorMap = {
        Resting: 'rgba(148, 163, 184, 0.08)',
        Sampling: 'rgba(16, 185, 129, 0.09)',
        Purging: 'rgba(245, 158, 11, 0.08)'
      };
      let segStart = 1;
      let curMode = referenceRowsForPhases[0].Mode;
      for (let i = 1; i <= referenceRowsForPhases.length; i++) {
        const m = i < referenceRowsForPhases.length ? referenceRowsForPhases[i].Mode : null;
        if (m !== curMode) {
          if (modeColorMap[curMode]) {
            shapes.push({
              type: 'rect',
              xref: 'x',
              yref: 'paper',
              x0: segStart,
              x1: i,
              y0: 0,
              y1: 1,
              fillcolor: modeColorMap[curMode],
              line: { width: 0 }
            });
          }
          segStart = i + 1;
          curMode = m;
        }
      }
    }

    const layout = {
      title: {
        text: state.tsMode === 'compare_samples'
          ? `📈 Dynamic Sensor Time-Series (${dsKey}_*) — Sensor ${state.tsSensor} Across Active Samples`
          : `📈 Multi-Sensor Response Curve (${dsKey}_*) — Sample: ${(state.sampleMap[state.tsSample]?.label) || state.tsSample}`,
        font: { size: 13.5, color: '#F8FAFC' }
      },
      paper_bgcolor: '#0B0F19',
      plot_bgcolor: '#0B0F19',
      margin: { l: 64, r: 28, t: 52, b: 52 },
      font: { family: "'Plus Jakarta Sans', 'Noto Sans Thai', sans-serif", color: '#F8FAFC' },
      legend: { orientation: 'h', y: 1.08, x: 0.5, xanchor: 'center', font: { size: 10.5 } },
      shapes,
      xaxis: {
        title: { text: 'Sequential Measurement Index (Resting = Gray • Sampling = Green • Purging = Amber)', font: { size: 11.5, color: '#94A3B8' } },
        gridcolor: 'rgba(148, 163, 184, 0.12)'
      },
      yaxis: {
        title: { text: state.baselineCorrection ? 'Sensor Signal (ΔBlank)' : 'Raw Sensor Signal', font: { size: 12, color: '#38BDF8' } },
        gridcolor: 'rgba(148, 163, 184, 0.14)'
      }
    };

    Plotly.react(plotEl, traces, layout, { responsive: true, displaylogo: false });
  }

  /**
   * TAB 7: Quality Control (QC) & Blank Baseline Stability Monitor
   */
  function renderQcTab(resultByClass) {
    const envEl = document.getElementById('plotQcEnv');
    const blankEl = document.getElementById('plotQcBlank');
    if (!envEl || !blankEl) return;

    // 1. Environmental Stability Chart (Temp & Flow Rate Mean ± SD per Active Sample)
    const sampleLabels = [];
    const tempMeans = [];
    const tempSds = [];
    const flowMeans = [];
    const flowSds = [];

    for (const s of state.sampleClasses) {
      if (!state.activeSamples.has(s.id)) continue;
      const rows = resultByClass[s.id] || [];
      if (rows.length === 0) continue;

      const tStats = computeDescriptiveStats(rows.map(r => r['Temp'] || 0));
      const fStats = computeDescriptiveStats(rows.map(r => r['Flow Rate'] || 0));

      sampleLabels.push(s.label);
      tempMeans.push(tStats.mean);
      tempSds.push(tStats.sd);
      flowMeans.push(fStats.mean);
      flowSds.push(fStats.sd);
    }

    const envTraces = [
      {
        type: 'scatter',
        mode: 'lines+markers',
        name: 'Temp (°C) Mean ± SD',
        x: sampleLabels,
        y: tempMeans,
        error_y: { type: 'data', array: tempSds, visible: true, color: '#F472B6' },
        line: { color: '#F472B6', width: 2.2 },
        marker: { size: 7, color: '#F472B6' },
        yaxis: 'y1'
      },
      {
        type: 'scatter',
        mode: 'lines+markers',
        name: 'Flow Rate (mL/min) Mean ± SD',
        x: sampleLabels,
        y: flowMeans,
        error_y: { type: 'data', array: flowSds, visible: true, color: '#38BDF8' },
        line: { color: '#38BDF8', width: 2.2 },
        marker: { size: 7, color: '#38BDF8', symbol: 'square' },
        yaxis: 'y2'
      }
    ];

    const envLayout = {
      title: { text: '🌡️ Environmental QC: Temperature (°C) & Flow Rate (mL/min) Stability', font: { size: 13, color: '#F8FAFC' } },
      paper_bgcolor: '#0B0F19',
      plot_bgcolor: '#0B0F19',
      margin: { l: 58, r: 62, t: 52, b: 58 },
      font: { family: "'Plus Jakarta Sans', 'Noto Sans Thai', sans-serif", color: '#F8FAFC' },
      legend: { orientation: 'h', y: 1.09, x: 0.5, xanchor: 'center', font: { size: 10.5 } },
      xaxis: { title: { text: 'Sample Classes', font: { size: 11.5, color: '#94A3B8' } }, gridcolor: 'rgba(148, 163, 184, 0.12)' },
      yaxis: {
        title: { text: 'Temperature (°C)', font: { size: 11.5, color: '#F472B6' } },
        tickfont: { color: '#F472B6' },
        gridcolor: 'rgba(148, 163, 184, 0.12)'
      },
      yaxis2: {
        title: { text: 'Flow Rate (mL/min)', font: { size: 11.5, color: '#38BDF8' } },
        tickfont: { color: '#38BDF8' },
        overlaying: 'y',
        side: 'right'
      }
    };

    Plotly.react(envEl, envTraces, envLayout, { responsive: true, displaylogo: false });

    // 2. Reference Blank Stability Monitor across Active Sensors
    const blankClasses = state.sampleClasses.filter(s => s.isBlank);
    const activeSensors = (state.pcaResult && state.pcaResult.validSensors.length > 0)
      ? state.pcaResult.validSensors
      : state.sensorCols;
    const rawDs = state.datasets[state.activeDataset] || {};

    const blankTraces = blankClasses.map((bObj, idx) => {
      const bRows = (rawDs[bObj.id] || []).filter(r => state.activeCycles.has(r['Cycle Loop']));
      const means = [];
      const sds = [];
      activeSensors.forEach(col => {
        const st = computeDescriptiveStats(bRows.map(r => r[col] || 0));
        means.push(st.mean);
        sds.push(st.sd);
      });

      const lineColor = idx === 0 ? '#34D399' : '#FBBF24';
      return {
        type: 'bar',
        name: `${bObj.label} (n=${bRows.length})`,
        x: activeSensors,
        y: means,
        error_y: { type: 'data', array: sds, visible: true, color: '#F8FAFC' },
        marker: { color: lineColor, opacity: 0.85 }
      };
    });

    const blankLayout = {
      title: { text: '⚪ Reference Blank Baseline Stability Across Active Sensors (Mean ± SD)', font: { size: 13, color: '#F8FAFC' } },
      paper_bgcolor: '#0B0F19',
      plot_bgcolor: '#0B0F19',
      barmode: 'group',
      margin: { l: 58, r: 28, t: 52, b: 58 },
      font: { family: "'Plus Jakarta Sans', 'Noto Sans Thai', sans-serif", color: '#F8FAFC' },
      legend: { orientation: 'h', y: 1.09, x: 0.5, xanchor: 'center', font: { size: 10.5 } },
      xaxis: { title: { text: 'Active Gas Sensors', font: { size: 11.5, color: '#38BDF8' } }, gridcolor: 'rgba(148, 163, 184, 0.12)' },
      yaxis: { title: { text: 'Baseline Sensor Response', font: { size: 11.5, color: '#34D399' } }, gridcolor: 'rgba(148, 163, 184, 0.12)' }
    };

    Plotly.react(blankEl, blankTraces, blankLayout, { responsive: true, displaylogo: false });
  }

  /**
   * Compute Centroids & Pairwise Euclidean Distance Matrix for Tab 8 & Excel Export
   */
  function computeEuclideanDistanceAnalysis(resultByClass, pca) {
    const activeClassObjs = state.sampleClasses.filter(
      s => state.activeSamples.has(s.id) && (resultByClass[s.id] || []).length > 0
    );

    const centroids = activeClassObjs.map(s => {
      const rows = resultByClass[s.id] || [];
      if (state.distanceSpace === 'zscore' && pca && pca.validSensors.length > 0) {
        const p = pca.validSensors.length;
        const vec = new Array(p).fill(0);
        for (let i = 0; i < rows.length; i++) {
          const z = pca.getZScoreRow(rows[i]);
          for (let j = 0; j < p; j++) vec[j] += z[j] || 0;
        }
        return { sampleObj: s, vec: vec.map(v => v / rows.length) };
      } else {
        let sx = 0, sy = 0, sz = 0;
        for (let i = 0; i < rows.length; i++) {
          const [x, y, z] = pca ? pca.projectRow(rows[i]) : [0, 0, 0];
          sx += x; sy += y; sz += z;
        }
        return { sampleObj: s, vec: [sx / rows.length, sy / rows.length, sz / rows.length] };
      }
    });

    const m = centroids.length;
    const matrix = Array.from({ length: m }, () => new Array(m).fill(0));
    const pairs = [];

    for (let i = 0; i < m; i++) {
      for (let j = 0; j < m; j++) {
        if (i === j) {
          matrix[i][j] = 0;
          continue;
        }
        const va = centroids[i].vec;
        const vb = centroids[j].vec;
        let sumSq = 0;
        for (let k = 0; k < va.length; k++) {
          const d = (va[k] || 0) - (vb[k] || 0);
          sumSq += d * d;
        }
        const dist = Math.sqrt(sumSq);
        matrix[i][j] = Number(dist.toFixed(3));
        if (i < j) {
          pairs.push({
            a: centroids[i].sampleObj.label,
            b: centroids[j].sampleObj.label,
            dist
          });
        }
      }
    }

    pairs.sort((p1, p2) => p1.dist - p2.dist);
    return {
      labels: centroids.map(c => c.sampleObj.label),
      matrix,
      mostSimilar: pairs.slice(0, 3),
      mostDistinct: [...pairs].reverse().slice(0, 3)
    };
  }

  /**
   * TAB 8: Pairwise Euclidean Distance Matrix & Comprehensive Statistics Table
   */
  function renderDistanceTab(resultByClass, pca, getXYZ, axisTitleX, axisTitleY, axisTitleZ) {
    const plotEl = document.getElementById('plotDistanceMatrix');
    const bannerEl = document.getElementById('distancePairsBanner');
    const tableEl = document.getElementById('fullStatsTableContainer');
    if (!plotEl) return;

    const { labels, matrix, mostSimilar, mostDistinct } = computeEuclideanDistanceAnalysis(resultByClass, pca);

    if (bannerEl) {
      const simHtml = mostSimilar
        .map(p => `<span class="pair-chip similar">ใกล้สุด: ${escapeHtml(p.a)}–${escapeHtml(p.b)} (${p.dist.toFixed(2)})</span>`)
        .join('');
      const distHtml = mostDistinct
        .slice(0, 2)
        .map(p => `<span class="pair-chip distinct">ต่างสุด: ${escapeHtml(p.a)}–${escapeHtml(p.b)} (${p.dist.toFixed(2)})</span>`)
        .join('');
      bannerEl.innerHTML = simHtml + distHtml;
    }

    const trace = {
      type: 'heatmap',
      x: labels,
      y: labels,
      z: matrix,
      colorscale: 'Viridis',
      text: matrix,
      texttemplate: '%{z:.2f}',
      textfont: { family: "'JetBrains Mono', monospace", size: 10.5 },
      hovertemplate: `Pair: <b>%{y} vs %{x}</b><br>Euclidean Distance: <b>%{z:.3f}</b><extra></extra>`
    };

    const layout = {
      title: {
        text: `📏 Pairwise Euclidean Distance Matrix (${state.distanceSpace === 'pca' ? 'PCA Space PC1–PC3' : 'Z-Score Sensor Space'})`,
        font: { size: 13, color: '#F8FAFC' }
      },
      paper_bgcolor: '#0B0F19',
      plot_bgcolor: '#0B0F19',
      margin: { l: 95, r: 28, t: 48, b: 60 },
      font: { family: "'Plus Jakarta Sans', 'Noto Sans Thai', sans-serif", color: '#F8FAFC' },
      yaxis: { autorange: 'reversed' }
    };

    Plotly.react(plotEl, [trace], layout, { responsive: true, displaylogo: false });

    // Render Comprehensive Statistics Table (Count n, Mean, SD, Median, Min, Max)
    if (tableEl) {
      const shortX = state.axisMode === 'pca' ? 'PC1' : state.customAxes.x;
      const shortY = state.axisMode === 'pca' ? 'PC2' : state.customAxes.y;
      const shortZ = state.axisMode === 'pca' ? 'PC3' : state.customAxes.z;

      const bodyRows = [];
      for (const s of state.sampleClasses) {
        if (!state.activeSamples.has(s.id)) continue;
        const rows = resultByClass[s.id] || [];
        if (rows.length === 0) continue;

        const xs = [];
        const ys = [];
        const zs = [];
        for (let i = 0; i < rows.length; i++) {
          const [x, y, z] = getXYZ(rows[i]);
          xs.push(x); ys.push(y); zs.push(z);
        }

        const axesSpecs = [
          { name: `X (${shortX})`, st: computeDescriptiveStats(xs) },
          { name: `Y (${shortY})`, st: computeDescriptiveStats(ys) },
          { name: `Z (${shortZ})`, st: computeDescriptiveStats(zs) }
        ];

        axesSpecs.forEach((ax, idx) => {
          bodyRows.push(`
            <tr>
              ${idx === 0 ? `<td rowspan="3"><strong style="color:${s.color}">${s.shapeIcon || '●'} ${escapeHtml(s.label)}</strong><br><small style="color:#94A3B8">n = ${rows.length}</small></td>` : ''}
              <td>${escapeHtml(ax.name)}</td>
              <td>${ax.st.mean.toFixed(3)}</td>
              <td>${ax.st.sd.toFixed(3)}</td>
              <td>${ax.st.median.toFixed(3)}</td>
              <td>${ax.st.min.toFixed(3)}</td>
              <td>${ax.st.max.toFixed(3)}</td>
            </tr>
          `);
        });
      }

      tableEl.innerHTML = `
        <table class="data-table">
          <thead>
            <tr>
              <th>Sample (n)</th>
              <th>Axis</th>
              <th>Mean</th>
              <th>SD</th>
              <th>Median</th>
              <th>Min</th>
              <th>Max</th>
            </tr>
          </thead>
          <tbody>${bodyRows.join('')}</tbody>
        </table>
      `;
    }
  }

  // --------------------------------------------------------------------------
  // 8. FOCUS HIGHLIGHT (SIDEBAR HOVER -> BRIGHTEN AREA HULL & DIM OTHERS)
  // --------------------------------------------------------------------------

  let focusRafId = null;

  function applyFocusHighlight(hoveredSampleId) {
    state.focusedSample = hoveredSampleId;
    if (focusRafId) cancelAnimationFrame(focusRafId);

    focusRafId = requestAnimationFrame(() => {
      const plotEl = document.getElementById('plot3d');
      if (!plotEl || !state.traceMeta.length) return;

      const bannerDot = document.getElementById('focusBannerDot');
      const bannerText = document.getElementById('focusBannerText');

      if (!hoveredSampleId) {
        const scatterIndices = [];
        const hullIndices = [];

        state.traceMeta.forEach((meta, idx) => {
          if (meta.traceType === 'scatter') scatterIndices.push(idx);
          else if (meta.traceType === 'hull') hullIndices.push(idx);
        });

        if (scatterIndices.length > 0) {
          Plotly.restyle(plotEl, {
            'marker.opacity': 0.92,
            'marker.size': state.markerSize
          }, scatterIndices);
        }
        if (hullIndices.length > 0) {
          Plotly.restyle(plotEl, {
            opacity: state.hullOpacity
          }, hullIndices);
        }

        if (bannerDot && bannerText) {
          bannerDot.style.background = '#38BDF8';
          bannerDot.style.boxShadow = '0 0 8px #38BDF8';
          bannerText.textContent = 'คลิกค้างเพื่อหมุน 360° • เลื่อนลูกกลิ้งเพื่อซูม • คลิกขวาค้างเพื่อเลื่อนแกน (Pan)';
        }
        return;
      }

      const sampleInfo = state.sampleMap[hoveredSampleId];
      if (bannerDot && bannerText && sampleInfo) {
        bannerDot.style.background = sampleInfo.color;
        bannerDot.style.boxShadow = `0 0 10px ${sampleInfo.color}`;
        bannerText.innerHTML = `กำลังเน้นพื้นที่ (Focus Highlight): <strong style="color:${sampleInfo.color}">${sampleInfo.shapeIcon || '●'} ${escapeHtml(sampleInfo.label)}</strong> (${escapeHtml(sampleInfo.groupLabel)})`;
      }

      state.traceMeta.forEach((meta, idx) => {
        const isTarget = meta.sampleId === hoveredSampleId;
        if (meta.traceType === 'scatter') {
          Plotly.restyle(plotEl, {
            'marker.opacity': isTarget ? 1.0 : 0.09,
            'marker.size': isTarget ? state.markerSize + 2.0 : Math.max(2, state.markerSize - 1.0)
          }, [idx]);
        } else if (meta.traceType === 'hull') {
          Plotly.restyle(plotEl, {
            opacity: isTarget ? Math.min(0.60, state.hullOpacity * 2.3) : 0.02
          }, [idx]);
        }
      });
    });
  }

  // --------------------------------------------------------------------------
  // 9. DYNAMIC REACTIVE UI RENDERERS & PART 4 EXCEL EXPORT
  // --------------------------------------------------------------------------

  function updateHeaderAndPipelineUI() {
    const badgeEl = document.getElementById('dataSourceBadge');
    if (badgeEl) {
      badgeEl.textContent = state.isDemoData ? 'DEMO DATA' : 'EXCEL LOADED';
      badgeEl.className = state.isDemoData ? 'badge badge-amber' : 'badge badge-emerald';
    }

    const fileNameEl = document.getElementById('loadedFileName');
    if (fileNameEl) {
      fileNameEl.textContent = state.fileName;
      fileNameEl.title = state.fileName;
    }

    const sidebarFileBadge = document.getElementById('sidebarActiveFileBadge');
    if (sidebarFileBadge) {
      sidebarFileBadge.textContent = `📄 ไฟล์ปัจจุบัน: ${state.fileName}`;
      sidebarFileBadge.title = state.fileName;
    }

    const sheetCountEl = document.getElementById('datasetSheetCount');
    if (sheetCountEl) {
      sheetCountEl.textContent = `${state.sampleClasses.length} ชนิด (${state.activeDataset}_*)`;
    }

    // Update detected samples & blanks summary banner
    const sampleCount = state.sampleClasses.filter(s => !s.isBlank).length;
    const blankCount = state.sampleClasses.filter(s => s.isBlank).length;
    const summaryBanner = document.getElementById('detectedSampleSummaryBanner');
    if (summaryBanner) {
      summaryBanner.innerHTML =
        `<span>พบข้อมูลตัวอย่าง <strong style="color:var(--accent-cyan)">${sampleCount}</strong> ชนิด และ Blank <strong style="color:#E2E8F0">${blankCount}</strong> ชนิด</span>` +
        `<span class="badge">${state.sensorCols.length} เซนเซอร์</span>`;
    }

    // Update PCA sensor range label
    const pcaRangeEl = document.getElementById('pcaSensorRangeLabel');
    if (pcaRangeEl && state.sensorCols.length > 0) {
      pcaRangeEl.textContent = `ลดมิติข้อมูล (${state.sensorCols[0]}–${state.sensorCols[state.sensorCols.length - 1]})`;
    }
  }

  function populateAxisDropdowns() {
    const selX = document.getElementById('selectAxisX');
    const selY = document.getElementById('selectAxisY');
    const selZ = document.getElementById('selectAxisZ');
    const allOptions = [...state.sensorCols, ...state.extraNumericCols];

    [selX, selY, selZ].forEach(sel => {
      sel.innerHTML = '';
      allOptions.forEach(col => {
        const opt = document.createElement('option');
        opt.value = col;
        opt.textContent = state.sensorCols.includes(col) ? `${col} (Gas Sensor)` : `${col} (Parameter)`;
        sel.appendChild(opt);
      });
    });

    selX.value = state.customAxes.x;
    selY.value = state.customAxes.y;
    selZ.value = state.customAxes.z;

    const defHint = document.getElementById('defaultAxesHintText');
    const resetBtn = document.getElementById('btnResetDefaultAxes');
    const { x, y, z } = state.defaultCustomAxes;
    if (defHint) defHint.textContent = `ค่าเริ่มต้น: X=${x}, Y=${y}, Z=${z}`;
    if (resetBtn) resetBtn.textContent = `รีเซ็ต ${x}/${y}/${z}`;
  }

  /**
   * Render Dynamic Quick Group Action Buttons based on discovered Samples, Blanks, and BlankAssignment / Date groups
   */
  function renderDynamicGroupButtons() {
    const container = document.getElementById('dynamicGroupActionsContainer');
    container.innerHTML = '';

    const addQuickBtn = (label, title, onClick, spanFull = false) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = `quick-btn ${spanFull ? 'span-full' : ''}`;
      btn.title = title;
      btn.innerHTML = `<span>${escapeHtml(label)}</span>`;
      btn.addEventListener('click', onClick);
      container.appendChild(btn);
    };

    // Standard Select All / Clear All
    addQuickBtn('✓ เปิดทั้งหมด (Select All)', 'เปิดแสดงข้อมูลทุกชนิดที่ตรวจพบในไฟล์', () => {
      state.activeSamples = new Set(state.sampleClasses.map(s => s.id));
      refreshDashboard(true);
    });

    addQuickBtn('✕ ปิดทั้งหมด (Clear All)', 'ซ่อนข้อมูลทุกชนิดชั่วคราว', () => {
      state.activeSamples.clear();
      refreshDashboard(true);
    });

    // Category filter buttons if both Samples and Blanks exist
    const sampleIds = state.sampleClasses.filter(s => !s.isBlank).map(s => s.id);
    const blankIds = state.sampleClasses.filter(s => s.isBlank).map(s => s.id);

    if (sampleIds.length > 0 && blankIds.length > 0) {
      addQuickBtn(`🧪 เฉพาะตัวอย่าง (${sampleIds.length})`, 'แสดงเฉพาะข้อมูลตัวอย่าง (ไม่รวม Blank)', () => {
        state.activeSamples = new Set(sampleIds);
        refreshDashboard(true);
      });
      addQuickBtn(`⚪ เฉพาะ Blank (${blankIds.length})`, 'แสดงเฉพาะข้อมูลอ้างอิง Blank', () => {
        state.activeSamples = new Set(blankIds);
        refreshDashboard(true);
      });
    }

    // Dynamic Group buttons extracted from BlankAssignment or Date column (uses live edited sample labels)
    state.dynamicGroups.forEach(grp => {
      const displayNames = grp.sampleIds.map(id => (state.sampleMap[id] ? state.sampleMap[id].label : id));
      const firstLabel = displayNames[0] || '';
      const lastLabel = displayNames[displayNames.length - 1] || '';
      const rangeHint = displayNames.length > 2 ? `${firstLabel}–${lastLabel}` : displayNames.join(', ');
      const btnLabel = `📅 กลุ่ม ${grp.cleanKey} (${rangeHint})`;
      const btnTitle = `เลือกเฉพาะกลุ่ม ${grp.id}: ${displayNames.join(', ')}`;

      addQuickBtn(btnLabel, btnTitle, () => {
        state.activeSamples = new Set(grp.sampleIds);
        refreshDashboard(true);
      });
    });
  }

  /**
   * Render Dynamic Sample & Blank Toggle Pills inside the Scrollable List, filtered by search input.
   * Supports inline renaming of samples in "ข้อมูลตัวอย่าง (Samples)" with live propagation to Summary Cards & 3D Plot.
   */
  function renderSampleTogglePills() {
    const container = document.getElementById('sampleTogglesContainer');
    container.innerHTML = '';

    const query = state.sampleSearchQuery.trim().toLowerCase();
    const matchedClasses = state.sampleClasses.filter(s => {
      if (!query) return true;
      return (
        s.label.toLowerCase().includes(query) ||
        s.id.toLowerCase().includes(query) ||
        (s.groupLabel && s.groupLabel.toLowerCase().includes(query)) ||
        (s.assignedBlank && s.assignedBlank.toLowerCase().includes(query))
      );
    });

    const samplesList = matchedClasses.filter(s => !s.isBlank);
    const blanksList = matchedClasses.filter(s => s.isBlank);

    const buildCategorySection = (titleText, items, allowRename = false) => {
      if (items.length === 0) return;

      const header = document.createElement('div');
      header.className = 'sample-category-header';
      header.innerHTML = `<span>${titleText}</span><span>${items.length} ชนิด</span>`;
      container.appendChild(header);

      const grid = document.createElement('div');
      grid.className = 'sample-toggles-grid';

      for (const s of items) {
        const pill = document.createElement('div');
        pill.className = `sample-pill ${state.activeSamples.has(s.id) ? 'active' : 'inactive'}`;
        pill.dataset.sample = s.id;
        pill.style.setProperty('--sample-color', s.color);

        const baseSubTag = s.isBlank
          ? 'Reference Blank'
          : (s.assignedBlank ? `−${s.assignedBlank}` : s.groupLabel);
        const subTag = (s.label !== s.id)
          ? `ชีต ${s.id} • ${baseSubTag}`
          : baseSubTag;

        pill.innerHTML = `
          <div class="sample-pill-left">
            <span class="sample-shape-icon" title="รูปทรงจุดข้อมูล 3D: ${escapeHtml(s.shapeName || '')}">${s.shapeIcon || '●'}</span>
            <div class="sample-name-wrap">
              <span class="sample-name" title="${escapeHtml(s.label)}">${escapeHtml(s.label)}</span>
              <span class="sample-group-tag" data-grouptag="${escapeHtml(s.id)}">${escapeHtml(subTag)}</span>
            </div>
          </div>
          <div class="sample-pill-right">
            ${allowRename ? `<button type="button" class="sample-edit-btn" title="แก้ไขชื่อข้อมูลตัวอย่าง (${escapeHtml(s.label)})" aria-label="แก้ไขชื่อตัวอย่าง">✎</button>` : ''}
            <span class="sample-count" data-samplecount="${escapeHtml(s.id)}">n = 0</span>
          </div>
        `;

        // Inline Rename Handler for "ข้อมูลตัวอย่าง (Samples)"
        const startInlineRename = ev => {
          ev.stopPropagation();
          ev.preventDefault();

          const nameWrap = pill.querySelector('.sample-name-wrap');
          const nameSpan = pill.querySelector('.sample-name');
          if (!nameWrap || !nameSpan) return;
          if (nameWrap.querySelector('.sample-rename-input')) return;

          const input = document.createElement('input');
          input.type = 'text';
          input.className = 'sample-rename-input';
          input.value = s.label;
          input.setAttribute('aria-label', `แก้ไขชื่อตัวอย่าง ${s.id}`);

          let committed = false;
          const commitRename = shouldSave => {
            if (committed) return;
            committed = true;

            if (shouldSave) {
              const nextName = input.value.trim();
              if (nextName.length > 0) {
                s.label = nextName;
              }
            }

            renderSampleTogglePills();
            renderDynamicGroupButtons();
            populateTimeSeriesControls();
            refreshDashboard(true);
          };

          ['click', 'mousedown', 'dblclick'].forEach(evtName => {
            input.addEventListener(evtName, e => e.stopPropagation());
          });

          input.addEventListener('keydown', e => {
            e.stopPropagation();
            if (e.key === 'Enter') {
              e.preventDefault();
              commitRename(true);
            } else if (e.key === 'Escape') {
              e.preventDefault();
              commitRename(false);
            }
          });

          input.addEventListener('blur', () => {
            commitRename(true);
          });

          nameSpan.replaceWith(input);
          input.focus();
          input.select();
        };

        if (allowRename) {
          const editBtn = pill.querySelector('.sample-edit-btn');
          if (editBtn) {
            editBtn.addEventListener('click', startInlineRename);
          }
          const nameSpan = pill.querySelector('.sample-name');
          if (nameSpan) {
            nameSpan.addEventListener('dblclick', startInlineRename);
          }
        }

        // Click to toggle both scatter3d and mesh3d layers for this sample
        pill.addEventListener('click', e => {
          if (e.target.closest('.sample-rename-input') || e.target.closest('.sample-edit-btn')) {
            return;
          }
          if (state.activeSamples.has(s.id)) {
            state.activeSamples.delete(s.id);
          } else {
            state.activeSamples.add(s.id);
          }
          refreshDashboard(true);
        });

        // Hover to trigger Focus Highlight
        pill.addEventListener('mouseenter', () => {
          if (state.activeSamples.has(s.id)) {
            applyFocusHighlight(s.id);
          }
        });
        pill.addEventListener('mouseleave', () => {
          applyFocusHighlight(null);
        });

        grid.appendChild(pill);
      }

      container.appendChild(grid);
    };

    buildCategorySection('ข้อมูลตัวอย่าง (Samples)', samplesList, true);
    buildCategorySection('ข้อมูลอ้างอิง (Blank / Reference)', blanksList, false);

    if (matchedClasses.length === 0) {
      container.innerHTML = `
        <div style="text-align:center; padding:14px 8px; color:var(--text-secondary); font-size:12px;">
          ไม่พบชื่อตัวอย่างที่ตรงกับ "${escapeHtml(state.sampleSearchQuery)}"
        </div>
      `;
    }
  }

  function updateSampleTogglesUI(resultByClass) {
    let activeCount = 0;
    for (const s of state.sampleClasses) {
      const isOn = state.activeSamples.has(s.id);
      if (isOn) activeCount++;
    }

    document.querySelectorAll('#sampleTogglesContainer .sample-pill').forEach(pill => {
      const sampleId = pill.dataset.sample;
      const isOn = state.activeSamples.has(sampleId);
      pill.classList.toggle('active', isOn);
      pill.classList.toggle('inactive', !isOn);
    });

    document.querySelectorAll('#sampleTogglesContainer [data-samplecount]').forEach(countEl => {
      const sampleId = countEl.dataset.samplecount;
      const pts = (resultByClass[sampleId] || []).length;
      countEl.textContent = `n = ${pts}`;
    });

    const badge = document.getElementById('activeSampleCountBadge');
    if (badge) badge.textContent = `${activeCount} / ${state.sampleClasses.length}`;
  }

  function renderCyclePills() {
    const container = document.getElementById('cyclePillsContainer');
    container.innerHTML = '';
    const cycles = state.availableCycles;

    const labelEl = document.getElementById('cycleFilterLabel');
    if (labelEl) {
      labelEl.textContent = cycles.length > 0
        ? `รอบการวัด (Loops ${cycles[0]}–${cycles[cycles.length - 1]})`
        : 'รอบการวัด (Cycle Loop)';
    }

    const warmupBtn = document.getElementById('btnExcludeFirstCycle');
    if (warmupBtn && cycles.length > 1) {
      warmupBtn.textContent = `⚡ ตัดรอบแรก (Loops ${cycles[1]}–${cycles[cycles.length - 1]})`;
      const isWarmupActive = !state.activeCycles.has(cycles[0]) && state.activeCycles.size === cycles.length - 1;
      warmupBtn.classList.toggle('active', isWarmupActive);
    }

    const selectAllBtn = document.getElementById('btnSelectAllCycles');
    if (selectAllBtn) {
      const isAllSelected = cycles.length > 0 && state.activeCycles.size === cycles.length;
      selectAllBtn.classList.toggle('active', isAllSelected);
    }

    const btnCycles3to5 = document.getElementById('btnSelectCycles3to5');
    if (btnCycles3to5) {
      const target3to5 = [3, 4, 5].filter(c => cycles.includes(c));
      const is3to5Active =
        target3to5.length > 0 &&
        state.activeCycles.size === target3to5.length &&
        target3to5.every(c => state.activeCycles.has(c));
      btnCycles3to5.classList.toggle('active', is3to5Active);
    }

    container.style.gridTemplateColumns = `repeat(${Math.min(7, Math.max(1, cycles.length))}, 1fr)`;

    cycles.forEach(c => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = `cycle-pill ${state.activeCycles.has(c) ? 'active' : ''}`;
      btn.dataset.cycle = String(c);
      btn.textContent = String(c);
      btn.addEventListener('click', () => {
        if (state.activeCycles.has(c)) {
          state.activeCycles.delete(c);
        } else {
          state.activeCycles.add(c);
        }
        renderCyclePills();
        refreshDashboard(true);
      });
      container.appendChild(btn);
    });
  }

  function renderModePills() {
    const container = document.getElementById('modePillsContainer');
    container.innerHTML = '';
    const modes = state.availableModes;
    container.style.gridTemplateColumns = `repeat(${Math.min(3, Math.max(1, modes.length))}, 1fr)`;

    modes.forEach(mode => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = `mode-pill ${state.activeModes.has(mode) ? 'active' : ''}`;
      btn.dataset.mode = mode;
      btn.innerHTML = `
        <span class="mode-dot"></span>
        <span>${mode}</span>
      `;
      btn.addEventListener('click', () => {
        if (state.activeModes.has(mode)) {
          state.activeModes.delete(mode);
        } else {
          state.activeModes.add(mode);
        }
        btn.classList.toggle('active', state.activeModes.has(mode));
        refreshDashboard(true);
      });
      container.appendChild(btn);
    });
  }

  function updateTopStatusHeader(axX, axY, axZ, totalPoints) {
    document.getElementById('topStatusDataset').textContent = `${state.activeDataset}_*`;

    const pcaVarPill = document.getElementById('topStatusPcaVariancePill');
    const pcaCumVarEl = document.getElementById('topStatusPcaCumVar');

    if (state.axisMode === 'pca' && state.pcaResult) {
      const [v1, v2, v3] = state.pcaResult.explainedVariance;
      const cumVar = state.pcaResult.cumulativeVariance || (v1 + v2 + v3);
      const scopeTag = state.pcaFitScope === 'active' ? 'Active Re-fit' : 'Global';
      const granTag = state.pcaGranularity === 'cycle'
        ? `Cycle ${state.pcaSummaryMethod.toUpperCase()}`
        : 'All Points';

      document.getElementById('topStatusAxes').textContent =
        `PC1 (${v1.toFixed(1)}%) / PC2 (${v2.toFixed(1)}%) / PC3 (${v3.toFixed(1)}%)`;

      if (pcaVarPill && pcaCumVarEl) {
        pcaVarPill.style.display = 'flex';
        pcaCumVarEl.textContent = `${cumVar.toFixed(1)}% (${scopeTag} • ${granTag})`;
      }
    } else {
      document.getElementById('topStatusAxes').textContent =
        `${state.customAxes.x} / ${state.customAxes.y} / ${state.customAxes.z}`;
      if (pcaVarPill) {
        pcaVarPill.style.display = 'none';
      }
    }

    document.getElementById('topStatusBaseline').textContent =
      state.baselineCorrection
        ? (state.perCycleBlank ? '−Blank (Per-Cycle)' : '−Blank (Overall)')
        : 'Raw Signal';
    document.getElementById('topStatusPoints').textContent = totalPoints.toLocaleString();

    const blBadge = document.getElementById('baselineStatusBadge');
    if (blBadge) blBadge.style.display = state.baselineCorrection ? 'inline-block' : 'none';
  }

  /**
   * Upgraded Part 4 Bottom Summary Cards:
   * Displays Count (n), Mean, SD, Median, Min, and Max for X, Y, Z (or PC1, PC2, PC3)
   */
  function renderBottomSummaryCards(resultByClass, getXYZ, axX, axY, axZ, totalPoints) {
    const container = document.getElementById('summaryCardsContainer');
    container.innerHTML = '';

    const shortX = state.axisMode === 'pca' ? 'PC1' : state.customAxes.x;
    const shortY = state.axisMode === 'pca' ? 'PC2' : state.customAxes.y;
    const shortZ = state.axisMode === 'pca' ? 'PC3' : state.customAxes.z;

    let activeClassesCount = 0;

    for (const s of state.sampleClasses) {
      if (!state.activeSamples.has(s.id)) continue;
      const rows = resultByClass[s.id] || [];
      if (rows.length === 0) continue;

      activeClassesCount++;

      const xs = new Array(rows.length);
      const ys = new Array(rows.length);
      const zs = new Array(rows.length);

      for (let i = 0; i < rows.length; i++) {
        const [x, y, z] = getXYZ(rows[i]);
        xs[i] = x;
        ys[i] = y;
        zs[i] = z;
      }

      const stX = computeDescriptiveStats(xs);
      const stY = computeDescriptiveStats(ys);
      const stZ = computeDescriptiveStats(zs);

      const card = document.createElement('div');
      card.className = 'stat-card';
      card.style.setProperty('--card-color', s.color);

      card.innerHTML = `
        <div class="stat-card-head">
          <div class="stat-card-title" title="${escapeHtml(s.label)} (${escapeHtml(s.shapeName || '')})">
            <span class="stat-shape-icon" style="color:${s.color}; text-shadow: 0 0 6px ${s.color};">${s.shapeIcon || '●'}</span>
            <span style="overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${escapeHtml(s.label)}</span>
          </div>
          <span class="stat-card-count">n = ${rows.length}</span>
        </div>
        <table class="stat-metrics-table">
          <thead>
            <tr>
              <th>แกน</th>
              <th>Mean</th>
              <th>SD</th>
              <th>Med</th>
              <th>Min</th>
              <th>Max</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>X (${shortX})</td>
              <td>${stX.mean.toFixed(2)}</td>
              <td>${stX.sd.toFixed(2)}</td>
              <td>${stX.median.toFixed(2)}</td>
              <td>${stX.min.toFixed(2)}</td>
              <td>${stX.max.toFixed(2)}</td>
            </tr>
            <tr>
              <td>Y (${shortY})</td>
              <td>${stY.mean.toFixed(2)}</td>
              <td>${stY.sd.toFixed(2)}</td>
              <td>${stY.median.toFixed(2)}</td>
              <td>${stY.min.toFixed(2)}</td>
              <td>${stY.max.toFixed(2)}</td>
            </tr>
            <tr>
              <td>Z (${shortZ})</td>
              <td>${stZ.mean.toFixed(2)}</td>
              <td>${stZ.sd.toFixed(2)}</td>
              <td>${stZ.median.toFixed(2)}</td>
              <td>${stZ.min.toFixed(2)}</td>
              <td>${stZ.max.toFixed(2)}</td>
            </tr>
          </tbody>
        </table>
      `;

      card.addEventListener('mouseenter', () => applyFocusHighlight(s.id));
      card.addEventListener('mouseleave', () => applyFocusHighlight(null));

      container.appendChild(card);
    }

    if (activeClassesCount === 0) {
      container.innerHTML = `
        <div style="display:flex; align-items:center; justify-content:center; width:100%; color:var(--text-secondary); font-size:12.5px;">
          ไม่มีข้อมูลที่ตรงกับตัวกรองปัจจุบัน กรุณาเปิดชนิดตัวอย่าง (Sample Toggles) หรือเลือกโหมด/รอบการวัดเพิ่มเติม
        </div>
      `;
    }

    document.getElementById('summaryActiveClasses').textContent = activeClassesCount;
    document.getElementById('summaryTotalPoints').textContent = totalPoints.toLocaleString();
    document.getElementById('summaryActiveAxesLabel').textContent = `${shortX}, ${shortY}, ${shortZ}`;
  }

  /**
   * Part 4: Export Processed Data & PCA Chemometrics Report (.xlsx with 4 Sheets)
   */
  function exportProcessedDataWorkbook() {
    const { resultByClass, allFilteredRows } = getProcessedDataByClass();
    const pca = state.pcaResult || computePCA3D(allFilteredRows, state.pcaStandardize);
    const wb = XLSX.utils.book_new();

    // Sheet 1: PCA_Scores
    const scoresHeader = [
      'Sample ID',
      'Display Label',
      'Category / Group',
      'Assigned Blank',
      'Cycle Loop',
      'Mode',
      'Time',
      'PC1',
      'PC2',
      'PC3',
      ...state.sensorCols,
      ...state.extraNumericCols
    ];
    const scoresAoA = [scoresHeader];

    for (const s of state.sampleClasses) {
      if (!state.activeSamples.has(s.id)) continue;
      const rows = resultByClass[s.id] || [];
      for (const r of rows) {
        const [pc1, pc2, pc3] = pca.projectRow(r);
        scoresAoA.push([
          s.id,
          s.label,
          s.groupLabel,
          state.blankAssignment[s.id] || '-',
          r['Cycle Loop'],
          r['Mode'],
          r['Time'],
          Number(pc1.toFixed(5)),
          Number(pc2.toFixed(5)),
          Number(pc3.toFixed(5)),
          ...state.sensorCols.map(c => Number((r[c] || 0).toFixed(5))),
          ...state.extraNumericCols.map(c => Number((r[c] || 0).toFixed(3)))
        ]);
      }
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(scoresAoA), 'PCA_Scores');

    // Sheet 2: PCA_Loadings_Variance
    const [vec1, vec2, vec3] = pca.vectors || [[], [], []];
    const loadingsAoA = [
      ['SCBT316 eNose Chemometrics — PCA Eigenvalues, Explained Variance & Sensor Loadings'],
      ['Protocol', state.protocolInfo.rawText || state.protocolInfo.code],
      ['Active Dataset', `${state.activeDataset}_*`, 'Baseline Correction', state.baselineCorrection ? 'Enabled (-Blank)' : 'Raw Signal'],
      [],
      ['Component', 'Eigenvalue', 'Explained Variance (%)', 'Cumulative Variance (%)']
    ];
    (pca.allExplainedVariance || []).forEach((evPct, idx) => {
      loadingsAoA.push([
        `PC${idx + 1}`,
        Number((pca.allEigenvalues[idx] || 0).toFixed(5)),
        Number(evPct.toFixed(3)),
        Number((pca.allCumulativeVariance[idx] || 0).toFixed(3))
      ]);
    });
    loadingsAoA.push([]);
    loadingsAoA.push(['Sensor', 'Status (Auto-Clean)', 'Mean (μ)', 'SD (σ)', 'PC1 Loading', 'PC2 Loading', 'PC3 Loading']);

    state.sensorCols.forEach(col => {
      const vIdx = (pca.validSensors || []).indexOf(col);
      if (vIdx === -1) {
        loadingsAoA.push([col, 'Excluded (Variance = 0)', 0, 0, 0, 0, 0]);
      } else {
        loadingsAoA.push([
          col,
          'Active',
          Number((pca.means[vIdx] || 0).toFixed(5)),
          Number((pca.stds[vIdx] || 1).toFixed(5)),
          Number((vec1[vIdx] || 0).toFixed(5)),
          Number((vec2[vIdx] || 0).toFixed(5)),
          Number((vec3[vIdx] || 0).toFixed(5))
        ]);
      }
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(loadingsAoA), 'PCA_Loadings_Variance');

    // Sheet 3: Sample_Centroids_Stats
    const statsHeader = [
      'Sample ID',
      'Display Label',
      'Group',
      'Count (n)',
      'Variable / Axis',
      'Mean',
      'SD',
      'Median',
      'Min',
      'Max'
    ];
    const statsAoA = [statsHeader];

    for (const s of state.sampleClasses) {
      if (!state.activeSamples.has(s.id)) continue;
      const rows = resultByClass[s.id] || [];
      if (rows.length === 0) continue;

      const pc1Arr = [], pc2Arr = [], pc3Arr = [];
      for (const r of rows) {
        const [p1, p2, p3] = pca.projectRow(r);
        pc1Arr.push(p1); pc2Arr.push(p2); pc3Arr.push(p3);
      }

      const varList = [
        { name: 'PC1', vals: pc1Arr },
        { name: 'PC2', vals: pc2Arr },
        { name: 'PC3', vals: pc3Arr },
        ...state.sensorCols.map(col => ({ name: col, vals: rows.map(r => r[col] || 0) }))
      ];

      varList.forEach(vItem => {
        const st = computeDescriptiveStats(vItem.vals);
        statsAoA.push([
          s.id,
          s.label,
          s.groupLabel,
          st.count,
          vItem.name,
          Number(st.mean.toFixed(4)),
          Number(st.sd.toFixed(4)),
          Number(st.median.toFixed(4)),
          Number(st.min.toFixed(4)),
          Number(st.max.toFixed(4))
        ]);
      });
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(statsAoA), 'Sample_Centroids_Stats');

    // Sheet 4: Euclidean_Distance_Matrix
    const distData = computeEuclideanDistanceAnalysis(resultByClass, pca);
    const distAoA = [
      [`Pairwise Euclidean Distance Matrix (${state.distanceSpace === 'pca' ? 'PCA Space PC1–PC3' : 'Z-Score Sensor Space'})`],
      ['Sample', ...distData.labels]
    ];
    distData.labels.forEach((lbl, rIdx) => {
      distAoA.push([lbl, ...distData.matrix[rIdx]]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(distAoA), 'Euclidean_Distance_Matrix');

    XLSX.writeFile(wb, `SCBT316-eNose-Analytics-Report-${state.activeDataset}.xlsx`);
  }

  /**
   * Render Dynamic BlankAssignment & Discovered Sheet Overview Modal
   */
  function openBlankAssignmentModal() {
    const modal = document.getElementById('blankModalBackdrop');
    const body = document.getElementById('blankModalBody');
    const badge = document.getElementById('modalWorkbookBadge');

    badge.textContent = state.fileName;
    const baselines = computeBlankBaselines(state.activeDataset);
    const { x, y, z } = state.defaultCustomAxes;

    const assignmentRowsHtml = state.sampleClasses
      .map(s => {
        const assigned = state.blankAssignment[s.id] || '-';
        const bVals = (baselines[assigned] && baselines[assigned].overall) || {};
        const nameDisplay = s.label !== s.id
          ? `${escapeHtml(s.label)} <span style="color:var(--text-secondary); font-size:11px;">(ชีต: ${escapeHtml(s.id)})</span>`
          : escapeHtml(s.label);
        return `
          <tr>
            <td><strong style="color:${s.color}">${s.shapeIcon || '●'} ${nameDisplay}</strong></td>
            <td>${s.isBlank ? 'ข้อมูลอ้างอิง (Blank)' : `ข้อมูลตัวอย่าง (${escapeHtml(s.groupLabel)})`}</td>
            <td><span class="badge">${escapeHtml(assigned)}</span></td>
            <td>${(bVals[x] || 0).toFixed(3)}</td>
            <td>${(bVals[y] || 0).toFixed(3)}</td>
            <td>${(bVals[z] || 0).toFixed(3)}</td>
            <td>${(state.datasets[state.activeDataset][s.id] || []).length} แถว</td>
          </tr>
        `;
      })
      .join('');

    body.innerHTML = `
      <p style="color:var(--text-secondary); margin-bottom:12px; font-size:12.5px;">
        ตารางจับคู่ตัวอย่างกับสัญญาณพื้นหลัง (<strong>Dynamic BlankAssignment</strong>) ที่สกัดได้จากไฟล์ในชุดข้อมูล <code>${state.activeDataset}_*</code>:
      </p>
      <table class="data-table">
        <thead>
          <tr>
            <th>ชื่อตัวอย่าง (Sample)</th>
            <th>ประเภท (Category)</th>
            <th>Blank ที่ใช้หักลบ</th>
            <th>Blank Mean ${x}</th>
            <th>Blank Mean ${y}</th>
            <th>Blank Mean ${z}</th>
            <th>จำนวนแถว (${state.activeDataset}_*)</th>
          </tr>
        </thead>
        <tbody>
          ${assignmentRowsHtml}
        </tbody>
      </table>
    `;

    modal.classList.add('open');
  }

  // --------------------------------------------------------------------------
  // 10. DEMO WORKBOOK GENERATOR (PASSED THROUGH THE 100% DYNAMIC PARSER)
  // --------------------------------------------------------------------------

  function createRng(seed) {
    let t = seed >>> 0;
    return function () {
      t += 0x6D2B79F5;
      let r = Math.imul(t ^ (t >>> 15), 1 | t);
      r ^= r + Math.imul(r ^ (r >>> 7), 61 | r);
      return ((r ^ (r >>> 14)) >>> 0) / 4294967296;
    };
  }

  function randNormal(rng) {
    const u1 = Math.max(1e-7, rng());
    const u2 = rng();
    return Math.sqrt(-2.0 * Math.log(u1)) * Math.cos(2.0 * Math.PI * u2);
  }

  /**
   * Build a sample multi-sheet .xlsx workbook so the dashboard has data on first open
   * and when clicking "⬇ ไฟล์ตัวอย่าง .xlsx". Everything is ingested via ingestWorkbook() dynamically.
   */
  function buildSampleWorkbook() {
    const rng = createRng(20260316);
    const wb = XLSX.utils.book_new();

    // Include ReadMe sheet matching real SCBT316 protocol metadata
    const readMeAoA = [
      ['SCBT316 Electronic Nose — Gas-Solid Adsorption and Sorbent Regeneration'],
      ['Protocol: R10–S120–P60–F200 (Resting 10 s, Sampling 120 s, Purging 60 s, Flow Rate setpoint 200)'],
      ['Active Sensors: 12 / 15 (G4, G9, G15 are inactive / zero-variance channels)']
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(readMeAoA), 'ReadMe');

    const demoSpecs = [
      { name: 'A',           date: '12-Oct', blank: 'Blank_12Oct', c1: 3.85, c2: 2.40, c3: 1.95, scale: 1.00 },
      { name: 'B',           date: '12-Oct', blank: 'Blank_12Oct', c1: 2.65, c2: 4.10, c3: 2.85, scale: 1.15 },
      { name: 'C',           date: '12-Oct', blank: 'Blank_12Oct', c1: 4.90, c2: 3.75, c3: 3.60, scale: 1.35 },
      { name: 'D',           date: '12-Oct', blank: 'Blank_12Oct', c1: 1.95, c2: 2.95, c3: 4.35, scale: 0.95 },
      { name: 'DF',          date: '12-Oct', blank: 'Blank_12Oct', c1: 3.20, c2: 1.65, c3: 3.90, scale: 1.05 },
      { name: 'E',           date: '12-Oct', blank: 'Blank_12Oct', c1: 5.25, c2: 1.95, c3: 2.75, scale: 1.25 },
      { name: 'EF',          date: '14-Oct', blank: 'Blank_14Oct', c1: 4.30, c2: 4.65, c3: 1.85, scale: 1.20 },
      { name: 'F',           date: '14-Oct', blank: 'Blank_14Oct', c1: 2.35, c2: 4.85, c3: 4.20, scale: 1.10 },
      { name: 'G',           date: '14-Oct', blank: 'Blank_14Oct', c1: 5.45, c2: 3.10, c3: 4.80, scale: 1.40 },
      { name: 'H',           date: '14-Oct', blank: 'Blank_14Oct', c1: 3.55, c2: 3.35, c3: 2.95, scale: 0.90 },
      { name: 'Blank_12Oct', date: '12-Oct', blank: 'Blank_12Oct', c1: 1.05, c2: 0.95, c3: 0.88, scale: 0.35 },
      { name: 'Blank_14Oct', date: '14-Oct', blank: 'Blank_14Oct', c1: 1.18, c2: 1.08, c3: 0.96, scale: 0.35 }
    ];

    const blankAssignmentAoA = [
      ['Sample', 'Group Date', 'Assigned Blank'],
      ...demoSpecs
        .filter(d => !isBlankSampleName(d.name))
        .map(d => [d.name, d.date, d.blank])
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(blankAssignmentAoA), 'BlankAssignment');

    const sensorNames = Array.from({ length: 15 }, (_, i) => `G${i + 1}`);
    const headers = ['Time', 'Cycle Loop', 'Mode', ...sensorNames, 'Temp', 'Flow Rate'];

    demoSpecs.forEach((spec, idx) => {
      const samplingAoA = [headers];
      const rawAoA = [headers];
      let elapsedSec = 0;

      // In Sampling_B, add a pre-start row with Cycle Loop = 0 to verify automatic Cycle Loop <= 0 filtering
      if (spec.name === 'B') {
        samplingAoA.push(['00:00:00', 0, 'Sampling', ...sensorNames.map(() => 1.0), 32.0, 200.0]);
      }

      for (let cycle = 1; cycle <= 7; cycle++) {
        const cycleDrift = (cycle - 4) * 0.025;
        const phases = [
          { mode: 'Resting ',  count: 3, rf: 0.08 },
          { mode: 'Sampling ', count: 6, rf: 1.00 },
          { mode: 'Purging',   count: 3, rf: 0.32 }
        ];

        for (const ph of phases) {
          for (let pt = 0; pt < ph.count; pt++) {
            elapsedSec += 5;
            const mins = Math.floor(elapsedSec / 60);
            const secs = elapsedSec % 60;
            const timeStr = `00:${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;

            const sensorVals = sensorNames.map((col, sIdx) => {
              // G4, G9, G15 have constant 0 (Variance = 0) matching SCBT316 dataset
              if (col === 'G4' || col === 'G9' || col === 'G15') return 0;
              const noise = randNormal(rng) * 0.11;
              if (col === 'G13') return Number((1.05 + (spec.c1 - 1.05) * ph.rf + cycleDrift + noise).toFixed(4));
              if (col === 'G14') return Number((0.95 + (spec.c2 - 0.95) * ph.rf + cycleDrift * 0.8 + noise).toFixed(4));
              if (col === 'G10') return Number((0.88 + (spec.c3 - 0.88) * ph.rf - cycleDrift * 0.6 + noise).toFixed(4));
              const angle = (idx + 1) * 0.55 + (sIdx + 1) * 0.42;
              const baseVal = 0.85 + 0.15 * Math.sin(sIdx);
              const delta = spec.scale * (1.1 + 0.8 * Math.sin(angle) + 0.5 * Math.cos(angle * 0.7));
              return Number((baseVal + delta * ph.rf + cycleDrift * 0.5 + noise * 0.85).toFixed(4));
            });

            const tempVal = Number((32.15 + randNormal(rng) * 0.08).toFixed(2));
            const flowVal = Number((200.0 + randNormal(rng) * 0.6).toFixed(1));
            const rowArr = [timeStr, cycle, ph.mode, ...sensorVals, tempVal, flowVal];

            rawAoA.push(rowArr);
            if (ph.mode.trim() === 'Sampling') {
              samplingAoA.push(rowArr);
            }
          }
        }
      }

      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(samplingAoA), `Sampling_${spec.name}`);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rawAoA), `Raw_${spec.name}`);
    });

    return wb;
  }

  // --------------------------------------------------------------------------
  // 11. EVENT LISTENERS & CAMERA SLIDER SYNC
  // --------------------------------------------------------------------------

  function hideWelcomeLandingOverlay() {
    const overlay = document.getElementById('welcomeLandingOverlay');
    if (overlay && !overlay.classList.contains('hidden')) {
      overlay.classList.add('hidden');
      setTimeout(() => {
        resizeActivePlots();
      }, 380);
    }
  }

  function showWelcomeLandingOverlay() {
    const overlay = document.getElementById('welcomeLandingOverlay');
    if (overlay) {
      overlay.classList.remove('hidden');
    }
  }

  function handleExcelFileUpload(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        ingestWorkbook(workbook, file.name, false);
        hideWelcomeLandingOverlay();
      } catch (err) {
        alert(`ไม่สามารถอ่านไฟล์ Excel ได้: ${err.message}`);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function resizeActivePlots() {
    const plotIds = [
      'plot3d',
      'plot2dProjection',
      'plotBiplot2d',
      'plotScree',
      'plotRadar',
      'plotHeatmap',
      'plotTimeSeries',
      'plotQcEnv',
      'plotQcBlank',
      'plotDistanceMatrix'
    ];
    plotIds.forEach(id => {
      const el = document.getElementById(id);
      if (el && el.offsetParent !== null && el.data) {
        try {
          Plotly.Plots.resize(el);
        } catch (_) {}
      }
    });
  }

  function setupEventListeners() {
    // 1. Excel Upload Input & Drag-and-Drop
    const fileInput = document.getElementById('excelFileInput');
    const dropzone = document.getElementById('uploadDropzone');

    fileInput.addEventListener('change', e => {
      if (e.target.files && e.target.files[0]) {
        handleExcelFileUpload(e.target.files[0]);
        e.target.value = ''; // Allow re-uploading same file name after edits
      }
    });

    dropzone.addEventListener('dragover', e => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });
    dropzone.addEventListener('dragleave', () => {
      dropzone.classList.remove('dragover');
    });
    dropzone.addEventListener('drop', e => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        handleExcelFileUpload(e.dataTransfer.files[0]);
      }
    });

    // Download Demo .xlsx file button
    document.getElementById('btnDownloadDemoXlsx').addEventListener('click', () => {
      const wb = buildSampleWorkbook();
      XLSX.writeFile(wb, 'SCBT316-eNose-InomeCafe-Rawdata-Luukmai.xlsx');
    });

    // 1b. Welcome / Landing Overlay Upload, Drag-and-Drop & Action Buttons
    const welcomeFileInput = document.getElementById('welcomeExcelFileInput');
    const welcomeDropzone = document.getElementById('welcomeUploadDropzone');
    const btnWelcomeEnterDemo = document.getElementById('btnWelcomeEnterDemo');
    const btnWelcomeDownloadDemo = document.getElementById('btnWelcomeDownloadDemo');
    const btnReopenWelcomeOverlay = document.getElementById('btnReopenWelcomeOverlay');

    if (welcomeFileInput) {
      welcomeFileInput.addEventListener('change', e => {
        if (e.target.files && e.target.files[0]) {
          handleExcelFileUpload(e.target.files[0]);
          e.target.value = '';
        }
      });
    }

    if (welcomeDropzone) {
      welcomeDropzone.addEventListener('dragover', e => {
        e.preventDefault();
        welcomeDropzone.classList.add('dragover');
      });
      welcomeDropzone.addEventListener('dragleave', () => {
        welcomeDropzone.classList.remove('dragover');
      });
      welcomeDropzone.addEventListener('drop', e => {
        e.preventDefault();
        welcomeDropzone.classList.remove('dragover');
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
          handleExcelFileUpload(e.dataTransfer.files[0]);
        }
      });
    }

    if (btnWelcomeEnterDemo) {
      btnWelcomeEnterDemo.addEventListener('click', () => {
        hideWelcomeLandingOverlay();
      });
    }

    if (btnWelcomeDownloadDemo) {
      btnWelcomeDownloadDemo.addEventListener('click', () => {
        const wb = buildSampleWorkbook();
        XLSX.writeFile(wb, 'SCBT316-eNose-InomeCafe-Rawdata-Luukmai.xlsx');
      });
    }

    if (btnReopenWelcomeOverlay) {
      btnReopenWelcomeOverlay.addEventListener('click', () => {
        showWelcomeLandingOverlay();
      });
    }

    // 1c. Collapsible Slide Toggle for Sidebar Sections 2, 3, and 4
    [2, 3, 4].forEach(secNum => {
      const secEl = document.getElementById(`controlSection${secNum}`);
      const headerEl = document.getElementById(`sectionHeader${secNum}`);
      const btnEl = document.getElementById(`btnToggleSection${secNum}`);
      if (secEl && headerEl && btnEl) {
        const toggleSection = e => {
          e.stopPropagation();
          const isCollapsed = secEl.classList.toggle('collapsed');
          btnEl.setAttribute('aria-expanded', String(!isCollapsed));
        };
        headerEl.addEventListener('click', toggleSection);
        btnEl.addEventListener('click', toggleSection);
      }
    });

    // Inspect BlankAssignment Modal
    document.getElementById('btnInspectSheets').addEventListener('click', openBlankAssignmentModal);
    document.getElementById('btnCloseBlankModal').addEventListener('click', () => {
      document.getElementById('blankModalBackdrop').classList.remove('open');
    });
    document.getElementById('blankModalBackdrop').addEventListener('click', e => {
      if (e.target.id === 'blankModalBackdrop') {
        e.target.classList.remove('open');
      }
    });

    // 2. Dataset Switcher: Sampling_* vs Raw_*
    document.querySelectorAll('#datasetSegmented .seg-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#datasetSegmented .seg-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.activeDataset = btn.dataset.dataset;
        if (state.activeDataset === 'Sampling' && state.availableModes.includes('Sampling') && !state.activeModes.has('Sampling')) {
          state.activeModes.add('Sampling');
          renderModePills();
        }
        updateHeaderAndPipelineUI();
        refreshDashboard(true);
      });
    });

    // 3. Baseline Correction Switch & Per-Cycle Blank Subtraction Toggle
    document.getElementById('toggleBaselineCorrection').addEventListener('change', e => {
      state.baselineCorrection = e.target.checked;
      const perCycleRow = document.getElementById('perCycleBlankRow');
      if (perCycleRow) {
        perCycleRow.style.display = state.baselineCorrection ? 'flex' : 'none';
      }
      refreshDashboard(true);
    });

    const togglePerCycleBlank = document.getElementById('togglePerCycleBlank');
    if (togglePerCycleBlank) {
      togglePerCycleBlank.addEventListener('change', e => {
        state.perCycleBlank = e.target.checked;
        refreshDashboard(true);
      });
    }

    // 4. Axis Mode Switcher: Custom Sensor Axes vs 3D PCA Mode
    document.querySelectorAll('#axisModeSegmented .seg-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#axisModeSegmented .seg-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.axisMode = btn.dataset.axismode;

        document.getElementById('customAxesPanel').style.display = state.axisMode === 'custom' ? 'flex' : 'none';
        document.getElementById('pcaAxesPanel').style.display = state.axisMode === 'pca' ? 'block' : 'none';
        refreshDashboard(true);
      });
    });

    ['X', 'Y', 'Z'].forEach(axisKey => {
      const sel = document.getElementById(`selectAxis${axisKey}`);
      sel.addEventListener('change', () => {
        state.customAxes[axisKey.toLowerCase()] = sel.value;
        refreshDashboard(true);
      });
    });

    document.getElementById('btnResetDefaultAxes').addEventListener('click', () => {
      state.customAxes = { ...state.defaultCustomAxes };
      document.getElementById('selectAxisX').value = state.customAxes.x;
      document.getElementById('selectAxisY').value = state.customAxes.y;
      document.getElementById('selectAxisZ').value = state.customAxes.z;
      refreshDashboard(true);
    });

    document.getElementById('togglePcaStandardize').addEventListener('change', e => {
      state.pcaStandardize = e.target.checked;
      refreshDashboard(true);
    });

    const toggleExcludeBlanksPca = document.getElementById('toggleExcludeBlanksPca');
    if (toggleExcludeBlanksPca) {
      toggleExcludeBlanksPca.addEventListener('change', e => {
        state.excludeBlanksPca = e.target.checked;
        refreshDashboard(true);
      });
    }

    // PCA Granularity Mode: All Time-Points vs Cycle Summary (1 point / cycle)
    document.querySelectorAll('#pcaGranularitySegmented .seg-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#pcaGranularitySegmented .seg-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.pcaGranularity = btn.dataset.granularity;

        const methodRow = document.getElementById('pcaCycleSummaryMethodRow');
        if (methodRow) {
          methodRow.style.display = state.pcaGranularity === 'cycle' ? 'block' : 'none';
        }
        refreshDashboard(true);
      });
    });

    // PCA Cycle Summary Feature Extraction Method (4 methods: mean, max, delta_mean, delta_max)
    document.querySelectorAll('#pcaSummaryMethodSegmented .seg-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#pcaSummaryMethodSegmented .seg-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.pcaSummaryMethod = btn.dataset.summarymethod;
        refreshDashboard(true);
      });
    });

    // PCA Recalculation Scope: Global PCA vs Dynamic Re-fit on Active Samples
    document.querySelectorAll('#pcaFitScopeSegmented .seg-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#pcaFitScopeSegmented .seg-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.pcaFitScope = btn.dataset.fitscope;
        refreshDashboard(true);
      });
    });

    // 5. Dynamic Sample Search Input
    const searchInput = document.getElementById('sampleSearchInput');
    const clearSearchBtn = document.getElementById('btnClearSampleSearch');
    if (searchInput) {
      searchInput.addEventListener('input', e => {
        state.sampleSearchQuery = e.target.value;
        if (clearSearchBtn) {
          clearSearchBtn.style.display = state.sampleSearchQuery ? 'block' : 'none';
        }
        renderSampleTogglePills();
        const { resultByClass } = getProcessedDataByClass();
        updateSampleTogglesUI(resultByClass);
      });
    }
    if (clearSearchBtn && searchInput) {
      clearSearchBtn.addEventListener('click', () => {
        state.sampleSearchQuery = '';
        searchInput.value = '';
        clearSearchBtn.style.display = 'none';
        renderSampleTogglePills();
        const { resultByClass } = getProcessedDataByClass();
        updateSampleTogglesUI(resultByClass);
      });
    }

    // 6. Select All Cycles button, Cycles 3-5 Quick Button & Warm-up Cycle 1 Exclusion Quick Button
    document.getElementById('btnSelectAllCycles').addEventListener('click', () => {
      if (state.activeCycles.size === state.availableCycles.length) {
        state.activeCycles.clear();
      } else {
        state.activeCycles = new Set(state.availableCycles);
      }
      renderCyclePills();
      refreshDashboard(true);
    });

    const btnSelectCycles3to5 = document.getElementById('btnSelectCycles3to5');
    if (btnSelectCycles3to5) {
      btnSelectCycles3to5.addEventListener('click', () => {
        state.activeCycles.clear();
        const targetCycles = [3, 4, 5].filter(c => state.availableCycles.includes(c));
        if (targetCycles.length > 0) {
          targetCycles.forEach(c => state.activeCycles.add(c));
        } else {
          state.availableCycles.slice(2, 5).forEach(c => state.activeCycles.add(c));
        }
        renderCyclePills();
        refreshDashboard(true);
      });
    }

    const btnExcludeFirstCycle = document.getElementById('btnExcludeFirstCycle');
    if (btnExcludeFirstCycle) {
      btnExcludeFirstCycle.addEventListener('click', () => {
        if (state.availableCycles.length > 1) {
          const firstCyc = state.availableCycles[0];
          const isAlreadyWarmupCut =
            !state.activeCycles.has(firstCyc) &&
            state.activeCycles.size === state.availableCycles.length - 1;
          if (isAlreadyWarmupCut) {
            state.activeCycles = new Set(state.availableCycles);
          } else {
            state.activeCycles = new Set(state.availableCycles.slice(1));
          }
          renderCyclePills();
          refreshDashboard(true);
        }
      });
    }

    // 6b. Top Status Bar Horizontal Wheel Scroll & 7-Tab Analysis Suite Toggle
    const topbarStatusScroll = document.getElementById('topbarStatusScroll');
    if (topbarStatusScroll) {
      topbarStatusScroll.addEventListener('wheel', e => {
        if (topbarStatusScroll.scrollWidth > topbarStatusScroll.clientWidth && Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
          e.preventDefault();
          topbarStatusScroll.scrollLeft += e.deltaY;
        }
      }, { passive: false });
    }

    const viewTabsScrollList = document.querySelector('#viewTabsBar .view-tabs-scroll-list');
    if (viewTabsScrollList) {
      viewTabsScrollList.addEventListener('wheel', e => {
        if (viewTabsScrollList.scrollWidth > viewTabsScrollList.clientWidth && Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
          e.preventDefault();
          viewTabsScrollList.scrollLeft += e.deltaY;
        }
      }, { passive: false });
    }

    const viewTabsBar = document.getElementById('viewTabsBar');
    const btnToggleAnalysisTabs = document.getElementById('btnToggleAnalysisTabs');
    const btnToggleAnalysisTabsLabel = document.getElementById('btnToggleAnalysisTabsLabel');
    const btnCloseAnalysisTabs = document.getElementById('btnCloseAnalysisTabs');

    const activateViewTab = targetTab => {
      if (!targetTab) return;
      state.activeViewTab = targetTab;
      if (targetTab !== 'pca3d') {
        state.lastSecondaryTab = targetTab;
      }

      document.querySelectorAll('#viewTabsBar .view-tab-btn').forEach(b => {
        b.classList.toggle('active', b.dataset.viewtab === targetTab);
      });

      document.querySelectorAll('.tab-view-pane').forEach(pane => {
        pane.classList.toggle('active', pane.id === `tabPane-${targetTab}`);
      });

      refreshDashboard(true);
      requestAnimationFrame(resizeActivePlots);
    };

    const setAnalysisSuiteOpen = shouldOpen => {
      state.analysisSuiteOpen = Boolean(shouldOpen);
      if (viewTabsBar) {
        viewTabsBar.classList.toggle('hidden', !state.analysisSuiteOpen);
      }
      if (btnToggleAnalysisTabs) {
        btnToggleAnalysisTabs.classList.toggle('active', state.analysisSuiteOpen);
        btnToggleAnalysisTabs.setAttribute('aria-expanded', String(state.analysisSuiteOpen));
      }
      if (btnToggleAnalysisTabsLabel) {
        btnToggleAnalysisTabsLabel.textContent = state.analysisSuiteOpen
          ? 'กลับหน้ากราฟ 3D หลัก'
          : 'แสดงผลการวิเคราะห์รูปแบบอื่นๆ';
      }

      if (state.analysisSuiteOpen) {
        const nextTab = state.activeViewTab === 'pca3d'
          ? (state.lastSecondaryTab || 'biplot')
          : state.activeViewTab;
        activateViewTab(nextTab);
      } else {
        state.scoreDimension = '3d';
        document.querySelectorAll('#scoreDimensionSegmented .submode-btn').forEach(b => {
          b.classList.toggle('active', b.dataset.scoredim === '3d');
        });
        activateViewTab('pca3d');
      }

      const startTime = performance.now();
      function animateTabBarResize(now) {
        resizeActivePlots();
        if (now - startTime < 340) requestAnimationFrame(animateTabBarResize);
      }
      requestAnimationFrame(animateTabBarResize);
    };

    if (btnToggleAnalysisTabs) {
      btnToggleAnalysisTabs.addEventListener('click', () => {
        setAnalysisSuiteOpen(!state.analysisSuiteOpen);
      });
    }

    if (btnCloseAnalysisTabs) {
      btnCloseAnalysisTabs.addEventListener('click', () => {
        setAnalysisSuiteOpen(false);
      });
    }

    document.querySelectorAll('#viewTabsBar .view-tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const targetTab = btn.dataset.viewtab;
        if (!targetTab) return;
        activateViewTab(targetTab);
      });
    });

    // Tab 1 Sub-Mode Controls (3D Space vs 2D Projection, 3D Biplot Vectors, Centroids)
    document.querySelectorAll('#scoreDimensionSegmented .submode-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#scoreDimensionSegmented .submode-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.scoreDimension = btn.dataset.scoredim || '3d';
        refreshDashboard(true);
        requestAnimationFrame(resizeActivePlots);
      });
    });

    const toggleShow3dBiplot = document.getElementById('toggleShow3dBiplot');
    if (toggleShow3dBiplot) {
      toggleShow3dBiplot.addEventListener('change', e => {
        state.show3dBiplot = e.target.checked;
        refreshDashboard(true);
      });
    }

    const toggleShowCentroids = document.getElementById('toggleShowCentroids');
    if (toggleShowCentroids) {
      toggleShowCentroids.addEventListener('change', e => {
        state.showCentroids = e.target.checked;
        refreshDashboard(true);
      });
    }

    // Tab 2 PCA Biplot Controls: Master Vector Switch, Select All / Clear All Sensors & Right Table Slide Toggle
    const toggleBiplotAllVectors = document.getElementById('toggleBiplotAllVectors');
    if (toggleBiplotAllVectors) {
      toggleBiplotAllVectors.addEventListener('change', e => {
        state.showBiplotVectors = e.target.checked;
        if (state.showBiplotVectors && state.activeBiplotSensors.size === 0) {
          state.activeBiplotSensors = new Set(state.sensorCols);
        }
        refreshDashboard(true);
      });
    }

    const btnBiplotSelectAllSensors = document.getElementById('btnBiplotSelectAllSensors');
    if (btnBiplotSelectAllSensors) {
      btnBiplotSelectAllSensors.addEventListener('click', () => {
        state.showBiplotVectors = true;
        state.activeBiplotSensors = new Set(state.sensorCols);
        if (toggleBiplotAllVectors) toggleBiplotAllVectors.checked = true;
        refreshDashboard(true);
      });
    }

    const btnBiplotClearAllSensors = document.getElementById('btnBiplotClearAllSensors');
    if (btnBiplotClearAllSensors) {
      btnBiplotClearAllSensors.addEventListener('click', () => {
        state.activeBiplotSensors.clear();
        refreshDashboard(true);
      });
    }

    const biplotSidePanel = document.getElementById('biplotSidePanel');
    const btnToggleBiplotTable = document.getElementById('btnToggleBiplotTable');
    if (biplotSidePanel && btnToggleBiplotTable) {
      btnToggleBiplotTable.addEventListener('click', e => {
        e.stopPropagation();
        const isCollapsed = biplotSidePanel.classList.toggle('collapsed');
        btnToggleBiplotTable.setAttribute('aria-expanded', String(!isCollapsed));

        const startTime = performance.now();
        function animateBiplotResize(now) {
          resizeActivePlots();
          if (now - startTime < 380) {
            requestAnimationFrame(animateBiplotResize);
          }
        }
        requestAnimationFrame(animateBiplotResize);
      });
    }

    // Tab 4 Radar Scale Switcher
    document.querySelectorAll('#radarScaleSegmented .submode-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#radarScaleSegmented .submode-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.radarScale = btn.dataset.radarscale || 'zscore';
        refreshDashboard(true);
      });
    });

    // Tab 5 Heatmap Sort Switcher
    document.querySelectorAll('#heatmapSortSegmented .submode-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#heatmapSortSegmented .submode-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.heatmapSort = btn.dataset.heatmapsort || 'pc1';
        refreshDashboard(true);
      });
    });

    // Tab 6 Time-Series Controls
    const tsSourceSelect = document.getElementById('tsSourceSelect');
    const tsModeSelect = document.getElementById('tsModeSelect');
    const tsSensorSelect = document.getElementById('tsSensorSelect');
    const tsSampleSelect = document.getElementById('tsSampleSelect');
    const tsSensorWrap = document.getElementById('tsSensorSelectWrap');
    const tsSampleWrap = document.getElementById('tsSampleSelectWrap');

    if (tsSourceSelect) {
      tsSourceSelect.addEventListener('change', e => {
        state.tsSource = e.target.value;
        renderTimeSeriesTab();
      });
    }
    if (tsModeSelect) {
      tsModeSelect.addEventListener('change', e => {
        state.tsMode = e.target.value;
        if (tsSensorWrap) tsSensorWrap.style.display = state.tsMode === 'compare_samples' ? 'inline-flex' : 'none';
        if (tsSampleWrap) tsSampleWrap.style.display = state.tsMode === 'multi_sensor' ? 'inline-flex' : 'none';
        renderTimeSeriesTab();
      });
    }
    if (tsSensorSelect) {
      tsSensorSelect.addEventListener('change', e => {
        state.tsSensor = e.target.value;
        renderTimeSeriesTab();
      });
    }
    if (tsSampleSelect) {
      tsSampleSelect.addEventListener('change', e => {
        state.tsSample = e.target.value;
        renderTimeSeriesTab();
      });
    }

    // Tab 7 QC Sub-Mode Switcher (Stability vs Distance Matrix & Stats)
    document.querySelectorAll('#qcSubmodeSegmented .submode-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#qcSubmodeSegmented .submode-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.qcSubmode = btn.dataset.qcsubmode || 'stability';

        const isDist = state.qcSubmode === 'distance';
        const qcDualGridWrap = document.getElementById('qcDualGridWrap');
        const qcDistanceMatrixWrap = document.getElementById('qcDistanceMatrixWrap');
        const distanceSpaceSeg = document.getElementById('distanceSpaceSegmented');
        const distancePairsBanner = document.getElementById('distancePairsBanner');

        if (qcDualGridWrap) qcDualGridWrap.style.display = isDist ? 'none' : 'grid';
        if (qcDistanceMatrixWrap) qcDistanceMatrixWrap.style.display = isDist ? 'grid' : 'none';
        if (distanceSpaceSeg) distanceSpaceSeg.style.display = isDist ? 'inline-flex' : 'none';
        if (distancePairsBanner) distancePairsBanner.style.display = isDist ? 'flex' : 'none';

        refreshDashboard(true);
        requestAnimationFrame(resizeActivePlots);
      });
    });

    // Tab 7 Distance Space Switcher
    document.querySelectorAll('#distanceSpaceSegmented .submode-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#distanceSpaceSegmented .submode-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.distanceSpace = btn.dataset.distspace || 'pca';
        refreshDashboard(true);
      });
    });

    // Top KPI Strip Collapse Button
    const btnToggleKpiStrip = document.getElementById('btnToggleKpiStrip');
    const kpiSummaryStrip = document.getElementById('kpiSummaryStrip');
    if (btnToggleKpiStrip && kpiSummaryStrip) {
      btnToggleKpiStrip.addEventListener('click', () => {
        const isCollapsed = kpiSummaryStrip.classList.toggle('collapsed');
        btnToggleKpiStrip.classList.toggle('collapsed', isCollapsed);
        btnToggleKpiStrip.setAttribute('aria-expanded', String(!isCollapsed));
        const startTime = performance.now();
        function animateKpiResize(now) {
          resizeActivePlots();
          if (now - startTime < 380) requestAnimationFrame(animateKpiResize);
        }
        requestAnimationFrame(animateKpiResize);
      });
    }

    // 7. Left Sidebar Slide Toggle (Covers Sections 1–4)
    const sidebarPanel = document.getElementById('sidebarPanel');
    const btnToggleSidebar = document.getElementById('btnToggleSidebar');
    if (btnToggleSidebar && sidebarPanel) {
      btnToggleSidebar.addEventListener('click', () => {
        const isCollapsed = sidebarPanel.classList.toggle('collapsed');
        btnToggleSidebar.setAttribute('aria-expanded', String(!isCollapsed));

        const startTime = performance.now();
        function animateResize(now) {
          resizeActivePlots();
          if (now - startTime < 420) {
            requestAnimationFrame(animateResize);
          }
        }
        requestAnimationFrame(animateResize);
      });
    }

    // 8. Unified 'มุมมองกราฟ 3D' Box Slide Toggle
    const view3dCard = document.getElementById('view3dControlCard');
    const view3dHeader = document.getElementById('view3dCardHeader');
    const btnToggleView3dCard = document.getElementById('btnToggleView3dCard');
    if (view3dCard && view3dHeader && btnToggleView3dCard) {
      const toggleView3dBox = e => {
        e.stopPropagation();
        const isCollapsed = view3dCard.classList.toggle('collapsed');
        btnToggleView3dCard.setAttribute('aria-expanded', String(!isCollapsed));
      };
      btnToggleView3dCard.addEventListener('click', toggleView3dBox);
      view3dHeader.addEventListener('click', toggleView3dBox);
    }

    // 8b. Bottom 'Active Sample Statistics' Box Slide Toggle (Right-side Arrow Button)
    const bottomSummaryBar = document.getElementById('bottomSummaryBar');
    const btnToggleBottomSummary = document.getElementById('btnToggleBottomSummary');
    if (bottomSummaryBar && btnToggleBottomSummary) {
      btnToggleBottomSummary.addEventListener('click', e => {
        e.stopPropagation();
        const isCollapsed = bottomSummaryBar.classList.toggle('collapsed');
        btnToggleBottomSummary.setAttribute('aria-expanded', String(!isCollapsed));

        const startTime = performance.now();
        function animateResize(now) {
          resizeActivePlots();
          if (now - startTime < 420) {
            requestAnimationFrame(animateResize);
          }
        }
        requestAnimationFrame(animateResize);
      });
    }

    // 9. Vertical (Up-Down) & Horizontal (Left-Right) 3D Rotation Sliders
    const sliderVert = document.getElementById('sliderRotateVertical');
    const sliderHoriz = document.getElementById('sliderRotateHorizontal');
    const btnVertUp = document.getElementById('btnRotVertUp');
    const btnVertDown = document.getElementById('btnRotVertDown');
    const btnHorizLeft = document.getElementById('btnRotHorizLeft');
    const btnHorizRight = document.getElementById('btnRotHorizRight');

    if (sliderVert && sliderHoriz) {
      const onSliderRotate = () => {
        stopAutoRotate();
        document.querySelectorAll('.cam-btn[data-cam]').forEach(b => b.classList.remove('active'));
        const azDeg = parseFloat(sliderHoriz.value) || 0;
        const elDeg = parseFloat(sliderVert.value) || 0;
        rotateCameraFromSliders(azDeg, elDeg);
      };

      sliderVert.addEventListener('input', onSliderRotate);
      sliderHoriz.addEventListener('input', onSliderRotate);

      if (btnVertUp) {
        btnVertUp.addEventListener('click', () => {
          sliderVert.value = String(Math.min(85, (parseFloat(sliderVert.value) || 0) + 10));
          onSliderRotate();
        });
      }
      if (btnVertDown) {
        btnVertDown.addEventListener('click', () => {
          sliderVert.value = String(Math.max(-85, (parseFloat(sliderVert.value) || 0) - 10));
          onSliderRotate();
        });
      }
      if (btnHorizLeft) {
        btnHorizLeft.addEventListener('click', () => {
          let nextAz = (parseFloat(sliderHoriz.value) || 0) - 15;
          if (nextAz < -180) nextAz += 360;
          sliderHoriz.value = String(nextAz);
          onSliderRotate();
        });
      }
      if (btnHorizRight) {
        btnHorizRight.addEventListener('click', () => {
          let nextAz = (parseFloat(sliderHoriz.value) || 0) + 15;
          if (nextAz > 180) nextAz -= 360;
          sliderHoriz.value = String(nextAz);
          onSliderRotate();
        });
      }
    }

    // 10. Floating Toolbar: Camera Presets & Auto-Rotate
    document.querySelectorAll('.cam-btn[data-cam]').forEach(btn => {
      btn.addEventListener('click', () => {
        stopAutoRotate();
        document.querySelectorAll('.cam-btn[data-cam]').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const camKey = btn.dataset.cam;
        const preset = CAMERA_PRESETS[camKey] || CAMERA_PRESETS.isometric;
        state.currentCamera = preset;
        syncSlidersFromCamera(preset);
        const plotEl = document.getElementById('plot3d');
        Plotly.relayout(plotEl, { 'scene.camera': preset });
      });
    });

    document.getElementById('btnAutoRotate3D').addEventListener('click', toggleAutoRotate);

    // 11. Floating Toolbar: 3D Area Highlight Controls
    document.getElementById('toggleShowAreaHull').addEventListener('change', e => {
      state.showAreaHull = e.target.checked;
      refreshDashboard(true);
    });

    document.querySelectorAll('#hullTypeSegmented .seg-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#hullTypeSegmented .seg-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.hullType = btn.dataset.hulltype;
        refreshDashboard(true);
      });
    });

    const sliderOpacity = document.getElementById('sliderHullOpacity');
    const opacityValEl = document.getElementById('hullOpacityValue');
    sliderOpacity.addEventListener('input', e => {
      state.hullOpacity = parseFloat(e.target.value);
      opacityValEl.textContent = state.hullOpacity.toFixed(2);

      const plotEl = document.getElementById('plot3d');
      const hullIndices = [];
      state.traceMeta.forEach((m, idx) => {
        if (m.traceType === 'hull') hullIndices.push(idx);
      });
      if (hullIndices.length > 0) {
        Plotly.restyle(plotEl, { opacity: state.hullOpacity }, hullIndices);
      }
    });

    const sliderMarker = document.getElementById('sliderMarkerSize');
    const markerValEl = document.getElementById('markerSizeValue');
    sliderMarker.addEventListener('input', e => {
      state.markerSize = parseFloat(e.target.value);
      markerValEl.textContent = `${state.markerSize.toFixed(1)} px`;

      const plotEl = document.getElementById('plot3d');
      const scatterIndices = [];
      state.traceMeta.forEach((m, idx) => {
        if (m.traceType === 'scatter') scatterIndices.push(idx);
      });
      if (scatterIndices.length > 0) {
        Plotly.restyle(plotEl, { 'marker.size': state.markerSize }, scatterIndices);
      }
    });

    // Export Processed Analytics Workbook (.xlsx)
    const btnExportProcessedXlsx = document.getElementById('btnExportProcessedXlsx');
    if (btnExportProcessedXlsx) {
      btnExportProcessedXlsx.addEventListener('click', exportProcessedDataWorkbook);
    }

    // Export Active Plot PNG
    document.getElementById('btnExportPlotPng').addEventListener('click', () => {
      const tabToPlotId = {
        pca3d: state.scoreDimension === '2d' ? 'plot2dProjection' : 'plot3d',
        biplot: 'plotBiplot2d',
        scree: 'plotScree',
        radar: 'plotRadar',
        heatmap: 'plotHeatmap',
        timeseries: 'plotTimeSeries',
        qc: state.qcSubmode === 'distance' ? 'plotDistanceMatrix' : 'plotQcEnv',
        distance: 'plotDistanceMatrix'
      };
      const targetId = tabToPlotId[state.activeViewTab] || 'plot3d';
      const plotEl = document.getElementById(targetId);
      if (plotEl) {
        Plotly.downloadImage(plotEl, {
          format: 'png',
          width: 1600,
          height: 1000,
          filename: `eNose-${state.activeViewTab}-${state.activeDataset}`
        });
      }
    });

    window.addEventListener('resize', resizeActivePlots);
  }

  function syncSlidersFromCamera(cam) {
    if (!cam || !cam.eye) return;
    const { x, y, z } = cam.eye;
    const r = Math.sqrt(x * x + y * y + z * z) || 2.47;
    const azDeg = Math.round((Math.atan2(y, x) * 180) / Math.PI);
    const clampedRatio = Math.max(-1, Math.min(1, z / r));
    const elDeg = Math.max(-85, Math.min(85, Math.round((Math.asin(clampedRatio) * 180) / Math.PI)));

    const sliderHoriz = document.getElementById('sliderRotateHorizontal');
    const sliderVert = document.getElementById('sliderRotateVertical');
    const horizReadout = document.getElementById('rotHorizontalValue');
    const vertReadout = document.getElementById('rotVerticalValue');

    if (sliderHoriz) sliderHoriz.value = String(azDeg);
    if (sliderVert) sliderVert.value = String(elDeg);
    if (horizReadout) horizReadout.textContent = `${azDeg}°`;
    if (vertReadout) vertReadout.textContent = `${elDeg}°`;
    state.autoRotateAngle = (azDeg * Math.PI) / 180;
  }

  function rotateCameraFromSliders(azimuthDeg, elevationDeg) {
    const curEye = (state.currentCamera && state.currentCamera.eye) || { x: 1.55, y: 1.55, z: 1.15 };
    const radius = Math.max(0.8, Math.sqrt(curEye.x * curEye.x + curEye.y * curEye.y + curEye.z * curEye.z) || 2.47);

    const azRad = (azimuthDeg * Math.PI) / 180;
    const elRad = (Math.max(-85, Math.min(85, elevationDeg)) * Math.PI) / 180;

    const newCam = {
      eye: {
        x: radius * Math.cos(elRad) * Math.cos(azRad),
        y: radius * Math.cos(elRad) * Math.sin(azRad),
        z: radius * Math.sin(elRad)
      },
      up: { x: 0, y: 0, z: 1 },
      center: (state.currentCamera && state.currentCamera.center) || { x: 0, y: 0, z: 0 }
    };

    state.currentCamera = newCam;
    state.autoRotateAngle = azRad;

    const horizReadout = document.getElementById('rotHorizontalValue');
    const vertReadout = document.getElementById('rotVerticalValue');
    if (horizReadout) horizReadout.textContent = `${Math.round(azimuthDeg)}°`;
    if (vertReadout) vertReadout.textContent = `${Math.round(elevationDeg)}°`;

    const plotEl = document.getElementById('plot3d');
    if (plotEl) {
      Plotly.relayout(plotEl, { 'scene.camera': newCam });
    }
  }

  function stopAutoRotate() {
    if (state.autoRotateTimer) {
      clearInterval(state.autoRotateTimer);
      state.autoRotateTimer = null;
      const btn = document.getElementById('btnAutoRotate3D');
      if (btn) btn.textContent = '↻ หมุน 360°';
    }
  }

  function toggleAutoRotate() {
    const btn = document.getElementById('btnAutoRotate3D');
    if (state.autoRotateTimer) {
      stopAutoRotate();
      return;
    }
    btn.textContent = '⏸ หยุดหมุน';
    const radius = 2.15;
    state.autoRotateTimer = setInterval(() => {
      state.autoRotateAngle += 0.035;
      const cam = {
        eye: {
          x: radius * Math.cos(state.autoRotateAngle),
          y: radius * Math.sin(state.autoRotateAngle),
          z: 1.15
        },
        up: { x: 0, y: 0, z: 1 },
        center: { x: 0, y: 0, z: 0 }
      };
      state.currentCamera = cam;
      syncSlidersFromCamera(cam);
      const plotEl = document.getElementById('plot3d');
      if (plotEl) Plotly.relayout(plotEl, { 'scene.camera': cam });
    }, 45);
  }

  async function initApp() {
    setupEventListeners();

    // Load initial demo workbook through the exact same 100% dynamic ingestWorkbook() pipeline
    const demoWb = buildSampleWorkbook();
    ingestWorkbook(demoWb, 'SCBT316-eNose-InomeCafe-Rawdata-Luukmai.xlsx (Demo)', true);

    // Attempt silent auto-load if an Excel file is placed alongside index.html on a local server
    try {
      const resp = await fetch('./SCBT316-eNose-InomeCafe-Rawdata-Luukmai.xlsx');
      if (resp.ok) {
        const buf = await resp.arrayBuffer();
        const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
        ingestWorkbook(wb, 'SCBT316-eNose-InomeCafe-Rawdata-Luukmai.xlsx', false);
      }
    } catch (_) {}
  }

  document.addEventListener('DOMContentLoaded', initApp);
})();
