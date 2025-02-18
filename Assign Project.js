// Global Configuration with corrected sheet names
const CONFIG = {
  SHEETS: {
    ZOMATO: 'Assign Zomato',
    SWIGGY: 'Assign Swiggy',
    SWIGGY_DATA: 'Swiggy Assigned Data',
  },
  COLUMNS: {
    ZOMATO: {
      RID: 'B', // Restaurant ID
      AGENT: 'L', // Agent Name (merged result) - Column 11
      LOOKUP1: 'M', // First lookup result - Column 12
      LOOKUP2: 'N', // Second lookup result - Column 13
      TYPE: 'J', // Restaurant Type
      FORMULA_RANGE: 14, // Total number of columns
    },
    SWIGGY: {
      REST_ID: 'C',
      AGENT: 'H',
      TYPE: 'B',
      LOOKUP1: 'K',
      LOOKUP2: 'L',
      FORMULA_RANGE: 12,
    },
  },
  BATCH_SIZE: 100,
  USER_INFO: {
    name: "MananPP",
    lastUpdated: "2025-01-29 07:02:41",
  },
};

// Utility Functions
function logEvent(action, details) {
  const timestamp = getUTCDateTime();
  const user = CONFIG.USER_INFO.name;
  Logger.log(`[${timestamp}] User: ${user} - Action: ${action} - Details: ${details}`);
}

function getUTCDateTime() {
  return Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd HH:mm:ss");
}

// Update DateTime Function
function updateDateTime() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');
  if (!dashboardSheet) return;

  const now = new Date();
  const timestampCell = dashboardSheet.getRange('A1'); // Example cell for timestamp
  timestampCell.setValue(now.toISOString().replace('T', ' ').substring(0, 19));
}

// Modified onOpen function with wrapper functions




// Wrapper functions for Zomato
function prepareZomatoData() {
  prepareData('zomato');
}

function reapplyZomatoFormulas() {
  reapplyFormulas('zomato');
}

function scanZomatoBlankAgents() {
  scanBlankAgents('zomato');
}

function clearZomatoAssignments() {
  clearAssignments('zomato');
}

// Wrapper functions for Swiggy
function prepareSwiggyData() {
  prepareData('swiggy');
}

function reapplySwiggyFormulas() {
  reapplyFormulas('swiggy');
}

function scanSwiggyBlankAgents() {
  scanBlankAgents('swiggy');
}

function clearSwiggyAssignments() {
  clearAssignments('swiggy');
}

// Prepare Data Function
function prepareData(platform) {
  logEvent('Data Preparation', `Starting ${platform} data preparation`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assignSheet = ss.getSheetByName(CONFIG.SHEETS[platform.toUpperCase()]);
  const dataSheet = ss.getSheetByName(CONFIG.SHEETS.SWIGGY_DATA);

  if (!assignSheet || !dataSheet) {
    const missingSheet = !assignSheet ? CONFIG.SHEETS[platform.toUpperCase()] : CONFIG.SHEETS.SWIGGY_DATA;
    logEvent('Error', `Sheet "${missingSheet}" not found`);
    SpreadsheetApp.getUi().alert(`Sheet "${missingSheet}" not found.`);
    return;
  }

  const lastRow = assignSheet.getLastRow();
  if (lastRow < 2) {
    logEvent('Error', `No data found in ${platform} sheet`);
    SpreadsheetApp.getUi().alert(`No data found in "${platform}" sheet.`);
    return;
  }

  try {
    const columns = CONFIG.COLUMNS[platform.toUpperCase()];
    const range = columns.FORMULA_RANGE;
    const assignRange = assignSheet.getRange(2, 1, lastRow - 1, range);
    const assignData = prepareFormulas(assignRange.getValues(), platform, lastRow);
    assignRange.setValues(assignData);
    SpreadsheetApp.flush();
    logEvent('Success', `${platform} data preparation completed successfully`);
    showAgentSelectionSidebar();
  } catch (e) {
    logEvent('Error', `Data preparation failed: ${e.message}`);
    SpreadsheetApp.getUi().alert('An error occurred while preparing the data. Please check the logs.');
  }
}

// Prepare Formulas Function
function prepareFormulas(data, platform, lastRow) {
  const columns = CONFIG.COLUMNS[platform.toUpperCase()];
  for (let i = 0; i < data.length; i++) {
    const currentRow = i + 2; // Adding 2 because data starts from row 2
    if (platform.toUpperCase() === 'ZOMATO') {
      // Duplicate check formula (Column A)
      data[i][0] = `=COUNTIF(${columns.RID}2:${columns.RID}${lastRow}, ${columns.RID}${currentRow})`;
      // First Lookup (Column M)
      data[i][12] = `=IFERROR(VLOOKUP(${columns.RID}${currentRow}, 'Swiggy Assigned Data'!A:D, 4, FALSE), "")`;
      // Second Lookup (Column N)
      data[i][13] = `=IFERROR(VLOOKUP(${columns.RID}${currentRow}, 'Assign Swiggy'!C:H, 6, FALSE), "")`;
      // Final Agent Name (Column L) - Merged result
      data[i][11] = `=IF(N${currentRow}<>"", N${currentRow}, M${currentRow})`;
    } else {
      // SWIGGY formulas remain unchanged
      data[i][0] = `=COUNTIF(${columns.REST_ID}2:${columns.REST_ID}${lastRow}, ${columns.REST_ID}${currentRow})`;
      data[i][10] = `=IFERROR(VLOOKUP(${columns.REST_ID}${currentRow}, 'Swiggy Assigned Data'!A:D, 4, FALSE), "")`;
      data[i][11] = `=IFERROR(VLOOKUP(${columns.REST_ID}${currentRow}, 'Assign Zomato'!B:L, 11, FALSE), "")`;
      data[i][7] = `=IF(L${currentRow}<>"", L${currentRow}, K${currentRow})`;
    }
  }
  return data;
}


function reapplyFormulas(platform) {
  const timestamp = "2025-01-30 09:15:37";
  logEvent('Formula Reapplication', `Starting ${platform} formula reapplication by MananPP`);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const assignSheet = ss.getSheetByName(CONFIG.SHEETS[platform.toUpperCase()]);

    if (!assignSheet) {
      throw new Error(`Sheet "${CONFIG.SHEETS[platform.toUpperCase()]}" not found`);
    }

    const lastRow = assignSheet.getLastRow();
    if (lastRow < 2) {
      throw new Error(`No data found in ${platform} sheet`);
    }

    const columns = CONFIG.COLUMNS[platform.toUpperCase()];
    const range = columns.FORMULA_RANGE;

    // Get lookup columns based on platform
    let assignmentCol, firstLookupCol, secondLookupCol;
    if (platform.toUpperCase() === 'ZOMATO') {
      assignmentCol = 'L';
      firstLookupCol = 'M';
      secondLookupCol = 'N';
    } else {
      assignmentCol = 'H';
      firstLookupCol = 'K';
      secondLookupCol = 'L';
    }

    // Store current backgrounds
    const currentBackgrounds = assignSheet.getRange(`${assignmentCol}2:${assignmentCol}${lastRow}`).getBackgrounds();

    // Get current assignments and lookup values
    const currentAssignments = assignSheet.getRange(`${assignmentCol}2:${assignmentCol}${lastRow}`).getValues();
    const currentFirstLookups = assignSheet.getRange(`${firstLookupCol}2:${firstLookupCol}${lastRow}`).getValues();
    const currentSecondLookups = assignSheet.getRange(`${secondLookupCol}2:${secondLookupCol}${lastRow}`).getValues();

    // Update formulas
    const assignRange = assignSheet.getRange(2, 1, lastRow - 1, range);
    const assignData = assignRange.getValues();
    const updatedData = prepareFormulas(assignData, platform, lastRow);
    assignRange.setValues(updatedData);
    SpreadsheetApp.flush();

    // Get new lookup values
    const newFirstLookups = assignSheet.getRange(`${firstLookupCol}2:${firstLookupCol}${lastRow}`).getValues();
    const newSecondLookups = assignSheet.getRange(`${secondLookupCol}2:${secondLookupCol}${lastRow}`).getValues();

    // Prepare background colors array
    const newBackgrounds = currentBackgrounds.map(row => row.slice()); // Clone existing backgrounds

    // Compare and highlight differences
    let changesCount = 0;
    for (let i = 0; i < currentAssignments.length; i++) {
      const currentAssignment = currentAssignments[i][0];
      const currentMergeValue = currentSecondLookups[i][0] || currentFirstLookups[i][0];
      const newMergeValue = newSecondLookups[i][0] || newFirstLookups[i][0];

      // If there's a manual assignment (not empty and not a formula)
      if (currentAssignment && !String(currentAssignment).startsWith('=')) {
        // Keep the manual assignment
        assignSheet.getRange(i + 2, assignmentCol.charCodeAt(0) - 64).setValue(currentAssignment);
        
        // Check if lookup values changed
        if (currentMergeValue !== newMergeValue && (currentMergeValue || newMergeValue)) {
          changesCount++;
          newBackgrounds[i][0] = '#ffd7d7';
        }
      }
    }

    // Apply all background colors at once
    assignSheet.getRange(`${assignmentCol}2:${assignmentCol}${lastRow}`).setBackgrounds(newBackgrounds);

    // Final flush and notify
    SpreadsheetApp.flush();
    
    const message = changesCount > 0 ? 
      `Formulas reapplied successfully for ${platform}. ${changesCount} changes in lookup values detected and highlighted.` :
      `Formulas reapplied successfully for ${platform}. No changes in lookup values detected.`;
    
    logEvent('Success', `${platform} formulas reapplied by MananPP at ${timestamp}. Changes: ${changesCount}`);
    SpreadsheetApp.getUi().alert(message);

  } catch (e) {
    logEvent('Error', `Formula reapplication failed: ${e.message}`);
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}

function scanBlankAgents(platform) {
  const timestamp = "2025-01-30 06:33:38";
  logEvent('Blank Agent Scan', `Starting ${platform} blank agent scan at ${timestamp}`);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS[platform.toUpperCase()]);
    
    if (!sheet) {
      throw new Error(`Sheet "${CONFIG.SHEETS[platform.toUpperCase()]}" not found`);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      throw new Error('No data to scan');
    }

    // Get the agent column based on platform
    const agentColumn = platform.toUpperCase() === 'ZOMATO' ? 
      CONFIG.COLUMNS.ZOMATO.AGENT : 
      CONFIG.COLUMNS.SWIGGY.AGENT;
    
    // Convert column letter to number
    const agentColNum = columnToNumber(agentColumn);

    // Get all agent values
    const agentRange = sheet.getRange(2, agentColNum, lastRow - 1, 1);
    const agentValues = agentRange.getValues();

    // Find blank rows
    const blankRows = [];
    agentValues.forEach((row, index) => {
      if (!row[0] || row[0].toString().trim() === '') {
        blankRows.push(index + 2); // Adding 2 because we started from row 2
      }
    });

    // Report results
    if (blankRows.length > 0) {
      const message = `Found ${blankRows.length} blank agent assignments:\nRows: ${blankRows.join(', ')}`;
      SpreadsheetApp.getUi().alert(message);
      logEvent('Scan Results', message);

      // Highlight blank cells
      blankRows.forEach(row => {
        sheet.getRange(row, agentColNum).setBackground('#ffd7d7');
      });

      // Store scan results for agent assignment
      const scanResults = {
        platform: platform,
        timestamp: timestamp,
        blankRows: blankRows,
        agentColumn: agentColumn,
        agentColNum: agentColNum,
        sheetName: CONFIG.SHEETS[platform.toUpperCase()]
      };

      // Store scan results in Script Properties
      PropertiesService.getScriptProperties().setProperty(
        'BLANK_SCAN_RESULTS',
        JSON.stringify(scanResults)
      );

      // Show agent selection sidebar in blank-fill mode
      showAgentSelectionSidebar(platform, true);

    } else {
      const message = 'No blank agent assignments found.';
      SpreadsheetApp.getUi().alert(message);
      logEvent('Scan Results', message);
      
      // Clear any stored scan results
      PropertiesService.getScriptProperties().deleteProperty('BLANK_SCAN_RESULTS');
    }

    // Update last modified
    CONFIG.USER_INFO.lastUpdated = timestamp;
    logEvent('Success', `Blank agent scan completed for ${platform}`);

  } catch (e) {
    // Clear any stored scan results on error
    PropertiesService.getScriptProperties().deleteProperty('BLANK_SCAN_RESULTS');
    logEvent('Error', `Blank agent scan failed: ${e.message}`);
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}

function showAgentSelectionSidebar(platform = 'zomato', isBlankFill = false) {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('AgentSelection')
    .setWidth(600)
    .setHeight(500)
    .setTitle('Agent Assignment Panel');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Agent Assignment Panel');
}

function updateSidebarData() {
  return {
    config: {
      user: CONFIG.USER_INFO.name,
      lastUpdated: "2025-01-28 13:07:38",
      agents: ["Komal", "Prithvi", "Dhruvi", "Mohit", "Kaushik", "Shibani", "Saahil", "Nirali"]
    }
  };
}

function assignAgents(platform, selectedAgents) {
  logEvent('Assignment Start', `Selected Agents for ${platform}: ${selectedAgents}`);

  try {
    if (!validateAgentSelection(selectedAgents)) {
      throw new Error('Invalid agent selection');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const assignSheet = ss.getSheetByName(CONFIG.SHEETS[platform.toUpperCase()]);
    const lastRow = assignSheet.getLastRow();

    if (lastRow < 2) {
      throw new Error(`No data found in ${platform} sheet`);
    }

    const columnRange = platform.toUpperCase() === 'ZOMATO' ? 'A2:N' : 'A2:L';
    const assignData = assignSheet.getRange(`${columnRange}${lastRow}`).getValues();
    
    const {availableAgents, absentAgents} = parseAgentSelection(selectedAgents);

    if (availableAgents.length === 0) {
      throw new Error('No present agents selected');
    }

    // Process and write data
    const processedData = processAgentAssignment(assignData, availableAgents, absentAgents, platform);
    writeDataInBatches(assignSheet, processedData, platform);

    updateLastModified();
    logEvent('Success', `${platform} data assignment completed successfully`);
    SpreadsheetApp.getUi().alert(`${platform} data assigned successfully.`);

  } catch (e) {
    handleAssignmentError(e);
  }
}

function validateAgentSelection(selectedAgents) {
  return selectedAgents && 
         selectedAgents !== 'undefined' && 
         selectedAgents !== '[]';
}

function parseAgentSelection(selectedAgents) {
  const parsed = JSON.parse(selectedAgents);
  return {
    availableAgents: parsed.filter(agent => agent.status === '✔').map(agent => agent.name),
    absentAgents: parsed.filter(agent => agent.status === '✘').map(agent => agent.name)
  };
}

function processAgentAssignment(assignData, availableAgents, absentAgents, platform) {
  // First pass: Initial assignment
  let processedData = assignAgentsToData(assignData, availableAgents, absentAgents, platform);
  
  // Second pass: Balance and adjust assignments
  processedData = balanceAssignments(processedData, availableAgents, absentAgents, platform);
  
  return processedData;
}

function writeDataInBatches(sheet, data, platform) {
  const batchSize = CONFIG.BATCH_SIZE;
  for (let i = 0; i < data.length; i += batchSize) {
    const batch = data.slice(i, Math.min(i + batchSize, data.length));
    const targetColumn = platform.toUpperCase() === 'ZOMATO' ? 11 : 7;
    sheet.getRange(2 + i, 1, batch.length, batch[0].length).setValues(batch);
    logEvent('Progress', `Batch write completed: ${i + batch.length}/${data.length} rows`);
  }
}

function assignAgentsToData(assignData, availableAgents, absentAgents, platform) {
  try {
    const agentLoad = initializeAgentLoad(availableAgents);
    const assignedRIDs = new Map();
    const targetColumn = platform.toUpperCase() === 'ZOMATO' ? 11 : 7;
    const typeColumn = platform.toUpperCase() === 'ZOMATO' ? 9 : 1;

    assignData.forEach((row, index) => {
      try {
        const RID = platform.toUpperCase() === 'ZOMATO' ? row[1] : row[2];
        const restType = row[typeColumn].toLowerCase();
        const count = row[0];
        const currentAgent = row[targetColumn];

        if (shouldRetainAgent(currentAgent, absentAgents)) {
          updateAgentLoad(agentLoad, currentAgent, restType, count);
          assignedRIDs.set(RID, currentAgent);
          return;
        }

        const newAgent = determineAssignedAgent(RID, restType, count, availableAgents, agentLoad, assignedRIDs);
        row[targetColumn] = newAgent;
        updateAgentLoad(agentLoad, newAgent, restType, count);
        assignedRIDs.set(RID, newAgent);
      } catch (e) {
        logEvent('Error', `Failed to process row ${index + 2}: ${e.message}`);
      }
    });

    logEvent('Info', `Agent Load Distribution: ${JSON.stringify(agentLoad)}`);
    return assignData;
  } catch (e) {
    logEvent('Error', `assignAgentsToData failed: ${e.message}`);
    throw e;
  }
}

function initializeAgentLoad(availableAgents) {
  return availableAgents.reduce((acc, agent) => {
    acc[agent] = {
      main: 0,
      virtual: 0,
      count1: 0,
      total: 0
    };
    return acc;
  }, {});
}

function shouldRetainAgent(currentAgent, absentAgents) {
  return currentAgent && !absentAgents.includes(currentAgent);
}

function updateAgentLoad(agentLoad, agent, restType, count) {
  agentLoad[agent].total++;
  if (restType === "main") {
    agentLoad[agent].main++;
  } else if (restType === "virtual" && count === 1) {
    agentLoad[agent].count1++;
  }
}

function determineAssignedAgent(RID, restType, count, availableAgents, agentLoad, assignedRIDs) {
  if (assignedRIDs.has(RID)) {
    return assignedRIDs.get(RID);
  }
  return determineAgent(restType, count, availableAgents, agentLoad);
}

function determineAgent(restType, count, availableAgents, agentLoad) {
  switch(true) {
    case restType === "main":
      return getLeastLoadedAgent(availableAgents, agentLoad, "main");
    case restType === "virtual" && count === 1:
      return getLeastLoadedAgent(availableAgents, agentLoad, "count1");
    default:
      return getModeratelyLoadedAgent(availableAgents, agentLoad);
  }
}

function getLeastLoadedAgent(availableAgents, agentLoad, type) {
  return availableAgents.reduce((leastLoaded, agent) => {
    if (!leastLoaded || agentLoad[agent][type] < agentLoad[leastLoaded][type]) {
      return agent;
    }
    return leastLoaded;
  }, null);
}

function getModeratelyLoadedAgent(availableAgents, agentLoad) {
  return availableAgents.reduce((selected, agent) => {
    if (!selected || agentLoad[agent].total < agentLoad[selected].total) {
      return agent;
    }
    return selected;
  }, null);
}

function balanceAssignments(assignData, availableAgents, absentAgents, platform) {
  try {
    const agentCounts = initializeAgentCounts(availableAgents);
    const targetColumn = platform.toUpperCase() === 'ZOMATO' ? 11 : 7;
    const typeColumn = platform.toUpperCase() === 'ZOMATO' ? 9 : 1;

    // Count current assignments
    countCurrentAssignments(assignData, agentCounts, targetColumn, typeColumn, absentAgents);

    // Balance assignments
    balanceAgentWorkload(assignData, availableAgents, agentCounts, targetColumn, typeColumn);

    // Ensure consistent RID assignments
    ensureConsistentAssignments(assignData, platform);

    return assignData;
  } catch (e) {
    logEvent('Error', `Balance adjustment failed: ${e.message}`);
    throw e;
  }
}

function initializeAgentCounts(availableAgents) {
  return availableAgents.reduce((acc, agent) => {
    acc[agent] = { main: 0, virtual: 0, count1: 0 };
    return acc;
  }, {});
}

function countCurrentAssignments(assignData, agentCounts, targetColumn, typeColumn, absentAgents) {
  assignData.forEach((row, index) => {
    try {
      const agent = row[targetColumn];
      const type = (row[typeColumn] || '').toLowerCase();
      if (agent && !absentAgents.includes(agent)) {
        agentCounts[agent][type]++;
      }
    } catch (e) {
      logEvent('Error', `Failed to count assignment at row ${index + 2}: ${e.message}`);
    }
  });
}

function balanceAgentWorkload(assignData, availableAgents, agentCounts, targetColumn, typeColumn) {
  availableAgents.forEach(agent => {
    availableAgents.forEach(compareAgent => {
      if (agent !== compareAgent) {
        adjustWorkload(assignData, agent, compareAgent, agentCounts, targetColumn, typeColumn);
      }
    });
  });
}

function adjustWorkload(assignData, agent, compareAgent, agentCounts, targetColumn, typeColumn) {
  try {
    let adjustmentsMade = 0;
    while (agentCounts[agent].main > agentCounts[compareAgent].main + 1) {
      if (!adjustSingleAssignment(assignData, agent, compareAgent, agentCounts, targetColumn, typeColumn)) {
        break;
      }
      adjustmentsMade++;
    }

    if (adjustmentsMade > 0) {
      logEvent('Info', `Adjusted ${adjustmentsMade} main assignments from ${agent} to ${compareAgent}`);
    }
  } catch (e) {
    logEvent('Error', `Failed to adjust workload: ${e.message}`);
    throw e;
  }
}

function adjustSingleAssignment(assignData, agent, compareAgent, agentCounts, targetColumn, typeColumn) {
  for (let i = 0; i < assignData.length; i++) {
    if (assignData[i][targetColumn] === agent && 
        assignData[i][typeColumn].toLowerCase() === 'main') {
      assignData[i][targetColumn] = compareAgent;
      agentCounts[agent].main--;
      agentCounts[compareAgent].main++;
      return true;
    }
  }
  return false;
}

function ensureConsistentAssignments(assignData, platform) {
  try {
    const assignments = new Map();
    const reassignments = trackAndUpdateAssignments(assignData, assignments, platform);
    
    if (reassignments > 0) {
      logEvent('Info', `Made ${reassignments} reassignments for consistency`);
    }
  } catch (e) {
    logEvent('Error', `Consistency check failed: ${e.message}`);
    throw e;
  }
}

function trackAndUpdateAssignments(assignData, assignments, platform) {
  let reassignmentCount = 0;
  const idColumn = platform.toUpperCase() === 'ZOMATO' ? 1 : 2;
  const agentColumn = platform.toUpperCase() === 'ZOMATO' ? 11 : 7;

  // First pass: Record assignments
  assignData.forEach(row => {
    const id = row[idColumn];
    const agent = row[agentColumn];
    if (id && agent) {
      assignments.set(id, agent);
    }
  });

  // Second pass: Ensure consistency
  assignData.forEach(row => {
    const id = row[idColumn];
    const agent = row[agentColumn];
    if (assignments.has(id) && assignments.get(id) !== agent) {
      row[agentColumn] = assignments.get(id);
      reassignmentCount++;
    }
  });

  return reassignmentCount;
}

function clearAssignments(platform) {
  const timestamp = "2025-01-30 09:46:41";
  try {
    logEvent('Progress', `Starting to clear ${platform} assignments by MananPP`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS[platform.toUpperCase()]);
    
    if (!sheet) {
      throw new Error(`${platform} sheet not found`);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // Define columns to clear based on platform
      let columnsToClear;
      if (platform.toUpperCase() === 'ZOMATO') {
        // Clear columns A, L, M, N
        columnsToClear = [1, 12, 13, 14];  // A, L, M, N
      } else {
        // Clear columns A, H, K, L
        columnsToClear = [1, 8, 11, 12];   // A, H, K, L
      }

      // Clear content and formatting for each column
      columnsToClear.forEach(col => {
        sheet.getRange(2, col, lastRow - 1, 1)
          .clearContent()
          .clearFormat();
      });

      logEvent('Success', `All ${platform} assignments and related data cleared successfully by MananPP at ${timestamp}`);
      SpreadsheetApp.getUi().alert(`Successfully cleared ${platform} data`);
    }
  } catch (e) {
    logEvent('Error', `Failed to clear ${platform} assignments: ${e.message}`);
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
    throw e;
  }
}

function updateLastModified() {
  try {
    CONFIG.USER_INFO.lastUpdated = getUTCDateTime();
    logEvent('Info', `Last modified info updated to ${CONFIG.USER_INFO.lastUpdated}`);
  } catch (e) {
    logEvent('Error', `Failed to update last modified info: ${e.message}`);
    throw e;
  }
}

function handleAssignmentError(error) {
  const errorMessage = error.message || 'Unknown error occurred';
  logEvent('Error', `Assignment failed: ${errorMessage}`);
  SpreadsheetApp.getUi().alert(`Error: ${errorMessage}`);
}

function getLogs() {
  try {
    return Logger.getLog();
  } catch (e) {
    console.error('Failed to retrieve logs:', e);
    return 'Failed to retrieve logs';
  }
}

// Helper function to validate sheet existence
function validateSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = [
    CONFIG.SHEETS.ZOMATO,
    CONFIG.SHEETS.SWIGGY,
    CONFIG.SHEETS.SWIGGY_DATA
  ];
  
  const missingSheets = requiredSheets.filter(sheetName => !ss.getSheetByName(sheetName));
  
  if (missingSheets.length > 0) {
    logEvent('Error', `Missing sheets: ${missingSheets.join(', ')}`);
    return false;
  }
  return true;
}

// Helper function to validate platform
function validatePlatform(platform) {
  const validPlatforms = ['ZOMATO', 'SWIGGY'];
  const upperPlatform = platform.toUpperCase();
  
  if (!validPlatforms.includes(upperPlatform)) {
    throw new Error(`Invalid platform: ${platform}. Must be either Zomato or Swiggy.`);
  }
  return upperPlatform;
}

// Optional: Add this helper function for cleaner sheet validation
function getSheetOrThrow(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  return sheet;
}

// Optional: Add this helper function for range validation
function validateRange(sheet, startRow, numRows, numCols) {
  if (startRow < 1 || numRows < 1 || numCols < 1) {
    throw new Error('Invalid range parameters');
  }
  
  const maxRows = sheet.getLastRow();
  if (startRow + numRows - 1 > maxRows) {
    throw new Error('Range extends beyond sheet boundaries');
  }
  return true;
}

// Helper function to convert column letter to number
function columnToNumber(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result *= 26;
    result += column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return result;
}

// Helper function to convert number to column letter (for reference)
function numberToColumn(number) {
  let result = '';
  while (number > 0) {
    const remainder = (number - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    number = Math.floor((number - 1) / 26);
  }
  return result;
}

// Optional: Add this function to clear highlighting if needed
function clearHighlighting(sheet, column, startRow, endRow) {
  if (endRow >= startRow) {
    sheet.getRange(startRow, column, endRow - startRow + 1, 1)
         .setBackground(null);
  }
}

// Error handling utility
function wrapErrorHandler(func) {
  return (...args) => {
    try {
      return func(...args);
    } catch (e) {
      logEvent('Error', `Function failed: ${e.message}`);
      throw e;
    }
  };
}
// Dashboard

// Modified onOpen function with wrapper functions
function onOpen() {
  logEvent('Menu Creation', 'Custom menu created');
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 Operations Dashboard')
    .addSubMenu(ui.createMenu('🟠 Zomato')
      .addItem('📊 Prepare Data', 'prepareZomatoData')
      .addItem('🔄 Reapply Formulas', 'reapplyZomatoFormulas')
      .addItem('🔍 Scan Blank Agents', 'scanZomatoBlankAgents')
      .addItem('🗑️ Clear Assignments', 'clearZomatoAssignments'))
    .addSubMenu(ui.createMenu('🟧 Swiggy')
      .addItem('📊 Prepare Data', 'prepareSwiggyData')
      .addItem('🔄 Reapply Formulas', 'reapplySwiggyFormulas')
      .addItem('🔍 Scan Blank Agents', 'scanSwiggyBlankAgents')
      .addItem('🗑️ Clear Assignments', 'clearSwiggyAssignments'))
    .addSeparator()
    .addItem('📈 Open Dashboard', 'showDashboard')
    .addToUi();
}

// Create a time-driven trigger to update the timestamp every minute
function createTimeDrivenTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'updateDateTime') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('updateDateTime')
    .timeBased()
    .everyMinutes(1)
    .create();
}

// Get Dashboard Data Function
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zomatoSheet = ss.getSheetByName(CONFIG.SHEETS.ZOMATO);
  const swiggySheet = ss.getSheetByName(CONFIG.SHEETS.SWIGGY);

  // Get agent distribution
  const agentDistribution = getAgentDistribution();

  // Log the agent distribution for debugging
  console.log("Agent Distribution:", agentDistribution);

  return {
    timestamp: new Date().toISOString(),
    currentUser: Session.getActiveUser().getEmail(),
    platforms: {
      zomato: calculatePlatformCounts(zomatoSheet, CONFIG.COLUMNS.ZOMATO),
      swiggy: calculatePlatformCounts(swiggySheet, CONFIG.COLUMNS.SWIGGY)
    },
    agentDistribution: agentDistribution // Include agentDistribution in the returned data
  };
}

// Calculate Platform Counts Function
function calculatePlatformCounts(sheet, columnConfig) {
  if (!sheet) return { total: 0, main: 0, virtual: 0, virtualCounts: {} };
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // Skip header row

  let counts = {
    total: 0,
    main: 0,
    virtual: 0,
    virtualCounts: {},
  };

  const typeColIndex = columnToIndex(columnConfig.TYPE);
  const agentColIndex = columnToIndex(columnConfig.AGENT);
  const countColIndex = 0; // Column A has count

  rows.forEach((row) => {
    if (row[agentColIndex] && row[agentColIndex].toString().trim()) {
      const count = parseInt(row[countColIndex]) || 0;
      const type = row[typeColIndex]?.toString().trim().toLowerCase();
      if (count > 0) {
        counts.total++;
        if (type === 'main') {
          counts.main++;
        } else if (type === 'virtual') {
          counts.virtual++;
          counts.virtualCounts[count] = (counts.virtualCounts[count] || 0) + 1;
        }
      }
    }
  });

  counts.loadPercentage = Math.round((counts.total / rows.length) * 100);
  return counts;
}

// Show Dashboard Function
function showDashboard() {
  const dashboardData = getDashboardData();
  const template = HtmlService.createTemplateFromFile('Dashboard');
  template.data = dashboardData;
  const html = template.evaluate()
    .setTitle('📊 Assignment Dashboard')
    .setWidth(1000)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Assignment Dashboard');
}

// Utility function to log events (if you're using it)
function logEvent(category, action) {
  console.log(`${category}: ${action}`);
}

// Function to get matrix data for the detailed table
function getMatrixData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zomatoSheet = ss.getSheetByName(CONFIG.SHEETS.ZOMATO);
  const swiggySheet = ss.getSheetByName(CONFIG.SHEETS.SWIGGY);
  
  // Get unique agents from both sheets
  const agents = new Set();
  [zomatoSheet, swiggySheet].forEach(sheet => {
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const agentColIndex = sheet === zomatoSheet ? 
      columnToIndex(CONFIG.COLUMNS.ZOMATO.AGENT) : 
      columnToIndex(CONFIG.COLUMNS.SWIGGY.AGENT);
    
    data.slice(1).forEach(row => {
      if (row[agentColIndex]) agents.add(row[agentColIndex].toString().trim());
    });
  });

  return Array.from(agents).sort().map(agent => {
    const zomatoCounts = getAgentCounts(agent, zomatoSheet, CONFIG.COLUMNS.ZOMATO);
    const swiggyCounts = getAgentCounts(agent, swiggySheet, CONFIG.COLUMNS.SWIGGY);
    
    return {
      agent,
      zomato: zomatoCounts,
      swiggy: swiggyCounts,
      total: {
        main: zomatoCounts.main + swiggyCounts.main,
        virtual: zomatoCounts.virtual + swiggyCounts.virtual,
        total: zomatoCounts.total + swiggyCounts.total
      }
    };
  });
}

function getAgentCounts(agent, sheet, columnConfig) {
  if (!sheet) return { main: 0, virtual: 0, virtualCounts: {}, total: 0 };

  const data = sheet.getDataRange().getValues();
  const agentColIndex = columnToIndex(columnConfig.AGENT);
  const typeColIndex = columnToIndex(columnConfig.TYPE);
  const countColIndex = 0;

  return data.slice(1).reduce((counts, row) => {
    if (row[agentColIndex]?.toString().trim() === agent) {
      const count = parseInt(row[countColIndex]) || 0;
      const type = row[typeColIndex]?.toString().trim().toLowerCase();

      if (count > 0) {
        if (type === 'main') counts.main++;
        else if (type === 'virtual') {
          counts.virtual++;
          counts.virtualCounts[count] = (counts.virtualCounts[count] || 0) + 1;
        }
        counts.total++;
      }
    }
    return counts;
  }, { main: 0, virtual: 0, virtualCounts: {}, total: 0 });
}

function getAgentDistribution() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zomatoSheet = ss.getSheetByName(CONFIG.SHEETS.ZOMATO);
  const swiggySheet = ss.getSheetByName(CONFIG.SHEETS.SWIGGY);

  // Get all agents and their counts
  const agentData = {};

  // Process Zomato data
  processSheetData(zomatoSheet, CONFIG.COLUMNS.ZOMATO, agentData, 'zomato');
  // Process Swiggy data
  processSheetData(swiggySheet, CONFIG.COLUMNS.SWIGGY, agentData, 'swiggy');

  return Object.entries(agentData)
    .map(([agent, data]) => ({
      agent,
      main: data.mainTotal,
      virtual: Object.values(data.virtualCounts).reduce((sum, count) => sum + count, 0),
      grandTotal: data.mainTotal + Object.values(data.virtualCounts).reduce((sum, count) => sum + count, 0)
    }))
    .sort((a, b) => a.agent.localeCompare(b.agent));
}

function processSheetData(sheet, columnConfig, agentData, platform) {
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const agentColIndex = columnToIndex(columnConfig.AGENT);
  const typeColIndex = columnToIndex(columnConfig.TYPE);
  const countColIndex = 0; // Column A has count

  data.slice(1).forEach((row, index) => {
    const agent = row[agentColIndex]?.toString().trim();
    if (!agent) return;

    if (!agentData[agent]) {
      agentData[agent] = {
        mainTotal: 0,
        virtualCounts: {},
        platforms: { zomato: 0, swiggy: 0 }
      };
    }

    const count = parseInt(row[countColIndex]) || 0;
    const type = row[typeColIndex]?.toString().trim().toLowerCase();

    if (count > 0) {
      agentData[agent].platforms[platform]++;
      if (type === 'main') {
        agentData[agent].mainTotal++;
      } else if (type === 'virtual') {
        agentData[agent].virtualCounts[count] = (agentData[agent].virtualCounts[count] || 0) + 1;
      }
    }

    Logger.log(`Processed row ${index + 2}: Agent=${agent}, Type=${type}, Count=${count}`);
  });
}

function getAgentDistribution() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zomatoSheet = ss.getSheetByName(CONFIG.SHEETS.ZOMATO);
  const swiggySheet = ss.getSheetByName(CONFIG.SHEETS.SWIGGY);

  // Get all agents and their counts
  const agentData = {};

  // Process Zomato data
  processSheetData(zomatoSheet, CONFIG.COLUMNS.ZOMATO, agentData, 'zomato');

  // Process Swiggy data
  processSheetData(swiggySheet, CONFIG.COLUMNS.SWIGGY, agentData, 'swiggy');

  return Object.entries(agentData)
    .map(([agent, data]) => ({
      agent,
      main: data.mainTotal || 0,
      virtual: Object.values(data.virtualCounts || {}).reduce((sum, count) => sum + count, 0),
      grandTotal: (data.mainTotal || 0) + Object.values(data.virtualCounts || {}).reduce((sum, count) => sum + count, 0),
      virtualBreakdown: Object.entries(data.virtualCounts || {})
        .sort(([a], [b]) => parseInt(a) - parseInt(b)) // Sort by count
        .map(([count, total]) => ({ count: parseInt(count), total }))
    }))
    .sort((a, b) => a.agent.localeCompare(b.agent));
}

function processSheetData(sheet, columnConfig, agentData, platform) {
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const agentColIndex = columnToIndex(columnConfig.AGENT);
  const typeColIndex = columnToIndex(columnConfig.TYPE);
  const countColIndex = 0; // Column A has count

  data.slice(1).forEach(row => {
    const agent = row[agentColIndex]?.toString().trim();
    if (!agent) return;

    if (!agentData[agent]) {
      agentData[agent] = {
        mainTotal: 0,
        virtualCounts: {}, // Tracks virtual assignments by count
        platforms: { zomato: 0, swiggy: 0 }
      };
    }

    const count = parseInt(row[countColIndex]) || 0;
    const type = row[typeColIndex]?.toString().trim().toLowerCase();

    if (count > 0) {
      agentData[agent].platforms[platform]++;
      if (type === 'main') {
        agentData[agent].mainTotal++;
      } else if (type === 'virtual') {
        agentData[agent].virtualCounts[count] = (agentData[agent].virtualCounts[count] || 0) + 1;
      }
    }
  });
}

// Helper Function: Convert Column Letter to Index
function columnToIndex(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result *= 26;
    result += column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return result - 1;
}

// Refresh Data Function
function refreshData() {
  const refreshButton = document.querySelector('.refresh-button');
  refreshButton.disabled = true;
  refreshButton.innerHTML = '🔄 Refreshing...';

  google.script.run
    .withSuccessHandler(function(newData) {
      console.log("Data received:", newData); // Debugging log
      if (!newData || !newData.agentDistribution) {
        console.error("Agent data is missing or undefined.");
      } else {
        updateDashboardCounts(newData);
      }
      refreshButton.disabled = false;
      refreshButton.innerHTML = '🔄 Refresh Data';
    })
    .withFailureHandler(function(error) {
      console.error('Refresh failed:', error);
      refreshButton.disabled = false;
      refreshButton.innerHTML = '🔄 Refresh Data';
    })
    .getAgentDistribution(); // Call the server-side function to get agent distribution
}

// Update Dashboard Counts Function
function updateDashboardCounts(data) {
  // Update timestamp
  updateDateTime();

  // Update platform counts
  ['zomato', 'swiggy'].forEach(platform => {
    const counts = data.platforms[platform];
    const card = document.querySelector(`.platform-card.${platform}`);

    // Update main stats
    card.querySelector('.stats').innerHTML = `
      Total: ${counts.total || 0}
      Main: ${counts.main || 0}
      Virtual: ${counts.virtual || 0}
    `;

    // Update virtual breakdown
    const virtualCounts = card.querySelector('.virtual-counts');
    virtualCounts.innerHTML = '';
    Object.entries(counts.virtualCounts || {})
      .sort(([a], [b]) => parseInt(a) - parseInt(b))
      .forEach(([count, total]) => {
        if (total > 0) {
          virtualCounts.innerHTML += `
            Count ${count}: ${total}
          `;
        }
      });
  });

  // Update agent assignment matrix
  if (data.agentDistribution && data.agentDistribution.length > 0) {
    updateAgentTable(data.agentDistribution);
  } else {
    console.error("Agent distribution data is missing or empty.");
    const tbody = document.getElementById('agentTableBody');
    tbody.innerHTML = '<tr><td colspan="4">No agent data available</td></tr>';
  }
}

function filterTable() {
  const filterValue = document.getElementById('platformFilter').value;
  const rows = document.querySelectorAll('#agentTableBody tr');
  rows.forEach(row => {
    const platform = row.getAttribute('data-platform');
    if (filterValue === 'all' || platform === filterValue) {
      row.style.display = ''; // Show the row
    } else {
      row.style.display = 'none'; // Hide the row
    }
  });
}

// Update Agent Table Function
function updateAgentTable(data) {
  console.log("Agent Data:", data); // Debugging log

  const tbody = document.getElementById('agentTableBody');
  if (!tbody) {
    console.error("Element with ID 'agentTableBody' not found.");
    return;
  }

  tbody.innerHTML = ''; // Clear existing rows

  if (!data || data.length === 0) {
    console.error("No agent data available.");
    tbody.innerHTML = '<tr><td colspan="4">No agent data available</td></tr>';
    return;
  }

  data.forEach(agentData => {
    const row = document.createElement('tr');
    row.innerHTML = `
      <td>${agentData.agent}</td>
      <td>${agentData.main}</td>
      <td>${agentData.virtual}</td>
      <td>${agentData.grandTotal}</td>
    `;
    tbody.appendChild(row);
  });
}

// Initial update
//updateDateTime();
// Update every second
//setInterval(updateDateTime, 1000);
