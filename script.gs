function updateDailyUsage() {
  const props = PropertiesService.getScriptProperties();
  const ADMIN_KEY = props.getProperty('OPENAI_ADMIN_KEY');
  if (!ADMIN_KEY) throw new Error('Set OPENAI_ADMIN_KEY in Script Properties.');

  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone() || 'Africa/Cairo';
  
  // Optional: You can specify specific projects to track
  // Leave empty to track ALL projects, or add specific IDs here
  const TRACK_SPECIFIC_PROJECTS = [
    // Add your specific project IDs here if you want to limit tracking
   
  ];

  // Create summary sheets
  const summaryUsageSheet = ss.getSheetByName('Summary Usage') || ss.insertSheet('Summary Usage');
  if (summaryUsageSheet.getLastRow() === 0) {
    summaryUsageSheet.appendRow([
      'Date', 'Project ID', 'Project Name', 'Model',
      'Total Tokens (Input+Output)', 'Input tokens', 'Output tokens',
      'Cached tokens', 'Requests', 'Cost (USD)'
    ]);
  }

  const summaryCostsSheet = ss.getSheetByName('Summary Costs') || ss.insertSheet('Summary Costs');
  if (summaryCostsSheet.getLastRow() === 0) {
    summaryCostsSheet.appendRow([
      'Date', 'Project ID', 'Project Name', 'Line Item', 'Amount (USD)', 'Cost (+30%)'
    ]);
  } else {
    summaryCostsSheet.getRange(1, 6).setValue('Cost (+30%)');
  }
  summaryCostsSheet.hideColumn(summaryCostsSheet.getRange(1, 5));

  // Use calendar days (yesterday only for usage, 2 days for costs)
  const endDate = new Date();
  endDate.setHours(23, 59, 59, 999); // End of today
  
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - 1); // Yesterday
  startDate.setHours(0, 0, 0, 0); // Start of yesterday
  
  // Time window for usage: yesterday only
  const usageStart = Math.floor(startDate.getTime() / 1000);
  const usageEnd = Math.floor(startDate.getTime() / 1000) + 86400 - 1; // End of yesterday (23:59:59)
  
  // Time window for costs: today and yesterday (2 days)
  const costsStart = Math.floor(startDate.getTime() / 1000);
  const costsEnd = Math.floor(endDate.getTime() / 1000);
  
  console.log('📅 Usage time window (yesterday only):', new Date(usageStart * 1000).toISOString(), 'to', new Date(usageEnd * 1000).toISOString());
  console.log('💰 Costs time window (2 days):', new Date(costsStart * 1000).toISOString(), 'to', new Date(costsEnd * 1000).toISOString());

  const headers = { 
    'Authorization': 'Bearer ' + ADMIN_KEY
    // Remove OpenAI-Organization header - let the API key determine the org
  };

  // --- Fetch project list for ID → name mapping ---
  const projectsResp = UrlFetchApp.fetch('https://api.openai.com/v1/organization/projects', {
    method: 'get',
    headers,
    muteHttpExceptions: true
  });
  
  console.log('Projects response code:', projectsResp.getResponseCode());
  console.log('Projects response:', projectsResp.getContentText());
  
  const projectsJson = safeJson(projectsResp);
  if (!projectsJson) throw new Error('Failed to fetch project list');
  
  const projectNameMap = {};
  (projectsJson.data || []).forEach(p => {
    projectNameMap[p.id] = p.name || '';
  });

  function getProjectName(id) {
    return projectNameMap[id] || '';
  }

  function cleanModelName(model) {
    if (!model) return '';
    return model.replace(/-\d{4}-\d{2}-\d{2}$/, '');
  }

  // -------- Usage API (Corrected endpoint) --------
  const usageParams = {
    start_time: usageStart,
    end_time: usageEnd,
    bucket_width: '1d',
    limit: 31, // Maximum allowed for daily bucket width
    group_by: ['project_id', 'model']
  };
  
  // If specific projects are specified, filter by them
  if (TRACK_SPECIFIC_PROJECTS.length > 0) {
    usageParams.project_ids = TRACK_SPECIFIC_PROJECTS;
  }

  const usageUrl = 'https://api.openai.com/v1/organization/usage/completions?' + encodeParams(usageParams);
  console.log('Usage URL:', usageUrl);
  
  const usageResp = UrlFetchApp.fetch(usageUrl, { 
    method: 'get', 
    headers, 
    muteHttpExceptions: true 
  });
  
  console.log('Usage response code:', usageResp.getResponseCode());
  console.log('Usage response:', usageResp.getContentText());
  
  const usageJson = safeJson(usageResp);
  if (!usageJson) throw new Error('Failed to fetch usage data');

  const usageBuckets = usageJson.data || [];
  const usageRows = [];
  
  usageBuckets.forEach(bucket => {
    const dateStr = Utilities.formatDate(new Date(bucket.start_time * 1000), tz, 'yyyy-MM-dd');
    (bucket.results || []).forEach(r => {
      if (!(r.input_tokens || r.output_tokens || r.input_cached_tokens || r.num_model_requests)) return;

      const totalInputTokens = Number(r.input_tokens || 0) + Number(r.input_cached_tokens || 0);
      const totalTokens = totalInputTokens + Number(r.output_tokens || 0);

      usageRows.push([
        dateStr,
        r.project_id || '',
        getProjectName(r.project_id || ''),
        cleanModelName(r.model),
        totalTokens,
        Number(r.input_tokens || 0),
        Number(r.output_tokens || 0),
        Number(r.input_cached_tokens || 0),
        Number(r.num_model_requests || 0),
        '' // Cost will be calculated separately
      ]);
    });
  });

  // Group new data by date+project+model and combine duplicates
  const combinedUsageMap = new Map();
  if (usageRows.length) {
    usageRows.forEach(row => {
      const key = row[0] + '|' + row[1] + '|' + row[3]; // date|projectId|model
      
      if (combinedUsageMap.has(key)) {
        // Combine with existing entry
        const existing = combinedUsageMap.get(key);
        existing[4] += Number(row[4]) || 0; // Total tokens
        existing[5] += Number(row[5]) || 0; // Input tokens
        existing[6] += Number(row[6]) || 0; // Output tokens
        existing[7] += Number(row[7]) || 0; // Cached tokens
        existing[8] += Number(row[8]) || 0; // Requests
        // Cost column (9) will be calculated later
      } else {
        // First entry for this combination
        combinedUsageMap.set(key, [...row]);
      }
    });
  }
  
  // Add data to summary sheets (combine duplicates and replace existing data)
  if (usageRows.length) {
    // Check if we already have data for these date+project+model combinations
    const existingData = summaryUsageSheet.getDataRange().getValues();
    const headers = existingData[0];
    const existingRows = existingData.slice(1);
    
    // Create a map of existing data by date+project+model
    const existingMap = new Map();
    existingRows.forEach(row => {
      const key = row[0] + '|' + row[1] + '|' + row[3]; // date|projectId|model
      existingMap.set(key, true);
    });
    
    // Convert combined map back to array
    const combinedUsageRows = Array.from(combinedUsageMap.values());
    
    // Find rows that need to be replaced (same date+project+model)
    const rowsToReplace = combinedUsageRows.filter(row => {
      const key = row[0] + '|' + row[1] + '|' + row[3]; // date|projectId|model
      return existingMap.has(key);
    });
    
    // Find completely new rows
    const newRows = combinedUsageRows.filter(row => {
      const key = row[0] + '|' + row[1] + '|' + row[3]; // date|projectId|model
      return !existingMap.has(key);
    });
    
    // Remove existing rows that will be replaced
    if (rowsToReplace.length > 0) {
      rowsToReplace.forEach(rowToReplace => {
        const key = rowToReplace[0] + '|' + rowToReplace[1] + '|' + rowToReplace[3];
        // Find and remove the existing row
        for (let i = existingRows.length - 1; i >= 0; i--) {
          const existingKey = existingRows[i][0] + '|' + existingRows[i][1] + '|' + existingRows[i][3];
          if (existingKey === key) {
            // Remove the existing row (add 2 because of 0-based index and header row)
            summaryUsageSheet.deleteRow(i + 2);
            break;
          }
        }
      });
      console.log('🔄 Replaced ' + rowsToReplace.length + ' existing usage records');
    }
    
    // Add all combined data (including replacements)
    if (combinedUsageRows.length > 0) {
      summaryUsageSheet.getRange(summaryUsageSheet.getLastRow() + 1, 1, combinedUsageRows.length, combinedUsageRows[0].length).setValues(combinedUsageRows);
      console.log('✅ Added ' + combinedUsageRows.length + ' combined usage records (' + newRows.length + ' new, ' + rowsToReplace.length + ' replaced)');
      console.log('📊 Combined ' + usageRows.length + ' original rows into ' + combinedUsageRows.length + ' unique entries');
    }
  }

  // -------- Costs API (Corrected endpoint) --------
  const costsParams = {
    start_time: costsStart,
    end_time: costsEnd,
    bucket_width: '1d',
    limit: 31, // Maximum allowed for daily bucket width
    group_by: ['project_id', 'line_item']
  };
  
  // If specific projects are specified, filter by them
  if (TRACK_SPECIFIC_PROJECTS.length > 0) {
    costsParams.project_ids = TRACK_SPECIFIC_PROJECTS;
  }

  const costsUrl = 'https://api.openai.com/v1/organization/costs?' + encodeParams(costsParams);
  console.log('Costs URL:', costsUrl);
  
  const costsResp = UrlFetchApp.fetch(costsUrl, { 
    method: 'get', 
    headers, 
    muteHttpExceptions: true 
  });
  
  console.log('Costs response code:', costsResp.getResponseCode());
  console.log('Costs response:', costsResp.getContentText());
  
  const costsJson = safeJson(costsResp);
  if (!costsJson) throw new Error('Failed to fetch costs data');

  const costBuckets = costsJson.data || [];
  const costRows = [];
  
  costBuckets.forEach(bucket => {
    const dateStr = Utilities.formatDate(new Date(bucket.start_time * 1000), tz, 'yyyy-MM-dd');
    (bucket.results || []).forEach(r => {
      if (!r.amount || !r.amount.value) return;
      
      const amount = Number(r.amount.value);
      const amountPlus30 = Math.round(amount * 1.3 * 100) / 100;
      
      costRows.push([
        dateStr,
        r.project_id || '',
        getProjectName(r.project_id || ''),
        r.line_item || '',
        amount,        // hidden
        amountPlus30   // shown as Cost (+30%)
      ]);
    });
  });

  // Add cost data to summary sheets (replace existing data for the same date+project+lineItem combinations)
  if (costRows.length) {
    // Check if we already have data for these date+project+lineItem combinations
    const existingCostData = summaryCostsSheet.getDataRange().getValues();
    const costHeaders = existingCostData[0];
    const existingCostRows = existingCostData.slice(1);
    
    // Create a map of existing cost data by date+project+lineItem
    const existingCostMap = new Map();
    existingCostRows.forEach(row => {
      const key = row[0] + '|' + row[1] + '|' + row[3]; // date|projectId|lineItem
      existingCostMap.set(key, true);
    });
    
    // Find rows that need to be replaced (same date+project+lineItem)
    const costRowsToReplace = costRows.filter(row => {
      const key = row[0] + '|' + row[1] + '|' + row[3]; // date|projectId|lineItem
      return existingCostMap.has(key);
    });
    
    // Find completely new rows
    const newCostRows = costRows.filter(row => {
      const key = row[0] + '|' + row[1] + '|' + row[3]; // date|projectId|lineItem
      return !existingCostMap.has(key);
    });
    
    // Remove existing rows that will be replaced
    if (costRowsToReplace.length > 0) {
      costRowsToReplace.forEach(rowToReplace => {
        const key = rowToReplace[0] + '|' + rowToReplace[1] + '|' + rowToReplace[3];
        // Find and remove the existing row
        for (let i = existingCostRows.length - 1; i >= 0; i--) {
          const existingKey = existingCostRows[i][0] + '|' + existingCostRows[i][1] + '|' + existingCostRows[i][3];
          if (existingKey === key) {
            // Remove the existing row (add 2 because of 0-based index and header row)
            summaryCostsSheet.deleteRow(i + 2);
            break;
          }
        }
      });
      console.log('🔄 Replaced ' + costRowsToReplace.length + ' existing cost records');
    }
    
    // Add all new data (including replacements)
    if (costRows.length > 0) {
      summaryCostsSheet.getRange(summaryCostsSheet.getLastRow() + 1, 1, costRows.length, costRows[0].length).setValues(costRows);
      console.log('✅ Added ' + costRows.length + ' cost records (' + newCostRows.length + ' new, ' + costRowsToReplace.length + ' replaced)');
    }
  }

  // Create individual project sheets
  try {
    // Use combined usage rows for project sheets to maintain consistency
    const combinedUsageRows = Array.from(combinedUsageMap.values());
    createProjectSheets(ss, combinedUsageRows, costRows, tz);
    console.log('✅ Project sheets created successfully');
  } catch (e) {
    console.error('❌ Error creating project sheets:', e.toString());
    // Continue execution even if project sheets fail
  }
  
  // Update summary usage sheet with calculated costs
  try {
    updateUsageWithCosts(summaryUsageSheet, summaryCostsSheet);
    console.log('✅ Summary sheets updated with costs');
  } catch (e) {
    console.error('❌ Error updating summary sheets:', e.toString());
  }
}

function createProjectSheets(ss, usageRows, costRows, tz) {
  // Group data by project
  const projectUsage = {};
  const projectCosts = {};
  
  // Group usage data by project
  usageRows.forEach(row => {
    const projectId = row[1];
    const projectName = row[2];
    if (!projectUsage[projectId]) {
      projectUsage[projectId] = {
        name: projectName,
        data: []
      };
    }
    projectUsage[projectId].data.push(row);
  });
  
  // Group cost data by project
  costRows.forEach(row => {
    const projectId = row[1];
    const projectName = row[2];
    if (!projectCosts[projectId]) {
      projectCosts[projectId] = {
        name: projectName,
        data: []
      };
    }
    projectCosts[projectId].data.push(row);
  });
  
  console.log('📊 Projects with usage data:', Object.keys(projectUsage).length);
  console.log('💰 Projects with cost data:', Object.keys(projectCosts).length);
  
  // Log project names for debugging
  Object.keys(projectUsage).forEach(projectId => {
    console.log('  - ' + projectUsage[projectId].name + ' (' + projectId + ')');
  });
  
  // Create or update sheets for each project
  Object.keys(projectUsage).forEach(projectId => {
    try {
      const projectName = projectUsage[projectId].name;
      const safeSheetName = getSafeSheetName(projectName);
      
      console.log('📝 Processing project: ' + projectName + ' (' + projectId + ')');
      
      // Create or get project usage sheet
      let projectUsageSheet = ss.getSheetByName(safeSheetName + ' - Usage');
      if (!projectUsageSheet) {
        projectUsageSheet = ss.insertSheet(safeSheetName + ' - Usage');
        projectUsageSheet.appendRow([
          'Date', 'Project ID', 'Project Name', 'Model',
          'Total Tokens (Input+Output)', 'Input tokens', 'Output tokens',
          'Cached tokens', 'Requests', 'Cost (USD)'
        ]);
      }
      
      // Add usage data
      if (projectUsage[projectId].data && projectUsage[projectId].data.length > 0) {
        projectUsageSheet.getRange(projectUsageSheet.getLastRow() + 1, 1, 
          projectUsage[projectId].data.length, projectUsage[projectId].data[0].length)
          .setValues(projectUsage[projectId].data);
        console.log('  ✅ Added ' + projectUsage[projectId].data.length + ' usage records');
      }
      
      // Create or get project costs sheet
      let projectCostsSheet = ss.getSheetByName(safeSheetName + ' - Costs');
      if (!projectCostsSheet) {
        projectCostsSheet = ss.insertSheet(safeSheetName + ' - Costs');
        projectCostsSheet.appendRow([
          'Date', 'Project ID', 'Project Name', 'Line Item', 'Amount (USD)', 'Cost (+30%)'
        ]);
        projectCostsSheet.hideColumn(projectCostsSheet.getRange(1, 5));
      }
      
      // Add cost data
      if (projectCosts[projectId] && projectCosts[projectId].data && projectCosts[projectId].data.length > 0) {
        projectCostsSheet.getRange(projectCostsSheet.getLastRow() + 1, 1, 
          projectCosts[projectId].data.length, projectCosts[projectId].data[0].length)
          .setValues(projectCosts[projectId].data);
        console.log('  ✅ Added ' + projectCosts[projectId].data.length + ' cost records');
      } else {
        console.log('  ⚠  No cost data for project ' + projectName);
      }
      
      // Update project usage sheet with calculated costs
      updateUsageWithCosts(projectUsageSheet, projectCostsSheet);
      console.log('  ✅ Updated costs for project ' + projectName);
      
    } catch (e) {
      console.error('❌ Error processing project ' + projectId + ':', e.toString());
    }
  });
  

}

function getSafeSheetName(name) {
  // Google Sheets has a 31 character limit for sheet names
  // Remove special characters and limit length
  return name.replace(/[\\\/\*\?\:\[\]]/g, '').substring(0, 25);
}



function updateUsageWithCosts(usageSheet, costsSheet) {
  // Get all usage data
  const usageData = usageSheet.getDataRange().getValues();
  const headers = usageData[0];
  const usageRows = usageData.slice(1);
  
  // Get all costs data
  const costsData = costsSheet.getDataRange().getValues();
  const costRows = costsData.slice(1);
  
  // Create cost lookup by date and project
  const costLookup = {};
  costRows.forEach(row => {
    const date = row[0];
    const projectId = row[1];
    const lineItem = row[3];
    const cost = row[5]; // Cost (+30%)
    
    if (!costLookup[date]) costLookup[date] = {};
    if (!costLookup[date][projectId]) costLookup[date][projectId] = {};
    costLookup[date][projectId][lineItem] = cost;
  });
  
  // Update usage rows with costs
  usageRows.forEach((row, index) => {
    const date = row[0];
    const projectId = row[1];
    const model = row[3];
    
    // Find matching cost for this date, project, and model
    let cost = 0;
    if (costLookup[date] && costLookup[date][projectId]) {
      // Try to find exact model match first
      if (costLookup[date][projectId][model]) {
        cost = costLookup[date][projectId][model];
      } else {
        // If no exact match, sum all costs for this project/date
        Object.values(costLookup[date][projectId]).forEach(c => {
          cost += Number(c);
        });
      }
    }
    
    // Update the cost column (column 10, index 9)
    usageSheet.getRange(index + 2, 10).setValue(cost);
  });
}

// --- helpers ---
function encodeParams(params) {
  const parts = [];
  Object.keys(params).forEach(k => {
    const v = params[k];
    if (Array.isArray(v)) {
      v.forEach(item => parts.push(encodeURIComponent(k) + '=' + encodeURIComponent(item)));
    } else {
      parts.push(encodeURIComponent(k) + '=' + encodeURIComponent(v));
    }
  });
  return parts.join('&');
}

function safeJson(resp) {
  try {
    const text = resp.getContentText();
    const code = resp.getResponseCode();
    console.log('Response code:', code);
    console.log('Response text:', text);
    
    if (code >= 200 && code < 300) {
      if (!text.trim()) {
        console.log('Empty response body');
        return null;
      }
      return JSON.parse(text);
    }
    console.error('HTTP ' + code + ': ' + text);
    return null;
  } catch (e) {
    console.error('JSON parse error: ' + e);
    return null;
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('OpenAI Usage')
    .addItem('Run now', 'updateDailyUsage')
    .addItem('Get yesterday data', 'getYesterdayData')
    .addItem('Install daily trigger', 'installDailyTrigger')
    .addItem('Test API connection', 'testAPIConnection')
    .addItem('Get Organization Info', 'getOrganizationInfo')

    .addItem('Show Tracked Projects', 'showTrackedProjects')
    .addItem('Clean Duplicate Data', 'cleanDuplicateData')
    .addToUi();
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'updateDailyUsage') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('updateDailyUsage').timeBased().everyDays(1).atHour(0).create(); // ✅ midnight
}

function testAPIConnection() {
  try {
    const props = PropertiesService.getScriptProperties();
    const ADMIN_KEY = props.getProperty('OPENAI_ADMIN_KEY');
    if (!ADMIN_KEY) {
      SpreadsheetApp.getUi().alert('Error: OPENAI_ADMIN_KEY not set in Script Properties');
      return;
    }
    
    const headers = { 
      'Authorization': 'Bearer ' + ADMIN_KEY
    };
    
    // Test projects API
    const projectsResp = UrlFetchApp.fetch('https://api.openai.com/v1/organization/projects', {
      method: 'get',
      headers,
      muteHttpExceptions: true
    });
    
    const code = projectsResp.getResponseCode();
    const text = projectsResp.getContentText();
    
    if (code >= 200 && code < 300) {
      const projectsData = JSON.parse(text);
      const projectCount = projectsData.data?.length || 0;
      SpreadsheetApp.getUi().alert('✅ API connection successful!\n\nResponse code: ' + code + '\n\nProjects found: ' + projectCount);
    } else {
      SpreadsheetApp.getUi().alert('❌ API connection failed!\n\nResponse code: ' + code + '\n\nResponse: ' + text);
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error testing API: ' + e.toString());
  }
}

function getOrganizationInfo() {
  try {
    const props = PropertiesService.getScriptProperties();
    const ADMIN_KEY = props.getProperty('OPENAI_ADMIN_KEY');
    if (!ADMIN_KEY) {
      SpreadsheetApp.getUi().alert('Error: OPENAI_ADMIN_KEY not set in Script Properties');
      return;
    }
    
    const headers = { 
      'Authorization': 'Bearer ' + ADMIN_KEY
    };
    
    // Get organization info
    const orgResp = UrlFetchApp.fetch('https://api.openai.com/v1/organizations', {
      method: 'get',
      headers,
      muteHttpExceptions: true
    });
    
    const code = orgResp.getResponseCode();
    const text = orgResp.getContentText();
    
    if (code >= 200 && code < 300) {
      const orgData = JSON.parse(text);
      const orgInfo = orgData.data?.[0] || {};
      SpreadsheetApp.getUi().alert('Organization Info:\n\nID: ' + orgInfo.id + '\nName: ' + orgInfo.name + '\n\nResponse: ' + text);
    } else {
      SpreadsheetApp.getUi().alert('❌ Failed to get organization info!\n\nResponse code: ' + code + '\n\nResponse: ' + text);
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error getting organization info: ' + e.toString());
  }
}



function getYesterdayData() {
  try {
    const props = PropertiesService.getScriptProperties();
    const ADMIN_KEY = props.getProperty('OPENAI_ADMIN_KEY');
    if (!ADMIN_KEY) {
      SpreadsheetApp.getUi().alert('Error: OPENAI_ADMIN_KEY not set in Script Properties');
      return;
    }
    
    const ss = SpreadsheetApp.getActive();
    const tz = ss.getSpreadsheetTimeZone() || 'Africa/Cairo';
    
    // Get yesterday's data specifically
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    yesterday.setHours(0, 0, 0, 0);
    
    const endDate = new Date(yesterday);
    endDate.setHours(23, 59, 59, 999);
    
    const start = Math.floor(yesterday.getTime() / 1000);
    const end = Math.floor(endDate.getTime() / 1000);
    
    console.log('📅 Yesterday time window:', new Date(start * 1000).toISOString(), 'to', new Date(end * 1000).toISOString());
    
    const headers = { 
      'Authorization': 'Bearer ' + ADMIN_KEY
    };
    
    // Fetch projects for name mapping
    const projectsResp = UrlFetchApp.fetch('https://api.openai.com/v1/organization/projects', {
      method: 'get',
      headers,
      muteHttpExceptions: true
    });
    
    const projectsJson = safeJson(projectsResp);
    if (!projectsJson) throw new Error('Failed to fetch project list');
    
    const projectNameMap = {};
    (projectsJson.data || []).forEach(p => {
      projectNameMap[p.id] = p.name || '';
    });
    
    function getProjectName(id) {
      return projectNameMap[id] || '';
    }
    
    function cleanModelName(model) {
      if (!model) return '';
      return model.replace(/-\d{4}-\d{2}-\d{2}$/, '');
    }
    
    // Get yesterday's usage data
    const usageParams = {
      start_time: start,
      end_time: end,
      bucket_width: '1d',
      limit: 31,
      group_by: ['project_id', 'model']
    };
    
    const usageUrl = 'https://api.openai.com/v1/organization/usage/completions?' + encodeParams(usageParams);
    console.log('Yesterday usage URL:', usageUrl);
    
    const usageResp = UrlFetchApp.fetch(usageUrl, { 
      method: 'get', 
      headers, 
      muteHttpExceptions: true 
    });
    
    const usageJson = safeJson(usageResp);
    if (!usageJson) throw new Error('Failed to fetch yesterday usage data');
    
    const usageBuckets = usageJson.data || [];
    const usageRows = [];
    
    usageBuckets.forEach(bucket => {
      const dateStr = Utilities.formatDate(new Date(bucket.start_time * 1000), tz, 'yyyy-MM-dd');
      (bucket.results || []).forEach(r => {
        if (!(r.input_tokens || r.output_tokens || r.input_cached_tokens || r.num_model_requests)) return;
        
        const totalInputTokens = Number(r.input_tokens || 0) + Number(r.input_cached_tokens || 0);
        const totalTokens = totalInputTokens + Number(r.output_tokens || 0);
        
        usageRows.push([
          dateStr,
          r.project_id || '',
          getProjectName(r.project_id || ''),
          cleanModelName(r.model),
          totalTokens,
          Number(r.input_tokens || 0),
          Number(r.output_tokens || 0),
          Number(r.input_cached_tokens || 0),
          Number(r.num_model_requests || 0),
          '' // Cost will be calculated separately
        ]);
      });
    });
    
    // Get yesterday's cost data
    const costsParams = {
      start_time: start,
      end_time: end,
      bucket_width: '1d',
      limit: 31,
      group_by: ['project_id', 'line_item']
    };
    
    const costsUrl = 'https://api.openai.com/v1/organization/costs?' + encodeParams(costsParams);
    console.log('Yesterday costs URL:', usageUrl);
    
    const costsResp = UrlFetchApp.fetch(costsUrl, { 
      method: 'get', 
      headers, 
      muteHttpExceptions: true 
    });
    
    const costsJson = safeJson(costsResp);
    if (!costsJson) throw new Error('Failed to fetch yesterday costs data');
    
    const costBuckets = costsJson.data || [];
    const costRows = [];
    
    costBuckets.forEach(bucket => {
      const dateStr = Utilities.formatDate(new Date(bucket.start_time * 1000), tz, 'yyyy-MM-dd');
      (bucket.results || []).forEach(r => {
        if (!r.amount || !r.amount.value) return;
        
        const amount = Number(r.amount.value);
        const amountPlus30 = Math.round(amount * 1.3 * 100) / 100;
        
        costRows.push([
          dateStr,
          r.project_id || '',
          getProjectName(r.project_id || ''),
          r.line_item || '',
          amount,
          amountPlus30
        ]);
      });
    });
    
    // Create or get yesterday summary sheets
    const yesterdayUsageSheet = ss.getSheetByName('Yesterday Usage') || ss.insertSheet('Yesterday Usage');
    if (yesterdayUsageSheet.getLastRow() === 0) {
      yesterdayUsageSheet.appendRow([
        'Date', 'Project ID', 'Project Name', 'Model',
        'Total Tokens (Input+Output)', 'Input tokens', 'Output tokens',
        'Cached tokens', 'Requests', 'Cost (USD)'
      ]);
    } else {
      // Clear existing data (keep headers)
      yesterdayUsageSheet.getRange(2, 1, yesterdayUsageSheet.getLastRow() - 1, yesterdayUsageSheet.getLastColumn()).clear();
    }
    
    const yesterdayCostsSheet = ss.getSheetByName('Yesterday Costs') || ss.insertSheet('Yesterday Costs');
    if (yesterdayCostsSheet.getLastRow() === 0) {
      yesterdayCostsSheet.appendRow([
        'Date', 'Project ID', 'Project Name', 'Line Item', 'Amount (USD)', 'Cost (+30%)'
      ]);
      yesterdayCostsSheet.hideColumn(yesterdayCostsSheet.getRange(1, 5));
    } else {
      // Clear existing data (keep headers)
      yesterdayCostsSheet.getRange(2, 1, yesterdayCostsSheet.getLastRow() - 1, yesterdayCostsSheet.getLastColumn()).clear();
    }
    
    // Add data to yesterday sheets
    if (usageRows.length) {
      yesterdayUsageSheet.getRange(2, 1, usageRows.length, usageRows[0].length).setValues(usageRows);
    }
    
    if (costRows.length) {
      yesterdayCostsSheet.getRange(2, 1, costRows.length, costRows[0].length).setValues(costRows);
    }
    
    // Update usage with costs
    updateUsageWithCosts(yesterdayUsageSheet, yesterdayCostsSheet);
    
         SpreadsheetApp.getUi().alert('✅ Yesterday\'s data retrieved successfully!\n\nUsage records: ' + usageRows.length + '\nCost records: ' + costRows.length);
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error getting yesterday data: ' + e.toString());
  }
}

function showTrackedProjects() {
  try {
    const props = PropertiesService.getScriptProperties();
    const ADMIN_KEY = props.getProperty('OPENAI_ADMIN_KEY');
    if (!ADMIN_KEY) {
      SpreadsheetApp.getUi().alert('Error: OPENAI_ADMIN_KEY not set in Script Properties');
      return;
    }
    
    const headers = { 
      'Authorization': 'Bearer ' + ADMIN_KEY
    };
    
    // Get all projects
    const projectsResp = UrlFetchApp.fetch('https://api.openai.com/v1/organization/projects', {
      method: 'get',
      headers,
      muteHttpExceptions: true
    });
    
    const code = projectsResp.getResponseCode();
    const text = projectsResp.getContentText();
    
    if (code >= 200 && code < 300) {
      const projectsData = JSON.parse(text);
      const projects = projectsData.data || [];
      
      let message = '📊 *All Projects in Your Organization:*\n\n';
      projects.forEach((project, index) => {
        const status = project.status === 'active' ? '✅' : '❌';
        message += (index + 1) + '. ' + status + ' **' + project.name + '**\n';
        message += '   ID: ' + project.id + '\n';
        message += '   Status: ' + project.status + '\n\n';
      });
      
      message += '\n**Total Projects:** ' + projects.length + '\n';
      message += '**Active Projects:** ' + projects.filter(p => p.status === 'active').length + '\n\n';
      message += 'The script will track **ALL active projects** by default.\n';
      message += 'To limit tracking to specific projects, edit the TRACK_SPECIFIC_PROJECTS array in the script.';
      
      SpreadsheetApp.getUi().alert(message);
    } else {
      SpreadsheetApp.getUi().alert('❌ Failed to fetch projects!\n\nResponse code: ' + code + '\n\nResponse: ' + text);
    }
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error showing tracked projects: ' + e.toString());
  }
}

function cleanDuplicateData() {
  try {
    const ss = SpreadsheetApp.getActive();
    
    // Clean Summary Usage sheet
    const summaryUsageSheet = ss.getSheetByName('Summary Usage');
    if (summaryUsageSheet) {
      const usageData = summaryUsageSheet.getDataRange().getValues();
      const headers = usageData[0];
      const usageRows = usageData.slice(1);
      
      if (usageRows.length > 0) {
        // Create a map to track unique combinations of date, project, and model
        const uniqueMap = new Map();
        const uniqueRows = [];
        
        usageRows.forEach(row => {
          const key = row[0] + '|' + row[1] + '|' + row[3]; // date|projectId|model
          if (!uniqueMap.has(key)) {
            uniqueMap.set(key, true);
            uniqueRows.push(row);
          } else {
            console.log('⚠️ Found duplicate: ' + row[0] + ' | ' + row[2] + ' | ' + row[3]);
          }
        });
        
        // Clear existing data and add unique rows
        summaryUsageSheet.getRange(2, 1, usageRows.length, headers.length).clear();
        if (uniqueRows.length > 0) {
          summaryUsageSheet.getRange(2, 1, uniqueRows.length, headers.length).setValues(uniqueRows);
        }
        
        console.log('✅ Cleaned Summary Usage sheet: Removed ' + (usageRows.length - uniqueRows.length) + ' duplicate rows');
      }
    }
    
    // Clean Summary Costs sheet
    const summaryCostsSheet = ss.getSheetByName('Summary Costs');
    if (summaryCostsSheet) {
      const costsData = summaryCostsSheet.getDataRange().getValues();
      const costHeaders = costsData[0];
      const costRows = costsData.slice(1);
      
      if (costRows.length > 0) {
        // Create a map to track unique combinations of date, project, and line item
        const uniqueCostMap = new Map();
        const uniqueCostRows = [];
        
        costRows.forEach(row => {
          const key = row[0] + '|' + row[1] + '|' + row[3]; // date|projectId|lineItem
          if (!uniqueCostMap.has(key)) {
            uniqueCostMap.set(key, true);
            uniqueCostRows.push(row);
          } else {
            console.log('⚠️ Found duplicate cost: ' + row[0] + ' | ' + row[2] + ' | ' + row[3]);
          }
        });
        
        // Clear existing data and add unique rows
        summaryCostsSheet.getRange(2, 1, costRows.length, costHeaders.length).clear();
        if (uniqueCostRows.length > 0) {
          summaryCostsSheet.getRange(2, 1, uniqueCostRows.length, costHeaders.length).setValues(uniqueCostRows);
        }
        
        console.log('✅ Cleaned Summary Costs sheet: Removed ' + (costRows.length - uniqueCostRows.length) + ' duplicate rows');
      }
    }
    
    SpreadsheetApp.getUi().alert('✅ Duplicate data cleaned successfully!');
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error cleaning duplicate data: ' + e.toString());
  }
}
