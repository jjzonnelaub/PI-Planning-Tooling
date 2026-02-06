// ===== VALUE STREAM CONFIGURATION =====
// Updated to match consolidated capacity planning format
const VALUE_STREAM_CONFIG = {
  'EMA Clinical': {
    filter: null,
    scrumTeams: [
      'Alchemist',
      'Avengers',
      'Explorers',
      'Eyefinity',
      'Mandalore',
      'Ordernauts',
      'Painkillers',
      'Artificially Intelligent',
      'Patience',
      'Embryonics',
      'Vesties',
      'Spice Runners',
      'Pain Killers'
    ]
  },
  'EMA RAC': {
    filter: null,
    alternateNames: ['EMA RaC'],
    scrumTeams: ['Achievers', 'Borg', 'Cyborg']
  },
  'RCM Genie': {
    filter: null,
    alternateNames: ['RCM'],
    scrumTeams: ['Claimbots', 'Frontliners', 'Integrators', 'Vajra']
  },
  'MMPM': {
    filter: null,
    scrumTeams: [
      'Billionaires',
      'Claimcraft',
      'Lynx',
      'Penny-Wise',
      'Time-Keepers',
      'Trailblazers',
      'ClaimCraft',
      'Kaizen'
    ]
  },
  'Patient Collaboration': {
    filter: null,
    scrumTeams: [
      'Agni',
      'Apollo',
      'Bheem',
      'Jupiter',
      'Rubber Ducks',
      'Sudo',
      'Vaayu',
      'Voyagers'
    ]
  },
  'AIMM': {
    filter: {
      scrumTeam: 'Artificially Intelligent',
      valueStreams: ['EMA Clinical', 'EMA RAC', 'MMPM', 'Patient Collaboration']
    },
    scrumTeams: ['Artificially Intelligent']
  }
};

function getExcludedTeamsForContext(issues, currentValueStream) {
  if (!issues || !currentValueStream) {
    console.log('No issues or value stream provided for exclusion check');
    return [];
  }
  
  const normalizedCurrent = normalizeValueStreamName(currentValueStream);
  console.log('\n=== Determining Excluded Teams for Context: ' + normalizedCurrent + ' ===');
  
  const teamIssues = {};
  issues.forEach(issue => {
    const team = issue.scrumTeam;
    if (!team) return;
    
    if (!teamIssues[team]) {
      teamIssues[team] = { dependencies: [], regularWork: [] };
    }
    
    if (issue.issueType === 'Dependency') {
      teamIssues[team].dependencies.push(issue);
    } else {
      teamIssues[team].regularWork.push(issue);
    }
  });
  
  const excludedTeams = [];
  
  Object.keys(teamIssues).forEach(team => {
    const teamData = teamIssues[team];
    const hasRegularWork = teamData.regularWork.length > 0;
    
    const crossStreamDependencies = teamData.dependencies.filter(dep => {
      const depValueStream = normalizeValueStreamName(dep.dependsOnValuestream || dep.valueStream);
      const isCrossStream = depValueStream && depValueStream !== normalizedCurrent;
      if (isCrossStream) {
        console.log('  Team "' + team + '" has cross-VS dependency: ' + dep.key + ' from ' + (dep.dependsOnValuestream || dep.valueStream));
      }
      return isCrossStream;
    });
    
    const shouldExclude = !hasRegularWork && 
                         teamData.dependencies.length > 0 && 
                         crossStreamDependencies.length === teamData.dependencies.length;
    
    if (shouldExclude) {
      const depValueStreams = [...new Set(crossStreamDependencies.map(d => d.dependsOnValuestream || d.valueStream))];
      console.log('  EXCLUDING team "' + team + '"');
      console.log('     Reason: Only has ' + teamData.dependencies.length + ' dependency(ies) from: ' + depValueStreams.join(', '));
      console.log('     No regular work in ' + normalizedCurrent);
      excludedTeams.push(team);
    } else if (hasRegularWork) {
      console.log('  INCLUDING team "' + team + '": Has ' + teamData.regularWork.length + ' regular work items in ' + normalizedCurrent);
    }
  });
  
  console.log('\nExcluded ' + excludedTeams.length + ' team(s) from ' + normalizedCurrent + ':', excludedTeams);
  console.log('=== End Exclusion Analysis ===\n');
  
  return excludedTeams;
}

function normalizeValueStreamName(name) {
  if (!name) return '';
  return name.toString().toUpperCase().trim().replace(/[-_\s]+/g, ' ').replace(/\s+/g, ' ');
}

function isExcludedTeam(teamName, excludedTeams) {
  if (!teamName || !excludedTeams || excludedTeams.length === 0) return false;
  const normalizedTeam = teamName.toString().toUpperCase().trim();
  return excludedTeams.some(excludedTeam => excludedTeam.toString().toUpperCase().trim() === normalizedTeam);
}

const JIRA_CONFIG = {
  get baseUrl() { const config = getJiraConfig(); return config.baseUrl; },
  get email() { const config = getJiraConfig(); return config.email; },
  get apiToken() { const config = getJiraConfig(); return config.apiToken; }
};

const FIELD_MAPPINGS = {
  summary: 'summary',
  status: 'status',
  storyPoints: 'customfield_10037',
  storyPointEstimate: 'customfield_10016',
  epicLink: 'customfield_10014',
  programIncrement: 'customfield_10113',
  valueStream: 'customfield_10046',
  orgField: 'customfield_11192',
  piCommitment: 'customfield_10063',
  scrumTeam: 'customfield_10040',
  piTargetIteration: 'customfield_10061',
  iterationStart: 'customfield_10069',
  iterationEnd: 'customfield_10070',
  allocation: 'customfield_10043',
  portfolioInitiative: 'customfield_10049',
  programInitiative: 'customfield_10050',
  featurePoints: 'customfield_10252',
  rag: 'customfield_10068',
  ragNote: 'customfield_10067',
  dependsOnValuestream: 'customfield_10114',
  dependsOnTeam: 'customfield_10120',
  costOfDelay: 'customfield_10065',
  labels: 'labels',
  sprint: 'customfield_10020',
  fixVersions: 'fixVersions'
};

const CACHE_EXPIRATION_MINUTES = 60;

const ALLOCATION_CATEGORIES = {
  FEATURES: 'Features (Product - Compliance & Feature)',
  TECH: 'Tech / Platform',
  KLO: 'Planned KLO',
  QUALITY: 'Planned Quality'
};

const COLORS = {
  HEADER_PRIMARY: '#1B365D',
  HEADER_SECONDARY: '#6B3FA0',
  GOLD_ACCENT: '#FFC72C',
  BACKGROUND_LIGHT: '#F5F5F5',
  BACKGROUND_WARNING: '#FFF9C4',
  BACKGROUND_ERROR: '#FFCDD2',
  BACKGROUND_SUCCESS: '#C8E6C9',
  BACKGROUND_DANGER: '#FFCDD2',
  BACKGROUND_PURPLE: '#E8DEF8',
  PLANNING_HEADER: '#E1D5E7',
  ALLOCATION_FEATURES: '#c9daf8',
  ALLOCATION_TECH: '#d9ead3',
  ALLOCATION_KLO: '#fce5cd',
  ALLOCATION_QUALITY: '#f4cccc'
};

const UI_CONFIG = {
  DEFAULT_FONT: 'Comfortaa',
  PROGRESS_UPDATE_INTERVAL: 500
};

const REPORT_LOG_CONFIG = {
  sheetName: 'Report Log',
  headers: [
    'Generated Date', 'PI', 'Value Stream',
    'Report Name', 'Spreadsheet URL', 'Spreadsheet ID',
    'Epic Count', 'Status'
  ]
};

function buildEpicJQL(programIncrement, displayValueStream) {
  const config = VALUE_STREAM_CONFIG[displayValueStream];
  
  if (!config) {
    console.warn('Unknown value stream: ' + displayValueStream + ', using default query');
    return 'issuetype = Epic AND cf[10113] = "' + programIncrement + '" AND cf[10046] = "' + displayValueStream + '" AND status != "Closed"';
  }
  
  let jql = 'issuetype = Epic AND cf[10113] = "' + programIncrement + '" AND status != "Closed"';
  
  if (displayValueStream === 'AIMM') {
    const validValueStreams = config.filter.valueStreams || ['EMA Clinical', 'EMA RAC', 'MMPM'];
    jql += ' AND cf[10046] in ("' + validValueStreams.join('","') + '")';
    jql += ' AND cf[10040] = "' + config.filter.scrumTeam + '"';
  } else {
    jql += ' AND cf[10046] = "' + displayValueStream + '"';
  }
  
  console.log('JQL for ' + displayValueStream + ': ' + jql);
  return jql;
}

function getAvailableValueStreams() {
  return ['AIMM', 'EMA Clinical', 'EMA RAC', 'MMPM', 'Patient Collaboration', 'RCM Genie'].sort();
}

function mapAllocationToCategory(allocation) {
  return getAllocationCategory(allocation);
}

function getAllocationCategory(allocation) {
  if (!allocation) return ALLOCATION_CATEGORIES.FEATURES;
  
  const allocationLower = allocation.toString().toLowerCase().trim();
  
  const mappingRules = {
    [ALLOCATION_CATEGORIES.FEATURES]: [
      'feature', 'product', 'compliance', 'capability', 'enhancement',
      'story', 'user story', 'requirement', 'func', 'new feature'
    ],
    [ALLOCATION_CATEGORIES.TECH]: [
      'tech', 'platform', 'infrastructure', 'architecture', 'technical',
      'system', 'framework', 'devops', 'tooling', 'engineering'
    ],
    [ALLOCATION_CATEGORIES.KLO]: [
      'klo', 'keep', 'lights', 'maintenance', 'support', 'operational',
      'ops', 'sustaining', 'keep lights on', 'bau', 'business as usual'
    ],
    [ALLOCATION_CATEGORIES.QUALITY]: [
      'quality', 'defect', 'bug', 'fix', 'issue', 'problem',
      'qa', 'test', 'testing', 'quality assurance', 'defect fix'
    ]
  };
  
  for (const [category, keywords] of Object.entries(mappingRules)) {
    for (const keyword of keywords) {
      if (allocationLower.includes(keyword)) {
        return category;
      }
    }
  }
  
  return ALLOCATION_CATEGORIES.FEATURES;
}

function getAllocationColor(allocationType) {
  const colorMap = {
    [ALLOCATION_CATEGORIES.FEATURES]: COLORS.ALLOCATION_FEATURES,
    [ALLOCATION_CATEGORIES.TECH]: COLORS.ALLOCATION_TECH,
    [ALLOCATION_CATEGORIES.KLO]: COLORS.ALLOCATION_KLO,
    [ALLOCATION_CATEGORIES.QUALITY]: COLORS.ALLOCATION_QUALITY
  };
  return colorMap[allocationType] || '#ffffff';
}

function getTeamsForValueStream(valueStream) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const registrySheet = spreadsheet.getSheetByName('Team Registry');
  
  if (!registrySheet) {
    console.log('Team Registry sheet not found, using hardcoded config');
    return getTeamsFromConfig(valueStream);
  }
  
  console.log('Reading teams for ' + valueStream + ' from Team Registry...');
  
  const data = registrySheet.getDataRange().getValues();
  if (data.length < 2) {
    console.warn('Team Registry sheet has no data');
    return [];
  }
  
  const headers = data[0];
  const vsCol = headers.indexOf('Value Stream');
  const teamCol = headers.indexOf('Scrum Team');
  const activeCol = headers.indexOf('Active');
  
  if (vsCol === -1 || teamCol === -1 || activeCol === -1) {
    console.error('Team Registry sheet missing required columns');
    return getTeamsFromConfig(valueStream);
  }
  
  const teams = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowVS = row[vsCol];
    const team = row[teamCol];
    const active = row[activeCol];
    
    if (team && team.toString().includes('[Add Team Name]')) continue;
    if (rowVS === valueStream && active === true && team) {
      teams.push(team.toString().trim());
    }
  }
  
  console.log('Found ' + teams.length + ' active teams for ' + valueStream + ':', teams.join(', '));
  return teams;
}

function getTeamsFromConfig(valueStream) {
  if (typeof VALUE_STREAM_CONFIG !== 'undefined' && VALUE_STREAM_CONFIG[valueStream]) {
    const config = VALUE_STREAM_CONFIG[valueStream];
    return config.scrumTeams || [];
  }
  
  const fallbackTeams = {
    'EMA Clinical': [
      'Alchemist', 'Avengers', 'Explorers', 'Eyefinity',
      'Mandalore', 'Ordernauts', 'Painkillers', 'Artificially Intelligent',
      'Patience', 'Embryonics', 'Vesties', 'Spice Runners', 'Pain Killers'
    ],
    'MMPM': [
      'Billionaires', 'Claimcraft', 'ClaimCraft', 'Lynx',
      'Penny-Wise', 'Time-Keepers', 'Trailblazers', 'Kaizen'
    ],
    'Patient Collaboration': [
      'Agni', 'Apollo', 'Bheem', 'Jupiter',
      'Rubber Ducks', 'Sudo', 'Vaayu', 'Voyagers'
    ],
    'RCM Genie': ['Claimbots', 'Frontliners', 'Integrators', 'Vajra'],
    'RCM': ['Claimbots', 'Frontliners', 'Integrators', 'Vajra'],
    'EMA RAC': ['Achievers', 'Borg', 'Cyborg'],
    'EMA RaC': ['Achievers', 'Borg', 'Cyborg'],
    'AIMM': ['Artificially Intelligent']
  };
  
  return fallbackTeams[valueStream] || [];
}

function getAllValueStreamsFromRegistry() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const registrySheet = spreadsheet.getSheetByName('Team Registry');
  
  if (!registrySheet) {
    return ['AIMM', 'EMA Clinical', 'EMA RAC', 'MMPM', 'Patient Collaboration', 'RCM Genie'];
  }
  
  const data = registrySheet.getDataRange().getValues();
  const headers = data[0];
  const vsCol = headers.indexOf('Value Stream');
  const activeCol = headers.indexOf('Active');
  
  const valueStreams = new Set();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const vs = row[vsCol];
    const active = row[activeCol];
    if (vs && active === true) {
      valueStreams.add(vs.toString().trim());
    }
  }
  
  return Array.from(valueStreams).sort();
}

function validateTeamRegistry(programIncrement, valueStream) {
  console.log('\n=== Validating Team Registry for ' + valueStream + ' in ' + programIncrement + ' ===');
  
  const registryTeams = getTeamsForValueStream(valueStream);
  console.log('Teams in registry: ' + registryTeams.length);
  
  const jql = 'cf[10113] = "' + programIncrement + '" AND cf[10046] = "' + valueStream + '"';
  console.log('Testing with JQL: ' + jql);
  
  try {
    const issues = searchJiraIssues(jql, 100);
    
    const jiraTeams = new Set();
    issues.forEach(issue => {
      if (issue.scrumTeam) jiraTeams.add(issue.scrumTeam);
    });
    
    console.log('Teams found in JIRA: ' + jiraTeams.size);
    console.log('JIRA teams:', Array.from(jiraTeams).sort().join(', '));
    
    const inRegistryNotJira = registryTeams.filter(t => !jiraTeams.has(t));
    const inJiraNotRegistry = Array.from(jiraTeams).filter(t => !registryTeams.includes(t));
    
    if (inRegistryNotJira.length > 0) {
      console.warn('Teams in registry but NOT in JIRA:', inRegistryNotJira.join(', '));
    }
    if (inJiraNotRegistry.length > 0) {
      console.warn('Teams in JIRA but NOT in registry:', inJiraNotRegistry.join(', '));
      console.warn('   Add these teams to Team Registry!');
    }
    if (inRegistryNotJira.length === 0 && inJiraNotRegistry.length === 0) {
      console.log('Team Registry matches JIRA perfectly!');
    }
    
  } catch (error) {
    console.error('Error validating teams:', error);
  }
}