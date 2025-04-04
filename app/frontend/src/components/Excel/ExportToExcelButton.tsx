// src/components/ExportToExcelButton.tsx
import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  Button, 
  MessageBar, 
  MessageBarType, 
  ProgressIndicator,
  DetailsList,
  IColumn,
  SelectionMode,
  Panel,
  PanelType,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  SearchBox,
  Toggle
} from '@fluentui/react';
import { ChatAppRequest, ResponseMessage } from '../../api/models';



interface PolicyHeaderInfo {
  insuredName: string;
  clientCode: string;
  policyNumber: string;
  policyDates: string;
  policyType: string;
  policyPremium: string;
  expiringPolicyPremium: string;
}

interface SectionFieldValue {
  fieldName: string;
  value: string;
  page: string;
}

interface PolicyData {
  headerInfo: PolicyHeaderInfo;
  sections: {
    [sectionName: string]: SectionFieldValue[];
  };
}

// Cache for policy data to avoid repeated queries
const policyDataCache: Record<string, PolicyData> = {};

// Debug logger
const debugLog = (message: string, data?: any) => {
  if (process.env.NODE_ENV !== 'production') {
    console.log(`[ExportButton] ${message}`, data !== undefined ? data : '');
  }
};

// Single query function - more reliable than batching
const queryRAG = async (prompt: string, policyId: string): Promise<string> => {
  try {
    debugLog(`Sending query: "${prompt.substring(0, 50)}..."`);
    
    const messages: ResponseMessage[] = [
      {
        content: `For policy ID ${policyId}: ${prompt}`,
        role: 'user'
      }
    ];

    const request: ChatAppRequest = {
      messages,
      context: {
        overrides: {
          temperature: 0,
          top: 3,
          retrieval_mode: undefined,
          semantic_ranker: true,
          suggest_followup_questions: false,
          vector_fields: [],
          language: 'en'
        }
      },
      session_state: null
    };

    const response = await fetch('/api/ask', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(request)
    });

    if (!response.ok) throw new Error(`API error ${response.statusText}`);

    const data = await response.json();
    const result = data.message?.content?.trim() ?? '';
    
    // Log result for debugging
    debugLog(`Query result: "${result.substring(0, 50)}..."`);
    
    return result || 'NF';
  } catch (err) {
    console.error('Query RAG error:', err);
    return 'NF';
  }
};

// Retry logic for critical queries
const queryWithRetry = async (prompt: string, policyId: string, maxRetries = 2): Promise<string> => {
  let result = 'NF';
  let attempts = 0;
  
  while (attempts <= maxRetries && (result === 'NF' || !result)) {
    if (attempts > 0) {
      debugLog(`Retrying query attempt ${attempts}/${maxRetries}`);
      // Wait a bit between retries
      await new Promise(r => setTimeout(r, 1000));
    }
    
    result = await queryRAG(prompt, policyId);
    attempts++;
    
    if (result && result !== 'NF') break;
  }
  
  return result;
};

// ================ NEW CODE FOR UI INTEGRATION ================
// Interface for UI display rows
interface DisplayRow {
  key: string;
  section: string;
  field: string;
  value: string;
  page: string;
  isNew?: boolean; // To highlight newly added rows
  rawResponse?: string; // For debugging - show the original response
}

// Convert policy data to display rows for UI
const convertToDisplayRows = (
  data: Partial<PolicyData>, 
  existingRows: DisplayRow[] = [],
  showAllResponses = false // Add flag to show all responses
): DisplayRow[] => {
  const rows: DisplayRow[] = [...existingRows];
  const existingKeys = new Set(existingRows.map(row => row.key));
  
  // Process header info if available
  if (data.headerInfo) {
    const headerFields = [
      { field: 'insuredName', label: 'Insured Name' },
      { field: 'clientCode', label: 'Client Code' },
      { field: 'policyNumber', label: 'Policy Number' },
      { field: 'policyDates', label: 'Policy Dates' },
      { field: 'policyType', label: 'Policy Type' },
      { field: 'policyPremium', label: 'Policy Premium' },
      { field: 'expiringPolicyPremium', label: 'Expiring Policy Premium' }
    ];
    
    for (const { field, label } of headerFields) {
      const value = data.headerInfo[field as keyof PolicyHeaderInfo];
      // If showing all responses, include even empty values
      if (showAllResponses || (value && 
          value !== 'NF' && 
          value !== 'I don\'t know.' &&
          value !== 'I don\'t know')) {
        
        const key = `Header_${field}`;
        if (!existingKeys.has(key)) {
          rows.push({
            key,
            section: 'Header',
            field: label,
            value: value || 'NF', // Show NF instead of empty string
            page: '',
            isNew: true,
            rawResponse: value
          });
          existingKeys.add(key);
        }
      }
    }
  }
  
  // Process sections if available
  if (data.sections) {
    for (const [section, fields] of Object.entries(data.sections)) {
      for (const field of fields) {
        // If showing all responses, include even empty values
        if (showAllResponses || (field.value && 
            field.value !== 'NF' && 
            field.value !== 'I don\'t know.' &&
            field.value !== 'I don\'t know' &&
            field.fieldName)) {
          
          const key = `${section}_${field.fieldName}`;
          if (!existingKeys.has(key)) {
            rows.push({
              key,
              section,
              field: field.fieldName,
              value: field.value || 'NF', // Show NF instead of empty string
              page: field.page,
              isNew: true,
              rawResponse: field.value
            });
            existingKeys.add(key);
          }
        }
      }
    }
  }
  
  return rows;
};

// Modified Original Function with Callbacks but Same Core Logic
const extractPolicyData = async (
  policyId: string,
  setProgress?: (progress: number) => void,
  onPartialResults?: (data: Partial<PolicyData>, displayRows: DisplayRow[], showAllResponses?: boolean) => void,
  showAllResponses = false
): Promise<PolicyData> => {
  debugLog(`Starting extraction for policy ID: ${policyId}`);
  
  // Check cache first
  if (policyDataCache[policyId]) {
    debugLog(`Using cached data for policy ID: ${policyId}`);
    if (setProgress) setProgress(100);
    if (onPartialResults) {
      const displayRows = convertToDisplayRows(policyDataCache[policyId], [], showAllResponses);
      onPartialResults(policyDataCache[policyId], displayRows, showAllResponses);
    }
    return policyDataCache[policyId];
  }

  const policyData: PolicyData = {
    headerInfo: {
      insuredName: '',
      clientCode: '',
      policyNumber: '',
      policyDates: '',
      policyType: '',
      policyPremium: '',
      expiringPolicyPremium: ''
    },
    sections: {}
  };
  
  // For tracking partial data for UI updates - initialize with undefined
  const partialData: Partial<PolicyData> = {
    headerInfo: {
      insuredName: '',
      clientCode: '',
      policyNumber: '',
      policyDates: '',
      policyType: '',
      policyPremium: '',
      expiringPolicyPremium: ''
    }, // Start with empty object to only show fields we've processed
    sections: {}
  };
  
  // For tracking display rows
  let currentDisplayRows: DisplayRow[] = [];
  
  // Stats
  let totalQueries = 0;
  let successfulQueries = 0;

  try {
    if (setProgress) setProgress(5);
    
    // OPTIMIZATION: Get header fields with specific, direct queries
    const headerPrompts = [
      ['insuredName', 'Extract ONLY the Insured Name from the policy. Return ONLY the value, no other text.'],
      ['clientCode', 'Extract ONLY the Client Code from the policy. Return ONLY the value, no other text.'],
      ['policyNumber', 'Extract ONLY the Policy Number from the policy. Return ONLY the value, no other text.'],
      ['policyDates', 'Extract ONLY the Policy Effective Date and Expiration Date from the policy. Format as MM/DD/YY - MM/DD/YY. Return ONLY the formatted date range, no other text.'],
      ['policyType', 'Extract ONLY the Policy Type from the policy. Return ONLY the value, no other text.'],
      ['policyPremium', 'Extract ONLY the Policy Premium from the policy. Include $ and commas. Return ONLY the amount, no other text.'],
      ['expiringPolicyPremium', 'Extract ONLY the Expiring Policy Premium from the policy. Include $ and commas. Return ONLY the amount, no other text.']
    ] as const;
    
    // Process header fields one by one to ensure accuracy
    for (let i = 0; i < headerPrompts.length; i++) {
      const [field, prompt] = headerPrompts[i];
      totalQueries++;
      
      debugLog(`Processing header field: ${field}`);
      const result = await queryRAG(prompt, policyId);
      
      policyData.headerInfo[field] = result.trim();
      partialData.headerInfo![field] = result.trim(); // Add to partial data for UI
      
      if (result && result !== 'NF') {
        successfulQueries++;
      }
      
      // UI Update after each header field - NEW
      if (onPartialResults) {
        currentDisplayRows = convertToDisplayRows(
          { headerInfo: { [field]: result.trim() } as any },
          currentDisplayRows,
          showAllResponses
        );
        onPartialResults({...partialData}, [...currentDisplayRows], showAllResponses);
      }
      
      // Update progress for header fields (0-30%)
      if (setProgress) {
        const headerProgress = 5 + Math.floor((i + 1) / headerPrompts.length * 25);
        setProgress(headerProgress);
      }
    }
    
    debugLog(`Completed header extraction. Success rate: ${successfulQueries}/${totalQueries}`);
    
    // Get all sections in the policy - USE RETRY FOR THIS CRITICAL QUERY
    debugLog("Requesting sections list");
    const sectionsPrompt = 'List all coverage sections in this policy. Respond ONLY with comma-separated section names, nothing else.';
    totalQueries++;
    
    // Use retry for this critical query
    const sectionsResponse = await queryWithRetry(sectionsPrompt, policyId);
    debugLog(`Sections response: "${sectionsResponse}"`);
    
    const sections = sectionsResponse.split(',').map(s => s.trim()).filter(s => s && s !== 'NF');
    
    if (sections.length > 0) {
      successfulQueries++;
      debugLog(`Found ${sections.length} sections: ${JSON.stringify(sections)}`);
    } else {
      debugLog("WARNING: No sections found in the policy");
    }
    
    if (setProgress) setProgress(35);
    
    // Process each section
    for (let i = 0; i < sections.length; i++) {
      const section = sections[i];
      debugLog(`Processing section ${i+1}/${sections.length}: ${section}`);
      
      // First get fields for this section - USE RETRY FOR THIS CRITICAL QUERY
      const fieldsPrompt = `List only the field names in the ${section} section. Format as comma-separated values with NO additional text.`;
      totalQueries++;
      
      // Use retry for this critical query
      const fieldsResponse = await queryWithRetry(fieldsPrompt, policyId);
      debugLog(`Fields for ${section}: "${fieldsResponse}"`);
      
      const fields = fieldsResponse.split(',').map(f => f.trim()).filter(f => f && f !== 'NF');
      
      if (fields.length > 0) {
        successfulQueries++;
        debugLog(`Found ${fields.length} fields in section ${section}: ${JSON.stringify(fields)}`);
      } else {
        debugLog(`WARNING: No fields found in section ${section}`);
      }
      
      policyData.sections[section] = [];
      
      // Then get each field's value
      for (let j = 0; j < fields.length; j++) {
        const field = fields[j];
        const valuePrompt = `For the ${section} section: Extract ONLY the value for ${field}. Return the value followed by the page number, separated by a pipe symbol: value|page`;
        totalQueries++;
        
        const valueResponse = await queryRAG(valuePrompt, policyId);
        debugLog(`${section}.${field} value: "${valueResponse}"`);
        
        // Parse value and page
        const [value = '', page = ''] = valueResponse.split('|').map(part => part.trim());
        
        const fieldData: SectionFieldValue = {
          fieldName: field,
          value: value,
          page: page
        };
        
        // Always add to policyData even if value is empty or NF
        policyData.sections[section].push(fieldData);
        
        if (value && value !== 'NF') {
          successfulQueries++;
        }
        
        // Add to partial data for UI updates
        if (!partialData.sections![section]) {
          partialData.sections![section] = [];
        }
        partialData.sections![section].push(fieldData);
        
        // UI Update after each field - NEW
        if (onPartialResults) {
          currentDisplayRows = convertToDisplayRows(
            { sections: { [section]: [fieldData] } },
            currentDisplayRows,
            showAllResponses
          );
          onPartialResults({...partialData}, [...currentDisplayRows], showAllResponses);
        }
        
        // Update progress - sections get 35-98% of the progress bar
        if (setProgress) {
          const sectionProgressWeight = (1.0 / sections.length) * 63;
          const fieldProgressWeight = sectionProgressWeight / (fields.length || 1);
          const progress = 35 + 
            (i * sectionProgressWeight) + 
            ((j + 1) * fieldProgressWeight);
          
          setProgress(Math.min(Math.floor(progress), 98));
        }
      }
    }
    
    debugLog(`Extraction complete. Total queries: ${totalQueries}, Successful: ${successfulQueries}`);
    
    // Cache the results for future use
    policyDataCache[policyId] = policyData;
    
    if (setProgress) setProgress(100);
    
    // Final update with all data
    if (onPartialResults) {
      currentDisplayRows = convertToDisplayRows(policyData, [], showAllResponses);
      onPartialResults(policyData, currentDisplayRows, showAllResponses);
    }
    
    return policyData;
  } catch (err) {
    debugLog(`Error extracting policy data: ${err}`);
    console.error('Error extracting policy data:', err);
    throw err;
  }
};

// Transform to Excel format (unchanged)
const transformToExcelFormat = (data: PolicyData, includeEmpty = false): any[] => {
  const rows: any[] = [];

  // Define header fields to include
  const headerFields = [
    { field: 'insuredName', label: 'Insured Name' },
    { field: 'clientCode', label: 'Client Code' },
    { field: 'policyNumber', label: 'Policy Number' },
    { field: 'policyDates', label: 'Policy Dates' },
    { field: 'policyType', label: 'Policy Type' },
    { field: 'policyPremium', label: 'Policy Premium' },
    { field: 'expiringPolicyPremium', label: 'Expiring Policy Premium' }
  ];
  
  // Create header section - only include fields with real values
  for (const { field, label } of headerFields) {
    const value = data.headerInfo[field as keyof PolicyHeaderInfo];
    // Skip empty values unless includeEmpty is true
    if (includeEmpty || (value && 
        value !== 'NF' && 
        value !== 'I don\'t know.' &&
        value !== 'I don\'t know')) {
      rows.push({
        'Section': 'Header',
        'Field': label,
        'Value': value || '',
        'Page': ''
      });
    }
  }

  // Add section fields
  for (const [section, fields] of Object.entries(data.sections)) {
    for (const field of fields) {
      // Skip fields with "I don't know" or empty values unless includeEmpty is true
      if (includeEmpty || (field.value && 
          field.value !== 'NF' && 
          field.value !== 'I don\'t know.' &&
          field.value !== 'I don\'t know' &&
          field.fieldName)) {
        rows.push({
          'Section': section,
          'Field': field.fieldName,
          'Value': field.value || '',
          'Page': field.page
        });
      }
    }
  }

  return rows;
};

// Debug helper to see what's actually in the data
const logDataToConsole = (data: PolicyData) => {
  console.log('===== POLICY DATA EXTRACTED =====');
  console.log('Header Info:', data.headerInfo);
  console.log('Sections:', Object.keys(data.sections));
  
  let totalFields = 0;
  for (const [section, fields] of Object.entries(data.sections)) {
    console.log(`Section ${section} has ${fields.length} fields`);
    totalFields += fields.length;
  }
  
  console.log(`Total fields across all sections: ${totalFields}`);
  console.log('=================================');
};

// ================ THE UI COMPONENT ================

const ExportToExcelButton: React.FC<{ policyId: string; fileName?: string }> = ({
  policyId,
  fileName = 'policy-data.xlsx'
}) => {
  const [isExporting, setIsExporting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [progress, setProgress] = useState<number>(0);
  const exportInProgress = useRef(false);
  const [debugMode, setDebugMode] = useState(false);
  const [showAllResponses, setShowAllResponses] = useState(false);
  const [estimatedTime, setEstimatedTime] = useState<number | null>(null);
  
  // Progressive UI states
  const [showProgressivePanel, setShowProgressivePanel] = useState(false);
  const [partialData, setPartialData] = useState<Partial<PolicyData>>({
    headerInfo: {
      insuredName: '',
      clientCode: '',
      policyNumber: '',
      policyDates: '',
      policyType: '',
      policyPremium: '',
      expiringPolicyPremium: ''
    }, // Initialize with empty strings for all required properties
    sections: {}
  });
  const [displayRows, setDisplayRows] = useState<DisplayRow[]>([]);
  const [filteredRows, setFilteredRows] = useState<DisplayRow[]>([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedSection, setSelectedSection] = useState<string | null>(null);
  
  // Stats
  const [stats, setStats] = useState({
    totalQueries: 0,
    successfulQueries: 0,
    startTime: 0
  });
  
  // Define columns for the data display
  const columns: IColumn[] = [
    {
      key: 'section',
      name: 'Section',
      fieldName: 'section',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
    },
    {
      key: 'field',
      name: 'Field',
      fieldName: 'field',
      minWidth: 120,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: 'value',
      name: 'Value',
      fieldName: 'value',
      minWidth: 200,
      isResizable: true,
    },
    {
      key: 'page',
      name: 'Page',
      fieldName: 'page',
      minWidth: 50,
      maxWidth: 80,
      isResizable: true,
    },
    // Show raw response in debug mode
    ...(debugMode ? [{
      key: 'rawResponse',
      name: 'Raw Response',
      fieldName: 'rawResponse',
      minWidth: 150,
      isResizable: true,
    }] : [])
  ];

  // Cleanup function to ensure state is reset properly
  useEffect(() => {
    return () => {
      exportInProgress.current = false;
    };
  }, []);
  
  // Effect to filter rows based on search and section filter
  useEffect(() => {
    let filtered = [...displayRows];
    
    // Apply section filter if selected
    if (selectedSection) {
      filtered = filtered.filter(row => row.section === selectedSection);
    }
    
    // Apply search filter if any
    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      filtered = filtered.filter(
        row => 
          row.section.toLowerCase().includes(term) ||
          row.field.toLowerCase().includes(term) ||
          row.value.toLowerCase().includes(term)
      );
    }
    
    setFilteredRows(filtered);
    
    // Set a timeout to remove "isNew" highlighting after a few seconds
    const timeoutId = setTimeout(() => {
      setDisplayRows(rows => rows.map(row => ({...row, isNew: false})));
    }, 3000);
    
    return () => clearTimeout(timeoutId);
  }, [displayRows, searchTerm, selectedSection]);

  // Handle partial results updates
  const handlePartialResults = (data: Partial<PolicyData>, rows: DisplayRow[], showAll = false) => {
    setPartialData(data);
    setDisplayRows(rows);
    
    // Auto-open panel if it's not already open and we have data
    if (!showProgressivePanel && rows.length > 0) {
      setShowProgressivePanel(true);
    }
  };

  // Handle export to Excel
  const exportToExcel = (data: DisplayRow[], includeAllRows = false) => {
    if (data.length === 0) {
      setError('No data available to export.');
      return;
    }
    
    try {
      // Format data for Excel
      const excelRows = data.map(row => ({
        'Section': row.section,
        'Field': row.field,
        'Value': row.value === 'NF' && !includeAllRows ? '' : row.value,
        'Page': row.page
      }));
      
      // Generate Excel file
      const sheet = XLSX.utils.json_to_sheet(excelRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, sheet, 'Policy Data');
      
      // Save file
      XLSX.writeFile(wb, fileName);
    } catch (err) {
      console.error('Export error:', err);
      setError(`Export failed: ${err instanceof Error ? err.message : String(err)}`);
    }
  };

  // Handle the main export process
  const handleExport = async () => {
    if (exportInProgress.current) return; // Prevent multiple exports
    
    setIsExporting(true);
    setError(null);
    setProgress(0);
    setEstimatedTime(null);
    setPartialData({ 
      headerInfo: {
        insuredName: '',
        clientCode: '',
        policyNumber: '',
        policyDates: '',
        policyType: '',
        policyPremium: '',
        expiringPolicyPremium: ''
      }, // Initialize with empty strings for all required properties 
      sections: {} 
    });
    setDisplayRows([]);
    setFilteredRows([]);
    exportInProgress.current = true;
    
    // Record start time for performance measurement
    const startTime = Date.now();
    setStats({
      totalQueries: 0,
      successfulQueries: 0,
      startTime
    });

    try {
      // First, clear cache for this policy to ensure fresh extraction
      // Comment this out if you want to persist cache across runs
      if (policyDataCache[policyId]) {
        debugLog(`Clearing cache for policy ${policyId}`);
        delete policyDataCache[policyId];
      }
      
      // Start extraction with progressive updates
      await extractPolicyData(
        policyId,
        (progressValue) => {
          setProgress(progressValue);
          
          // Update estimated time after we have some progress
          if (progressValue > 10 && progressValue < 95) {
            const elapsedMs = Date.now() - startTime;
            const estimatedTotalMs = (elapsedMs / progressValue) * 100;
            const remainingSeconds = Math.round((estimatedTotalMs - elapsedMs) / 1000);
            setEstimatedTime(remainingSeconds);
          }
        },
        handlePartialResults,
        showAllResponses
      );
      
      // Log performance information
      const totalTime = (Date.now() - startTime) / 1000;
      debugLog(`Export completed in ${totalTime.toFixed(1)} seconds`);
      
      setIsExporting(false);
      setEstimatedTime(null);
      exportInProgress.current = false;
      
    } catch (err) {
      console.error('Export error:', err);
      setError(`Export failed: ${err instanceof Error ? err.message : String(err)}`);
      setIsExporting(false);
      setEstimatedTime(null);
      exportInProgress.current = false;
    }
  };
  
  // Get unique sections for the filter dropdown
  const uniqueSections = React.useMemo(() => {
    const sections = new Set(displayRows.map(row => row.section));
    return Array.from(sections);
  }, [displayRows]);

  return (
    <div>
      <Button
        primary
        disabled={isExporting}
        onClick={handleExport}
        iconProps={{ iconName: 'ExcelDocument' }}
      >
        {isExporting ? 'Extracting Data...' : 'Extract Policy Data'}
      </Button>

      {isExporting && (
        <div style={{ marginTop: '10px', width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <ProgressIndicator 
            label="Extracting policy data..." 
            description={
              estimatedTime 
                ? `${progress}% complete (about ${estimatedTime > 60 
                    ? `${Math.floor(estimatedTime / 60)} min ${estimatedTime % 60} sec` 
                    : `${estimatedTime} seconds`} remaining)`
                : `${progress}% complete`
            }
            percentComplete={progress / 100} 
          />
        </div>
      )}

      {error && (
        <MessageBar 
          messageBarType={MessageBarType.error} 
          isMultiline
          style={{ marginTop: '10px' }}
        >
          {error}
        </MessageBar>
      )}
      
      {/* Progressive UI Panel */}
      <Panel
        isOpen={showProgressivePanel}
        onDismiss={() => setShowProgressivePanel(false)}
        headerText="Policy Data Explorer"
        type={PanelType.large}
        closeButtonAriaLabel="Close"
        onRenderFooterContent={() => (
          <div style={{ display: 'flex', justifyContent: 'space-between', width: '100%' }}>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton onClick={() => setShowProgressivePanel(false)}>
                Close
              </DefaultButton>
              {debugMode && (
                <DefaultButton 
                  onClick={() => {
                    if (policyId) {
                      delete policyDataCache[policyId];
                      alert('Cache cleared for this policy.');
                    }
                  }}
                >
                  Clear Cache
                </DefaultButton>
              )}
            </Stack>
            <PrimaryButton 
              onClick={() => exportToExcel(filteredRows, showAllResponses)}
              disabled={filteredRows.length === 0}
              iconProps={{ iconName: 'ExcelDocument' }}
            >
              Export Current Data to Excel
            </PrimaryButton>
          </div>
        )}
        isFooterAtBottom={true}
      >
        <Stack tokens={{ childrenGap: 16 }}>
          {/* Status and controls section */}
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus">
              {isExporting 
                ? `Extracting data... (${progress}% complete)`
                : `Found ${displayRows.length} fields in ${uniqueSections.length} sections`
              }
            </Text>
            
            {isExporting && (
              <Spinner 
                size={SpinnerSize.small} 
                label="Extracting..." 
                ariaLive="assertive"
                labelPosition="right"
              />
            )}
          </Stack>
          
          {/* Filter and search controls */}
          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <Stack.Item grow={1}>
              <SearchBox
                placeholder="Search in fields and values..."
                onChange={(_, newValue) => setSearchTerm(newValue || '')}
                disabled={isExporting && displayRows.length === 0}
              />
            </Stack.Item>
            
            <Stack.Item>
              <DefaultButton
                text={selectedSection ? `Section: ${selectedSection}` : "All Sections"}
                menuProps={{
                  items: [
                    {
                      key: 'all',
                      text: 'All Sections',
                      onClick: () => setSelectedSection(null)
                    },
                    ...uniqueSections.map(section => ({
                      key: section,
                      text: section,
                      onClick: () => setSelectedSection(section)
                    }))
                  ]
                }}
                disabled={isExporting && displayRows.length === 0 || uniqueSections.length === 0}
              />
            </Stack.Item>
            
            {/* Debug controls */}
            {debugMode && (
              <Stack.Item>
                <Toggle 
                  label="Show All Responses"
                  checked={showAllResponses}
                  onChange={(_, checked) => setShowAllResponses(checked || false)}
                />
              </Stack.Item>
            )}
          </Stack>
          
          {/* Data table */}
          <DetailsList
            items={filteredRows}
            columns={columns}
            selectionMode={SelectionMode.none}
            onRenderItemColumn={(item, _, column) => {
              if (!column) return null;
              
              const content = item[column.fieldName as keyof DisplayRow];
              
              // For value column, display empty responses nicely
              if (column.key === 'value' && (content === 'NF' || content === '')) {
                return (
                  <span style={{ color: '#999', fontStyle: 'italic' }}>
                    {content === 'NF' ? 'Not Found' : 'Empty'}
                  </span>
                );
              }
              
              // Highlight newly added rows
              if (item.isNew) {
                return (
                  <div style={{ 
                    backgroundColor: '#fffde6', 
                    padding: '4px', 
                    borderRadius: '2px',
                    animation: 'fadeBackgroundColor 2s forwards'
                  }}>
                    {content}
                  </div>
                );
              }
              
              return content;
            }}
            styles={{
              root: {
                // Add styles for the scrollable area
                overflowY: 'auto',
                height: 'calc(100vh - 240px)'
              }
            }}
          />
          
          {filteredRows.length === 0 && !isExporting && (
            <Stack horizontalAlign="center" style={{ padding: '40px 0' }}>
              <Text>No data available yet. Please wait for extraction to progress or adjust your filters.</Text>
            </Stack>
          )}
        </Stack>
        
        {/* CSS for animations */}
        <style>{`
          @keyframes fadeBackgroundColor {
            from { background-color: #fffde6; }
            to { background-color: transparent; }
          }
        `}</style>
      </Panel>
      
      {/* Debug panel (optional) */}
      <div style={{ marginTop: '10px', fontSize: '12px' }}>
        <a href="#" onClick={(e) => { e.preventDefault(); setDebugMode(!debugMode); }}>
          {debugMode ? "Hide Debug" : "Show Debug"}
        </a>
        
        {debugMode && (
          <div style={{ marginTop: '10px', padding: '8px', background: '#f5f5f5', borderRadius: '4px' }}>
            <p>Policy ID: {policyId}</p>
            <p>Export Status: {isExporting ? 'In Progress' : 'Ready'}</p>
            <p>Progress: {progress}%</p>
            <p>Displayed Rows: {displayRows.length}</p>
            <p>Sections Found: {uniqueSections.join(', ')}</p>
            <p>Cache Status: {policyDataCache[policyId] ? 'Cached' : 'Not Cached'}</p>
            <p>Show All Responses: {showAllResponses ? 'Yes' : 'No'}</p>
            {stats.startTime > 0 && (
              <p>Time Elapsed: {Math.round((Date.now() - stats.startTime) / 1000)}s</p>
            )}
            <div style={{ display: 'flex', gap: '8px', marginTop: '8px' }}>
              <Button 
                onClick={() => {
                  delete policyDataCache[policyId];
                  alert('Cache cleared for this policy.');
                }}
              >
                Clear Cache
              </Button>
              <Button 
                onClick={() => setShowProgressivePanel(true)}
              >
                Open Data Panel
              </Button>
              <Button 
                onClick={() => {
                  const toggledValue = !showAllResponses;
                  setShowAllResponses(toggledValue);
                  
                  // Refresh the display rows with all responses
                  if (policyDataCache[policyId]) {
                    const rows = convertToDisplayRows(policyDataCache[policyId], [], toggledValue);
                    setDisplayRows(rows);
                  }
                }}
              >
                {showAllResponses ? 'Hide Empty Responses' : 'Show All Responses'}
              </Button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default ExportToExcelButton;