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
  Dialog,
  DialogType,
  DialogFooter,
  Label,
  IconButton,
  SearchBox
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

// For progressive display in the UI
interface DisplayRow {
  key: string;
  section: string;
  field: string;
  value: string;
  page: string;
  isNew?: boolean; // To highlight newly added rows
}

// Cache for policy data to avoid repeated queries
const policyDataCache: Record<string, PolicyData> = {};

// Debug log function
const debug = (message: string, data?: any) => {
  if (process.env.NODE_ENV !== 'production') {
    console.log(`[ExportButton] ${message}`, data || '');
  }
};

// Single query function - more reliable than batching
const queryRAG = async (prompt: string, policyId: string): Promise<string> => {
  try {
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
    return result || 'NF';
  } catch (err) {
    console.error('Query RAG error:', err);
    return 'NF';
  }
};

// Add retry logic for more reliability
const queryRAGWithRetry = async (prompt: string, policyId: string, maxRetries = 2): Promise<string> => {
  let attempts = 0;
  let lastError = null;
  
  while (attempts <= maxRetries) {
    try {
      const result = await queryRAG(prompt, policyId);
      if (result && result !== 'NF') {
        if (attempts > 0) {
          debug(`Query succeeded after ${attempts} retries: "${prompt.substring(0, 50)}..."`);
        }
        return result;
      }
      
      // If we got 'NF', wait and retry
      attempts++;
      debug(`Query returned NF, retrying (${attempts}/${maxRetries}): "${prompt.substring(0, 50)}..."`);
      await new Promise(r => setTimeout(r, 800)); // Wait before retry
    } catch (err) {
      lastError = err;
      attempts++;
      debug(`Query failed, retrying (${attempts}/${maxRetries}): "${prompt.substring(0, 50)}..."`);
      await new Promise(r => setTimeout(r, 1000)); // Wait longer after error
    }
  }
  
  debug(`Query failed after ${maxRetries} retries: "${prompt.substring(0, 50)}..."`);
  return 'NF';
};

// Convert policy data to display rows for UI
const convertToDisplayRows = (
  data: Partial<PolicyData>, 
  existingRows: DisplayRow[] = []
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
      if (value && 
          value !== 'NF' && 
          value !== 'I don\'t know.' &&
          value !== 'I don\'t know') {
        
        const key = `Header_${field}`;
        if (!existingKeys.has(key)) {
          rows.push({
            key,
            section: 'Header',
            field: label,
            value,
            page: '',
            isNew: true
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
        if (field.value && 
            field.value !== 'NF' && 
            field.value !== 'I don\'t know.' &&
            field.value !== 'I don\'t know' &&
            field.fieldName) {
          
          const key = `${section}_${field.fieldName}`;
          if (!existingKeys.has(key)) {
            rows.push({
              key,
              section,
              field: field.fieldName,
              value: field.value,
              page: field.page,
              isNew: true
            });
            existingKeys.add(key);
          }
        }
      }
    }
  }
  
  return rows;
};

// Progressive policy data extraction using the original query pattern that worked well
const extractPolicyDataProgressively = async (
  policyId: string,
  setProgress?: (progress: number) => void,
  onPartialResults?: (data: Partial<PolicyData>, displayRows: DisplayRow[]) => void
): Promise<PolicyData> => {
  debug(`Starting extraction for policy ID: ${policyId}`);
  
  // Check cache first
  if (policyDataCache[policyId]) {
    debug(`Using cached data for policy ID: ${policyId}`);
    if (setProgress) setProgress(100);
    if (onPartialResults) {
      const displayRows = convertToDisplayRows(policyDataCache[policyId]);
      onPartialResults(policyDataCache[policyId], displayRows);
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
  
  // Track partial results for the UI
  const partialData: Partial<PolicyData> = {
    headerInfo: { ...policyData.headerInfo },
    sections: {}
  };
  
  // Track display rows
  let currentDisplayRows: DisplayRow[] = [];
  
  // Track stats
  let totalQueries = 0;
  let successfulQueries = 0;

  try {
    if (setProgress) setProgress(5);
    
    // HEADER FIELDS - Using original query approach that was working well
    // IMPORTANT: We're maintaining sequential queries as in the original implementation
    // but providing progressive UI updates after each one
    
    debug("Processing header fields sequentially");
    const headerPrompts = [
      ['insuredName', 'Extract ONLY the Insured Name from the policy. Return ONLY the value, no other text.'],
      ['clientCode', 'Extract ONLY the Client Code from the policy. Return ONLY the value, no other text.'],
      ['policyNumber', 'Extract ONLY the Policy Number from the policy. Return ONLY the value, no other text.'],
      ['policyDates', 'Extract ONLY the Policy Effective Date and Expiration Date from the policy. Format as MM/DD/YY - MM/DD/YY. Return ONLY the formatted date range, no other text.'],
      ['policyType', 'Extract ONLY the Policy Type from the policy. Return ONLY the value, no other text.'],
      ['policyPremium', 'Extract ONLY the Policy Premium from the policy. Include $ and commas. Return ONLY the amount, no other text.'],
      ['expiringPolicyPremium', 'Extract ONLY the Expiring Policy Premium from the policy. Include $ and commas. Return ONLY the amount, no other text.']
    ] as const;
    
    // Process header fields one by one to ensure accuracy - USING ORIGINAL APPROACH
    for (let i = 0; i < headerPrompts.length; i++) {
      const [field, prompt] = headerPrompts[i];
      totalQueries++;
      
      debug(`Querying header field: ${field}`);
      const result = await queryRAGWithRetry(prompt, policyId);
      
      if (result && result !== 'NF') {
        successfulQueries++;
        policyData.headerInfo[field] = result.trim();
        partialData.headerInfo![field] = result.trim();
        
        // Update UI with this header field
        if (onPartialResults) {
          currentDisplayRows = convertToDisplayRows(
            { headerInfo: { [field]: result.trim() } as any },
            currentDisplayRows
          );
          onPartialResults({ ...partialData }, [...currentDisplayRows]);
        }
      }
      
      // Update progress for header fields (5-30%)
      if (setProgress) {
        const headerProgress = 5 + Math.floor((i + 1) / headerPrompts.length * 25);
        setProgress(headerProgress);
      }
      
      // Small delay between queries to avoid rate limiting
      await new Promise(r => setTimeout(r, 250));
    }
    
    debug("Completed header fields extraction");
    if (setProgress) setProgress(30);
    
    // Get all sections in the policy - USING ORIGINAL APPROACH
    debug("Querying for policy sections");
    const sectionsPrompt = 'List all coverage sections in this policy. Respond ONLY with comma-separated section names, nothing else.';
    totalQueries++;
    const sectionsResponse = await queryRAGWithRetry(sectionsPrompt, policyId);
    const sections = sectionsResponse.split(',').map(s => s.trim()).filter(s => s && s !== 'NF');
    
    if (sections.length > 0) {
      successfulQueries++;
      debug(`Found ${sections.length} sections: ${sections.join(', ')}`);
    } else {
      debug("No sections found or section query failed");
    }
    
    if (setProgress) setProgress(35);
    
    // Process each section sequentially - USING ORIGINAL APPROACH
    for (let i = 0; i < sections.length; i++) {
      const section = sections[i];
      debug(`Processing section ${i+1}/${sections.length}: ${section}`);
      
      // First get fields for this section
      const fieldsPrompt = `List only the field names in the ${section} section. Format as comma-separated values with NO additional text.`;
      totalQueries++;
      const fieldsResponse = await queryRAGWithRetry(fieldsPrompt, policyId);
      const fields = fieldsResponse.split(',').map(f => f.trim()).filter(f => f && f !== 'NF');
      
      if (fields.length > 0) {
        successfulQueries++;
        debug(`Found ${fields.length} fields in section ${section}`);
      } else {
        debug(`No fields found in section ${section} or query failed`);
      }
      
      policyData.sections[section] = [];
      partialData.sections![section] = [];
      
      // Then get each field's value sequentially - USING ORIGINAL APPROACH
      for (let j = 0; j < fields.length; j++) {
        const field = fields[j];
        const valuePrompt = `For the ${section} section: Extract ONLY the value for ${field}. Return the value followed by the page number, separated by a pipe symbol: value|page`;
        totalQueries++;
        
        debug(`Querying value for field ${j+1}/${fields.length}: ${section}.${field}`);
        const valueResponse = await queryRAGWithRetry(valuePrompt, policyId);
        
        // Parse value and page
        const [value = '', page = ''] = valueResponse.split('|').map(part => part.trim());
        
        if (value && value !== 'NF') {
          successfulQueries++;
          
          const fieldValue = {
            fieldName: field,
            value: value,
            page: page
          };
          
          policyData.sections[section].push(fieldValue);
          
          // Make sure section exists in partial data
          if (!partialData.sections![section]) {
            partialData.sections![section] = [];
          }
          
          partialData.sections![section].push(fieldValue);
          
          // Update UI with this new field
          if (onPartialResults) {
            currentDisplayRows = convertToDisplayRows(
              { sections: { [section]: [fieldValue] } },
              currentDisplayRows
            );
            onPartialResults({ ...partialData }, [...currentDisplayRows]);
          }
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
        
        // Small delay between field queries to avoid rate limiting
        await new Promise(r => setTimeout(r, 250));
      }
      
      // Small delay between sections to avoid rate limiting
      await new Promise(r => setTimeout(r, 500));
    }
    
    debug(`Extraction completed. Total queries: ${totalQueries}, Successful: ${successfulQueries}`);
    
    // Cache the results for future use
    policyDataCache[policyId] = policyData;
    
    if (setProgress) setProgress(100);
    
    // Final update with all data
    if (onPartialResults) {
      currentDisplayRows = currentDisplayRows.map(row => ({ ...row, isNew: false }));
      onPartialResults(policyData, currentDisplayRows);
    }
    
    return policyData;
  } catch (err) {
    debug(`Error extracting policy data: ${err}`);
    console.error('Error extracting policy data:', err);
    throw err;
  }
};

// Transform to Excel format (unchanged)
const transformToExcelFormat = (data: PolicyData): any[] => {
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
  
  // Create header section
  for (const { field, label } of headerFields) {
    const value = data.headerInfo[field as keyof PolicyHeaderInfo];
    if (value && 
        value !== 'NF' && 
        value !== 'I don\'t know.' &&
        value !== 'I don\'t know') {
      rows.push({
        'Section': 'Header',
        'Field': label,
        'Value': value,
        'Page': ''
      });
    }
  }

  // Add section fields
  for (const [section, fields] of Object.entries(data.sections)) {
    for (const field of fields) {
      if (field.value && 
          field.value !== 'NF' && 
          field.value !== 'I don\'t know.' &&
          field.value !== 'I don\'t know' &&
          field.fieldName) {
        rows.push({
          'Section': section,
          'Field': field.fieldName,
          'Value': field.value,
          'Page': field.page
        });
      }
    }
  }

  return rows;
};

// Debug helper 
const logDataToConsole = (data: PolicyData) => {
  debug('===== POLICY DATA EXTRACTED =====');
  debug(`Header Info: ${JSON.stringify(data.headerInfo)}`);
  debug(`Sections: ${Object.keys(data.sections).join(', ')}`);
  
  let totalFields = 0;
  for (const [section, fields] of Object.entries(data.sections)) {
    debug(`Section ${section} has ${fields.length} fields`);
    totalFields += fields.length;
  }
  
  debug(`Total fields across all sections: ${totalFields}`);
  debug('=================================');
};

const ExportToExcelButton: React.FC<{ policyId: string; fileName?: string }> = ({
  policyId,
  fileName = 'policy-data.xlsx'
}) => {
  const [isExporting, setIsExporting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [progress, setProgress] = useState<number>(0);
  const exportInProgress = useRef(false);
  const [debugMode, setDebugMode] = useState(false);
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
    },
    sections: {}
  });
  const [displayRows, setDisplayRows] = useState<DisplayRow[]>([]);
  const [filteredRows, setFilteredRows] = useState<DisplayRow[]>([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedSection, setSelectedSection] = useState<string | null>(null);
  
  // Stats tracking
  const [extractionStats, setExtractionStats] = useState({
    startTime: 0,
    totalFields: 0,
    fieldsExtracted: 0
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
    }
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
    
    // Remove "isNew" flag after 3 seconds
    const updatedRows = filtered.map(row => {
      if (row.isNew) {
        return { ...row, isNew: false };
      }
      return row;
    });
    
    setFilteredRows(updatedRows);
  }, [displayRows, searchTerm, selectedSection]);

  // Handle partial results updates
  const handlePartialResults = (data: Partial<PolicyData>, rows: DisplayRow[]) => {
    setPartialData(data);
    setDisplayRows(rows);
    
    // Update stats
    setExtractionStats(prevStats => ({
      ...prevStats,
      fieldsExtracted: rows.length,
    }));
    
    // Auto-open panel if it's not already open and we have data
    if (!showProgressivePanel && rows.length > 0) {
      setShowProgressivePanel(true);
    }
  };

  // Handle export to Excel
  const exportToExcel = (data: DisplayRow[]) => {
    if (data.length === 0) {
      setError('No data available to export.');
      return;
    }
    
    try {
      // Format data for Excel
      const excelRows = data.map(row => ({
        'Section': row.section,
        'Field': row.field,
        'Value': row.value,
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
      }, 
      sections: {} 
    });
    setDisplayRows([]);
    setFilteredRows([]);
    exportInProgress.current = true;
    
    // Record start time for performance measurement
    const startTime = Date.now();
    setExtractionStats({
      startTime,
      totalFields: 0,
      fieldsExtracted: 0
    });

    try {
      // Start extraction with progressive updates
      await extractPolicyDataProgressively(
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
        handlePartialResults
      );
      
      // Log performance information
      const totalTime = (Date.now() - startTime) / 1000;
      debug(`Export completed in ${totalTime.toFixed(1)} seconds`);
      
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
  
  // Calculate extraction rate and estimated completion
  const extractionRate = React.useMemo(() => {
    if (!isExporting || extractionStats.fieldsExtracted === 0) return null;
    
    const elapsedSeconds = (Date.now() - extractionStats.startTime) / 1000;
    if (elapsedSeconds < 5) return null; // Need some time to get a meaningful rate
    
    return extractionStats.fieldsExtracted / elapsedSeconds;
  }, [isExporting, extractionStats]);

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
        headerText={`Policy Data Explorer${isExporting ? ' (Extraction in Progress)' : ''}`}
        type={PanelType.large}
        closeButtonAriaLabel="Close"
        onRenderFooterContent={() => (
          <div style={{ display: 'flex', justifyContent: 'space-between', width: '100%' }}>
            <DefaultButton onClick={() => setShowProgressivePanel(false)}>
              Close
            </DefaultButton>
            <PrimaryButton 
              onClick={() => exportToExcel(filteredRows)}
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
                ? `Extracted ${displayRows.length} fields so far (${progress}% complete)`
                : `Found ${displayRows.length} fields in ${uniqueSections.length} sections`
              }
            </Text>
            
            {isExporting && (
              <Stack horizontal verticalAlign="center">
                <Spinner 
                  size={SpinnerSize.small} 
                  ariaLive="assertive"
                />
                <Text variant="small" style={{ marginLeft: 8 }}>
                  {extractionRate && extractionRate > 0
                    ? `${extractionRate.toFixed(1)} fields/sec`
                    : 'Extracting...'
                  }
                </Text>
              </Stack>
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
          </Stack>
          
          {/* Data table */}
          <DetailsList
            items={filteredRows}
            columns={columns}
            selectionMode={SelectionMode.none}
            onRenderItemColumn={(item, _, column) => {
              if (!column) return null;
              
              const content = item[column.fieldName as keyof DisplayRow];
              
              // Highlight newly added rows
              if (item.isNew) {
                return (
                  <div style={{ 
                    backgroundColor: '#fffde6', 
                    padding: '4px', 
                    borderRadius: '2px',
                    animation: 'fadeBackgroundColor 3s forwards'
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
          
          {filteredRows.length === 0 && (
            <Stack horizontalAlign="center" style={{ padding: '40px 0' }}>
              <Text>
                {isExporting 
                  ? "Extracting data... the results will appear here as they become available." 
                  : "No data available yet. Start the extraction process or adjust your filters."}
              </Text>
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
          {debugMode ? "Hide Debug Info" : "Show Debug Info"}
        </a>
        
        {debugMode && (
          <div style={{ marginTop: '10px', padding: '8px', background: '#f5f5f5', borderRadius: '4px' }}>
            <p>Policy ID: {policyId}</p>
            <p>Export Status: {isExporting ? 'In Progress' : 'Ready'}</p>
            <p>Progress: {progress}%</p>
            <p>Fields Extracted: {displayRows.length}</p>
            <p>Sections Found: {uniqueSections.join(', ')}</p>
            <p>Cache Status: {policyDataCache[policyId] ? 'Cached' : 'Not Cached'}</p>
            <p>Elapsed Time: {extractionStats.startTime > 0 ? `${Math.round((Date.now() - extractionStats.startTime) / 1000)}s` : 'N/A'}</p>
            {extractionRate && <p>Extraction Rate: {extractionRate.toFixed(2)} fields/sec</p>}
            <Button 
              onClick={() => {
                delete policyDataCache[policyId];
                alert('Cache cleared for this policy.');
              }}
              style={{ marginTop: '8px' }}
            >
              Clear Cache
            </Button>
            <Button 
              onClick={() => setShowProgressivePanel(true)}
              style={{ marginTop: '8px', marginLeft: '8px' }}
            >
              Open Data Panel
            </Button>
          </div>
        )}
      </div>
    </div>
  );
};

export default ExportToExcelButton;