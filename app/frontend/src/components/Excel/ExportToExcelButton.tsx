// src/components/ExportToExcelButton.tsx
import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Button, Spinner, MessageBar, MessageBarType, ProgressIndicator } from '@fluentui/react';
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

// Optimized policy data extraction with better reliability
const extractPolicyData = async (
  policyId: string,
  setProgress?: (progress: number) => void
): Promise<PolicyData> => {
  // Check cache first
  if (policyDataCache[policyId]) {
    if (setProgress) setProgress(100);
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
      const result = await queryRAG(prompt, policyId);
      policyData.headerInfo[field] = result.trim();
      
      // Update progress for header fields (0-30%)
      if (setProgress) {
        const headerProgress = 5 + Math.floor((i + 1) / headerPrompts.length * 25);
        setProgress(headerProgress);
      }
    }
    
    // Get all sections in the policy
    const sectionsPrompt = 'List all coverage sections in this policy. Respond ONLY with comma-separated section names, nothing else.';
    const sectionsResponse = await queryRAG(sectionsPrompt, policyId);
    const sections = sectionsResponse.split(',').map(s => s.trim()).filter(s => s && s !== 'NF');
    
    if (setProgress) setProgress(35);
    
    // Process each section
    for (let i = 0; i < sections.length; i++) {
      const section = sections[i];
      
      // First get fields for this section
      const fieldsPrompt = `List only the field names in the ${section} section. Format as comma-separated values with NO additional text.`;
      const fieldsResponse = await queryRAG(fieldsPrompt, policyId);
      const fields = fieldsResponse.split(',').map(f => f.trim()).filter(f => f && f !== 'NF');
      
      policyData.sections[section] = [];
      
      // Then get each field's value
      for (let j = 0; j < fields.length; j++) {
        const field = fields[j];
        const valuePrompt = `For the ${section} section: Extract ONLY the value for ${field}. Return the value followed by the page number, separated by a pipe symbol: value|page`;
        const valueResponse = await queryRAG(valuePrompt, policyId);
        
        // Parse value and page
        const [value = '', page = ''] = valueResponse.split('|').map(part => part.trim());
        
        if (value && value !== 'NF') {
          policyData.sections[section].push({
            fieldName: field,
            value: value,
            page: page
          });
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
    
    // Cache the results for future use
    policyDataCache[policyId] = policyData;
    
    if (setProgress) setProgress(100);
    return policyData;
  } catch (err) {
    console.error('Error extracting policy data:', err);
    throw err;
  }
};

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
  
  // Create header section - only include fields with real values
  for (const { field, label } of headerFields) {
    const value = data.headerInfo[field as keyof PolicyHeaderInfo];
    // Skip empty values, NF, or "I don't know" values
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
      // Skip fields with "I don't know" or empty values
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

const ExportToExcelButton: React.FC<{ policyId: string; fileName?: string }> = ({
  policyId,
  fileName = 'policy-data.xlsx'
}) => {
  const [isExporting, setIsExporting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [progress, setProgress] = useState<number>(0);
  const exportInProgress = useRef(false);
  const [debugMode, setDebugMode] = useState(false);

  // Cleanup function to ensure state is reset properly
  useEffect(() => {
    return () => {
      exportInProgress.current = false;
    };
  }, []);

  const handleExport = async () => {
    if (exportInProgress.current) return; // Prevent multiple exports
    
    setIsExporting(true);
    setError(null);
    setProgress(0);
    exportInProgress.current = true;

    try {
      // Start extraction with progress updates
      const data = await extractPolicyData(policyId, (progressValue) => {
        setProgress(progressValue);
      });
      
      // Log data for debugging if needed
      logDataToConsole(data);
      
      // Check if we have any actual data
      const hasHeaderData = Object.values(data.headerInfo).some(value => value && value !== 'NF');
      const hasSectionData = Object.values(data.sections).some(section => section.length > 0);
      
      if (!hasHeaderData && !hasSectionData) {
        throw new Error('No policy data was found. Please try again or contact support.');
      }
      
      // Transform and prepare Excel data
      const rows = transformToExcelFormat(data);
      
      if (rows.length === 0) {
        throw new Error('No data was extracted for export.');
      }
      
      // Generate Excel file
      const sheet = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, sheet, 'Policy Data');
      
      // Save file with short timeout to ensure UI updates
      setTimeout(() => {
        XLSX.writeFile(wb, fileName);
        setIsExporting(false);
        exportInProgress.current = false;
      }, 100);
      
    } catch (err) {
      console.error('Export error:', err);
      setError(`Export failed: ${err instanceof Error ? err.message : String(err)}`);
      setIsExporting(false);
      exportInProgress.current = false;
    }
  };

  return (
    <div>
      <Button
        primary
        disabled={isExporting}
        onClick={handleExport}
        iconProps={{ iconName: 'ExcelDocument' }}
      >
        {isExporting ? 'Exporting...' : 'Export to Excel'}
      </Button>

      {isExporting && (
        <div style={{ marginTop: '10px', width: '100%',  display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <ProgressIndicator 
            label="Exporting policy data..." 
            description={`${progress}% complete`}
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
      
      {/* Optional debug toggle - remove in production */}
      <div style={{ marginTop: '10px', fontSize: '12px' }}>
        <a href="#" onClick={(e) => { e.preventDefault(); setDebugMode(!debugMode); }}>
          {debugMode ? "Hide Debug" : "Show Debug"}
        </a>
        
        {debugMode && (
          <div style={{ marginTop: '10px', padding: '8px', background: '#f5f5f5', borderRadius: '4px' }}>
            <p>Policy ID: {policyId}</p>
            <p>Export Status: {isExporting ? 'In Progress' : 'Ready'}</p>
            <p>Progress: {progress}%</p>
            <p>Cache Status: {policyDataCache[policyId] ? 'Cached' : 'Not Cached'}</p>
            <Button 
              onClick={() => {
                delete policyDataCache[policyId];
                alert('Cache cleared for this policy.');
              }}
              style={{ marginTop: '8px' }}
            >
              Clear Cache
            </Button>
          </div>
        )}
      </div>
    </div>
  );
};

export default ExportToExcelButton;