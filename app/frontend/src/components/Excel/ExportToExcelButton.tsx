// src/components/ExportToExcelButton.tsx
import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Button, Spinner, MessageBar, MessageBarType } from '@fluentui/react';
import { useAuthContext } from '../../hooks/useAuthContext';
import { ChatAppRequest, ResponseMessage } from '../../api/models';

// Define types for policy extraction data
interface PolicyHeaderInfo {
  insuredName: string;
  clientCode: string;
  policyNumber: string;
  policyDates: string; // Format: MM/DD/YY - MM/DD/YY
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

// Function to query the RAG model with the extraction prompts
async function extractPolicyData(
  policyId: string, 
  idToken: string | undefined
): Promise<PolicyData> {
  // Initialize the policy data structure
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

  // Phase 1: Extract Header Information
  const headerPrompts = [
    'Extract the Insured Name from the policy. Return only the value, nothing else.',
    'Extract the Client Code from the policy. Return only the value, nothing else.',
    'Extract the Policy Number from the policy. Return only the value, nothing else.',
    'Extract the Policy Effective Date and Expiration Date from the policy. Return only the value in format MM/DD/YY - MM/DD/YY, nothing else.',
    'Extract the Policy Type from the policy. Return only the value, nothing else.',
    'Extract the Policy Premium from the policy. Return only the value with $ and commas, nothing else.',
    'Extract the Expiring Policy Premium from the policy. Return only the value with $ and commas, nothing else.'
  ];

  const headerFields = [
    'insuredName',
    'clientCode',
    'policyNumber',
    'policyDates',
    'policyType',
    'policyPremium',
    'expiringPolicyPremium'
  ];

  // Query for header information
  for (let i = 0; i < headerPrompts.length; i++) {
    const prompt = headerPrompts[i];
    const field = headerFields[i];
    
    const response = await queryRAG(prompt, policyId, idToken);
    // @ts-ignore (we know the field exists in headerInfo)
    policyData.headerInfo[field] = response;
  }

  // Phase 2: Identify Coverage Sections
  const sectionGroupPrompts = [
    'Check if any of these coverage sections exist in the policy: Common Declarations, Schedules, General Liability, Employee Benefits Liability, Cyber Liability, Property, Inland Marine, Crime, Auto. Respond with comma-separated list of only those that exist.',
    'Check if any of these coverage sections exist in the policy: Garage/Garage Keepers, Workers Compensation, Umbrella/Excess, Professional Liability / E&O, Accident & Health, Animal Mortality, Equipment Breakdown, Directors & Officers. Respond with comma-separated list of only those that exist.',
    'Check if any of these coverage sections exist in the policy: Earthquake/Flood, Employment Practices Liability, Fiduciary Liability, Foreign, Group Travel Accident, Kidnap & Ransom, Liquor, Motor Carrier/Truckers. Respond with comma-separated list of only those that exist.',
    'Check if any of these coverage sections exist in the policy: Motor Truck Cargo, Ocean Marine, Pollution, Products Liability, Wind/Hail, Yacht & Hull, Other, Notes. Respond with comma-separated list of only those that exist.'
  ];

  let allSections: string[] = [];

  // Query for sections
  for (const prompt of sectionGroupPrompts) {
    const response = await queryRAG(prompt, policyId, idToken);
    if (response && response !== 'NF') {
      const sections = response.split(',').map(s => s.trim());
      allSections = [...allSections, ...sections];
    }
  }

  // Phase 3: For each section, identify fields
  for (const section of allSections) {
    const fieldPrompt = `List only the field names in the ${section} section of this policy. Format as comma-separated values. Use the exact field names as they would appear in the policy.`;
    const fieldsResponse = await queryRAG(fieldPrompt, policyId, idToken);
    
    if (fieldsResponse && fieldsResponse !== 'NF') {
      const fields = fieldsResponse.split(',').map(f => f.trim());
      policyData.sections[section] = [];
      
      // Phase 4: Extract values for each field
      for (const field of fields) {
        const valuePrompt = `Extract the ${field} value from the ${section} section of the policy. Return only the value and page number as: value|page`;
        const valueResponse = await queryRAG(valuePrompt, policyId, idToken);
        
        if (valueResponse && valueResponse !== 'NF') {
          const [value, page] = valueResponse.split('|');
          policyData.sections[section].push({
            fieldName: field,
            value: value.trim(),
            page: page.trim()
          });
        }
      }
    }
  }

  return policyData;
}

// Function to send a prompt to the RAG model
async function queryRAG(prompt: string, policyId: string, idToken: string | undefined): Promise<string> {
  try {
    // Create message for the RAG model
    const messages: ResponseMessage[] = [
      {
        content: `For policy ID ${policyId}: ${prompt}`,
        role: 'user'
      }
    ];

    // Create the request object
    const request: ChatAppRequest = {
      messages: messages,
      context: {
        overrides: {
          // Set appropriate overrides for policy extraction
          temperature: 0, // Lower temperature for factual extraction
          top: 3, // Limit to top 3 chunks for precision
          retrieval_mode: 'hybrid' as any, // Use hybrid retrieval for better results
          semantic_ranker: true, // Enable semantic ranking
          suggest_followup_questions: false, // No need for follow-up questions
          vector_fields: [], // Add appropriate vector fields if required
          language: 'en' // Specify the language, e.g., English
        }
      },
      session_state: null
    };

    // Make the API call to the backend
    // This would use one of the existing API methods like askApi or chatApi
    // For simplicity, we're assuming a direct fetch call here
    const response = await fetch('/api/ask', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        ...(idToken ? { Authorization: `Bearer ${idToken}` } : {})
      },
      body: JSON.stringify(request)
    });

    if (!response.ok) {
      throw new Error(`Error querying RAG model: ${response.statusText}`);
    }

    const data = await response.json();
    
    // Extract the response text
    const responseText = data.message.content.trim();
    return responseText === '' ? 'NF' : responseText;
  } catch (error) {
    console.error('Error querying RAG model:', error);
    return 'NF'; // Default to "Not Found" on error
  }
}

// Function to transform policy data to Excel format
function transformToExcelFormat(policyData: PolicyData): any[] {
  // Create Excel rows based on the CSV template structure
  const rows: any[] = [];

  // Add header row (you would customize this based on your CSV structure)
  rows.push({
    'Insured Name': policyData.headerInfo.insuredName,
    'Client Code': policyData.headerInfo.clientCode,
    'Policy Number': policyData.headerInfo.policyNumber,
    'Policy Dates': policyData.headerInfo.policyDates,
    'Policy Type': policyData.headerInfo.policyType,
    'Policy Premium': policyData.headerInfo.policyPremium,
    'Expiring Policy Premium': policyData.headerInfo.expiringPolicyPremium
  });

  // Add section data
  // This would need to be customized based on your exact Excel template structure
  Object.entries(policyData.sections).forEach(([sectionName, fields]) => {
    fields.forEach(field => {
      rows.push({
        'Section': sectionName,
        'Field': field.fieldName,
        'Value': field.value,
        'Page': field.page
      });
    });
  });

  return rows;
}

// The export button component
const ExportToExcelButton: React.FC<{
  policyId: string;
  fileName?: string;
}> = ({ policyId, fileName = 'policy-data.xlsx' }) => {
  const [isExporting, setIsExporting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const { idToken } = useAuthContext();

  const handleExport = async () => {
    setIsExporting(true);
    setError(null);

    try {
      // 1. Extract policy data
      const policyData = await extractPolicyData(policyId, idToken);
      
      // 2. Transform to Excel format
      const excelData = transformToExcelFormat(policyData);
      
      // 3. Generate Excel file
      const worksheet = XLSX.utils.json_to_sheet(excelData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Policy Data');
      
      // 4. Download the file
      XLSX.writeFile(workbook, fileName);
    } catch (err) {
      setError(`Failed to export: ${err instanceof Error ? err.message : String(err)}`);
    } finally {
      setIsExporting(false);
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
      
      {isExporting && <Spinner label="Exporting policy data to Excel..." />}
      
      {error && (
        <MessageBar messageBarType={MessageBarType.error}>
          {error}
        </MessageBar>
      )}
    </div>
  );
};

export default ExportToExcelButton;