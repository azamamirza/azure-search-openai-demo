// src/components/ExportToExcelButton.tsx
import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Button, Spinner, MessageBar, MessageBarType } from '@fluentui/react';
import { ChatAppRequest, ResponseMessage } from '../../api/models';

// Define types for policy extraction data
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

// Function to query the RAG model with the extraction prompts
async function extractPolicyData(policyId: string): Promise<PolicyData> {
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

  for (let i = 0; i < headerPrompts.length; i++) {
    const prompt = headerPrompts[i];
    const field = headerFields[i];
    const response = await queryRAG(prompt, policyId);
    // @ts-ignore
    policyData.headerInfo[field] = response;
  }

  const sectionGroupPrompts = [
    'Check if any of these coverage sections exist in the policy: Common Declarations, Schedules, General Liability, Employee Benefits Liability, Cyber Liability, Property, Inland Marine, Crime, Auto. Respond with comma-separated list of only those that exist.',
    'Check if any of these coverage sections exist in the policy: Garage/Garage Keepers, Workers Compensation, Umbrella/Excess, Professional Liability / E&O, Accident & Health, Animal Mortality, Equipment Breakdown, Directors & Officers. Respond with comma-separated list of only those that exist.',
    'Check if any of these coverage sections exist in the policy: Earthquake/Flood, Employment Practices Liability, Fiduciary Liability, Foreign, Group Travel Accident, Kidnap & Ransom, Liquor, Motor Carrier/Truckers. Respond with comma-separated list of only those that exist.',
    'Check if any of these coverage sections exist in the policy: Motor Truck Cargo, Ocean Marine, Pollution, Products Liability, Wind/Hail, Yacht & Hull, Other, Notes. Respond with comma-separated list of only those that exist.'
  ];

  let allSections: string[] = [];

  for (const prompt of sectionGroupPrompts) {
    const response = await queryRAG(prompt, policyId);
    if (response && response !== 'NF') {
      const sections = response.split(',').map(s => s.trim());
      allSections = [...allSections, ...sections];
    }
  }

  for (const section of allSections) {
    const fieldPrompt = `List only the field names in the ${section} section of this policy. Format as comma-separated values. Use the exact field names as they would appear in the policy.`;
    const fieldsResponse = await queryRAG(fieldPrompt, policyId);

    if (fieldsResponse && fieldsResponse !== 'NF') {
      const fields = fieldsResponse.split(',').map(f => f.trim());
      policyData.sections[section] = [];

      for (const field of fields) {
        const valuePrompt = `Extract the ${field} value from the ${section} section of the policy. Return only the value and page number as: value|page`;
        const valueResponse = await queryRAG(valuePrompt, policyId);

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

async function queryRAG(prompt: string, policyId: string): Promise<string> {
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
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(request)
    });

    if (!response.ok) {
      throw new Error(`Error querying RAG model: ${response.statusText}`);
    }

    const data = await response.json();
    const responseText = data.message.content.trim();
    return responseText === '' ? 'NF' : responseText;
  } catch (error) {
    console.error('Error querying RAG model:', error);
    return 'NF';
  }
}

function transformToExcelFormat(policyData: PolicyData): any[] {
  const rows: any[] = [];

  rows.push({
    'Insured Name': policyData.headerInfo.insuredName,
    'Client Code': policyData.headerInfo.clientCode,
    'Policy Number': policyData.headerInfo.policyNumber,
    'Policy Dates': policyData.headerInfo.policyDates,
    'Policy Type': policyData.headerInfo.policyType,
    'Policy Premium': policyData.headerInfo.policyPremium,
    'Expiring Policy Premium': policyData.headerInfo.expiringPolicyPremium
  });

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

const ExportToExcelButton: React.FC<{
  policyId: string;
  fileName?: string;
}> = ({ policyId, fileName = 'policy-data.xlsx' }) => {
  const [isExporting, setIsExporting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleExport = async () => {
    setIsExporting(true);
    setError(null);

    try {
      const policyData = await extractPolicyData(policyId);
      const excelData = transformToExcelFormat(policyData);

      const worksheet = XLSX.utils.json_to_sheet(excelData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Policy Data');

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
