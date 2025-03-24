// src/api/policyExtractionService.ts
import { getHeaders } from './api';
import { ChatAppRequest, ChatAppResponse, ResponseMessage, RetrievalMode } from './models';

export interface PolicyExtractionOptions {
  useSemanticRanker?: boolean;
  useHybridSearch?: boolean;
  maxTokens?: number;
  temperature?: number;
}

export interface PolicySection {
  sectionName: string;
  fields: PolicyField[];
}

export interface PolicyField {
  name: string;
  value: string;
  page?: string;
}

export interface PolicyData {
  headerInfo: {
    insuredName: string;
    clientCode: string;
    policyNumber: string;
    policyDates: string;
    policyType: string;
    policyPremium: string;
    expiringPolicyPremium: string;
  };
  sections: PolicySection[];
}

/**
 * Class that handles policy data extraction using the RAG backend
 */
export class PolicyExtractionService {
  // Default options for policy extraction
  private static defaultOptions: PolicyExtractionOptions = {
    useSemanticRanker: true,
    useHybridSearch: true,
    maxTokens: 1000,
    temperature: 0
  };

  /**
   * Extracts policy data using the prompting schema from the policy document
   * 
   * @param policyNumber - The policy number to extract data for
   * @param idToken - Authentication token for the request
   * @param options - Optional extraction options
   * @returns Extracted policy data
   */
  public static async extractPolicyData(
    policyNumber: string,
    idToken?: string,
    options?: Partial<PolicyExtractionOptions>
  ): Promise<PolicyData> {
    // Merge options with defaults
    const extractOptions = { ...this.defaultOptions, ...options };
    
    // Initialize the policy data
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
      sections: []
    };

    try {
      // Phase 1: Extract header information
      await this.extractHeaderInformation(policyData, policyNumber, idToken, extractOptions);
      
      // Phase 2: Identify coverage sections
      const sections = await this.identifyCoverageSections(policyNumber, idToken, extractOptions);
      
      // Phase 3 & 4: Extract fields and values for each section
      for (const sectionName of sections) {
        await this.extractSectionData(policyData, sectionName, policyNumber, idToken, extractOptions);
      }
      
      return policyData;
    } catch (error) {
      console.error('Error extracting policy data:', error);
      throw new Error(`Failed to extract policy data: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * Extracts header information from the policy
   */
  private static async extractHeaderInformation(
    policyData: PolicyData,
    policyNumber: string,
    idToken?: string,
    options?: PolicyExtractionOptions
  ): Promise<void> {
    const headerFields = [
      { field: 'insuredName', prompt: 'Extract the Insured Name from the policy. Return only the value, nothing else.' },
      { field: 'clientCode', prompt: 'Extract the Client Code from the policy. Return only the value, nothing else.' },
      { field: 'policyNumber', prompt: 'Extract the Policy Number from the policy. Return only the value, nothing else.' },
      { field: 'policyDates', prompt: 'Extract the Policy Effective Date and Expiration Date from the policy. Return only the value in format MM/DD/YY - MM/DD/YY, nothing else.' },
      { field: 'policyType', prompt: 'Extract the Policy Type from the policy. Return only the value, nothing else.' },
      { field: 'policyPremium', prompt: 'Extract the Policy Premium from the policy. Return only the value with $ and commas, nothing else.' },
      { field: 'expiringPolicyPremium', prompt: 'Extract the Expiring Policy Premium from the policy. Return only the value with $ and commas, nothing else.' }
    ];

    for (const { field, prompt } of headerFields) {
      try {
        const value = await this.queryChatAPI(prompt, policyNumber, idToken, options);
        // @ts-ignore (we know these fields exist)
        policyData.headerInfo[field] = value || 'NF';
      } catch (error) {
        console.error(`Error extracting ${field}:`, error);
        // @ts-ignore
        policyData.headerInfo[field] = 'NF';
      }
    }
  }

  /**
   * Identifies coverage sections in the policy
   */
  private static async identifyCoverageSections(
    policyNumber: string,
    idToken?: string,
    options?: PolicyExtractionOptions
  ): Promise<string[]> {
    const sectionGroupPrompts = [
      'Check if any of these coverage sections exist in the policy: Common Declarations, Schedules, General Liability, Employee Benefits Liability, Cyber Liability, Property, Inland Marine, Crime, Auto. Respond with comma-separated list of only those that exist.',
      'Check if any of these coverage sections exist in the policy: Garage/Garage Keepers, Workers Compensation, Umbrella/Excess, Professional Liability / E&O, Accident & Health, Animal Mortality, Equipment Breakdown, Directors & Officers. Respond with comma-separated list of only those that exist.',
      'Check if any of these coverage sections exist in the policy: Earthquake/Flood, Employment Practices Liability, Fiduciary Liability, Foreign, Group Travel Accident, Kidnap & Ransom, Liquor, Motor Carrier/Truckers. Respond with comma-separated list of only those that exist.',
      'Check if any of these coverage sections exist in the policy: Motor Truck Cargo, Ocean Marine, Pollution, Products Liability, Wind/Hail, Yacht & Hull, Other, Notes. Respond with comma-separated list of only those that exist.'
    ];

    let allSections: string[] = [];

    for (const prompt of sectionGroupPrompts) {
      try {
        const response = await this.queryChatAPI(prompt, policyNumber, idToken, options);
        if (response && response !== 'NF') {
          const sections = response.split(',').map(s => s.trim()).filter(s => s);
          allSections = [...allSections, ...sections];
        }
      } catch (error) {
        console.error('Error identifying sections:', error);
      }
    }

    return allSections;
  }

  /**
   * Extracts field data for a specific section
   */
  private static async extractSectionData(
    policyData: PolicyData,
    sectionName: string,
    policyNumber: string,
    idToken?: string,
    options?: PolicyExtractionOptions
  ): Promise<void> {
    try {
      // Phase 3: Identify fields for this section
      const fieldsPrompt = `List only the field names in the ${sectionName} section of this policy. Format as comma-separated values. Use the exact field names as they would appear in the policy.`;
      const fieldsResponse = await this.queryChatAPI(fieldsPrompt, policyNumber, idToken, options);
      
      if (!fieldsResponse || fieldsResponse === 'NF') {
        return;
      }
      
      const fields = fieldsResponse.split(',').map(f => f.trim()).filter(f => f);
      const sectionData: PolicySection = {
        sectionName,
        fields: []
      };
      
      // Phase 4: Extract values for each field
      for (const fieldName of fields) {
        const valuePrompt = `Extract the ${fieldName} value from the ${sectionName} section of the policy. Return only the value and page number as: value|page`;
        const valueResponse = await this.queryChatAPI(valuePrompt, policyNumber, idToken, options);
        
        if (valueResponse && valueResponse !== 'NF') {
          // Split value and page
          const parts = valueResponse.split('|');
          const value = parts[0]?.trim() || 'NF';
          const page = parts[1]?.trim() || '';
          
          sectionData.fields.push({
            name: fieldName,
            value,
            page
          });
        } else {
          sectionData.fields.push({
            name: fieldName,
            value: 'NF'
          });
        }
      }
      
      policyData.sections.push(sectionData);
    } catch (error) {
      console.error(`Error extracting section data for ${sectionName}:`, error);
    }
  }

  /**
   * Queries the chat API with a specific prompt
   */
  private static async queryChatAPI(
    prompt: string,
    policyNumber: string,
    idToken?: string,
    options?: PolicyExtractionOptions
  ): Promise<string> {
    try {
      const headers = await getHeaders(idToken);
      
      const messages: ResponseMessage[] = [
        {
          role: 'system',
          content: 'You are an insurance policy extraction assistant. Extract specific data points from policy documents using precise, factual language. When asked to extract values, provide only the requested information in the specified format without explanations or additional text. For non-existent information use "NF" (Not Found), for inapplicable fields use "NA" (Not Applicable), for schedules with more than 2 items use "SS" (See Schedule), and for data that should exist but isn\'t found use "NL" (Not Listed). Format monetary values with $ and commas (e.g., "$1,000,000") and dates in MM/DD/YY format.'
        },
        {
          role: 'user',
          content: `For policy ${policyNumber}: ${prompt}`
        }
      ];
      
      const request: ChatAppRequest = {
        messages,
        context: {
          overrides: {
            retrieval_mode: options?.useHybridSearch ? 'hybrid' as RetrievalMode : 'text' as RetrievalMode,
            semantic_ranker: options?.useSemanticRanker,
            temperature: options?.temperature,
            top: 5,
            suggest_followup_questions: false,
            vector_fields: [], // Add appropriate vector fields if required
            language: 'en' // Specify the language, e.g., 'en' for English
          }
        },
        session_state: null
      };
      
      const response = await fetch('/ask', {
        method: 'POST',
        headers: { ...headers, 'Content-Type': 'application/json' },
        body: JSON.stringify(request)
      });
      
      if (!response.ok) {
        throw new Error(`API request failed with status ${response.status}`);
      }
      
      const data = await response.json() as ChatAppResponse;
      return data.message.content.trim() || 'NF';
    } catch (error) {
      console.error('Error querying chat API:', error);
      return 'NF';
    }
  }
}
