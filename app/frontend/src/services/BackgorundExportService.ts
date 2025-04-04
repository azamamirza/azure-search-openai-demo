// src/services/BackgroundExportService.ts
import { EventEmitter } from 'events';
import { endpoints } from './endpoints';
import { ChatAppRequest, ResponseMessage } from '../api/models';

// Reuse existing types
export interface PolicyHeaderInfo {
  insuredName: string;
  clientCode: string;
  policyNumber: string;
  policyDates: string;
  policyType: string;
  policyPremium: string;
  expiringPolicyPremium: string;
}

export interface SectionFieldValue {
  fieldName: string;
  value: string;
  page: string;
}

export interface PolicyData {
  headerInfo: PolicyHeaderInfo;
  sections: {
    [sectionName: string]: SectionFieldValue[];
  };
}

export interface ExportJob {
  id: string;
  policyId: string;
  fileName: string;
  status: 'running' | 'completed' | 'failed';
  progress: number;
  error?: string;
  result?: PolicyData;
  startTime: Date;
}

// Singleton service for background policy exports
class BackgroundExportService extends EventEmitter {
  private static instance: BackgroundExportService;
  private activeJobs: Map<string, ExportJob> = new Map();
  private apiEndpoint: string = 'https://capps-backend-qpblwwfjsavwq.bluepebble-f9101277.centralus.azurecontainerapps.io/api/v1/query/';

  private constructor() {
    super();
  }

  public static getInstance(): BackgroundExportService {
    if (!BackgroundExportService.instance) {
      BackgroundExportService.instance = new BackgroundExportService();
    }
    return BackgroundExportService.instance;
  }

  /**
   * Start a background policy export
   */
  public startExport(policyId: string, fileName: string, idToken?: string): string {
    // Create a unique job ID
    const jobId = `job_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;
    
    // Create job object
    const job: ExportJob = {
      id: jobId,
      policyId,
      fileName,
      status: 'running',
      progress: 0,
      startTime: new Date()
    };
    
    // Store the job
    this.activeJobs.set(jobId, job);
    
    // Emit job started event
    this.emit('job_started', job);
    
    // Start the extraction process in the background
    this.runExtractionInBackground(job, idToken);
    
    return jobId;
  }

  /**
   * Get current job status
   */
  public getJobStatus(jobId: string): ExportJob | undefined {
    return this.activeJobs.get(jobId);
  }

  /**
   * Get all active jobs
   */
  public getAllJobs(): ExportJob[] {
    return Array.from(this.activeJobs.values());
  }

  /**
   * Run the extraction process in the background
   */
  private async runExtractionInBackground(job: ExportJob, idToken?: string): Promise<void> {
    try {
      // Start the extraction process - using the same code as the original ExportToExcelButton
      const policyData = await this.extractPolicyData(job.policyId, idToken, (progress) => {
        // Update job progress
        this.updateJobProgress(job.id, progress);
      });
      
      // Update job with results
      const updatedJob: ExportJob = {
        ...job,
        status: 'completed',
        progress: 100,
        result: policyData
      };
      
      this.activeJobs.set(job.id, updatedJob);
      
      // Emit job completed event
      this.emit('job_completed', updatedJob);
    } catch (error) {
      // Update job with error
      const updatedJob: ExportJob = {
        ...job,
        status: 'failed',
        error: error instanceof Error ? error.message : String(error)
      };
      
      this.activeJobs.set(job.id, updatedJob);
      
      // Emit job failed event
      this.emit('job_failed', updatedJob);
    }
  }

  /**
   * Update job progress
   */
  private updateJobProgress(jobId: string, progress: number): void {
    const job = this.activeJobs.get(jobId);
    if (job) {
      const updatedJob = { ...job, progress };
      this.activeJobs.set(jobId, updatedJob);
      this.emit('job_updated', updatedJob);
    }
  }

  /**
   * Extract policy data - reusing the existing extraction logic
   */
  private async extractPolicyData(
    policyId: string, 
    idToken?: string,
    onProgress?: (progress: number) => void
  ): Promise<PolicyData> {
    // Initialize the policy data structure (same as original code)
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

    if (onProgress) onProgress(5);

    // Query for header information
    for (let i = 0; i < headerPrompts.length; i++) {
      const prompt = headerPrompts[i];
      const field = headerFields[i];
      
      const response = await this.queryRAG(prompt, policyId, idToken);
      // @ts-ignore (we know the field exists in headerInfo)
      policyData.headerInfo[field] = response;
      
      if (onProgress) onProgress(10 + (i * 5));
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
    for (let i = 0; i < sectionGroupPrompts.length; i++) {
      const prompt = sectionGroupPrompts[i];
      const response = await this.queryRAG(prompt, policyId, idToken);
      
      if (response && response !== 'NF') {
        const sections = response.split(',').map(s => s.trim());
        allSections = [...allSections, ...sections];
      }
      
      if (onProgress) onProgress(45 + (i * 5));
    }

    // Phase 3: For each section, identify fields
    const progressPerSection = allSections.length ? (40 / allSections.length) : 0;
    
    for (let i = 0; i < allSections.length; i++) {
      const section = allSections[i];
      const fieldPrompt = `List only the field names in the ${section} section of this policy. Format as comma-separated values. Use the exact field names as they would appear in the policy.`;
      const fieldsResponse = await this.queryRAG(fieldPrompt, policyId, idToken);
      
      if (fieldsResponse && fieldsResponse !== 'NF') {
        const fields = fieldsResponse.split(',').map(f => f.trim());
        policyData.sections[section] = [];
        
        // Phase 4: Extract values for each field
        for (const field of fields) {
          const valuePrompt = `Extract the ${field} value from the ${section} section of the policy. Return only the value and page number as: value|page`;
          const valueResponse = await this.queryRAG(valuePrompt, policyId, idToken);
          
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
      
      if (onProgress) {
        const progress = Math.min(65 + (i * progressPerSection), 95);
        onProgress(progress);
      }
    }

    if (onProgress) onProgress(98);
    return policyData;
  }

  /**
   * Query the RAG model - similar to the original code
   */
  private async queryRAG(prompt: string, policyId: string, idToken?: string): Promise<string> {
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
            temperature: 0,
            top: 3,
            retrieval_mode: 'hybrid' as any,
            semantic_ranker: true,
            suggest_followup_questions: false,
            vector_fields: [],
            language: 'en'
          }
        },
        session_state: null
      };

      // Make the API call to the backend
      const response = await fetch(this.apiEndpoint, {
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
}

export const backgroundExportService = BackgroundExportService.getInstance();
export default backgroundExportService;