// src/utils/excelExport.ts
import * as XLSX from 'xlsx';
import { PolicyData, PolicySection, PolicyField } from '../api/policyExtraction';

/**
 * Interface for configuring the Excel export
 */
export interface ExcelExportOptions {
  fileName?: string;
  sheetName?: string;
  includeHeaderRow?: boolean;
  dateFormat?: string;
}

/**
 * Utility class for exporting policy data to Excel
 */
export class PolicyExcelExporter {
  private static defaultOptions: ExcelExportOptions = {
    fileName: 'policy-data.xlsx',
    sheetName: 'Policy Data',
    includeHeaderRow: true,
    dateFormat: 'MM/DD/YY'
  };

  /**
   * Exports policy data to an Excel file
   * 
   * @param policyData - The policy data to export
   * @param options - Optional configuration for the export
   */
  public static exportToExcel(policyData: PolicyData, options?: Partial<ExcelExportOptions>): void {
    // Merge options with defaults
    const exportOptions = { ...this.defaultOptions, ...options };
    
    try {
      // Transform the policy data into a 2D array for Excel
      const excelData = this.transformPolicyDataToExcelFormat(policyData);
      
      // Create a new workbook
      const workbook = XLSX.utils.book_new();
      
      // Create a worksheet from the data
      const worksheet = XLSX.utils.aoa_to_sheet(excelData);
      
      // Add the worksheet to the workbook
      XLSX.utils.book_append_sheet(workbook, worksheet, exportOptions.sheetName);
      
      // Write the workbook to a file and trigger download
      XLSX.writeFile(workbook, exportOptions.fileName!);
    } catch (error) {
      console.error('Error exporting to Excel:', error);
      throw new Error(`Failed to export to Excel: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * Transforms policy data into a format suitable for Excel
   * 
   * @param policyData - The policy data to transform
   * @returns A 2D array representation of the data
   */
  private static transformPolicyDataToExcelFormat(policyData: PolicyData): any[][] {
    // Create the header and policy information rows
    const headerRow = [
      'Type', 'Property', 'Value', 'Page', 'Notes'
    ];
    
    const rows: any[][] = [headerRow];
    
    // Add header information
    rows.push(['Header', 'Insured Name', policyData.headerInfo.insuredName || 'NF', '', '']);
    rows.push(['Header', 'Client Code', policyData.headerInfo.clientCode || 'NF', '', '']);
    rows.push(['Header', 'Policy Number', policyData.headerInfo.policyNumber || 'NF', '', '']);
    rows.push(['Header', 'Policy Dates', policyData.headerInfo.policyDates || 'NF', '', '']);
    rows.push(['Header', 'Policy Type', policyData.headerInfo.policyType || 'NF', '', '']);
    rows.push(['Header', 'Policy Premium', policyData.headerInfo.policyPremium || 'NF', '', '']);
    rows.push(['Header', 'Expiring Policy Premium', policyData.headerInfo.expiringPolicyPremium || 'NF', '', '']);
    
    // Add a blank row for separation
    rows.push(['', '', '', '', '']);
    
    // Add section information
    policyData.sections.forEach(section => {
      // Add section header
      rows.push([section.sectionName, '', '', '', '']);
      
      // Add fields for this section
      section.fields.forEach(field => {
        rows.push([
          section.sectionName,
          field.name,
          field.value || 'NF',
          field.page || '',
          ''  // Notes column left blank
        ]);
      });
      
      // Add a blank row after each section
      rows.push(['', '', '', '', '']);
    });
    
    return rows;
  }
  
  /**
   * Maps policy data to match the specific CSV template structure
   * 
   * @param policyData - The policy data to format
   * @param templateHeaders - The headers from the template CSV
   * @returns Data formatted according to the template
   */
  public static mapToTemplateFormat(policyData: PolicyData, templateHeaders: string[]): any[] {
    // This method would map the policy data to match the exact structure
    // of your CSV template. Since the full template structure isn't clear
    // from the provided CSV (it has 1019 columns), this would need to be
    // customized to your specific requirements.
    
    // Example placeholder implementation:
    const mappedData: any[] = [];
    const row: any = {};
    
    // Map known header fields
    templateHeaders.forEach(header => {
      switch (header) {
        case 'Insured Name':
          row[header] = policyData.headerInfo.insuredName || 'NF';
          break;
        case 'Client Code':
          row[header] = policyData.headerInfo.clientCode || 'NF';
          break;
        case 'Policy Number':
          row[header] = policyData.headerInfo.policyNumber || 'NF';
          break;
        // Map other fields as needed
        default:
          // Try to find matching section/field
          let found = false;
          policyData.sections.forEach(section => {
            const matchingField = section.fields.find(field => field.name === header);
            if (matchingField) {
              row[header] = matchingField.value;
              found = true;
            }
          });
          
          if (!found) {
            row[header] = '';
          }
      }
    });
    
    mappedData.push(row);
    return mappedData;
  }
  
  /**
   * Generates an Excel file from the template with policy data
   * 
   * @param policyData - The policy data to include
   * @param templatePath - Path to the template file
   * @param options - Export options
   */
  public static async exportUsingTemplate(
    policyData: PolicyData, 
    templatePath: string, 
    options?: Partial<ExcelExportOptions>
  ): Promise<void> {
    const exportOptions = { ...this.defaultOptions, ...options };
    
    try {
      // Fetch the template file
      const response = await fetch(templatePath);
      if (!response.ok) {
        throw new Error(`Failed to fetch template: ${response.status}`);
      }
      
      const templateBuffer = await response.arrayBuffer();
      
      // Load the template workbook
      const workbook = XLSX.read(templateBuffer, { type: 'array' });
      
      // Get the first sheet
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      // Convert to JSON to manipulate the data
      const templateData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      // Assuming first row contains headers
      const headers = templateData[0] as string[];
      
      // Map policy data to template format
      const mappedData = this.mapToTemplateFormat(policyData, headers);
      
      // Create a new worksheet with the mapped data
      const newWorksheet = XLSX.utils.json_to_sheet(mappedData, { header: headers });
      
      // Add the worksheet to a new workbook
      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, exportOptions.sheetName);
      
      // Write to file and download
      XLSX.writeFile(newWorkbook, exportOptions.fileName!);
    } catch (error) {
      console.error('Error exporting using template:', error);
      throw new Error(`Failed to export using template: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
}
