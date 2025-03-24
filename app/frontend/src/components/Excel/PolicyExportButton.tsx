// src/components/PolicyExportButton.tsx
import React, { useState } from 'react';
import { Button, Spinner, Dialog, DialogType, DialogFooter, TextField, MessageBar, MessageBarType, DefaultButton, PrimaryButton, ProgressIndicator } from '@fluentui/react';
import { LoginContext } from '../../loginContext';
import { PolicyExtractionService } from '../../api/policyExtraction';
import { PolicyExcelExporter } from '../../utils/excelExport';

export interface PolicyExportButtonProps {
  className?: string;
  buttonText?: string;
  disabled?: boolean;
  onExportComplete?: () => void;
  onExportError?: (error: Error) => void;
}

// Define LoginContext interface to avoid TypeScript errors
interface LoginContextType {
  loggedIn: boolean;
  idToken?: string;
  setLoggedIn: (_: boolean) => void; // Add the missing property
}

const PolicyExportButton: React.FC<PolicyExportButtonProps> = ({
  className,
  buttonText = 'Export to Excel',
  disabled = false,
  onExportComplete,
  onExportError
}) => {
  // State
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [exportProgress, setExportProgress] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [policyNumber, setPolicyNumber] = useState('');
  const [fileName, setFileName] = useState('policy-data.xlsx');
  
  // Get authentication token
  // Cast the context to the expected type to avoid TypeScript errors
  const { loggedIn, idToken } = React.useContext(LoginContext as React.Context<LoginContextType>);
  
  const handleExportClick = () => {
    setIsDialogOpen(true);
  };
  
  const closeDialog = () => {
    setIsDialogOpen(false);
    setError(null);
    setExportProgress(0);
  };
  
  const handleExport = async () => {
    if (!policyNumber) {
      setError('Please enter a policy number');
      return;
    }
    
    setIsExporting(true);
    setError(null);
    setExportProgress(10);
    
    try {
      // Start the extraction process
      setExportProgress(20);
      const policyData = await PolicyExtractionService.extractPolicyData(
        policyNumber, 
        idToken,
        {
          useSemanticRanker: true,
          useHybridSearch: true,
          temperature: 0
        }
      );
      
      setExportProgress(80);
      
      // Export to Excel
      PolicyExcelExporter.exportToExcel(policyData, {
        fileName: fileName || 'policy-data.xlsx',
        sheetName: 'Policy Data'
      });
      
      setExportProgress(100);
      
      // Notify success
      if (onExportComplete) {
        onExportComplete();
      }
      
      // Close dialog after a short delay to show 100% progress
      setTimeout(() => {
        closeDialog();
      }, 1000);
    } catch (error) {
      console.error('Export failed:', error);
      setError(`Export failed: ${error instanceof Error ? error.message : String(error)}`);
      
      if (onExportError && error instanceof Error) {
        onExportError(error);
      }
    } finally {
      setIsExporting(false);
    }
  };
  
  return (
    <>
      <Button
        className={className}
        iconProps={{ iconName: 'ExcelDocument' }}
        onClick={handleExportClick}
        disabled={disabled || !loggedIn}
        primary
      >
        {buttonText}
      </Button>
      
      <Dialog
        hidden={!isDialogOpen}
        onDismiss={closeDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Export Policy Data to Excel',
          subText: 'Enter the policy number to extract data for.'
        }}
        modalProps={{
          isBlocking: isExporting,
          styles: { main: { maxWidth: 450 } }
        }}
      >
        <TextField
          label="Policy Number"
          required
          value={policyNumber}
          onChange={(e, newValue) => setPolicyNumber(newValue || '')}
          disabled={isExporting}
        />
        
        <TextField
          label="File Name"
          value={fileName}
          onChange={(e, newValue) => setFileName(newValue || 'policy-data.xlsx')}
          disabled={isExporting}
          placeholder="policy-data.xlsx"
        />
        
        {isExporting && (
          <ProgressIndicator
            label="Extracting policy data..."
            description={`This may take a few minutes. Please wait.`}
            percentComplete={exportProgress / 100}
          />
        )}
        
        {error && (
          <MessageBar messageBarType={MessageBarType.error}>
            {error}
          </MessageBar>
        )}
        
        <DialogFooter>
          <PrimaryButton onClick={handleExport} disabled={isExporting || !policyNumber}>
            Export
          </PrimaryButton>
          <DefaultButton onClick={closeDialog} disabled={isExporting}>
            Cancel
          </DefaultButton>
        </DialogFooter>
      </Dialog>
    </>
  );
};

export default PolicyExportButton;