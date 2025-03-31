// src/components/PolicyExportButton.tsx
import React, { useState } from 'react';
import {
  Button,
  Spinner,
  Dialog,
  DialogType,
  DialogFooter,
  TextField,
  MessageBar,
  MessageBarType,
  DefaultButton,
  PrimaryButton,
  ProgressIndicator
} from '@fluentui/react';
import { PolicyExtractionService } from '../../api/policyExtraction';
import { PolicyExcelExporter } from '../../utils/excelExport';

export interface PolicyExportButtonProps {
  className?: string;
  buttonText?: string;
  disabled?: boolean;
  onExportComplete?: () => void;
  onExportError?: (error: Error) => void;
}

const PolicyExportButton: React.FC<PolicyExportButtonProps> = ({
  className,
  buttonText = 'Export to Excel',
  disabled = false,
  onExportComplete,
  onExportError
}) => {
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [exportProgress, setExportProgress] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [policyNumber, setPolicyNumber] = useState('');
  const [fileName, setFileName] = useState('policy-data.xlsx');

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
      setExportProgress(20);
      const policyData = await PolicyExtractionService.extractPolicyData(
        policyNumber,
        undefined, // idToken not needed
        {
          useSemanticRanker: true,
          useHybridSearch: true,
          temperature: 0
        }
      );

      setExportProgress(80);

      PolicyExcelExporter.exportToExcel(policyData, {
        fileName: fileName || 'policy-data.xlsx',
        sheetName: 'Policy Data'
      });

      setExportProgress(100);

      if (onExportComplete) {
        onExportComplete();
      }

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
        disabled={disabled}
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
            description="This may take a few minutes. Please wait."
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
