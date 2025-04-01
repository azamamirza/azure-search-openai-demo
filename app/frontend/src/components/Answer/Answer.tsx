import { useMemo, useState, useEffect } from "react";
import { Stack, IconButton } from "@fluentui/react";
import { useTranslation } from "react-i18next";
import DOMPurify from "dompurify";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import rehypeRaw from "rehype-raw";

import styles from "./Answer.module.css";
import { ChatAppResponse, getCitationFilePath, SpeechConfig } from "../../api";
import { parseAnswerToHtml } from "./AnswerParser";
import { AnswerIcon } from "./AnswerIcon";
import { SpeechOutputBrowser } from "./SpeechOutputBrowser";
import { SpeechOutputAzure } from "./SpeechOutputAzure";
import ExportToExcelButton from "../../components/Excel/ExportToExcelButton";

interface Props {
    answer: ChatAppResponse;
    index: number;
    speechConfig: SpeechConfig;
    isSelected?: boolean;
    isStreaming: boolean;
    onCitationClicked: (filePath: string) => void;
    onThoughtProcessClicked: () => void;
    onSupportingContentClicked: () => void;
    onFollowupQuestionClicked?: (question: string) => void;
    showFollowupQuestions?: boolean;
    showSpeechOutputBrowser?: boolean;
    showSpeechOutputAzure?: boolean;
}

export const Answer = ({
    answer,
    index,
    speechConfig,
    isSelected,
    isStreaming,
    onCitationClicked,
    onThoughtProcessClicked,
    onSupportingContentClicked,
    onFollowupQuestionClicked,
    showFollowupQuestions,
    showSpeechOutputAzure,
    showSpeechOutputBrowser,
}: Props) => {
    const followupQuestions = answer.context?.followup_questions;
    const parsedAnswer = useMemo(() => parseAnswerToHtml(answer, isStreaming, onCitationClicked), [answer]);
    const { t } = useTranslation();
    const sanitizedAnswerHtml = DOMPurify.sanitize(parsedAnswer.answerHtml);
    const [copied, setCopied] = useState(false);
    const [policyId, setPolicyId] = useState<string>("");
    const [showExportButton, setShowExportButton] = useState(false);

    // Extract policy ID from the answer or data points
    useEffect(() => {
        // Method 1: Try to find policy ID in the answer text
        const extractPolicyIdFromText = () => {
            // Common patterns for policy IDs in text
            const patterns = [
                /policy\s+id[:\s]+([A-Za-z0-9-_]+)/i,
                /policy[:\s#]+([A-Za-z0-9-_]+)/i,
                /policy\s+number[:\s]+([A-Za-z0-9-_]+)/i,
                /id[:\s]+([A-Za-z0-9-_]+)/i,
            ];

            // Try each pattern
            for (const pattern of patterns) {
                const match = sanitizedAnswerHtml.match(pattern);
                if (match && match[1]) {
                    return match[1].trim();
                }
            }
            return "";
        };

        // Method 2: Look for policy ID in data points/context
        const extractPolicyIdFromContext = () => {
            const dataPoints = answer.context?.data_points;
            if (dataPoints && Array.isArray(dataPoints)) {
                for (const dataPoint of dataPoints) {
                    if (typeof dataPoint === 'string') {
                        // Check if the data point contains policy ID information
                        const policyIdMatch = dataPoint.match(/policy\s+id[:\s]+([A-Za-z0-9-_]+)/i) || 
                                             dataPoint.match(/policy[:\s#]+([A-Za-z0-9-_]+)/i) ||
                                             dataPoint.match(/policy\s+number[:\s]+([A-Za-z0-9-_]+)/i);
                        
                        if (policyIdMatch && policyIdMatch[1]) {
                            return policyIdMatch[1].trim();
                        }
                    }
                }
            }
            return "";
        };

        // Try both methods
        const idFromText = extractPolicyIdFromText();
        const idFromContext = extractPolicyIdFromContext();
        
        const foundPolicyId = idFromText || idFromContext;
        
        if (foundPolicyId) {
            setPolicyId(foundPolicyId);
            setShowExportButton(true);
        } else {
            // If no policy ID found, check if answer is about a policy at all
            const isPolicyRelated = 
                sanitizedAnswerHtml.match(/policy|insurance|coverage|premium|insured/i) ||
                (answer.context?.data_points && Array.isArray(answer.context.data_points) && 
                  answer.context.data_points.some(dp => 
                    typeof dp === 'string' && dp.match(/policy|insurance|coverage|premium|insured/i)
                ));
                
            setShowExportButton(!!isPolicyRelated);
        }
    }, [answer, sanitizedAnswerHtml]);

    const handleCopy = () => {
        // Single replace to remove all HTML tags to remove the citations
        const textToCopy = sanitizedAnswerHtml.replace(/<a [^>]*><sup>\d+<\/sup><\/a>|<[^>]+>/g, "");

        navigator.clipboard
            .writeText(textToCopy)
            .then(() => {
                setCopied(true);
                setTimeout(() => setCopied(false), 2000);
            })
            .catch(err => console.error("Failed to copy text: ", err));
    };

    return (
        <Stack className={`${styles.answerContainer} ${isSelected && styles.selected}`} verticalAlign="space-between">
            <Stack.Item>
                <Stack horizontal horizontalAlign="space-between">
                    <AnswerIcon />
                    <div>
                        <IconButton
                            style={{ color: "black" }}
                            iconProps={{ iconName: copied ? "CheckMark" : "Copy" }}
                            title={copied ? t("tooltips.copied") : t("tooltips.copy")}
                            ariaLabel={copied ? t("tooltips.copied") : t("tooltips.copy")}
                            onClick={handleCopy}
                        />
                        <IconButton
                            style={{ color: "black" }}
                            iconProps={{ iconName: "Lightbulb" }}
                            title={t("tooltips.showThoughtProcess")}
                            ariaLabel={t("tooltips.showThoughtProcess")}
                            onClick={() => onThoughtProcessClicked()}
                            disabled={!answer.context.thoughts?.length}
                        />
                        <IconButton
                            style={{ color: "black" }}
                            iconProps={{ iconName: "ClipboardList" }}
                            title={t("tooltips.showSupportingContent")}
                            ariaLabel={t("tooltips.showSupportingContent")}
                            onClick={() => onSupportingContentClicked()}
                            disabled={!answer.context.data_points}
                        />
                        {showSpeechOutputAzure && (
                            <SpeechOutputAzure answer={sanitizedAnswerHtml} index={index} speechConfig={speechConfig} isStreaming={isStreaming} />
                        )}
                        {showSpeechOutputBrowser && <SpeechOutputBrowser answer={sanitizedAnswerHtml} />}
                    </div>
                </Stack>
            </Stack.Item>

            <Stack.Item grow>
                <div className={styles.answerText}>
                    <ReactMarkdown children={sanitizedAnswerHtml} rehypePlugins={[rehypeRaw]} remarkPlugins={[remarkGfm]} />
                </div>
            </Stack.Item>

            {!!parsedAnswer.citations.length && (
                <Stack.Item>
                    <Stack horizontal wrap tokens={{ childrenGap: 5 }}>
                        <span className={styles.citationLearnMore}>{t("citationWithColon")}</span>
                        {parsedAnswer.citations.map((x, i) => {
                            const path = getCitationFilePath(x);
                            return (
                                <a key={i} className={styles.citation} title={x} onClick={() => onCitationClicked(path)}>
                                    {`${++i}. ${x}`}
                                </a>
                            );
                        })}
                    </Stack>
                </Stack.Item>
            )}

            {/* Export to Excel section - always show if policy-related */}
            {showExportButton && (
                <Stack.Item>
                    <div style={{ marginTop: 12, display: "flex", justifyContent: "center", alignItems: "center" }}>
                        <ExportToExcelButton 
                            policyId={policyId} 
                            fileName={`Policy_${policyId || 'Data'}_${new Date().toISOString().split('T')[0]}.xlsx`}
                        />
                    </div>
                </Stack.Item>
            )}

            {!!followupQuestions?.length && showFollowupQuestions && onFollowupQuestionClicked && (
                <Stack.Item>
                    <Stack horizontal wrap className={`${!!parsedAnswer.citations.length ? styles.followupQuestionsList : ""}`} tokens={{ childrenGap: 6 }}>
                        <span className={styles.followupQuestionLearnMore}>{t("followupQuestions")}</span>
                        {followupQuestions.map((x, i) => {
                            return (
                                <a key={i} className={styles.followupQuestion} title={x} onClick={() => onFollowupQuestionClicked(x)}>
                                    {`${x}`}
                                </a>
                            );
                        })}
                    </Stack>
                </Stack.Item>
            )}
        </Stack>
    );
};