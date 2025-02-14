import { Stack, Pivot, PivotItem } from "@fluentui/react";
import { useTranslation } from "react-i18next";
import styles from "./AnalysisPanel.module.css";
import { SupportingContent } from "../SupportingContent";
import { ChatAppResponse } from "../../api";
import { AnalysisPanelTabs } from "./AnalysisPanelTabs";
import { ThoughtProcess } from "./ThoughtProcess";
import { MarkdownViewer } from "../MarkdownViewer";
import { GraphVisualization } from "../GraphVisualization/";
import { useState, useEffect } from "react";

interface Props {
    className: string;
    activeTab: AnalysisPanelTabs;
    onActiveTabChanged: (tab: AnalysisPanelTabs) => void;
    activeCitation?: string;
    citationHeight: string;
    answer: ChatAppResponse;
    retrievalMode: string;
}

export const AnalysisPanel = ({ answer, activeTab, activeCitation, citationHeight, className, onActiveTabChanged, retrievalMode }: Props) => {
    const { t } = useTranslation();

    const isDisabledGraphTab = retrievalMode !== 'Graph';
    const isDisabledThoughtProcessTab = !answer.context.thoughts;
    const isDisabledSupportingContentTab = !answer.context.data_points;
    const isDisabledCitationTab = !activeCitation;

    return (
        <Pivot className={className} selectedKey={activeTab} onLinkClick={pivotItem => pivotItem && onActiveTabChanged(pivotItem.props.itemKey as AnalysisPanelTabs)}>
            <PivotItem itemKey={AnalysisPanelTabs.ThoughtProcessTab} headerText={t("headerTexts.thoughtProcess")} disabled={isDisabledThoughtProcessTab}>
                <ThoughtProcess thoughts={answer.context.thoughts || []} />
            </PivotItem>
            <PivotItem itemKey={AnalysisPanelTabs.SupportingContentTab} headerText={t("headerTexts.supportingContent")} disabled={isDisabledSupportingContentTab}>
                <SupportingContent supportingContent={answer.context.data_points} />
            </PivotItem>
            <PivotItem itemKey={AnalysisPanelTabs.CitationTab} headerText={t("headerTexts.citation")} disabled={isDisabledCitationTab}>
                {activeCitation && <MarkdownViewer src={activeCitation} />}
            </PivotItem>
            <PivotItem itemKey={AnalysisPanelTabs.GraphVisualization} headerText={t("Graph")} disabled={isDisabledGraphTab}>
                <GraphVisualization answer={answer} />
            </PivotItem>
            <PivotItem itemKey={AnalysisPanelTabs.GraphVisualization} headerText={t("Graph")}>
                <GraphVisualization
                    relations={answer.context.data_points || []}
                />
            </PivotItem>


        </Pivot>
    );
};
