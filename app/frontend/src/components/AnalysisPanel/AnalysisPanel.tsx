import { Stack, Pivot, PivotItem } from "@fluentui/react";
import { useTranslation } from "react-i18next";
import styles from "./AnalysisPanel.module.css";
import { SupportingContent } from "../SupportingContent";
import { ThoughtProcess } from "./ThoughtProcess";
import { MarkdownViewer } from "../MarkdownViewer";
import { GraphVisualization } from "../GraphVisualization";
import { ChatAppResponse } from "../../api";
import { AnalysisPanelTabs } from "./AnalysisPanelTabs";
import { useState, useEffect } from "react";

interface Props {
    className: string;
    activeTab: AnalysisPanelTabs;
    onActiveTabChanged: (tab: AnalysisPanelTabs) => void;
    activeCitation?: string;
    citationHeight: string;
    answer: ChatAppResponse;
}

export const AnalysisPanel = ({ answer, activeTab, activeCitation, citationHeight, className, onActiveTabChanged }: Props) => {
    const { t } = useTranslation();
    const isDisabledGraphTab = !answer.context.graphData;

    return (
        <Pivot
            className={className}
            selectedKey={activeTab}
            onLinkClick={(item) => item && onActiveTabChanged(item.props.itemKey as AnalysisPanelTabs)}
        >
            <PivotItem itemKey={AnalysisPanelTabs.ThoughtProcessTab} headerText={t("headerTexts.thoughtProcess")}> 
                <ThoughtProcess thoughts={answer.context.thoughts ?? []} />
            </PivotItem>

            <PivotItem itemKey={AnalysisPanelTabs.SupportingContentTab} headerText={t("headerTexts.supportingContent")}> 
                <SupportingContent supportingContent={answer.context.data_points} />
            </PivotItem>
            <PivotItem itemKey={AnalysisPanelTabs.CitationTab} headerText={t("headerTexts.citation")} disabled={isDisabledCitationTab}>
                {activeCitation && <MarkdownViewer src={activeCitation} />}
            </PivotItem>
            <PivotItem itemKey={AnalysisPanelTabs.GraphVisualization} headerText={t("Graph")}>
                <GraphVisualization
                    relations={answer.context.data_points || []}
                />
            </PivotItem>


            <PivotItem itemKey={AnalysisPanelTabs.GraphVisualization} headerText={t("Graph")} 
                headerButtonProps={isDisabledGraphTab ? { disabled: false, style: { color: 'grey' } } : undefined}>
                <GraphVisualization nodes={answer.context.graphData?.nodes ?? []} edges={answer.context.graphData?.edges ?? []} />
            </PivotItem>
        </Pivot>
    );
};
