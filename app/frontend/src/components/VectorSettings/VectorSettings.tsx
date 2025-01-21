import { useEffect, useState } from "react";
import { Stack, IDropdownOption, Dropdown, IDropdownProps } from "@fluentui/react";
import { useId } from "@fluentui/react-hooks";
import { useTranslation } from "react-i18next";

import styles from "./VectorSettings.module.css";
import { HelpCallout } from "../../components/HelpCallout";
import { ChatAppRequest, graphRagApi, RetrievalMode, VectorFieldOptions } from "../../api";
import { getToken } from "../../authConfig";

interface Props {
    showImageOptions?: boolean;
    defaultRetrievalMode: RetrievalMode;
    updateRetrievalMode: (retrievalMode: RetrievalMode) => void;
    updateVectorFields: (options: VectorFieldOptions[]) => void;
}

export const VectorSettings = ({ updateRetrievalMode, updateVectorFields, showImageOptions, defaultRetrievalMode }: Props) => {
    const [retrievalMode, setRetrievalMode] = useState<RetrievalMode>(RetrievalMode.Hybrid);
    const [vectorFieldOption, setVectorFieldOption] = useState<VectorFieldOptions>(VectorFieldOptions.Both);

    const onRetrievalModeChange = async (
        _ev: React.FormEvent<HTMLDivElement>,
        option?: IDropdownOption<RetrievalMode> | undefined
    ) => {
        const selectedMode = option?.data || RetrievalMode.Hybrid;
        setRetrievalMode(selectedMode);
        updateRetrievalMode(selectedMode);
    
        if (selectedMode === RetrievalMode.Graph) {
            try {
                // Ensure proper authentication token retrieval
                const idToken = await getToken();
    
                // Prepare request payload
                const request: ChatAppRequest = {
                    messages: [{ content: "example query", role: "user" }],
                    context: {
                        overrides: {
                            retrieval_mode: RetrievalMode.Graph,
                            vector_fields: [VectorFieldOptions.Embedding], // or appropriate value
                            language: "en" // or appropriate value
                        }
                    },
                    session_state: null
                };
    
                console.log("Sending Graph RAG API request:", request);
    
                // Call the Graph RAG API
                const response = await graphRagApi(request, selectedMode, idToken);
    
                console.log("Graph RAG API response:", response);
            } catch (error) {
                if (error instanceof Error) {
                    console.error("Error fetching Graph RAG data:", error.message);
                } else {
                    console.error("Unknown error occurred during Graph RAG API call.");
                }
            }
        }
    };
    

    const onVectorFieldsChange = (_ev: React.FormEvent<HTMLDivElement>, option?: IDropdownOption<RetrievalMode> | undefined) => {
        setVectorFieldOption(option?.key as VectorFieldOptions);
        updateVectorFields([option?.key as VectorFieldOptions]);
    };

    useEffect(() => {
        showImageOptions
            ? updateVectorFields([VectorFieldOptions.Embedding, VectorFieldOptions.ImageEmbedding])
            : updateVectorFields([VectorFieldOptions.Embedding]);
    }, [showImageOptions]);

    const retrievalModeId = useId("retrievalMode");
    const retrievalModeFieldId = useId("retrievalModeField");
    const vectorFieldsId = useId("vectorFields");
    const vectorFieldsFieldId = useId("vectorFieldsField");
    const { t } = useTranslation();

    return (
        <Stack className={styles.container} tokens={{ childrenGap: 10 }}>
            <Dropdown
    id={retrievalModeFieldId}
    label={t("labels.retrievalMode.label")}
    selectedKey={retrievalMode.toString()}
    options={[
        { key: "hybrid", text: t("labels.retrievalMode.options.hybrid"), data: RetrievalMode.Hybrid },
        { key: "vectors", text: t("labels.retrievalMode.options.vectors"), data: RetrievalMode.Vectors },
        { key: "text", text: t("labels.retrievalMode.options.texts"), data: RetrievalMode.Text },
        { key: "graph", text: t("labels.retrievalMode.options.graph"), data: RetrievalMode.Graph }
    ]}
    onChange={onRetrievalModeChange}
    required
    aria-labelledby={retrievalModeId}
/>

            {showImageOptions && [RetrievalMode.Vectors, RetrievalMode.Hybrid].includes(retrievalMode) && (
                <Dropdown
                    id={vectorFieldsFieldId}
                    label={t("labels.vector.label")}
                    options={[
                        {
                            key: VectorFieldOptions.Embedding,
                            text: t("labels.vector.options.embedding"),
                            selected: vectorFieldOption === VectorFieldOptions.Embedding
                        },
                        {
                            key: VectorFieldOptions.ImageEmbedding,
                            text: t("labels.vector.options.imageEmbedding"),
                            selected: vectorFieldOption === VectorFieldOptions.ImageEmbedding
                        },
                        { key: VectorFieldOptions.Both, text: t("labels.vector.options.both"), selected: vectorFieldOption === VectorFieldOptions.Both }
                    ]}
                    onChange={onVectorFieldsChange}
                    aria-labelledby={vectorFieldsId}
                    onRenderLabel={(props: IDropdownProps | undefined) => (
                        <HelpCallout labelId={vectorFieldsId} fieldId={vectorFieldsFieldId} helpText={t("helpTexts.vectorFields")} label={props?.label} />
                    )}
                />
            )}
        </Stack>
    );
};
