import React, { useEffect, useRef, useState } from "react";
import { Network } from "vis-network";
import { DataSet } from "vis-data";
import styles from "./GraphVisualization.module.css";

interface GraphVisualizationProps {
    nodes?: { id: string; label: string; group?: string; title?: string }[];
    edges?: { from: string; to: string; label?: string }[];
}

export const GraphVisualization: React.FC<GraphVisualizationProps> = ({ 
    nodes = [], 
    edges = [] 
}) => {
    const networkContainer = useRef<HTMLDivElement>(null);
    const [network, setNetwork] = useState<Network | null>(null);
    const nodesDataSet = useRef(new DataSet<{ id: string; label: string; group?: string; title?: string }>([]));
    const edgesDataSet = useRef(new DataSet<{ id: string; from: string; to: string; label?: string }>([]));

    useEffect(() => {
        if (networkContainer.current && !network) {
            const newNetwork = new Network(
                networkContainer.current,
                { nodes: nodesDataSet.current, edges: edgesDataSet.current },
                {
                    nodes: {
                        shape: "dot",
                        scaling: { min: 10, max: 30 },
                        font: { face: "Arial", size: 12 },
                    },
                    edges: {
                        smooth: true,
                        width: 1.5,
                    },
                    physics: {
                        enabled: true,
                        barnesHut: {
                            gravitationalConstant: -2000,
                            centralGravity: 0.3,
                            springLength: 120,
                            springConstant: 0.05,
                        },
                    },
                    interaction: {
                        hover: true,
                        dragNodes: true,
                        zoomView: true,
                    },
                }
            );
            setNetwork(newNetwork);
        }
    }, [network]);

    useEffect(() => {
        nodesDataSet.current.update(
            nodes.map(node => ({
                ...node,
                shape: "dot",
                size: node.group === "important" ? 20 : 10,
                color: node.group === "important" ? "#ff6b6b" : "#4c6ef5",
                font: { color: "#000", size: 14 },
                borderWidth: 2,
                title: node.title || `ID: ${node.id}`,
            }))
        );
    }, [nodes]);

    useEffect(() => {
        edgesDataSet.current.update(
            edges.map(edge => ({
                ...edge,
                id: `${edge.from}-${edge.to}`,
                arrows: "to",
                font: { align: "middle" },
                color: { color: "#ccc", highlight: "#f03e3e" },
            }))
        );
    }, [edges]);

    return <div 
        ref={networkContainer} 
        className={styles.graphVisualization} 
    />;
};