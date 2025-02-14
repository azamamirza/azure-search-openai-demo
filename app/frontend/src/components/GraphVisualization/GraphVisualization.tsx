import React, { useEffect, useRef } from "react";
import { Network } from "vis-network";
import { DataSet } from "vis-data";
import "vis-network/styles/vis-network.css";
import styles from "./GraphVisualization.module.css";

interface GraphVisualizationProps {
    relations: string[];
}

const parseGraph = (relations: string[]) => {
    const nodesMap = new Map<string, { id: string; label: string }>();
    const edges: { id: string; from: string; to: string; label: string }[] = [];

    relations.forEach((relation, index) => {
        // Improved regex to handle complex node names
        const match = relation.match(/^\s*([^->]+?)\s*->\s*([^->]+?)\s*->\s*(.+?)\s*$/);
        if (match) {
            const [, source, label, target] = match.map(s => s.trim());
            
            // Add nodes
            if (!nodesMap.has(source)) {
                nodesMap.set(source, { id: source, label: source });
            }
            if (!nodesMap.has(target)) {
                nodesMap.set(target, { id: target, label: target });
            }
            
            // Add edge with unique ID
            edges.push({
                id: `${source}-${label}-${target}-${index}`,
                from: source,
                to: target,
                label: label
            });
        }
    });

    return {
        nodes: Array.from(nodesMap.values()),
        edges: edges.filter((edge, index) => 
            edges.findIndex(e => 
                e.from === edge.from && 
                e.to === edge.to && 
                e.label === edge.label
            ) === index
        )
    };
};

export const GraphVisualization: React.FC<GraphVisualizationProps> = ({ relations }) => {
    const networkContainer = useRef<HTMLDivElement>(null);
    const networkInstance = useRef<Network | null>(null);

    useEffect(() => {
        if (!networkContainer.current || relations.length === 0) return;

        // Destroy existing network
        if (networkInstance.current) {
            networkInstance.current.destroy();
        }

        const { nodes, edges } = parseGraph(relations);
        
        const options = {
            nodes: {
                shape: "box",
                margin: 10,
                font: { size: 14 },
                color: {
                    background: "#e6f3ff",
                    border: "#2B7CE9",
                    highlight: { background: "#fff966", border: "#FFA500" }
                },
                shadow: true
            },
            edges: {
                arrows: "to",
                font: { size: 12, strokeWidth: 0 },
                color: "#6c757d",
                smooth: { type: "cubicBezier" },
                length: 250
            },
            physics: {
                stabilization: true,
                barnesHut: {
                    gravitationalConstant: -2000,
                    springLength: 200,
                    springConstant: 0.04
                }
            },
            interaction: { hover: true },
            layout: {
                improvedLayout: true
            }
        };

        networkInstance.current = new Network(
            networkContainer.current,
            {
                nodes: new DataSet(nodes),
                edges: new DataSet(edges)
            },
            options
        );

        return () => {
            if (networkInstance.current) {
                networkInstance.current.destroy();
            }
        };
    }, [relations]);

    return (
        <div className={styles.container}>
            <div 
                ref={networkContainer} 
                style={{ 
                    width: "100%", 
                    height: "600px",
                    border: "1px solid #e1e1e1",
                    borderRadius: "8px",
                    backgroundColor: "#f9f9f9"
                }}
            />
            {relations.length === 0 && (
                <div className={styles.emptyState}>
                    No graph data available. Relationships will appear here as they're generated.
                </div>
            )}
        </div>
    );
};