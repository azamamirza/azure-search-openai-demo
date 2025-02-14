import React, { useEffect, useRef, useState } from "react";
import { Network, DataSet } from "vis-network";
import styles from "./GraphVisualization.module.css";

interface GraphVisualizationProps {
    relations: string[];
}

const parseGraph = (relations: string[]) => {
    const nodesMap = new Map<string, { id: string; label: string }>();
    const edges: { id: string; from: string; to: string; label?: string }[] = [];

    relations.forEach(relation => {
        // Improved regex to handle node labels with spaces
        const match = relation.match(/^\s*([^>]+?)\s*->\s*([^>]+?)\s*->\s*([^>]+?)\s*$/);
        if (match) {
            const [, source, label, target] = match.map(s => s.trim());

            if (!source || !target) return;

            nodesMap.set(source, { id: source, label: source });
            nodesMap.set(target, { id: target, label: target });

            edges.push({
                id: `${source}-${label}-${target}`,
                from: source,
                to: target,
                label: label || undefined,
            });
        }
    });

    return { 
        nodes: Array.from(nodesMap.values()),
        edges: edges.filter((e, i) => edges.findIndex(ee => ee.id === e.id) === i)
    };
};

export const GraphVisualization: React.FC<GraphVisualizationProps> = ({ relations }) => {
    const networkContainer = useRef<HTMLDivElement>(null);
    const [network, setNetwork] = useState<Network | null>(null);
    const [dimensions, setDimensions] = useState({ width: 0, height: 0 });

    useEffect(() => {
        const updateDimensions = () => {
            if (networkContainer.current) {
                setDimensions({
                    width: networkContainer.current.offsetWidth,
                    height: Math.max(400, networkContainer.current.offsetHeight)
                });
            }
        };

        updateDimensions();
        window.addEventListener("resize", updateDimensions);
        return () => window.removeEventListener("resize", updateDimensions);
    }, []);

    useEffect(() => {
        if (!networkContainer.current || relations.length === 0) return;

        const { nodes, edges } = parseGraph(relations);
        if (nodes.length === 0 || edges.length === 0) return;

        const data = {
            nodes: new DataSet(nodes),
            edges: new DataSet(edges)
        };

        const options = {
            nodes: {
                shape: "box",
                margin: 10,
                font: { size: 14 },
                color: {
                    background: "#e6f3ff",
                    border: "#2B7CE9",
                    highlight: { background: "#ffd966", border: "#ff9900" }
                }
            },
            edges: {
                arrows: "to",
                font: { size: 12, strokeWidth: 0 },
                color: "#6c757d",
                smooth: { type: "cubicBezier", forceDirection: "horizontal" },
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
            },
            width: `${dimensions.width}px`,
            height: `${dimensions.height}px`
        };

        const visNetwork = new Network(networkContainer.current, data, options);
        setNetwork(visNetwork);

        return () => {
            visNetwork.destroy();
            setNetwork(null);
        };
    }, [relations, dimensions]);

    return (
        <div className={styles.graphContainer}>
            <div 
                ref={networkContainer} 
                className={styles.graphVisualization}
                style={{ 
                    width: "100%",
                    height: "600px",
                    minHeight: "400px",
                    border: "1px solid #eee",
                    borderRadius: "4px"
                }}
            >
                {relations.length === 0 && (
                    <div className={styles.emptyMessage}>
                        Graph visualization will appear here as the AI generates relationships
                    </div>
                )}
            </div>
        </div>
    );
};