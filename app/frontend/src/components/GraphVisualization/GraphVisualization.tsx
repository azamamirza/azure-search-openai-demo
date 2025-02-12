import React, { useEffect, useRef } from "react";
import { Network } from "vis-network";
import { DataSet } from "vis-data";
import styles from "./GraphVisualization.module.css";

interface GraphVisualizationProps {
    nodes: { id: string; label: string }[];
    edges: { from: string; to: string; label?: string }[];
}

const GraphVisualization: React.FC<GraphVisualizationProps> = ({ nodes, edges }) => {
    const networkContainer = useRef<HTMLDivElement>(null);

    useEffect(() => {
        if (networkContainer.current) {
            const network = new Network(
                networkContainer.current,
                {
                    nodes: new DataSet(nodes),
                    edges: new DataSet(edges),
                },
                {
                    nodes: {
                        shape: "box",
                        font: { size: 14 },
                        margin: { top: 8, right: 8, bottom: 8, left: 8 },
                        widthConstraint: { maximum: 200 },
                    },
                    edges: {
                        arrows: "to",
                        smooth: true,
                    },
                    physics: {
                        stabilization: true,
                        barnesHut: {
                            gravitationalConstant: -2000,
                            springLength: 150,
                            springConstant: 0.04,
                        },
                    }
                }
            );

            return () => {
                network.destroy();
            };
        }
    }, [nodes, edges]);

    return <div ref={networkContainer} className={styles.graphVisualization} />;
};

export default GraphVisualization;
