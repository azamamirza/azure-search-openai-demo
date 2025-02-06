import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// https://vitejs.dev/config/
export default defineConfig({
    plugins: [react()],
    resolve: {
        preserveSymlinks: true
    },
    build: {
        outDir: "../backend/static",
        emptyOutDir: true,
        sourcemap: true,
        rollupOptions: {
            output: {
                manualChunks: id => {
                    if (id.includes("@fluentui/react-icons")) {
                        return "fluentui-icons";
                    } else if (id.includes("@fluentui/react")) {
                        return "fluentui-react";
                    } else if (id.includes("node_modules")) {
                        return "vendor";
                    }
                }
            }
        },
        target: "esnext"
    },
    server: {
        proxy: {
            "/content/": "http://localhost:50505",
            "/auth_setup": "http://localhost:50505",
            "/.auth/me": "http://localhost:50505",
            "/ask": "http://localhost:50505",
            "/chat": "http://localhost:50505",
            "/speech": "http://localhost:50505",
            "/config": "http://localhost:50505",
            "/upload": "http://localhost:50505",
            "/delete_uploaded": "http://localhost:50505",
            "/list_uploaded": "http://localhost:50505",
            "/chat_history": "http://localhost:50505",
            "/graph": {
                target: "https://bg-backend-app1.azurewebsites.net",
                changeOrigin: true,
                secure: false,
                rewrite: path => path.replace(/^\/graph/, ""), // ✅ Keep endpoint flexibility
                headers: {
                    "X-Forwarded-Proto": "https"
                }
            },
            "/graph-stream": {
                target: "https://bg-backend-app1.azurewebsites.net",
                changeOrigin: true,
                secure: false,
                ws: true, // ✅ Enable WebSockets support (sometimes required for SSE)
                rewrite: path => path.replace(/^\/graph-stream/, "/api/v1/stream_chat/"),
                headers: {
                    "X-Forwarded-Proto": "https",
                    Accept: "text/event-stream",
                    Connection: "keep-alive",
                    "Cache-Control": "no-cache"
                },
                configure: proxy => {
                    proxy.on("proxyReq", (proxyReq, req, res) => {
                        console.log(`Proxying request from ${req.url} to backend: ${proxyReq.path}`);
                        proxyReq.setHeader("Accept", "text/event-stream");
                        proxyReq.setHeader("Connection", "keep-alive");
                    });

                    proxy.on("proxyRes", (proxyRes, req, res) => {
                        console.log(`Backend responded for ${req.url} with status: ${proxyRes.statusCode}`);
                    });

                    proxy.on("error", (err, req, res) => {
                        console.error("Proxy error for /graph-stream:", err);
                        res.writeHead(500, { "Content-Type": "text/plain" });
                        res.end("Proxy error in Vite for /graph-stream");
                    });
                }
            },

            "/api": {
                target: "https://capps-backend-2775otfh6oiva.calmsand-dc0a0904.centralus.azurecontainerapps.io",
                changeOrigin: true,
                secure: false,
                ws: true,
                rewrite: path => path.replace(/^\/api/, "") // Remove "/api" prefix if needed
            }
        }
    }
});
