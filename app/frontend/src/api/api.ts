import {
    ChatAppResponse,
    ChatAppResponseOrError,
    ChatAppRequest,
    Config,
    SimpleAPIResponse,
    HistoryListApiResponse,
    HistroyApiResponse,
    RetrievalMode
} from "./models";
import { useLogin, getToken, isUsingAppServicesLogin } from "../authConfig";

const BACKEND_PROXY = "/api"; // Instead of hardcoding backend URL

interface GraphRagResponse {
    response: string;
    nodes?: any[];
} // Instead of hardcoding backend URL
export async function getHeaders(idToken: string | undefined): Promise<Record<string, string>> {
    // If using login and not using app services, add the id token of the logged in account as the authorization
    if (useLogin && !isUsingAppServicesLogin) {
        if (idToken) {
            return { Authorization: `Bearer ${idToken}` };
        }
    }

    return {};
}

export async function configApi(): Promise<Config> {
    const response = await fetch(`${BACKEND_PROXY}/config`, {
        method: "GET"
    });

    return (await response.json()) as Config;
}

export async function askApi(request: ChatAppRequest, idToken: string | undefined): Promise<ChatAppResponse> {
    const headers = await getHeaders(idToken);

    const response = await fetch(`${BACKEND_PROXY}/chat/stream`, {
        method: "POST",
        headers: { ...headers, "Content-Type": "application/json" },
        body: JSON.stringify(request)
    });

    if (response.status > 299 || !response.ok) {
        throw Error(`Request failed with status ${response.status}`);
    }
    const parsedResponse: ChatAppResponseOrError = await response.json();
    if (parsedResponse.error) {
        throw Error(parsedResponse.error);
    }

    return parsedResponse as ChatAppResponse;
}

export async function chatApi(request: ChatAppRequest, shouldStream: boolean, idToken: string | undefined): Promise<Response> {
    let url = `${BACKEND_PROXY}/chat`;
    if (shouldStream) {
        url += "/stream";
    }
    const headers = await getHeaders(idToken);
    return await fetch(url, {
        method: "POST",
        headers: { ...headers, "Content-Type": "application/json" },
        body: JSON.stringify(request)
    });
}
// export async function graphRagApi(requestData: ChatAppRequest, shouldStream: boolean, idToken: string | undefined): Promise<Response> {
//     const headers: HeadersInit = {
//         "Content-Type": "application/json",
//         ...(idToken ? { Authorization: `Bearer ${idToken}` } : {})
//     };

//     try {
//         const lastUserMessage =
//             requestData.messages
//                 .slice()
//                 .reverse()
//                 .find(m => m.role === "user")?.content || "";

//         const response = await fetch("/graph", {
//             method: "POST",
//             headers,
//             body: JSON.stringify({ query: lastUserMessage })
//         });

//         if (!response.ok) {
//             const errorText = await response.text();
//             throw new Error(`Graph RAG request failed: ${response.status} - ${errorText}`);
//         }

//         const apiData: GraphRagResponse = await response.json();

//         return new Response(apiData.response?.trim() || "No response available", {
//             status: 200,
//             headers: { "Content-Type": "text/plain" } // 👈 Sets response as plain text
//         });
//     } catch (error) {
//         console.error("Graph RAG API Error:", error.message || error);
//         throw new Error(`Graph RAG request failed: ${error.message || "Unknown error"}`);
//     }
// }
export async function graphRagApi(requestData: ChatAppRequest, shouldStream: boolean, idToken: string | undefined): Promise<Response> {
    const headers: HeadersInit = {
        "Content-Type": "application/json",
        ...(idToken ? { Authorization: `Bearer ${idToken}` } : {}),
        ...(shouldStream ? { "Accept": "text/event-stream" } : {})
    };

    try {
        const lastUserMessage =
            requestData.messages
                .slice()
                .reverse()
                .find(m => m.role === "user")?.content || "";

        const endpoint = shouldStream ? "/graph-stream" : "/graph";
        console.log(`Requesting: ${endpoint}`);

        const response = await fetch(endpoint, {
            method: "POST",
            headers,
            body: JSON.stringify({ query: lastUserMessage })
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.error(`Graph RAG API error response: ${errorText}`);
            throw new Error(`Graph RAG request failed: ${response.status} - ${errorText}`);
        }

        if (shouldStream) {
            // ✅ Handle Server-Sent Events (SSE) Stream
            const stream = new ReadableStream({
                async start(controller) {
                    const reader = response.body?.getReader();
                    if (!reader) {
                        controller.close();
                        return;
                    }

                    const decoder = new TextDecoder();
                    try {
                        while (true) {
                            const { value, done } = await reader.read();
                            if (done) break;
                            const chunk = decoder.decode(value, { stream: true });

                            // Stream text to frontend
                            controller.enqueue(chunk);
                        }
                    } catch (error) {
                        console.error("Error reading stream:", error);
                    } finally {
                        controller.close();
                    }
                }
            });

            return new Response(stream, {
                status: 200,
                headers: {
                    "Content-Type": "text/event-stream",
                    "Cache-Control": "no-cache",
                    "Connection": "keep-alive"
                }
            });
        } else {
            // ✅ Ensure Response is Correctly Parsed
            const text = await response.text();
            console.log("Non-streaming API Response:", text);

            let parsedData;
            try {
                parsedData = JSON.parse(text);
            } catch (error) {
                console.warn("Response is not valid JSON, returning raw text.");
                return new Response(text.trim(), {
                    status: 200,
                    headers: { "Content-Type": "text/plain" }
                });
            }

            return new Response(parsedData.response?.trim() || "No response available", {
                status: 200,
                headers: { "Content-Type": "application/json" }
            });
        }
    } catch (error) {
        if (error instanceof Error) {
            console.error("Graph RAG API Error:", error.message || error);
        } else {
            console.error("Graph RAG API Error:", error);
        }
        if (error instanceof Error) {
            throw new Error(`Graph RAG request failed: ${error.message || "Unknown error"}`);
        } else {
            throw new Error("Graph RAG request failed: Unknown error");
        }
    }
}


export async function getSpeechApi(text: string): Promise<string | null> {
    return await fetch("/speech", {
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify({
            text: text
        })
    })
        .then(response => {
            if (response.status == 200) {
                return response.blob();
            } else if (response.status == 400) {
                console.log("Speech synthesis is not enabled.");
                return null;
            } else {
                console.error("Unable to get speech synthesis.");
                return null;
            }
        })
        .then(blob => (blob ? URL.createObjectURL(blob) : null));
}

export function getCitationFilePath(citation: string): string {
    return `${BACKEND_PROXY}/content/${citation}`;
}

export async function uploadFileApi(request: FormData, idToken: string): Promise<SimpleAPIResponse> {
    const response = await fetch("/upload", {
        method: "POST",
        headers: await getHeaders(idToken),
        body: request
    });

    if (!response.ok) {
        throw new Error(`Uploading files failed: ${response.statusText}`);
    }

    const dataResponse: SimpleAPIResponse = await response.json();
    return dataResponse;
}

export async function deleteUploadedFileApi(filename: string, idToken: string): Promise<SimpleAPIResponse> {
    const headers = await getHeaders(idToken);
    const response = await fetch("/delete_uploaded", {
        method: "POST",
        headers: { ...headers, "Content-Type": "application/json" },
        body: JSON.stringify({ filename })
    });

    if (!response.ok) {
        throw new Error(`Deleting file failed: ${response.statusText}`);
    }

    const dataResponse: SimpleAPIResponse = await response.json();
    return dataResponse;
}

export async function listUploadedFilesApi(idToken: string): Promise<string[]> {
    const response = await fetch(`/list_uploaded`, {
        method: "GET",
        headers: await getHeaders(idToken)
    });

    if (!response.ok) {
        throw new Error(`Listing files failed: ${response.statusText}`);
    }

    const dataResponse: string[] = await response.json();
    return dataResponse;
}

export async function postChatHistoryApi(item: any, idToken: string): Promise<any> {
    const headers = await getHeaders(idToken);
    const response = await fetch("/chat_history", {
        method: "POST",
        headers: { ...headers, "Content-Type": "application/json" },
        body: JSON.stringify(item)
    });

    if (!response.ok) {
        throw new Error(`Posting chat history failed: ${response.statusText}`);
    }

    const dataResponse: any = await response.json();
    return dataResponse;
}

export async function getChatHistoryListApi(count: number, continuationToken: string | undefined, idToken: string): Promise<HistoryListApiResponse> {
    const headers = await getHeaders(idToken);
    const response = await fetch("/chat_history/items", {
        method: "POST",
        headers: { ...headers, "Content-Type": "application/json" },
        body: JSON.stringify({ count: count, continuation_token: continuationToken })
    });

    if (!response.ok) {
        throw new Error(`Getting chat histories failed: ${response.statusText}`);
    }

    const dataResponse: HistoryListApiResponse = await response.json();
    return dataResponse;
}

export async function getChatHistoryApi(id: string, idToken: string): Promise<HistroyApiResponse> {
    const headers = await getHeaders(idToken);
    const response = await fetch(`/chat_history/items/${id}`, {
        method: "GET",
        headers: { ...headers, "Content-Type": "application/json" }
    });

    if (!response.ok) {
        throw new Error(`Getting chat history failed: ${response.statusText}`);
    }

    const dataResponse: HistroyApiResponse = await response.json();
    return dataResponse;
}

export async function deleteChatHistoryApi(id: string, idToken: string): Promise<any> {
    const headers = await getHeaders(idToken);
    const response = await fetch(`/chat_history/items/${id}`, {
        method: "DELETE",
        headers: { ...headers, "Content-Type": "application/json" }
    });

    if (!response.ok) {
        throw new Error(`Deleting chat history failed: ${response.statusText}`);
    }

    const dataResponse: any = await response.json();
    return dataResponse;
}
