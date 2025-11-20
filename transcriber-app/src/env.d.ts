/// <reference types="vite/client" />

declare module '@google/genai' {
  export interface LiveServerMessage {
    serverContent?: {
      inputTranscription?: { text: string };
      turnComplete?: boolean;
    };
  }

  export interface LiveSession {
    sendRealtimeInput(input: { media: { data: string; mimeType: string } }): void;
    close(): void;
  }

  export interface LiveConnectionOptions {
    model: string;
    callbacks: {
      onopen?: () => void;
      onmessage?: (message: LiveServerMessage) => void;
      onerror?: (error: ErrorEvent) => void;
      onclose?: () => void;
    };
    config?: Record<string, unknown>;
  }

  export interface LiveConnection {
    then: (onfulfilled: (session: LiveSession) => void) => Promise<void>;
  }

  export class GoogleGenAI {
    constructor(config: { apiKey: string });
    live: {
      connect(options: LiveConnectionOptions): Promise<LiveSession>;
    };
    models: {
      generateContent(options: any): Promise<{ text: string } & Record<string, unknown>>;
    };
  }

  export const Modality: { AUDIO: string } & Record<string, string>;
}
