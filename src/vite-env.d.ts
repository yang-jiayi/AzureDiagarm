// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_APP_TITLE: string
  readonly VITE_AZURE_OPENAI_DEPLOYMENT_GPT61SOL?: string
  readonly VITE_AZURE_FOUNDRY_ENDPOINT?: string
  readonly VITE_AZURE_FOUNDRY_DEPLOYMENT_OPUS55?: string
  // add more env variables here
}

interface ImportMeta {
  readonly env: ImportMetaEnv
  readonly glob: (pattern: string, options?: { eager?: boolean; as?: string }) => Record<string, any>
}
