// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_APP_TITLE: string
  readonly VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA?: string
  // add more env variables here
}

interface ImportMeta {
  readonly env: ImportMetaEnv
  readonly glob: (pattern: string, options?: { eager?: boolean; as?: string }) => Record<string, any>
}
