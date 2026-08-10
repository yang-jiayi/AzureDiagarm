// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.tsx'
import { initTelemetry } from './services/telemetryService'
import { LanguageProvider } from './i18n/LanguageContext'
import './styles/design-tokens.css'
import './index.css'
import './styles/control-primitives.css'
import './styles/modal-primitives.css'
import './styles/high-contrast.css'
import './styles/a11y.css'

// Initialize Application Insights telemetry (no-ops if not configured)
initTelemetry()

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <LanguageProvider>
      <App />
    </LanguageProvider>
  </React.StrictMode>,
)
