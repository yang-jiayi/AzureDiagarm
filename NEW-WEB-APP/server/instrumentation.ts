
import { useAzureMonitor as configureAzureMonitor } from '@azure/monitor-opentelemetry';

if (process.env.APPLICATIONINSIGHTS_CONNECTION_STRING) {
  configureAzureMonitor({
    azureMonitorExporterOptions: {
      connectionString: process.env.APPLICATIONINSIGHTS_CONNECTION_STRING,
    },
  });
}