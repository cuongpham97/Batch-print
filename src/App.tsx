import { useEffect, useMemo, useState } from 'react';
import { BatCodePanel } from './components/BatCodePanel';
import { ConfigForm } from './components/ConfigForm';
import { SheetPreview } from './components/SheetPreview';
import {
  DocTypeId,
  PrintConfig,
  cloneDefaultPrintConfig,
  parseImportedPrintConfig,
} from './types/printConfig';
import { generateBatScript } from './utils/batGenerator';

const CONFIG_STORAGE_KEY = 'printbat.config.v1';

function loadStoredPrintConfig(): PrintConfig {
  try {
    const raw = window.localStorage.getItem(CONFIG_STORAGE_KEY);
    if (!raw) return cloneDefaultPrintConfig();
    const parsed = parseImportedPrintConfig(JSON.parse(raw) as unknown);
    return parsed ?? cloneDefaultPrintConfig();
  } catch {
    return cloneDefaultPrintConfig();
  }
}

export default function App() {
  const [config, setConfig] = useState<PrintConfig>(loadStoredPrintConfig);
  const [activeTab, setActiveTab] = useState<DocTypeId>('word');

  useEffect(() => {
    try {
      window.localStorage.setItem(CONFIG_STORAGE_KEY, JSON.stringify(config));
    } catch {
      // Ignore storage quota or restricted context errors
    }
  }, [config]);

  const batScript = useMemo(() => generateBatScript(config), [config]);

  return (
    <div className="min-h-screen lg:h-screen lg:overflow-hidden flex flex-col app-bg-gradient text-slate-900">
      <main className="flex-1 min-h-0 max-w-7xl w-full mx-auto px-4 sm:px-6 py-5 lg:py-0">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 items-start lg:items-stretch lg:h-full lg:min-h-0">
          <div className="lg:col-span-7 lg:h-full lg:min-h-0 lg:overflow-y-auto lg:py-5 panel-scroll">
            <ConfigForm
              config={config}
              activeTab={activeTab}
              onTabChange={setActiveTab}
              onChange={setConfig}
            />
          </div>

          <div className="lg:col-span-5 lg:h-full lg:min-h-0 lg:overflow-y-auto lg:py-5 panel-scroll space-y-4">
            <SheetPreview config={config} activeTab={activeTab} />

            <BatCodePanel
              script={batScript}
              scriptFileName={config.scriptFileName}
              onFileNameChange={(name) =>
                setConfig((prev) => ({ ...prev, scriptFileName: name }))
              }
            />
          </div>
        </div>
      </main>
    </div>
  );
}
