import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { HashRouter } from "react-router-dom";
import { QueryClientProvider } from "@tanstack/react-query";
import { FluentProvider, webLightTheme, type Theme } from "@fluentui/react-components";
import { TooltipProvider } from "@/components/ui/tooltip";
import { queryClient } from "@/lib/queryClient";
import "./index.css";
import App from "./App.tsx";

/* Fluent 2 tokens tuned to the studio look: rounder controls, system type. */
const studioTheme: Theme = {
  ...webLightTheme,
  fontFamilyBase:
    '"Segoe UI Variable Text", "Segoe UI", -apple-system, BlinkMacSystemFont, "SF Pro Text", system-ui, sans-serif',
  fontFamilyMonospace: '"Cascadia Code", "SF Mono", ui-monospace, Menlo, Consolas, monospace',
  borderRadiusMedium: "8px",
  borderRadiusLarge: "12px",
  borderRadiusXLarge: "16px",
  colorNeutralBackground1: "#ffffff",
  colorNeutralBackground2: "#f7f8fa",
  colorNeutralBackground3: "#eef0f3",
  colorNeutralStroke1: "#d5d8de",
  colorNeutralStroke2: "#e6e8ec",
  colorNeutralForeground1: "#1b1b1f",
  colorNeutralForeground2: "#424349",
  colorNeutralForeground3: "#6b6e76",
};

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <QueryClientProvider client={queryClient}>
      <FluentProvider theme={studioTheme} className="summary-studio-theme">
        <HashRouter>
          <TooltipProvider>
            <App />
          </TooltipProvider>
        </HashRouter>
      </FluentProvider>
    </QueryClientProvider>
  </StrictMode>,
);
