/**
 * M365 Copilot Adapter Component
 *
 * This file implements the MCP popover button for Microsoft 365 Copilot website.
 */

import { insertToggleButtonsCommon, initializeAdapter, AdapterConfig } from './common';
import { logMessage } from '../../utils/helpers';

// Configuration for M365 Copilot adapter
export const m365CopilotAdapterConfig = {
  adapterName: 'M365Copilot',
  hostname: 'copilot.microsoft.com',
  
  // Selector for the chat input container
  getChatInputContainer: (): HTMLElement | null => {
    // Look for the main chat input area
    const textarea = document.querySelector('textarea[class*="userInput"]') ||
                    document.querySelector('textarea[placeholder*="Message"]') ||
                    document.querySelector('textarea[placeholder*="Copilot"]');
    
    if (textarea) {
      // Find the parent container that holds the input and buttons
      const container = textarea.closest('form, div[class*="input"], div[class*="chat"]');
      return container as HTMLElement;
    }
    
    return null;
  },
  
  // Function to find where to insert the MCP button
  findInsertionPoint: (): HTMLElement | null => {
    const chatContainer = m365CopilotAdapterConfig.getChatInputContainer();
    if (!chatContainer) {
      logMessage('[M365 Copilot Adapter] Chat input container not found');
      return null;
    }
    
    // Look for existing button containers or action areas
    let insertionPoint = chatContainer.querySelector('div[class*="button"], div[class*="action"], div[class*="control"]');
    
    if (!insertionPoint) {
      // If no specific button container, look for any div that might contain buttons
      const allDivs = chatContainer.querySelectorAll('div');
      for (const div of allDivs) {
        const buttons = div.querySelectorAll('button');
        if (buttons.length > 0) {
          insertionPoint = div;
          break;
        }
      }
    }
    
    if (!insertionPoint) {
      // Fallback: use the chat container itself
      insertionPoint = chatContainer;
      logMessage('[M365 Copilot Adapter] Using chat container as fallback insertion point');
    }
    
    logMessage('[M365 Copilot Adapter] Found insertion point for MCP button');
    return insertionPoint as HTMLElement;
  },
  
  // Custom styling for M365 Copilot
  customStyles: `
    .mcp-button-wrapper {
      display: inline-flex;
      align-items: center;
      margin-left: 8px;
      margin-right: 8px;
    }
    
    .mcp-main-button {
      background: #0078d4 !important;
      border: 1px solid #106ebe !important;
      color: white !important;
      border-radius: 4px !important;
      padding: 6px 12px !important;
      font-size: 14px !important;
      font-weight: 500 !important;
      cursor: pointer !important;
      transition: all 0.2s ease !important;
    }
    
    .mcp-main-button:hover {
      background: #106ebe !important;
      border-color: #005a9e !important;
    }
    
    .mcp-main-button:active {
      background: #005a9e !important;
    }
    
    .mcp-main-button.inactive {
      background: #f3f2f1 !important;
      border-color: #d2d0ce !important;
      color: #605e5c !important;
    }
  `
};

// Helper functions
function getChatInputContainer(): HTMLElement | null {
    return m365CopilotAdapterConfig.getChatInputContainer();
}

function findInsertionPoint(): { container: Element; insertAfter: Element | null } | Element | null {
    const insertionPoint = m365CopilotAdapterConfig.findInsertionPoint();
    if (insertionPoint) {
        // Return the element directly - the common function will handle it
        return insertionPoint;
    }
    return null;
}

// Initialize the M365 Copilot adapter
export function initializeM365CopilotAdapter() {
    logMessage('Initializing M365 Copilot adapter component');
    
    const config: AdapterConfig = {
         adapterName: 'M365Copilot',
         storageKeyPrefix: 'm365copilot',
         findButtonInsertionPoint: findInsertionPoint,
         getStorage: () => localStorage,
         getCurrentURLKey: () => window.location.hostname + window.location.pathname
     };
    
    // Initialize the adapter with common functionality
    const stateManager = initializeAdapter(config);
    
    // Wait for the chat interface to be ready
     const observer = new MutationObserver((mutations, obs) => {
         const chatContainer = getChatInputContainer();
         if (chatContainer) {
             obs.disconnect();
             
             // Insert toggle buttons using common function
             insertToggleButtonsCommon(config, stateManager);
             logMessage('M365 Copilot MCP buttons inserted successfully');
         }
     });
     
     observer.observe(document.body, {
         childList: true,
         subtree: true
     });
     
     // Also try immediate insertion in case the interface is already ready
     const chatContainer = getChatInputContainer();
     if (chatContainer) {
         insertToggleButtonsCommon(config, stateManager);
         logMessage('M365 Copilot MCP buttons inserted immediately');
     }
}

// Global function for manual injection (debugging)
(window as any)[`injectMCPButtons_${m365CopilotAdapterConfig.adapterName}`] = () => {
  logMessage('[M365 Copilot Adapter] Manual MCP button injection triggered');
  const insertFn = (window as any)[`injectMCPButtons_${m365CopilotAdapterConfig.adapterName}`];
  if (insertFn) {
    insertFn();
  }
};

// Set up global injection function
window.injectMCPButtons = () => {
  logMessage('[M365 Copilot Adapter] Global MCP button injection triggered');
  const insertFn = (window as any)[`injectMCPButtons_${m365CopilotAdapterConfig.adapterName}`];
  if (insertFn) {
    insertFn();
  }
};

// Auto-initialize when the script loads
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initializeM365CopilotAdapter);
} else {
  initializeM365CopilotAdapter();
}