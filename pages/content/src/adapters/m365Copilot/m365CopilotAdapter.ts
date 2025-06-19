import { BaseAdapter } from "../common/baseAdapter";
import { SiteAdapter } from "../../utils/siteAdapter";
import { SidebarManager } from "../../components/sidebar";
import { logMessage } from "../../utils/helpers";
import { BaseAdapter } from "../common/baseAdapter";
import { initializeM365CopilotAdapter } from "../adaptercomponents/m365copilot";
import { insertToolResultToChatInput, attachFileToChatInput, submitChatInput } from '../../components/websites/m365copilot';

export class M365CopilotAdapter extends BaseAdapter implements SiteAdapter {
    name = "M365Copilot";
    hostname = ["copilot.microsoft.com"];
    public readonly siteIdentifier = "m365copilot";

    // Property to store the last URL
    private lastUrl: string = '';
    // Property to store the interval ID
    private urlCheckInterval: number | null = null;

    constructor() {
        super();
        // Create the sidebar manager instance
        this.sidebarManager = SidebarManager.getInstance('m365copilot');
        logMessage('Created M365Copilot sidebar manager instance');
    }

    protected initializeSidebarManager(): void {
        if (this.sidebarManager) {
            this.sidebarManager.initialize();
            logMessage('Initialized M365Copilot sidebar manager');
        }
    }

    protected initializeObserver(forceReset: boolean = false): void {
        logMessage("M365CopilotAdapter: Initializing observer");
        
        // Check the current URL immediately
        this.checkCurrentUrl();
        
        // Initialize M365 Copilot components
        initializeM365CopilotAdapter();
        
        // Start URL checking to handle navigation within M365 Copilot
        if (!this.urlCheckInterval) {
            this.lastUrl = window.location.href;
            this.urlCheckInterval = window.setInterval(() => {
                const currentUrl = window.location.href;
                
                if (currentUrl !== this.lastUrl) {
                    logMessage(`URL changed from ${this.lastUrl} to ${currentUrl}`);
                    this.lastUrl = currentUrl;
                    
                    initializeM365CopilotAdapter();
                    // Check if we should show or hide the sidebar based on URL
                    this.checkCurrentUrl();
                }
            }, 1000); // Check every second
        }
    }

    getChatInput(): HTMLElement | null {
        // Look for textarea with userInput class (it has many classes)
        return document.querySelector('textarea[class*="userInput"]') || 
               document.querySelector('textarea[placeholder*="Message"]') ||
               document.querySelector('textarea[placeholder*="Copilot"]');
    }

    getSubmitButton(): HTMLElement | null {
        // Look for submit button near the textarea
        const chatInput = this.getChatInput();
        if (chatInput) {
            // Try to find a button in the same container
            const container = chatInput.closest('form, div');
            if (container) {
                const submitBtn = container.querySelector('button[type="submit"], button[aria-label*="Send"], button[aria-label*="Submit"]');
                if (submitBtn) return submitBtn as HTMLElement;
            }
        }
        // Fallback: look for any button that might be a submit button
        return document.querySelector('button[aria-label*="Send"], button[type="submit"]');
    }

    /**
     * Insert text into the M365 Copilot input field
     * @param text Text to insert
     */
    insertTextIntoInput(text: string): void {
        insertToolResultToChatInput(text);
        logMessage(`M365CopilotAdapter: Inserted text into input: ${text.substring(0, 20)}...`);
    }

    /**
     * Trigger submission of the M365 Copilot input form
     */
    triggerSubmission(): void {
        // Use the function to submit the form
        submitChatInput()
            .then((success: boolean) => {
                logMessage(`M365CopilotAdapter: Triggered form submission: ${success ? 'success' : 'failed'}`);
            })
            .catch((error: Error) => {
                logMessage(`M365CopilotAdapter: Error triggering form submission: ${error}`);
            });
    }

    /**
     * Check if M365 Copilot supports file upload
     * @returns true if file upload is supported
     */
    supportsFileUpload(): boolean {
        return true;
    }

    /**
     * Attach a file to the M365 Copilot input
     * @param file The file to attach
     * @returns Promise that resolves to true if successful
     */
    async attachFile(file: File): Promise<boolean> {
        try {
            const result = await attachFileToChatInput(file);
            return result;
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : String(error);
            logMessage(`M365CopilotAdapter: Error in adapter when attaching file: ${errorMessage}`);
            console.error('M365CopilotAdapter: Error in adapter when attaching file:', error);
            return false;
        }
    }

    cleanup(): void {
        // Clear interval for URL checking
        if (this.urlCheckInterval) {
            window.clearInterval(this.urlCheckInterval);
            this.urlCheckInterval = null;
        }

        // Call the parent cleanup method
        super.cleanup();
    }

    /**
     * Check the current URL and show/hide sidebar accordingly
     */
    private checkCurrentUrl(): void {
        const currentUrl = window.location.href;
        logMessage(`M365CopilotAdapter: Checking current URL: ${currentUrl}`);

        // For M365 Copilot, we want to show the sidebar on all pages
        // You can customize this with specific URL patterns if needed
        if (this.sidebarManager && !this.sidebarManager.getIsVisible()) {
            logMessage('M365CopilotAdapter: Showing sidebar for M365 Copilot URL');
            this.sidebarManager.showWithToolOutputs();
        }
    }
}