/**
 * Chat Input Handler for M365 Copilot
 *
 * Utility functions for interacting with the M365 Copilot chat input area
 */

import { logMessage } from '@src/utils/helpers';

/**
 * Find the M365 Copilot chat input textarea element
 * @returns The chat input textarea element or null if not found
 */
export const findChatInputElement = (): HTMLTextAreaElement | null => {
  // Try to find the main input with userInput class
  let chatInput = document.querySelector('textarea[class*="userInput"]');

  if (chatInput) {
    logMessage('Found M365 Copilot input with userInput class');
    return chatInput as HTMLTextAreaElement;
  }

  // Try to find input with Message placeholder
  chatInput = document.querySelector('textarea[placeholder*="Message"]');

  if (chatInput) {
    logMessage('Found M365 Copilot input with Message placeholder');
    return chatInput as HTMLTextAreaElement;
  }

  // Try to find input with Copilot placeholder
  chatInput = document.querySelector('textarea[placeholder*="Copilot"]');

  if (chatInput) {
    logMessage('Found M365 Copilot input with Copilot placeholder');
    return chatInput as HTMLTextAreaElement;
  }

  // Try to find any textarea in the chat area
  chatInput = document.querySelector('textarea[aria-label*="chat"], textarea[aria-label*="message"]');

  if (chatInput) {
    logMessage('Found M365 Copilot input with chat/message aria-label');
    return chatInput as HTMLTextAreaElement;
  }

  // Fallback: find any textarea
  chatInput = document.querySelector('textarea');

  if (chatInput) {
    logMessage('Found M365 Copilot input as fallback textarea');
    return chatInput as HTMLTextAreaElement;
  }

  logMessage('Could not find any M365 Copilot chat input textarea');
  return null;
};

/**
 * Wrap content in tool_output tags
 * @param content The content to wrap
 * @returns The wrapped content
 */
export const wrapInToolOutput = (content: string): string => {
  return `<tool_output>\n${content}\n</tool_output>`;
};

/**
 * Format an object as a JSON string
 * @param data The data to format
 * @returns Formatted JSON string
 */
export const formatAsJson = (data: any): string => {
  return JSON.stringify(data, null, 2);
};

/**
 * Insert text into the M365 Copilot chat input
 * @param text The text to insert
 * @returns True if successful, false otherwise
 */
export const insertTextToChatInput = (text: string): boolean => {
  try {
    const chatInput = findChatInputElement();

    if (chatInput) {
      // Append the text to the existing text in the textarea
      const currentText = chatInput.value;
      // Add new line before and after the current text if there's existing content
      const formattedText = currentText ? `${currentText}\n\n${text}` : text;

      chatInput.value = formattedText;

      // Trigger input event to make M365 Copilot recognize the change
      const inputEvent = new Event('input', { bubbles: true });
      chatInput.dispatchEvent(inputEvent);

      // Also trigger change event
      const changeEvent = new Event('change', { bubbles: true });
      chatInput.dispatchEvent(changeEvent);

      // Focus the textarea
      chatInput.focus();

      logMessage('Appended text to M365 Copilot chat input');
      return true;
    } else {
      logMessage('Could not find M365 Copilot chat input');
      console.error('Could not find M365 Copilot chat input textarea');
      return false;
    }
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    logMessage(`Error inserting text into chat input: ${errorMessage}`);
    console.error('Error inserting text into chat input:', error);
    return false;
  }
};

/**
 * Insert tool result into the M365 Copilot chat input
 * @param result The tool result to insert
 * @returns True if successful, false otherwise
 */
export const insertToolResultToChatInput = (result: any): boolean => {
  try {
    if (typeof result !== 'string') {
      result = JSON.stringify(result, null, 2);
      logMessage('Converted tool result to string format');
    }

    return insertTextToChatInput(result);
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    logMessage(`Error formatting tool result: ${errorMessage}`);
    console.error('Error formatting tool result:', error);
    return false;
  }
};

/**
 * Attach a file to the M365 Copilot input
 * @param file The file to attach
 * @returns Promise that resolves to true if successful
 */
export const attachFileToChatInput = async (file: File): Promise<boolean> => {
  try {
    // Find the M365 Copilot input element
    const chatInput = findChatInputElement();

    if (!chatInput) {
      logMessage('Could not find M365 Copilot input element for file attachment');
      return false;
    }

    // Look for file upload button or input
    const fileInput = document.querySelector('input[type="file"]') as HTMLInputElement;
    const uploadButton = document.querySelector('button[aria-label*="attach"], button[aria-label*="upload"], button[title*="attach"], button[title*="upload"]') as HTMLButtonElement;

    if (fileInput) {
      // Create a DataTransfer object and add the file
      const dataTransfer = new DataTransfer();
      dataTransfer.items.add(file);
      fileInput.files = dataTransfer.files;

      // Trigger change event
      const changeEvent = new Event('change', { bubbles: true });
      fileInput.dispatchEvent(changeEvent);

      logMessage(`Attached file ${file.name} to M365 Copilot input via file input`);
      return true;
    }

    if (uploadButton) {
      // Click the upload button first
      uploadButton.click();
      
      // Wait a bit for the file input to appear
      await new Promise(resolve => setTimeout(resolve, 500));
      
      const newFileInput = document.querySelector('input[type="file"]') as HTMLInputElement;
      if (newFileInput) {
        const dataTransfer = new DataTransfer();
        dataTransfer.items.add(file);
        newFileInput.files = dataTransfer.files;

        const changeEvent = new Event('change', { bubbles: true });
        newFileInput.dispatchEvent(changeEvent);

        logMessage(`Attached file ${file.name} to M365 Copilot input via upload button`);
        return true;
      }
    }

    // Fallback: Try drag and drop simulation
    const dataTransfer = new DataTransfer();
    dataTransfer.items.add(file);

    const dragOverEvent = new DragEvent('dragover', {
      bubbles: true,
      cancelable: true,
      dataTransfer: dataTransfer,
    });

    const dropEvent = new DragEvent('drop', {
      bubbles: true,
      cancelable: true,
      dataTransfer: dataTransfer,
    });

    // Prevent default on dragover to enable drop
    chatInput.addEventListener('dragover', e => e.preventDefault(), { once: true });
    chatInput.dispatchEvent(dragOverEvent);

    // Simulate the drop event
    chatInput.dispatchEvent(dropEvent);

    logMessage(`Attempted to attach file ${file.name} to M365 Copilot input via drag and drop`);
    return true;
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    logMessage(`Error attaching file to M365 Copilot input: ${errorMessage}`);
    console.error('Error attaching file to M365 Copilot input:', error);
    return false;
  }
};

/**
 * Submit the current chat input (equivalent to pressing Enter)
 * @returns True if submission was successful, false otherwise
 */
export const submitChatInput = (maxWaitTime = 5000): Promise<boolean> => {
  return new Promise(resolve => {
    try {
      const chatInput = findChatInputElement();

      if (!chatInput) {
        logMessage('Could not find chat input to submit');
        resolve(false);
        return;
      }

      // Define a function to find the submit button
      const findSubmitButton = (): HTMLButtonElement | null => {
        const submitButton =
          document.querySelector('button[aria-label*="Send"]') ||
          document.querySelector('button[aria-label*="Submit"]') ||
          document.querySelector('button[type="submit"]') ||
          // Look for a button next to the textarea
          chatInput.parentElement?.querySelector('button') ||
          chatInput.closest('form, div')?.querySelector('button[type="submit"], button[aria-label*="Send"]') ||
          // Common pattern: button with send icon
          document.querySelector('button svg[data-icon="send"]')?.closest('button') ||
          document.querySelector('button [data-testid="send-button"]')?.closest('button');

        return submitButton as HTMLButtonElement | null;
      };

      // Try to find and check the submit button
      const submitButton = findSubmitButton();

      if (submitButton) {
        logMessage(`Found submit button (${submitButton.getAttribute('aria-label') || 'unknown'})`);

        // Function to check if button is enabled and click it
        const tryClickingButton = () => {
          const button = findSubmitButton();
          if (!button) {
            logMessage('Submit button no longer found');
            resolve(false);
            return;
          }

          // Check if the button is disabled
          const isDisabled =
            button.disabled ||
            button.getAttribute('disabled') !== null ||
            button.getAttribute('aria-disabled') === 'true' ||
            button.classList.contains('disabled');

          if (!isDisabled) {
            logMessage('Submit button is enabled, clicking it');
            button.click();
            resolve(true);
          } else {
            logMessage('Submit button is disabled, waiting...');
          }
        };

        // Set up a timer to periodically check if the button becomes enabled
        let elapsedTime = 0;
        const checkInterval = 200; // Check every 200ms

        const intervalId = setInterval(() => {
          elapsedTime += checkInterval;

          tryClickingButton();

          // If we've waited too long, try alternative methods
          if (elapsedTime >= maxWaitTime) {
            clearInterval(intervalId);
            logMessage(`Button remained disabled for ${maxWaitTime}ms, trying alternative methods`);

            // Method 2: Simulate Enter key press
            logMessage('Simulating Enter key press as fallback');

            // Focus the textarea first
            chatInput.focus();

            // Create and dispatch keydown event (Enter key)
            const keydownEvent = new KeyboardEvent('keydown', {
              key: 'Enter',
              code: 'Enter',
              keyCode: 13,
              which: 13,
              bubbles: true,
              cancelable: true,
            });

            // Create and dispatch keypress event
            const keypressEvent = new KeyboardEvent('keypress', {
              key: 'Enter',
              code: 'Enter',
              keyCode: 13,
              which: 13,
              bubbles: true,
              cancelable: true,
            });

            // Create and dispatch keyup event
            const keyupEvent = new KeyboardEvent('keyup', {
              key: 'Enter',
              code: 'Enter',
              keyCode: 13,
              which: 13,
              bubbles: true,
              cancelable: true,
            });

            // Dispatch all events in sequence
            chatInput.dispatchEvent(keydownEvent);
            chatInput.dispatchEvent(keypressEvent);
            chatInput.dispatchEvent(keyupEvent);

            // Try to find and submit a form as a last resort
            const form = chatInput.closest('form');
            if (form) {
              logMessage('Found form element, submitting it');
              form.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));
            }

            logMessage('Attempted all fallback methods to submit chat input');
            resolve(true);
          }
        }, checkInterval);

        // Initial check - maybe it's already enabled
        tryClickingButton();

        // If the button is already enabled and clicked, clear the interval
        if (submitButton && !submitButton.disabled) {
          clearInterval(intervalId);
        }
      } else {
        // If no button found, proceed with alternative methods immediately
        logMessage('No submit button found, trying alternative methods');

        // Method 2: Simulate Enter key press
        logMessage('Simulating Enter key press as fallback');

        // Focus the textarea first
        chatInput.focus();

        // Create and dispatch keydown event (Enter key)
        const keydownEvent = new KeyboardEvent('keydown', {
          key: 'Enter',
          code: 'Enter',
          keyCode: 13,
          which: 13,
          bubbles: true,
          cancelable: true,
        });

        // Create and dispatch keypress event
        const keypressEvent = new KeyboardEvent('keypress', {
          key: 'Enter',
          code: 'Enter',
          keyCode: 13,
          which: 13,
          bubbles: true,
          cancelable: true,
        });

        // Create and dispatch keyup event
        const keyupEvent = new KeyboardEvent('keyup', {
          key: 'Enter',
          code: 'Enter',
          keyCode: 13,
          which: 13,
          bubbles: true,
          cancelable: true,
        });

        // Dispatch all events in sequence
        chatInput.dispatchEvent(keydownEvent);
        chatInput.dispatchEvent(keypressEvent);
        chatInput.dispatchEvent(keyupEvent);

        // Try to find and submit a form as a last resort
        const form = chatInput.closest('form');
        if (form) {
          logMessage('Found form element, submitting it');
          form.dispatchEvent(new SubmitEvent('submit', { bubbles: true, cancelable: true }));
        }

        logMessage('Attempted all methods to submit chat input');
        resolve(true);
      }
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      logMessage(`Error submitting chat input: ${errorMessage}`);
      console.error('Error submitting chat input:', error);
      resolve(false);
    }
  });
};