import type { ParamValueElement } from '../core/types';
import { StabilizedBlock } from '../core/types';
import { CONFIG } from '../core/config';
import { safelySetContent } from '../utils/index';
import { storeExecutedFunction, generateContentSignature } from '../mcpexecute/storage';
import { checkAndDisplayFunctionHistory, createHistoryPanel, updateHistoryPanel } from './functionHistory';

// Add type declarations for the global adapter access
declare global {
  interface Window {
    mcpAdapter?: any;
    getCurrentAdapter?: () => any;
  }
}

// Performance optimizations: Cache constants and pre-compile regexes
const MAX_INSERT_LENGTH = 39000;
const MAX_INSERT_LENGTH_M365_COPILOT = 8000;
const WEBSITE_NAME_FOR_MAX_INSERT_LENGTH_CHECK = ['perplexity'];
const WEBSITE_NAME_FOR_M365_COPILOT_CHECK = ['m365copilot'];
const websiteName = window.location.hostname
  .toLowerCase()
  .replace(/^www\./i, '')
  .split('.')[0];

// Pre-compiled regexes for better performance
const INVOKE_REGEX = /<invoke name="([^"]+)"(?:\s+call_id="([^"]+)")?>/;
const PARAM_REGEX = /<parameter\s+name="([^"]+)"\s*(?:type="([^"]+)")?\s*>(.*?)<\/parameter>/gs;
const CDATA_REGEX = /<!\[CDATA\[([\s\S]*?)\]\]>/;
const NUMBER_REGEX = /^-?\d+(\.\d+)?$/;
const BOOLEAN_REGEX = /^(true|false)$/i;

// SVG icons as constants for reuse
const ICONS = {
  CODE: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M9.4 16.6L4.8 12l4.6-4.6L8 6l-6 6 6 6 1.4-1.4zm5.2 0l4.6-4.6-4.6-4.6L16 6l6 6-6 6-1.4-1.4z" fill="currentColor"/></svg>',
  PLAY: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M8 5v14l11-7z" fill="currentColor"/></svg>',
  INSERT:
    '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M16 12l-7-7v4H2v6h7v4l7-7z" fill="currentColor"/></svg>',
  ATTACH:
    '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="4" y="2" width="16" height="20" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><path d="M8 6h8M8 10h8M8 14h4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>',
  SPINNER:
    '<svg width="16" height="16" viewBox="0 0 50 50"><circle cx="25" cy="25" r="20" fill="none" stroke="currentColor" stroke-width="5" stroke-linecap="round" stroke-dasharray="31.4 31.4" transform="rotate(-90 25 25)"><animateTransform attributeName="transform" type="rotate" from="0 25 25" to="360 25 25" dur="0.8s" repeatCount="indefinite"/></circle></svg>',
};

// Performance utility: Object pooling for DOM elements
class ElementPool {
  private static pools = new Map<string, HTMLElement[]>();

  static get(tagName: string, className?: string): HTMLElement {
    const key = `${tagName}:${className || ''}`;
    const pool = this.pools.get(key) || [];

    if (pool.length > 0) {
      const element = pool.pop()!;
      // Reset element state
      element.innerHTML = '';
      element.className = className || '';
      element.removeAttribute('style');
      return element;
    }

    const element = document.createElement(tagName);
    if (className) element.className = className;
    return element;
  }

  static release(element: HTMLElement): void {
    const key = `${element.tagName.toLowerCase()}:${element.className}`;
    const pool = this.pools.get(key) || [];

    if (pool.length < 10) {
      // Limit pool size
      pool.push(element);
      this.pools.set(key, pool);
    }
  }
}

// Performance utility: Create optimized DOM elements with minimal reflows
const createOptimizedElement = (
  tagName: string,
  options: {
    className?: string;
    textContent?: string;
    innerHTML?: string;
    styles?: Record<string, string>;
    attributes?: Record<string, string>;
  } = {},
): HTMLElement => {
  const element = ElementPool.get(tagName, options.className);

  // Batch DOM operations to minimize reflows
  if (options.textContent) element.textContent = options.textContent;
  if (options.innerHTML) element.innerHTML = options.innerHTML;

  if (options.styles) {
    Object.assign(element.style, options.styles);
  }

  if (options.attributes) {
    Object.entries(options.attributes).forEach(([key, value]) => {
      element.setAttribute(key, value);
    });
  }

  return element;
};
/**
 * Add the raw XML toggle button and pre element to a function block
 * Performance optimized version using ElementPool and batch DOM operations
 *
 * @param blockDiv Function block div container
 * @param rawContent Raw XML content to display when toggled
 */
export const addRawXmlToggle = (blockDiv: HTMLDivElement, rawContent: string): void => {
  // Check for existing toggle to avoid duplicates
  if (blockDiv.querySelector('.raw-toggle')) {
    return;
  }

  // Get the original pre element that contains the function call
  const blockId = blockDiv.getAttribute('data-block-id');

  if (blockId) {
    // Try to find the original element with the complete XML
    const originalPre = document.querySelector(`pre[data-block-id="${blockId}"]`);
    if (originalPre) {
      // Use the original content directly
      rawContent = originalPre.textContent?.trim() || rawContent;
    }
  }

  // Use DocumentFragment for efficient DOM construction
  const fragment = document.createDocumentFragment();

  // Create container for raw XML content using optimized element creation
  const rawXmlContainer = createOptimizedElement('div', {
    className: 'function-results-panel xml-results-panel',
    styles: {
      display: 'none',
      marginTop: '12px',
      marginBottom: '4px',
    },
  });

  // Create the pre element for displaying raw XML with batch style assignment
  const rawXmlPre = createOptimizedElement('pre', {
    textContent: rawContent,
    styles: {
      whiteSpace: 'pre-wrap',
      margin: '0',
      padding: '12px',
      fontFamily: 'inherit',
      fontSize: '13px',
      lineHeight: '1.5',
    },
  });

  // Create toggle button with optimized element creation
  const toggleBtn = createOptimizedElement('button', {
    className: 'raw-toggle',
    innerHTML: `${ICONS.CODE}<span>Show Raw XML</span>`,
    styles: {
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      gap: '6px',
    },
  });

  // Cache references for performance
  const textSpan = toggleBtn.querySelector('span')!;
  let isVisible = false;

  // Optimized toggle handler with requestAnimationFrame for smooth transitions
  toggleBtn.onclick = () => {
    isVisible = !isVisible;

    // Use requestAnimationFrame for smooth visual updates
    requestAnimationFrame(() => {
      rawXmlContainer.style.display = isVisible ? 'block' : 'none';
      textSpan.textContent = isVisible ? 'Hide Raw XML' : 'Show Raw XML';
    });
  };

  // Batch DOM operations using fragment
  rawXmlContainer.appendChild(rawXmlPre);
  fragment.appendChild(toggleBtn);

  // Efficiently determine parent and append
  const targetParent = blockDiv.classList.contains('function-buttons')
    ? blockDiv.closest('.function-block') || blockDiv
    : blockDiv;

  // Single batch DOM update
  blockDiv.appendChild(fragment);
  targetParent.appendChild(rawXmlContainer);
};

/**
 * Setup auto-scroll functionality for parameter value divs
 *
 * @param paramValueDiv Parameter value div element
 */
export const setupAutoScroll = (paramValueDiv: ParamValueElement): void => {
  // Auto scroll disabled.
};

/**
 * Stabilize block height to prevent layout shifts during updates
 *
 * @param block The block element
 */
export const stabilizeBlock = (block: HTMLElement): void => {
  if (block.style.height === '') {
    const rect = block.getBoundingClientRect();
    block.style.height = `${rect.height}px`;
    block.style.overflow = 'hidden'; // Optional: prevent content overflow during transition
    if (CONFIG.debug) console.debug(`Stabilized block height: ${rect.height}px`);
  }
};

/**
 * Remove stabilized height
 *
 * @param block The block element
 */
export const unstabilizeBlock = (block: HTMLElement): void => {
  if (block.style.height !== '') {
    block.style.height = '';
    block.style.overflow = ''; // Reset overflow
    if (CONFIG.debug) console.debug('Unstabilized block height');
  }
};

/**
 * Optimized smooth block content update with improved performance
 * Reduces DOM operations and uses efficient transition management
 *
 * @param block The function block to update
 * @param newContent New HTML content to place inside the block
 * @param isStreaming Whether the content is still streaming
 */
export const smoothlyUpdateBlockContent = (
  block: HTMLElement,
  newContent: string | HTMLElement,
  isStreaming: boolean = false,
): void => {
  if (!block) return;

  // Performance: Check update lock more efficiently
  if (!isStreaming && block.hasAttribute('data-smooth-updating')) return;

  // Skip updates for completed blocks to prevent jitter
  const blockId = block.getAttribute('data-block-id');
  if (blockId && (window as any).completedStreams?.has(blockId)) {
    if (CONFIG.debug) console.debug(`Skipping update for completed block ${blockId}`);
    return;
  }

  // Skip updates if block is currently resyncing
  if (blockId && (window as any).resyncingBlocks?.has(blockId)) {
    if (CONFIG.debug) console.debug(`Skipping update for resyncing block ${blockId}`);
    return;
  }

  // Cache essential properties for performance
  const originalClasses = Array.from(block.classList);
  const originalParent = block.parentNode;

  // Mark block as updating
  block.setAttribute('data-smooth-updating', 'true');

  // Performance: Cache dimensions and scroll state
  const originalRect = block.getBoundingClientRect();
  const scrollState = {
    top: block.scrollTop,
    height: block.scrollHeight,
    clientHeight: block.clientHeight,
    wasScrollable: block.scrollHeight > block.clientHeight,
  };

  // Optimized shadow tracker creation
  const shadowTracker = createOptimizedElement('div', {
    styles: { display: 'none' },
    attributes: {
      'data-shadow-for': blockId || 'unknown-block',
      'data-update-in-progress': 'true',
    },
  });

  if (originalParent) {
    originalParent.insertBefore(shadowTracker, block.nextSibling);
  }

  // Performance: Use more efficient content parsing
  const parseNewContent = (content: string | HTMLElement): DocumentFragment => {
    const fragment = document.createDocumentFragment();

    if (typeof content === 'string') {
      // Optimized parsing with better error handling
      try {
        const template = document.createElement('template');
        template.innerHTML = content;
        fragment.appendChild(template.content);
      } catch (e) {
        // Fallback for CSP-restricted environments
        const div = document.createElement('div');
        div.textContent = content;
        fragment.appendChild(div);
      }
    } else {
      fragment.appendChild(content.cloneNode(true));
    }

    return fragment;
  };

  // Pre-calculate new content dimensions efficiently
  const measureNewContent = (fragment: DocumentFragment): number => {
    const tempContainer = createOptimizedElement('div', {
      styles: {
        position: 'absolute',
        visibility: 'hidden',
        width: `${originalRect.width}px`,
        left: '-9999px',
      },
    });

    tempContainer.appendChild(fragment.cloneNode(true));
    document.body.appendChild(tempContainer);
    const height = tempContainer.offsetHeight;
    document.body.removeChild(tempContainer);
    ElementPool.release(tempContainer);

    return height;
  };

  const newContentFragment = parseNewContent(newContent);
  const newHeight = measureNewContent(newContentFragment);

  // Optimized transition setup with batch style assignment
  const transitionDuration = isStreaming ? 150 : 250;
  Object.assign(block.style, {
    height: `${originalRect.height}px`,
    overflow: 'hidden',
    transition: `height ${transitionDuration}ms ease-in-out`,
  });

  // Efficient content wrapper management
  const createContentWrapper = (className: string, opacity: string = '1'): HTMLElement => {
    return createOptimizedElement('div', {
      className: `function-content-wrapper ${className}`,
      styles: {
        opacity,
        transform: opacity === '0' ? 'translateY(10px)' : 'translateY(0)',
        transition: `opacity ${transitionDuration * 0.6}ms ease-out, transform ${transitionDuration * 0.6}ms ease-out`,
      },
    });
  };

  // Move existing content to wrapper
  const oldWrapper = createContentWrapper('function-content-old');
  while (block.firstChild) {
    oldWrapper.appendChild(block.firstChild);
  }

  // Create new content wrapper
  const newWrapper = createContentWrapper('function-content-new', '0');
  newWrapper.appendChild(newContentFragment);

  // Batch DOM operations
  const fragment = document.createDocumentFragment();
  fragment.appendChild(oldWrapper);
  fragment.appendChild(newWrapper);
  block.appendChild(fragment);

  // Use requestAnimationFrame for smoother animations
  requestAnimationFrame(() => {
    // Start height transition
    block.style.height = `${newHeight}px`;

    // Fade out old content
    Object.assign(oldWrapper.style, {
      opacity: '0',
      transform: 'translateY(-10px)',
    });

    // Delayed fade in of new content
    setTimeout(
      () => {
        requestAnimationFrame(() => {
          Object.assign(newWrapper.style, {
            opacity: '1',
            transform: 'translateY(0)',
          });
        });
      },
      isStreaming ? 30 : 50,
    );
  });

  // Optimized mutation observer for DOM changes
  let blockRemoved = false;
  let replacementBlock: HTMLElement | null = null;

  const observer = new MutationObserver(mutations => {
    for (const mutation of mutations) {
      if (mutation.type === 'childList' && Array.from(mutation.removedNodes).includes(block)) {
        blockRemoved = true;

        // Efficiently find replacement block
        const addedElement = Array.from(mutation.addedNodes).find(
          (node): node is HTMLElement =>
            node.nodeType === Node.ELEMENT_NODE &&
            (node as HTMLElement).classList.contains('function-block') &&
            (node as HTMLElement).getAttribute('data-block-id') === blockId,
        ) as HTMLElement | undefined;

        if (addedElement) {
          replacementBlock = addedElement;
        }

        observer.disconnect();
        return;
      }
    }
  });

  observer.observe(originalParent || document.body, { childList: true, subtree: true });

  // Optimized cleanup with better performance
  setTimeout(() => {
    observer.disconnect();

    if (blockRemoved && replacementBlock) {
      // Efficiently transfer state to replacement
      originalClasses.forEach(cls => {
        if (!replacementBlock!.classList.contains(cls)) {
          replacementBlock!.classList.add(cls);
        }
      });

      // Batch style cleanup
      Object.assign(replacementBlock.style, {
        transition: '',
        height: '',
        overflow: '',
      });

      replacementBlock.removeAttribute('data-smooth-updating');

      if (scrollState.wasScrollable) {
        replacementBlock.scrollTop = scrollState.top;
      }
    } else if (document.body.contains(block)) {
      // Efficient content finalization
      while (newWrapper.firstChild) {
        block.appendChild(newWrapper.firstChild);
      }

      // Batch remove old elements
      [oldWrapper, newWrapper].forEach(wrapper => {
        if (block.contains(wrapper)) {
          ElementPool.release(wrapper);
          block.removeChild(wrapper);
        }
      });

      // Batch style reset
      Object.assign(block.style, {
        height: '',
        overflow: '',
        transition: '',
      });

      if (scrollState.wasScrollable) {
        block.scrollTop = scrollState.top;
      }

      block.removeAttribute('data-smooth-updating');
    }

    // Cleanup shadow tracker
    if (shadowTracker.parentNode) {
      shadowTracker.parentNode.removeChild(shadowTracker);
      ElementPool.release(shadowTracker);
    }
  }, transitionDuration + 50);
};

/**
 * Optimized execute button creation with efficient DOM operations
 * Performance improvements: batch DOM operations, use ElementPool, cache queries
 *
 * @param blockDiv Function block div container
 * @param rawContent Raw XML content containing the function call
 */
export const addExecuteButton = (blockDiv: HTMLDivElement, rawContent: string): void => {
  // Check for existing execute button to avoid duplicates
  if (blockDiv.querySelector('.execute-button')) {
    return;
  }

  // Parse the raw XML to extract function name and parameters
  const functionName = extractFunctionName(rawContent);
  const parameters = extractFunctionParameters(rawContent);

  // If we couldn't extract a function name, don't add the button
  if (!functionName) {
    return;
  }

  // Extract call_id from the raw content if available using pre-compiled regex
  const callIdMatch = INVOKE_REGEX.exec(rawContent);
  const callId = callIdMatch?.[2] || `call-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`;

  // Generate content signature for this function call
  const contentSignature = generateContentSignature(functionName, parameters);

  // Use DocumentFragment for efficient DOM construction
  const fragment = document.createDocumentFragment();

  // Create optimized execute button
  const executeButton = createOptimizedElement('button', {
    className: 'execute-button',
    innerHTML: `${ICONS.PLAY}<span>Run</span>`,
    styles: {
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      gap: '6px',
      marginLeft: '0',
    },
  }) as HTMLButtonElement;

  // Create optimized results panel
  const resultsPanel = createOptimizedElement('div', {
    className: 'function-results-panel',
    styles: {
      display: 'none',
      maxHeight: '200px',
      overflow: 'auto',
    },
    attributes: {
      'data-call-id': callId,
      'data-function-name': functionName,
    },
  }) as HTMLDivElement;

  // Create optimized loading indicator
  const loadingIndicator = createOptimizedElement('div', {
    className: 'function-loading',
    styles: {
      display: 'none',
      marginTop: '12px',
      padding: '10px',
      borderRadius: '8px',
      backgroundColor: 'rgba(0, 0, 0, 0.03)',
      border: '1px solid rgba(0, 0, 0, 0.06)',
    },
  }) as HTMLDivElement;

  // Cache DOM references for performance
  const buttonText = executeButton.querySelector('span')!;

  // Optimized click handler with better performance
  executeButton.onclick = () => {
    // Batch button state changes
    executeButton.disabled = true;
    buttonText.style.display = 'none';

    // Add spinner efficiently
    const spinner = createOptimizedElement('span', {
      className: 'execute-spinner',
      innerHTML: ICONS.SPINNER,
      styles: {
        display: 'inline-flex',
        marginLeft: '8px',
      },
    });

    executeButton.appendChild(spinner);

    // Reset results panel state efficiently
    resultsPanel.style.display = 'none';
    resultsPanel.innerHTML = '';

    // Execute the function using mcpHandler
    try {
      const mcpHandler = (window as any).mcpHandler;

      if (!mcpHandler) {
        resetButtonState();
        displayResult(resultsPanel, loadingIndicator, false, 'Error: mcpHandler not found');
        resultsPanel.style.display = 'block';
        return;
      }

      console.debug(`Executing function ${functionName}, call_id: ${callId} with arguments:`, parameters);

      mcpHandler.callTool(functionName, parameters, (result: any, error: any) => {
        resetButtonState();

        // Show results panel
        resultsPanel.style.display = 'block';
        resultsPanel.innerHTML = '';
        resultsPanel.appendChild(loadingIndicator);

        if (error) {
          displayResult(resultsPanel, loadingIndicator, false, error);
        } else {
          displayResult(resultsPanel, loadingIndicator, true, result);

          // Store execution and update history efficiently
          const executionData = storeExecutedFunction(functionName, callId, parameters, contentSignature);
          const historyPanel = (blockDiv.querySelector('.function-history-panel') ||
            createHistoryPanel(blockDiv, callId, contentSignature)) as HTMLDivElement;
          updateHistoryPanel(historyPanel, executionData, mcpHandler);
        }
      });
    } catch (error) {
      resetButtonState();
      resultsPanel.style.display = 'block';
      displayResult(
        resultsPanel,
        loadingIndicator,
        false,
        `Error: ${error instanceof Error ? error.message : String(error)}`,
      );
    }

    // Optimized button reset function
    function resetButtonState() {
      executeButton.disabled = false;
      buttonText.style.display = '';

      if (executeButton.contains(spinner)) {
        executeButton.removeChild(spinner);
        ElementPool.release(spinner);
      }
    }
  };

  // Batch DOM operations
  fragment.appendChild(executeButton);
  blockDiv.appendChild(fragment);

  // Efficiently determine target parent
  const targetParent = blockDiv.classList.contains('function-buttons')
    ? blockDiv.closest('.function-block') || blockDiv
    : blockDiv;

  targetParent.appendChild(resultsPanel);

  // Check for previous executions efficiently
  checkAndDisplayFunctionHistory(blockDiv, functionName, callId, contentSignature);
};

/**
 * Extract function name from raw XML content
 *
 * @param rawContent Raw XML content
 * @returns The function name or null if not found
 */
const extractFunctionName = (rawContent: string): string | null => {
  const invokeMatch = rawContent.match(/<invoke name="([^"]+)"(?:\s+call_id="([^"]+)")?>/);
  return invokeMatch && invokeMatch[1] ? invokeMatch[1] : null;
};

/**
 * Optimized function parameter extraction with pre-compiled regex patterns
 * Performance improvements: use pre-compiled regexes, reduce string operations
 *
 * @param rawContent Raw XML content
 * @returns Object with parameter names and values
 */
export const extractFunctionParameters = (rawContent: string): Record<string, any> => {
  const parameters: Record<string, any> = {};

  // Use pre-compiled regex for better performance
  let match;
  while ((match = PARAM_REGEX.exec(rawContent)) !== null) {
    const name = match[1];
    const type = match[2] || 'string';
    let value: any = match[3].trim();

    // Check for CDATA using pre-compiled regex
    const cdataMatch = CDATA_REGEX.exec(value);
    if (cdataMatch) {
      try {
        value = cdataMatch[1].trim();
        if (CONFIG.debug) console.debug(`Extracted CDATA content for parameter ${name}`);
      } catch (e) {
        console.error(`Failed to extract CDATA content for parameter ${name}:`, e);
        // value already set to original
      }
    }

    // Optimized type parsing with pre-compiled regexes
    switch (type) {
      case 'json':
        try {
          value = JSON.parse(value);
        } catch (e) {
          console.warn(`Failed to parse JSON for parameter '${name}'.`, e);
        }
        break;

      case 'number':
        const num = parseFloat(value);
        if (!isNaN(num)) value = num;
        break;

      case 'boolean':
        value = value.toLowerCase() === 'true';
        break;

      default:
        // Auto-detect only numeric and boolean values for better performance
        // Note: JSON parsing should only happen when explicitly specified with type="json"
        if (NUMBER_REGEX.test(value)) {
          value = parseFloat(value);
        } else if (BOOLEAN_REGEX.test(value)) {
          value = value.toLowerCase() === 'true';
        }
      // Removed automatic JSON parsing to prevent string parameters from being converted to objects
    }

    parameters[name] = value;
  }

  // Reset regex lastIndex for reuse (important for global regexes)
  PARAM_REGEX.lastIndex = 0;
  CDATA_REGEX.lastIndex = 0;

  return parameters;
};

/**
 * Optimized file attachment helper with improved performance
 * Performance improvements: reduce DOM operations, batch state changes, efficient event handling
 */
const attachResultAsFile = async (
  adapter: any,
  functionName: string,
  callId: string,
  rawResultText: string,
  button: HTMLButtonElement,
  iconSpan: HTMLElement,
  skipAutoInsertCheck: boolean = false,
): Promise<{ success: boolean; message: string | null }> => {
  // Early validation for better performance
  if (!adapter?.supportsFileUpload?.() || typeof adapter.attachFile !== 'function') {
    // Efficient error state handling
    const handleUnsupported = () => {
      const originalText = button.classList.contains('insert-result-button') ? 'Insert' : 'Attach File';
      button.textContent = 'Attach Not Supported';
      button.classList.add('attach-error');

      setTimeout(() => {
        button.textContent = originalText;
        button.prepend(iconSpan);
        button.classList.remove('attach-error');
      }, 2000);
    };

    handleUnsupported();
    console.error('Adapter not available or does not support file attachment.');
    return { success: false, message: null };
  }

  const fileName = `${functionName}_result_call_id_${callId}.txt`;
  const file = new File([rawResultText], fileName, { type: 'text/plain' });
  const originalButtonText = button.textContent || 'Attach File';
  let confirmationText: string | null = null;

  // Optimized button state management
  const setButtonState = (text: string, className?: string, disabled: boolean = true) => {
    button.textContent = text;
    button.prepend(iconSpan);
    button.disabled = disabled;

    if (className) {
      button.classList.add(className);
    }
  };

  const resetButtonState = (delay: number = 2000) => {
    setTimeout(() => {
      button.textContent = originalButtonText;
      button.prepend(iconSpan);
      button.classList.remove('attach-success', 'attach-error');
      button.disabled = false;
    }, delay);
  };

  try {
    setButtonState('Attaching...', undefined, true);

    // Skip actual adapter.attachFile call as handled by Perplexity adapter
    const success = true;

    if (success) {
      confirmationText = `Result attached as file: ${fileName}`;
      setButtonState('Attached!', 'attach-success', true);

      // Efficient event dispatch
      const eventDetail = {
        file,
        result: confirmationText,
        isFileAttachment: true,
        fileName,
        confirmationText,
        skipAutoInsertCheck,
      };

      // Use requestAnimationFrame for better performance
      requestAnimationFrame(() => {
        document.dispatchEvent(new CustomEvent('mcp:tool-execution-complete', { detail: eventDetail }));
      });

      resetButtonState();
      return { success: true, message: confirmationText };
    } else {
      setButtonState('Failed', 'attach-error', true);
      resetButtonState();
    }
  } catch (e) {
    setButtonState('Failed', 'attach-error', true);
    resetButtonState();
  }

  return { success: false, message: null };
};

/**
 * Optimized result display with efficient DOM operations and batch processing
 * Performance improvements: reduce DOM queries, batch operations, efficient element creation
 *
 * @param resultsPanel Results panel element
 * @param loadingIndicator Loading indicator element
 * @param success Whether the execution was successful
 * @param result Result or error message
 */
export const displayResult = (
  resultsPanel: HTMLDivElement,
  loadingIndicator: HTMLDivElement,
  success: boolean,
  result: any,
): void => {
  // Cache attributes for performance
  const callId = resultsPanel.getAttribute('data-call-id') || '';
  const functionName = resultsPanel.getAttribute('data-function-name') || '';

  // Efficient cleanup of previous results
  const cleanupPreviousResults = () => {
    // Remove loading indicator if present
    if (loadingIndicator.parentNode === resultsPanel) {
      resultsPanel.removeChild(loadingIndicator);
    }

    // Batch remove existing result content
    const existingResults = resultsPanel.querySelectorAll('.function-result-success, .function-result-error');
    existingResults.forEach(el => resultsPanel.removeChild(el));

    // Remove previous button container
    const existingButtonContainer = resultsPanel.nextElementSibling;
    if (existingButtonContainer?.classList.contains('insert-button-container')) {
      existingButtonContainer.parentNode?.removeChild(existingButtonContainer);
    }
  };

  // Optimized error message processing
  const processErrorMessage = (errorResult: any): string => {
    let errorMessage = '';

    if (typeof errorResult === 'string') {
      errorMessage = errorResult;
    } else if (errorResult && typeof errorResult === 'object') {
      errorMessage = errorResult.message || 'An unknown error occurred';
    } else {
      errorMessage = 'An unknown error occurred';
    }

    // Optimize server error message handling
    if (typeof errorMessage === 'string') {
      const errorMap = {
        SERVER_UNAVAILABLE: 'Server is disconnected. Please check your connection settings.',
        CONNECTION_ERROR: 'Connection to server failed. Please try reconnecting.',
        RECONNECT_ERROR: 'Connection to server failed. Please try reconnecting.',
        SERVER_ERROR: 'Server error occurred. Please check server status.',
      };

      for (const [key, message] of Object.entries(errorMap)) {
        if (errorMessage.includes(key)) {
          return message;
        }
      }
    }

    return errorMessage;
  };

  // Clean up previous results
  cleanupPreviousResults();
  loadingIndicator.style.display = 'none';

  if (success) {
    // Optimized success result processing
    let rawResultText = '';

    // Create result content efficiently
    const resultContent = createOptimizedElement('div', {
      className: 'function-result-success',
    });

    // Process result data efficiently
    if (typeof result === 'object') {
      try {
        rawResultText = JSON.stringify(result, null, 2);
        const pre = createOptimizedElement('pre', {
          textContent: rawResultText,
          styles: {
            fontFamily: 'inherit',
            fontSize: '13px',
            lineHeight: '1.5',
            padding: '0',
            margin: '0',
          },
        });
        resultContent.appendChild(pre);
      } catch (e) {
        rawResultText = String(result);
        resultContent.textContent = rawResultText;
      }
    } else {
      rawResultText = String(result);
      resultContent.textContent = rawResultText;
    }

    // Add result to panel
    resultsPanel.appendChild(resultContent);

    // Create button container efficiently using DocumentFragment
    const fragment = document.createDocumentFragment();
    const buttonContainer = createOptimizedElement('div', {
      className: 'function-buttons insert-button-container',
      styles: {
        display: 'flex',
        justifyContent: 'flex-end',
        marginTop: '10px',
        marginBottom: '10px',
      },
    });

    // Create optimized insert button
    const insertButton = createOptimizedElement('button', {
      className: 'insert-result-button',
      innerHTML: `${ICONS.INSERT}<span>Insert</span>`,
      styles: {
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        gap: '6px',
      },
      attributes: {
        'data-result-id': `result-${callId}-${Date.now()}`,
      },
    }) as HTMLButtonElement;

    // Cache button text element
    const insertButtonText = insertButton.querySelector('span')!;

    // Optimized insert button click handler
    insertButton.onclick = async () => {
      const adapter = window.mcpAdapter || window.getCurrentAdapter?.();

      if (!adapter) {
        const setErrorState = () => {
          insertButton.textContent = 'Failed (No Adapter)';
          insertButton.classList.add('insert-error');
          setTimeout(() => {
            insertButton.innerHTML = `${ICONS.INSERT}<span>Insert</span>`;
            insertButton.classList.remove('insert-error');
          }, 2000);
        };

        setErrorState();
        console.error('Adapter not available.');
        return;
      }

      const wrapperText = `<function_result call_id="${callId}">\n${rawResultText}\n</function_result>`;

      // Check result length and handle accordingly
      const shouldAttachAsFile = 
        (rawResultText.length > MAX_INSERT_LENGTH && WEBSITE_NAME_FOR_MAX_INSERT_LENGTH_CHECK.includes(websiteName)) ||
        (rawResultText.length > MAX_INSERT_LENGTH_M365_COPILOT && WEBSITE_NAME_FOR_M365_COPILOT_CHECK.includes(websiteName));
      
      if (shouldAttachAsFile) {
        const limit = WEBSITE_NAME_FOR_M365_COPILOT_CHECK.includes(websiteName) ? MAX_INSERT_LENGTH_M365_COPILOT : MAX_INSERT_LENGTH;
        console.log(`Result length (${wrapperText.length}) exceeds ${limit}. Attaching as file.`);
        await attachResultAsFile(
          adapter,
          functionName,
          callId,
          wrapperText,
          insertButton,
          insertButton.querySelector('span') as HTMLElement,
          true,
        );
      } else {
        if (typeof adapter.insertTextIntoInput === 'function') {
          // Efficient event dispatch with requestAnimationFrame
          requestAnimationFrame(() => {
            document.dispatchEvent(
              new CustomEvent('mcp:tool-execution-complete', {
                detail: {
                  result: wrapperText,
                  isFileAttachment: false,
                  fileName: '',
                  skipAutoInsertCheck: true,
                },
              }),
            );
          });

          // Optimized success state handling
          insertButton.textContent = 'Inserted!';
          insertButton.classList.add('insert-success');
          insertButton.disabled = true;

          setTimeout(() => {
            insertButton.innerHTML = `${ICONS.INSERT}<span>Insert</span>`;
            insertButton.classList.remove('insert-success');
            insertButton.disabled = false;
          }, 2000);
        } else {
          // Optimized error state
          console.error('Adapter insertTextIntoInput method not found');
          insertButton.textContent = 'Failed (No Insert Method)';
          insertButton.classList.add('insert-error');

          setTimeout(() => {
            insertButton.innerHTML = `${ICONS.INSERT}<span>Insert</span>`;
            insertButton.classList.remove('insert-error');
          }, 2000);
        }
      }
    };

    // Create attach button efficiently
    const attachButton = createOptimizedElement('button', {
      className: 'attach-file-button',
      innerHTML: `${ICONS.ATTACH}<span>Attach File</span>`,
      styles: {
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        gap: '6px',
      },
      attributes: {
        'data-result-id': `attach-${callId}-${Date.now()}`,
      },
    }) as HTMLButtonElement;

    // Optimized attach button handler
    attachButton.onclick = async () => {
      const adapter = window.mcpAdapter || window.getCurrentAdapter?.();
      await attachResultAsFile(
        adapter,
        functionName,
        callId,
        rawResultText,
        attachButton,
        attachButton.querySelector('span') as HTMLElement,
        true,
      );
    };

    // Efficiently build button container
    buttonContainer.appendChild(insertButton);

    // Only add attach button if supported
    const adapter = window.mcpAdapter || window.getCurrentAdapter?.();
    if (adapter?.supportsFileUpload?.()) {
      buttonContainer.appendChild(attachButton);
    }

    // Batch DOM update
    fragment.appendChild(buttonContainer);
    resultsPanel.parentNode?.insertBefore(fragment, resultsPanel.nextSibling);

    // Handle auto-attachment for large results
    const shouldAutoAttach = adapter?.supportsFileUpload?.() && (
      (rawResultText.length > MAX_INSERT_LENGTH && WEBSITE_NAME_FOR_MAX_INSERT_LENGTH_CHECK.includes(websiteName)) ||
      (rawResultText.length > MAX_INSERT_LENGTH_M365_COPILOT && WEBSITE_NAME_FOR_M365_COPILOT_CHECK.includes(websiteName))
    );
    
    if (shouldAutoAttach) {
      const limit = WEBSITE_NAME_FOR_M365_COPILOT_CHECK.includes(websiteName) ? MAX_INSERT_LENGTH_M365_COPILOT : MAX_INSERT_LENGTH;
      console.debug(`Auto-attaching file: Result length (${rawResultText.length}) exceeds ${limit}`);

      // Create efficient fake button for auto-attachment
      const fakeElements = {
        button: createOptimizedElement('button', {
          className: 'insert-result-button',
          styles: { display: 'none' },
        }) as HTMLButtonElement,
        icon: createOptimizedElement('span', {
          styles: { display: 'none' },
        }),
      };

      attachResultAsFile(adapter, functionName, callId, rawResultText, fakeElements.button, fakeElements.icon, false)
        .then(({ success, message }) => {
          if (success) {
            console.debug(`Auto-attached file successfully: ${message}`);
          } else {
            console.error('Failed to auto-attach file.');
            // Fallback to manual attach button
            setTimeout(() => attachButton.click(), 100);
          }

          // Cleanup fake elements
          ElementPool.release(fakeElements.button);
          ElementPool.release(fakeElements.icon);
        })
        .catch(err => {
          console.error('Error auto-attaching file:', err);
          ElementPool.release(fakeElements.button);
          ElementPool.release(fakeElements.icon);
        });
    } else {
      // Dispatch event for normal-sized results
      const wrappedResult = `<function_result call_id="${callId}">\n${rawResultText}\n</function_result>`;
      requestAnimationFrame(() => {
        document.dispatchEvent(
          new CustomEvent('mcp:tool-execution-complete', {
            detail: { result: wrappedResult, skipAutoInsertCheck: false },
          }),
        );
      });
    }
  } else {
    // Optimized error result handling
    const errorMessage = processErrorMessage(result);

    const resultContent = createOptimizedElement('div', {
      className: 'function-result-error',
      textContent: errorMessage,
    });

    resultsPanel.appendChild(resultContent);
  }
};
