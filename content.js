(function() {
  'use strict';

  // ===== CONFIGURATION =====
  const CONFIG = {
    MAX_RETRIES: 3,
    SCROLL_Dwell: 600,
    MAX_STAGNANT: 12,
    OBSERVER_TIMEOUT: 25000
  };

  const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

  // ===== ROBUST SELECTORS =====
  const SELECTORS = {
    transcriptContainer: [
      '#OneTranscript',
      '.ms-List',
      '[data-tid="transcript-panel"]',
      '[data-tid="transcript-container"]',
      '.transcript-container',
      '[class*="transcript-list" i]',
      '[class*="transcript-content" i]',
      '[role="log"][aria-label*="transcript" i]'
    ],
    transcriptEntries: [
      '.ms-List-cell',
      '[data-automationid="ListCell"]',
      '[data-testid="validate-measured-transcript-item-height"]',
      '[id^="listItem-"]',
      '[id^="sub-entry-"]',
      '[role="listitem"][class*="entryText"]'
    ],
    speaker: [
      '.itemDisplayName-501',
      '.eventSpeakerName-495',
      '[class*="itemDisplayName" i]',
      '[id^="speakerProfile-"]',
      '[data-tid="speaker-name"]',
      '[data-speaker]'
    ],
    timestamp: [
      '[id^="Header-timestamp-"]',
      '.baseTimestamp-497',
      '[class*="baseTimestamp" i]',
      '[data-tid="timestamp"]'
    ]
  };

  function findElement(selectors) {
    for (const selector of selectors) {
      const el = document.querySelector(selector);
      if (el) return el;
    }
    return null;
  }

  function findAllElements(selectors) {
    for (const selector of selectors) {
      const els = document.querySelectorAll(selector);
      if (els.length > 0) return Array.from(els);
    }
    return [];
  }

  // ===== DETECTION =====
  function isTranscriptPage() {
    const url = window.location.href;
    const isTeamsDomain = url.includes('teams.microsoft.com') || 
                          url.includes('sharepoint.com') ||
                          url.includes('office.com') ||
                          url.includes('microsoftstream.com') ||
                          url.includes('stream.microsoft.com');
    
    // Check for transcript panel in DOM
    const hasTranscriptPanel = !!document.querySelector('#OneTranscript') ||
                               !!document.querySelector('.ms-List') ||
                               !!document.querySelector('[aria-label*="transcript" i]') ||
                               !!document.querySelector('[class*="itemDisplayName"]') ||
                               !!document.querySelector('[class*="entryText"]') ||
                               !!document.querySelector('[id^="sub-entry-"]');
    
    if (hasTranscriptPanel) return true;
    
    if (!isTeamsDomain) return false;

    // Check for transcript indicators
    const indicators = [
      () => findElement(SELECTORS.transcriptContainer) !== null,
      () => findAllElements(SELECTORS.transcriptEntries).length > 0,
      () => document.querySelector('[data-tid*="transcript" i]') !== null,
      () => document.querySelector('[role="tab"][aria-label*="transcript" i]') !== null
    ];

    return indicators.some(check => check());
  }

  // ===== METADATA =====
  function getMeetingTitle() {
    const selectors = [
      'h1[data-tid="meeting-title"]',
      '[data-tid="video-title"]',
      '[data-tid="call-title"]',
      'h1',
      'title'
    ];

    for (const selector of selectors) {
      const el = document.querySelector(selector);
      if (el?.textContent?.trim()) {
        const text = el.textContent.trim();
        if (text && text !== 'Microsoft Teams') {
          return text.replace(/\s*\|\s*Microsoft Teams$/, '').replace(/\s+-\s+Microsoft Teams$/, '');
        }
      }
    }

    return 'Unknown Meeting';
  }

  function getMeetingDuration() {
    const durationSelectors = [
      '[data-tid="video-duration"]',
      '.video-duration',
      '[class*="duration" i]'
    ];

    for (const selector of durationSelectors) {
      const el = document.querySelector(selector);
      if (el?.textContent?.trim()) {
        return el.textContent.trim();
      }
    }

    // Calculate from transcript
    const entries = findAllElements(SELECTORS.transcriptEntries);
    if (entries.length >= 2) {
      const first = extractTimestamp(entries[0]);
      const last = extractTimestamp(entries[entries.length - 1]);
      
      if (first && last) {
        const d1 = new Date(first);
        const d2 = new Date(last);
        if (!isNaN(d1) && !isNaN(d2)) {
          const diff = d2 - d1;
          const mins = Math.floor(diff / 60000);
          const secs = Math.floor((diff % 60000) / 1000);
          return `${mins}:${secs.toString().padStart(2, '0')}`;
        }
      }
    }

    return 'Unknown';
  }

  // ===== EXTRACTION =====
  function extractTimestamp(entry) {
    // Data attributes
    const timeAttrs = ['data-start-time', 'data-timestamp', 'data-time'];
    for (const attr of timeAttrs) {
      const val = entry.getAttribute(attr) || 
                  entry.querySelector(`[${attr}]`)?.getAttribute(attr);
      if (val) {
        // Check if milliseconds
        const ms = parseInt(val);
        if (!isNaN(ms) && ms > 1000000000000) {
          return new Date(ms).toISOString();
        }
        return val;
      }
    }

    // Time element
    const timeEl = entry.querySelector('time[datetime]');
    if (timeEl) return timeEl.getAttribute('datetime');

    // Text content
    for (const selector of SELECTORS.timestamp) {
      const el = entry.querySelector(selector);
      if (el?.textContent?.trim()) {
        return el.textContent.trim();
      }
    }

    return null;
  }

  function extractSpeaker(entry) {
    // Data attribute
    const speaker = entry.getAttribute('data-speaker');
    if (speaker) return speaker.trim();

    // Specific selectors
    for (const selector of SELECTORS.speaker) {
      const el = entry.querySelector(selector);
      if (el?.textContent?.trim()) {
        return el.textContent.trim().replace(/:$/, '');
      }
    }

    // Text pattern
    const text = entry.textContent || '';
    const match = text.match(/^([^:]+):/);
    if (match) return match[1].trim();

    return 'Unknown';
  }

  function extractText(entry) {
    // If the entry itself is the text element (id^="sub-entry-")
    if (entry.id?.startsWith('sub-entry-')) {
      return entry.textContent.trim();
    }

    // Try speech entry format first (class suffix changes with Teams updates)
    const textEl = entry.querySelector('[class*="entryText"]');
    if (textEl?.textContent?.trim()) {
      return textEl.textContent.trim();
    }

    // Try event entry format
    const eventTextEl = entry.querySelector('[class*="eventText"]');
    if (eventTextEl?.textContent?.trim()) {
      return eventTextEl.textContent.trim();
    }

    // Fallback: sub-entry children or textHoverColor
    const textSelectors = [
      '[id^="sub-entry-"]',
      '[class*="textHoverColor" i]'
    ];

    for (const selector of textSelectors) {
      const el = entry.querySelector(selector);
      if (el?.textContent?.trim()) {
        return el.textContent.trim();
      }
    }

    return '';
  }

  // ===== LOADING WITH OBSERVER =====
  async function loadAllEntries(onProgress) {
    const container = findElement(SELECTORS.transcriptContainer);
    
    if (!container) {
      // Try page-level scrolling
      return await loadByPageScroll(onProgress);
    }

    return new Promise((resolve, reject) => {
      let previousCount = 0;
      let stagnantCount = 0;
      let timeoutId;

      const checkProgress = () => {
        const entries = findAllElements(SELECTORS.transcriptEntries);
        const currentCount = entries.length;
        
        onProgress?.('Loading transcript...', currentCount);

        if (currentCount === previousCount) {
          stagnantCount++;
          if (stagnantCount >= CONFIG.MAX_STAGNANT) {
            clearTimeout(timeoutId);
            observer.disconnect();
            resolve(currentCount);
            return;
          }
        } else {
          stagnantCount = 0;
          previousCount = currentCount;
        }

        // Scroll to load more
        container.scrollTo({ top: container.scrollHeight, behavior: 'auto' });
      };

      const observer = new MutationObserver(() => {
        setTimeout(checkProgress, CONFIG.SCROLL_Dwell);
      });

      observer.observe(container, { childList: true, subtree: true });

      timeoutId = setTimeout(() => {
        observer.disconnect();
        resolve(previousCount);
      }, CONFIG.OBSERVER_TIMEOUT);

      // Start
      checkProgress();
    });
  }

  async function loadByPageScroll(onProgress) {
    let previousCount = 0;
    let stagnantCount = 0;

    onProgress?.('Loading transcript...', 0);

    while (stagnantCount < CONFIG.MAX_STAGNANT) {
      window.scrollTo({ top: document.body.scrollHeight, behavior: 'auto' });
      await sleep(CONFIG.SCROLL_Dwell);

      const currentCount = findAllElements(SELECTORS.transcriptEntries).length;
      onProgress?.('Loading transcript...', currentCount);

      if (currentCount === previousCount) {
        stagnantCount++;
      } else {
        stagnantCount = 0;
        previousCount = currentCount;
      }
    }

    window.scrollTo({ top: 0, behavior: 'auto' });
    return previousCount;
  }

  // ===== MAIN EXTRACTION =====
  async function extractTranscript(options = {}, onProgress) {
    if (!isTranscriptPage()) {
      throw new Error('No transcript found. Open a Teams meeting recording with transcript enabled.');
    }

    // Load all entries
    const totalEntries = await loadAllEntries(onProgress);

    if (totalEntries === 0) {
      throw new Error('No transcript entries found');
    }

    onProgress?.('Processing transcript...');

    const entries = [];
    const speakers = new Set();
    const elements = findAllElements(SELECTORS.transcriptEntries);

    for (const el of elements) {
      const speaker = extractSpeaker(el);
      const timestamp = extractTimestamp(el);
      const text = extractText(el);

      if (!text) continue;

      speakers.add(speaker);

      const entry = { speaker, text };
      if (options.includeTimestamps !== false && timestamp) {
        entry.timestamp = timestamp;
        entry.startTime = timestamp; // For consistency
      }

      entries.push(entry);
    }

    // Merge consecutive same-speaker entries
    const merged = [];
    let current = null;

    for (const entry of entries) {
      if (current && current.speaker === entry.speaker) {
        current.text += ' ' + entry.text;
        if (entry.timestamp) current.endTime = entry.timestamp;
      } else {
        if (current) merged.push(current);
        current = { ...entry };
      }
    }
    if (current) merged.push(current);

    // Clean text
    if (options.cleanText) {
      merged.forEach(e => {
        e.text = cleanText(e.text);
      });
    }

    return {
      metadata: {
        title: getMeetingTitle(),
        duration: getMeetingDuration(),
        url: window.location.href,
        extractedAt: new Date().toISOString()
      },
      stats: {
        totalEntries: entries.length,
        mergedEntries: merged.length,
        uniqueSpeakers: speakers.size,
        speakers: Array.from(speakers)
      },
      entries: options.includeSpeakerNames !== false ? merged : 
               merged.map(e => ({ text: e.text, timestamp: e.timestamp }))
    };
  }

  function cleanText(text) {
    if (!text) return '';
    
    // Remove filler words
    let cleaned = text
      .replace(/\b(um|uh|er|ah|hm)\b/gi, '')
      .replace(/\b(you know,)\b/gi, '')
      .replace(/\b(sort of|kind of|like,)\b/gi, '')
      .replace(/\s+/g, ' ')
      .replace(/\s+([.,!?])/g, '$1')
      .trim();
    
    return cleaned;
  }

  // ===== MESSAGE HANDLER =====
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    const handlers = {
      CHECK_TRANSCRIPT_PAGE: () => {
        const isTranscript = isTranscriptPage();
        return {
          success: true,
          isTranscript,
          info: isTranscript ? {
            title: getMeetingTitle(),
            duration: getMeetingDuration(),
            entryCount: findAllElements(SELECTORS.transcriptEntries).length
          } : null
        };
      },

      EXTRACT_TRANSCRIPT: async () => {
        try {
          const result = await extractTranscript(request.options, (status, count) => {
            chrome.runtime.sendMessage({
              type: 'TRANSCRIPT_PROGRESS',
              status,
              count
            }).catch(() => {});
          });
          return { success: true, data: result };
        } catch (error) {
          return { success: false, error: error.message };
        }
      }
    };

    const handler = handlers[request.action];
    if (handler) {
      (async () => {
        try {
          const result = await handler();
          sendResponse(result);
        } catch (error) {
          sendResponse({ success: false, error: error.message });
        }
      })();
      return true;
    }
  });

  console.log('[Teams Transcript Exporter] v2.0 loaded');
})();
