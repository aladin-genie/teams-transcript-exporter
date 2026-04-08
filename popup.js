document.addEventListener('DOMContentLoaded', () => {
  // ===== UI ELEMENTS =====
  const statusIndicator = document.getElementById('statusIndicator');
  const notTranscriptView = document.getElementById('notTranscriptView');
  const transcriptView = document.getElementById('transcriptView');
  const exportBtn = document.getElementById('exportBtn');
  const progressDiv = document.getElementById('progress');
  const progressText = document.getElementById('progressText');
  const resultDiv = document.getElementById('result');

  const meetingTitle = document.getElementById('meetingTitle');
  const meetingDuration = document.getElementById('meetingDuration');
  const entryCount = document.getElementById('entryCount');

  const includeTimestamps = document.getElementById('includeTimestamps');
  const includeSpeakerNames = document.getElementById('includeSpeakerNames');
  const cleanText = document.getElementById('cleanText');
  const exportFormat = document.getElementById('exportFormat');

  let currentTab = null;
  let currentFrameId = 0;
  let isExporting = false;

  // ===== PAGE CHECK =====
  async function checkPage() {
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      currentTab = tab;

      if (!tab?.id) {
        showNotReady('No active tab');
        return;
      }

      const isValidUrl = tab.url?.includes('teams.microsoft.com') ||
                         tab.url?.includes('sharepoint.com') ||
                         tab.url?.includes('office.com') ||
                         tab.url?.includes('microsoftstream.com') ||
                         tab.url?.includes('stream.microsoft.com');
      
      if (!isValidUrl) {
        showNotReady('Not a Teams/SharePoint page');
        return;
      }

      statusIndicator.textContent = 'Detecting transcript...';
      statusIndicator.className = 'status-indicator not-ready';

      // Method 1: Try direct script injection to detect transcript
      const detected = await detectTranscriptDirectly(tab.id);
      if (detected) {
        showTranscriptView(detected);
        return;
      }

      // Method 2: Try content script messaging
      await injectContentScript(tab.id);
      
      let frames = [];
      try {
        frames = await chrome.webNavigation.getAllFrames({ tabId: tab.id });
      } catch (e) {
        frames = [{ frameId: 0 }];
      }

      for (const frame of frames) {
        try {
          const response = await sendMessage(tab.id, { action: 'CHECK_TRANSCRIPT_PAGE' }, 3000, frame.frameId);
          if (response?.success && response.isTranscript) {
            currentFrameId = frame.frameId;
            showTranscriptView(response.info);
            return;
          }
        } catch (e) {
          // This frame didn't respond, try next
        }
      }

      showNotTranscriptView();
    } catch (error) {
      console.error('Check error:', error);
      showNotReady('Refresh page & try again');
    }
  }

  // Direct detection via script injection
  async function detectTranscriptDirectly(tabId) {
    try {
      const results = await chrome.scripting.executeScript({
        target: { tabId: tabId, allFrames: true },
        func: () => {
          // Check for Microsoft Stream transcript elements
          const entries = document.querySelectorAll('.ms-List-cell');
          const hasTranscript = entries.length > 0 && 
            document.querySelector('.itemDisplayName-501') !== null &&
            document.querySelector('.entryText-489') !== null;
          
          if (hasTranscript) {
            return {
              found: true,
              entryCount: entries.length,
              title: document.title || 'Microsoft Stream Meeting',
              duration: 'Unknown'
            };
          }
          return { found: false };
        }
      });

      for (const result of results) {
        if (result?.result?.found) {
          return result.result;
        }
      }
    } catch (e) {
      console.log('Direct detection failed:', e);
    }
    return null;
  }

  // Inject content script into page and all frames
  async function injectContentScript(tabId) {
    try {
      await chrome.scripting.executeScript({
        target: { tabId: tabId, allFrames: true },
        files: ['content.js']
      });
      await sleep(300);
    } catch (e) {
      // Script might already be injected
    }
  }

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  function showNotReady(message) {
    statusIndicator.textContent = message;
    statusIndicator.className = 'status-indicator not-ready';
    notTranscriptView.style.display = 'block';
    transcriptView.style.display = 'none';
  }

  function showNotTranscriptView() {
    statusIndicator.textContent = 'No transcript detected';
    statusIndicator.className = 'status-indicator not-ready';
    notTranscriptView.style.display = 'block';
    transcriptView.style.display = 'none';
  }

  function showTranscriptView(info) {
    statusIndicator.textContent = '✓ Transcript ready';
    statusIndicator.className = 'status-indicator ready';
    notTranscriptView.style.display = 'none';
    transcriptView.style.display = 'block';

    meetingTitle.textContent = info?.title || 'Unknown';
    meetingDuration.textContent = info?.duration || 'Unknown';
    entryCount.textContent = info?.entryCount || '0';
  }

  function showProgress(text) {
    progressText.textContent = text;
    progressDiv.style.display = 'flex';
    exportBtn.disabled = true;
    isExporting = true;
  }

  function hideProgress() {
    progressDiv.style.display = 'none';
    exportBtn.disabled = false;
    isExporting = false;
  }

  function showResult(message, isError = false) {
    resultDiv.textContent = message;
    resultDiv.className = isError ? 'result error' : 'result success';
    resultDiv.style.display = 'block';
  }

  // ===== UTILITIES =====
  function sendMessage(tabId, message, timeout = 30000, frameId = undefined) {
    return new Promise((resolve, reject) => {
      const timeoutId = setTimeout(() => {
        reject(new Error('Request timed out'));
      }, timeout);

      const opts = frameId !== undefined ? { frameId } : {};
      chrome.tabs.sendMessage(tabId, message, opts, (response) => {
        clearTimeout(timeoutId);
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else {
          resolve(response);
        }
      });
    });
  }

  async function download(data, filename, mimeType) {
    const blob = new Blob([data], { type: mimeType });
    const url = URL.createObjectURL(blob);
    
    try {
      await chrome.downloads.download({ url, filename, saveAs: false });
    } finally {
      setTimeout(() => URL.revokeObjectURL(url), 60000);
    }
  }

  // ===== EXPORT HANDLER =====
  exportBtn.addEventListener('click', async () => {
    if (isExporting || !currentTab?.id) return;

    showProgress('Starting export...');

    try {
      const options = {
        includeTimestamps: includeTimestamps.checked,
        includeSpeakerNames: includeSpeakerNames.checked,
        cleanText: cleanText.checked
      };

      const response = await sendMessage(currentTab.id, {
        action: 'EXTRACT_TRANSCRIPT',
        options
      }, 60000, currentFrameId);

      if (!response?.success) {
        throw new Error(response?.error || 'Export failed');
      }

      const data = response.data;
      showProgress('Preparing download...');

      const safeTitle = data.metadata.title.replace(/[^a-zA-Z0-9\-_]/g, '-').replace(/-+/g, '-');
      const date = new Date().toISOString().split('T')[0];
      const baseFilename = `${safeTitle}-transcript-${date}`;

      let content, filename, mimeType;

      switch(exportFormat.value) {
        case 'json':
          content = JSON.stringify(data, null, 2);
          filename = `${baseFilename}.json`;
          mimeType = 'application/json';
          break;
        case 'txt':
          content = formatPlainText(data);
          filename = `${baseFilename}.txt`;
          mimeType = 'text/plain';
          break;
        case 'markdown':
          content = formatMarkdown(data);
          filename = `${baseFilename}.md`;
          mimeType = 'text/markdown';
          break;
        default:
          throw new Error('Unknown format');
      }

      await download(content, filename, mimeType);

      hideProgress();
      showResult(`✅ Exported ${data.stats.mergedEntries} entries from ${data.stats.uniqueSpeakers} speakers`);

    } catch (error) {
      console.error('Export error:', error);
      hideProgress();
      showResult('Error: ' + error.message, true);
    }
  });

  // ===== FORMAT HELPERS =====
  function formatPlainText(data) {
    const lines = [
      'MEETING TRANSCRIPT',
      '='.repeat(50),
      '',
      `Title: ${data.metadata.title}`,
      `Duration: ${data.metadata.duration}`,
      `URL: ${data.metadata.url}`,
      `Extracted: ${new Date(data.metadata.extractedAt).toLocaleString()}`,
      '',
      `Speakers: ${data.stats.speakers.join(', ')}`,
      `Total Segments: ${data.stats.mergedEntries}`,
      '',
      '='.repeat(50),
      ''
    ];

    data.entries.forEach(entry => {
      const time = entry.timestamp 
        ? new Date(entry.timestamp).toLocaleTimeString() 
        : '';
      
      if (time) lines.push(`[${time}]`);
      lines.push(`${entry.speaker}:`);
      lines.push(entry.text);
      lines.push('');
    });

    return lines.join('\n');
  }

  function formatMarkdown(data) {
    const lines = [
      `# ${data.metadata.title}`,
      '',
      `- **Duration:** ${data.metadata.duration}`,
      `- **Speakers:** ${data.stats.speakers.join(', ')}`,
      `- **Segments:** ${data.stats.mergedEntries}`,
      '',
      '---',
      ''
    ];

    let currentSpeaker = null;

    data.entries.forEach(entry => {
      if (entry.speaker !== currentSpeaker) {
        lines.push(`\n## ${entry.speaker}\n`);
        currentSpeaker = entry.speaker;
      }

      const time = entry.timestamp 
        ? new Date(entry.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
        : '';

      if (time) {
        lines.push(`**${time}** ${entry.text}\n`);
      } else {
        lines.push(`${entry.text}\n`);
      }
    });

    lines.push('');
    lines.push('---');
    lines.push('');
    lines.push(`*Generated by Teams Transcript Exporter on ${new Date().toLocaleString()}*`);

    return lines.join('\n');
  }

  // ===== INIT =====
  checkPage();
});
