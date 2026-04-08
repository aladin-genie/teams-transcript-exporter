# Teams Transcript Exporter

A Chrome extension to export Microsoft Teams meeting transcripts to JSON, plain text, or Markdown format.

## Features

- ✅ **Works on Teams/SharePoint recordings** - Export transcripts from meeting recordings
- ✅ **Complete transcript export** - Auto-scrolls to load all entries
- ✅ **Speaker detection** - Identifies and groups speakers
- ✅ **Timestamp extraction** - Optional timestamps for each entry
- ✅ **Smart merging** - Combines consecutive entries from same speaker
- ✅ **Text cleaning** - Remove filler words (um, uh, etc.)
- ✅ **Multiple formats** - JSON (structured), Plain Text, or Markdown
- ✅ **Privacy-first** - All data stays local, no external API calls

## Installation

### Load Unpacked (Developer Mode)

1. Open Chrome and navigate to `chrome://extensions/`
2. Enable "Developer mode" (toggle in top right)
3. Click "Load unpacked"
4. Select the `teams-transcript-exporter` folder
5. The extension icon will appear in your toolbar

## Usage

### Prerequisites

- You must have **read access** to the meeting transcript
- The meeting must have **transcription enabled** and the transcript must be visible

### Steps

1. Open a Microsoft Teams meeting recording in Chrome
2. Click on the **"Transcript"** tab to view the transcript
3. Click the extension icon in your toolbar
4. The extension will detect if a transcript is available
5. Select your export options:
   - Include timestamps (on/off)
   - Include speaker names (on/off)
   - Clean text - remove filler words (on/off)
   - Format: JSON, Plain Text, or Markdown
6. Click **"Export Transcript"**
7. The file will download automatically

## Output Formats

### JSON Format (Structured)

```json
{
  "metadata": {
    "title": "Weekly Team Sync",
    "duration": "45:30",
    "url": "https://teams.microsoft.com/...",
    "extractedAt": "2026-04-07T20:00:00.000Z"
  },
  "stats": {
    "totalEntries": 156,
    "mergedEntries": 89,
    "uniqueSpeakers": 4,
    "speakers": ["John Doe", "Jane Smith", "Bob Wilson", "Alice Brown"]
  },
  "entries": [
    {
      "speaker": "John Doe",
      "timestamp": "2026-04-07T14:00:15.000Z",
      "text": "Welcome everyone to our weekly sync. Let's go through the agenda."
    },
    {
      "speaker": "Jane Smith",
      "timestamp": "2026-04-07T14:00:45.000Z",
      "text": "Thanks John. I'll start with the project updates."
    }
  ]
}
```

### Plain Text Format

```
Meeting Transcript
==================

Title: Weekly Team Sync
Duration: 45:30
URL: https://teams.microsoft.com/...
Extracted: 2026-04-07T20:00:00.000Z

Speakers: John Doe, Jane Smith, Bob Wilson, Alice Brown
Total Entries: 89

==================

[2:00:15 PM] John Doe:
Welcome everyone to our weekly sync. Let's go through the agenda.

[2:00:45 PM] Jane Smith:
Thanks John. I'll start with the project updates.
```

### Markdown Format

```markdown
# Weekly Team Sync

**Duration:** 45:30

**Speakers:** John Doe, Jane Smith, Bob Wilson, Alice Brown

---

### John Doe

**2:00:15 PM** Welcome everyone to our weekly sync. Let's go through the agenda.

### Jane Smith

**2:00:45 PM** Thanks John. I'll start with the project updates.

---

*Extracted on 4/7/2026, 8:00:00 PM*
```

## How It Works

1. **Detection**: The extension checks if you're viewing a Teams/SharePoint meeting transcript
2. **Loading**: It auto-scrolls through the transcript panel to load all entries
3. **Extraction**: Parses each entry to extract speaker name, timestamp, and text
4. **Merging**: Groups consecutive entries from the same speaker
5. **Export**: Formats and downloads the complete transcript

## Files

- `manifest.json` - Extension configuration
- `popup.html` - Extension UI
- `popup.css` - UI styles
- `popup.js` - UI logic
- `content.js` - Transcript scraper
- `icon*.png` - Extension icons

## Limitations

- **Read access required**: You must have permission to view the transcript
- **Visible transcript**: The transcript tab must be open and visible
- **Teams/SharePoint only**: Works on official Microsoft platforms
- **Large transcripts**: Very long meetings may take time to scroll through

## Troubleshooting

### "No transcript detected"

- Make sure you're on a meeting recording page
- Click the "Transcript" tab in Teams
- Wait for the transcript to fully load
- Refresh the page and try again

### "Extension not loaded. Refresh the page."

- The content script needs to be injected
- Refresh the Teams/SharePoint page
- Then click the extension icon again

### Partial/missing transcript

- Some transcript entries may not load until you scroll
- The extension auto-scrolls, but very large transcripts may need manual scrolling first
- Try scrolling through the entire transcript before exporting

### Wrong speaker names

- The extension tries multiple methods to detect speakers
- Some transcript formats may not clearly separate speaker names
- Check the raw JSON export to see what was detected

## Privacy & Security

**🔒 100% Local Processing**
- No data leaves your browser
- No external API calls
- No analytics or tracking
- Exports saved directly to Downloads folder

## Comparison with Existing Tools

| Tool | License | Features | Our Approach |
|------|---------|----------|--------------|
| `ms-teams-sharepoint-downloader` | MIT | Video + Transcript | Simpler, transcript-only |
| `MS-Teams-Transcripter` | CC BY-NC | Live captions | Post-meeting transcripts |
| **Our Extension** | MIT | Focused transcript export | Clean, simple, local |

## License

MIT - For personal use only. Respect your organization's data policies.

## Future Enhancements

- [ ] Bulk export multiple meeting transcripts
- [ ] AI-powered meeting summaries
- [ ] Action item extraction
- [ ] Search within transcripts
- [ ] Integration with note-taking apps

## Credits

Inspired by:
- [brendangooden/ms-teams-sharepoint-downloader](https://github.com/brendangooden/ms-teams-sharepoint-downloader) - MIT License
