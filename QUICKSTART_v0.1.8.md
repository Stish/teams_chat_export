# Teams Chat Export v0.2.0 - Quick Start Guide

## What's New in v0.1.8?

Your Teams Chat Export has been completely redesigned! Instead of a single massive HTML file, it's now a **multi-file system** (SPA - Single Page Application) that's faster, smarter, and more organized.

## Quick Setup (5 minutes)

### 1. Configure Your Token
```bash
# Edit config.py
ACCESS_TOKEN = "your_graph_api_token_here"
```

### 2. Run the Export
```bash
python teams_chat_export.py
```

### 3. Open the Export
- Navigate to `output/index.html` in your browser
- Done! All your Teams chats are ready to explore

## What You'll See

```
┌─────────────────────────────────────────────┐
│ 🔍 Search Box                              │
│                                             │
│ 💻 Code 🔗 URLs 🖼️ Images | Clear          │
├─────────────────────────────────────────────┤
│ 📋 ONE-ON-ONE CHATS                        │
│   • Chat with John Doe                      │
│   • Chat with Jane Smith                    │
│                                             │
│ 👥 GROUP CHATS                             │
│   • Project Team                            │
│   • Design Squad                            │
│                                             │
│ 📅 MEETING CHATS                           │
│   • Q1 Planning                             │
│                                             │
│ 📢 CHANNELS                                │
│   • Team Alpha > General                   │
│   • Team Alpha > Announcements             │
│                                             │
│ ↑ ↓ 🌙                                      │
└─────────────────────────────────────────────┘
        ↓ Click a chat to view messages
```

## Main Features

### 🔍 **Universal Search**
- Search text across ALL chats at once
- Filter by: Code blocks, URLs, Images
- Real-time highlighting in results

### 🔄 **Smart Navigation**
- Sidebar always visible
- Click any chat to load it
- Breadcrumbs to jump back
- Remember last viewed chat

### 📊 **Statistics Dashboard**
- Total message count
- Chat breakdown by type
- Storage usage calculation
- Top 10 active participants
- Message date range

### 🌙 **Theme Toggle**
- Dark mode / Light mode
- Automatic theme switching
- Preference saved in browser

### 📸 **Image Viewer**
- Click any image to zoom
- Fullscreen view
- Press ESC to close

## File Structure

```
output/
├── index.html              👈 Open this file in browser
├── stats.html              📊 Statistics dashboard
├── chats/                  💬 Individual chat files
│   ├── Chat_with_John_Doe.html
│   ├── Project_Team.html
│   └── ...
└── assets/
    ├── style.css           🎨 Styling
    ├── script.js           ⚙️ Functionality
    ├── search-index.json   🔍 Search data
    └── img/                🖼️ Downloaded images
```

## How Search Works

### Basic Search
1. Type in the search box
2. Results update in real-time
3. Matched text highlights in green
4. Works across ALL chats

### Filter by Type
- **💻 Code**: Messages containing code blocks
- **🔗 URLs**: Messages with web links
- **🖼️ Images**: Messages with pictures

### Examples
- Search for `"bug fix"` → Find all bug-related discussions
- Search for `"deadline"` + filter `URLs` → Find deadlines with links
- Filter `Images` → See all messages with attachments

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| ↑ | Scroll to top |
| ↓ | Scroll to bottom |
| ESC | Close image viewer |
| Click 🌙 | Toggle dark mode |

## Troubleshooting

### "Can't find index.html"
- Make sure you opened `output/index.html`
- Not `output/stats.html` or `output/chats/*`

### "Search not working"
- Check browser developer console (F12)
- Ensure JavaScript is enabled
- Try a different browser

### "Some chats not showing"
- Make sure export completed successfully
- Check `output/chats/` folder
- Look for error messages in console

### "Images not loading"
- Verify `output/assets/img/` folder exists
- Check image files are there
- Try clearing browser cache

## Tips & Tricks

### 💡 Pro Tips

1. **Save Export to Cloud**
   - Upload entire `output/` folder to Google Drive, Dropbox, or OneDrive
   - Share the folder for team access
   - Export is fully self-contained

2. **Print for Archival**
   - Open any chat or stats page
   - Use browser print (Ctrl+P or Cmd+P)
   - Save as PDF for archival

3. **Search Before Reading**
   - Use search to find specific conversations
   - Combine filters (e.g., Code + specific name)
   - Saves time browsing large exports

4. **Share Individual Chats**
   - Open a chat file (`output/chats/*.html`)
   - Can be shared standalone
   - Still includes search functionality

5. **Backup Your Export**
   - Keep export folder in version control (Git)
   - Or zip entire `output/` folder
   - Or upload to cloud storage

## Performance

### Initial Load
- ~1-2 seconds to open index
- Sidebar loads immediately
- Content loads on demand

### Search Performance
- ~0.1-0.5 seconds for typical search
- Works smoothly with 10,000+ messages
- No server required (fully client-side)

### Large Exports (100,000+ messages)
- Individual chat files still load quickly
- Search index optimized for speed
- Consider breaking into multiple exports

## Common Questions

### Q: Can I edit the export?
A: Not recommended. HTML files are generated fresh each export. Any edits will be lost on next export. Instead, modify configuration before exporting.

### Q: Can I share the export?
A: Yes! Upload entire `output/` folder to cloud storage or shared drive. Anyone can open `index.html` in a browser.

### Q: How do I update an old export?
A: Run export again. New files will overwrite old ones in the `output/` folder.

### Q: Can I combine multiple exports?
A: Not easily. Each export is self-contained. Consider using multiple browser tabs or windows instead.

### Q: What happens if I delete a chat file?
A: It will disappear from sidebar on next refresh. Re-run export to restore it.

### Q: Can I use this on mobile?
A: Yes! The export is fully responsive. Use any mobile browser to open `index.html`.

### Q: Is my data secure?
A: Yes. Everything stays local on your device. No data is sent to servers. Works fully offline after loading.

## Getting Help

1. **Check Console for Errors**
   - Press F12 to open Developer Tools
   - Click "Console" tab
   - Look for red error messages

2. **Verify Export Completed**
   - Check console output for success message
   - Verify `output/stats.html` exists
   - Check `output/assets/` has files

3. **Try Different Browser**
   - Chrome, Firefox, or Safari
   - Clear cache and cookies
   - Disable extensions

## Next Steps

- 🔧 Check [README.md](README.md) for configuration options
- 🐛 Report issues or feature requests on GitHub

---

**Version**: v0.1.8  
**Last Updated**: July 06, 2026  
**Ready to explore your Teams data!** 🚀
