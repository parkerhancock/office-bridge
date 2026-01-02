# Office Bridge

A Claude Code plugin for live Microsoft Office automation via Office Add-ins.

## The Problem

Programmatic Office automation typically requires COM/AppleScript (platform-specific), VBA macros (security concerns), or offline file manipulation (no live editing). None give you real-time bidirectional communication with open documents.

**Office Bridge** connects Claude Code directly to running Office apps through the Office.js API, enabling live document reading and editing.

## Features

| App | Capabilities |
|-----|--------------|
| **Word** | Read document structure, edit by paragraph reference, tracked changes |
| **Excel** | Read/write cells and ranges, get sheet structure |
| **PowerPoint** | Create slides, fill placeholders, capture slide images |
| **Outlook** | Read emails, compose replies (limited for Gmail accounts) |

## Installation

### As Claude Code Plugin

```bash
claude plugins add SanctionedCodeList/office-bridge
```

Or via the SCL marketplace (includes other tools):

```bash
claude plugins add SanctionedCodeList/SCL_marketplace
```

### Manual Setup

1. Clone and install dependencies:
```bash
git clone https://github.com/SanctionedCodeList/office-bridge.git
cd office-bridge
./install.sh
```

2. Start the bridge server:
```bash
./server.sh &
```

3. Start add-in dev servers (as needed):
```bash
cd addins/word && npm run dev-server &        # Port 3000
cd addins/excel && npm run dev-server &       # Port 3001
cd addins/powerpoint && npm run dev-server &  # Port 3002
cd addins/outlook && npm run dev-server &     # Port 3003
```

4. Sideload add-ins into Office apps (see `references/setup.md`)

## Usage

Once connected, Claude Code can interact with your Office documents:

```typescript
import { connect } from "./src/client.js";

const bridge = await connect();
const docs = await bridge.documents();  // Word documents
const ppts = await bridge.powerpoint(); // PowerPoint presentations

// Get document tree
const tree = await docs[0].getTree();

// Edit by reference
await docs[0].replaceByRef({ p: 3 }, "New text");

// Capture slide image
const slide = await ppts[0].getSlideImage(1);
```

## How It Works

```
┌─────────────┐     WebSocket     ┌─────────────┐     Office.js    ┌─────────────┐
│ Claude Code │ ◄──────────────► │   Bridge    │ ◄──────────────► │   Office    │
│   (Client)  │                  │   Server    │                  │    Apps     │
└─────────────┘                  └─────────────┘                  └─────────────┘
```

1. Bridge server runs locally, listens for WebSocket connections
2. Office Add-ins connect to bridge via WebSocket
3. Claude Code connects to bridge, discovers available documents
4. Commands flow: Claude → Bridge → Add-in → Office.js → Document

## Documentation

- `references/setup.md` — Detailed setup and sideloading instructions
- `references/word.md` — Word API reference
- `references/powerpoint-api.md` — PowerPoint API reference

## Requirements

- Node.js 18+
- Microsoft Office (Word, Excel, PowerPoint, or Outlook)
- macOS or Windows

## License

MIT
