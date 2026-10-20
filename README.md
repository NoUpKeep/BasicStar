# BasicStar Word Processor
 
A full-screen plain-text word processor written in QB64-PE, inspired by
the classic word processors of the 1980s — particularly WordStar.
 
![Language](https://img.shields.io/badge/language-QB64--PE-blue)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)
 
---
 
## Features
 
- Full-screen editing (80x25 text mode)
- Shift+Arrow key text selection with visual highlight
- Cut, copy and paste (internal clipboard)
- Save and open plain text files
- Unsaved changes indicator
- Line and column display
- Up to 5,000 lines per document
---
 
## Controls
 
| Key | Action |
|-----|--------|
| Arrow keys | Move cursor |
| Home / End | Start / end of line |
| Page Up / Down | Scroll by screen |
| Shift + Arrow | Select text |
| Escape | Cancel selection |
| Backspace | Delete left / join lines |
| Delete | Delete right / join lines |
| Enter | New line |
| Ctrl+C | Copy selection |
| Ctrl+X | Cut selection |
| Ctrl+V | Paste |
| Ctrl+S | Save |
| Ctrl+O | Open |
| Ctrl+N | New document |
| Ctrl+Q | Quit |
 
---
 
## Requirements
 
- [QB64-PE](https://github.com/QB64-Phoenix-Edition/QB64pe) (QB64 Phoenix Edition)
No external libraries or dependencies required.
 
---
 
## Getting Started
 
1. Download or clone this repository.
2. Open `BasicStar.bas` in QB64-PE.
3. Press F5 to compile and run.
---
 
## Roadmap
 
- [ ] Undo / redo
- [ ] Find and replace
- [ ] Word wrap
- [x] "Save first?" prompt on New and Quit
---
 
## Acknowledgements
 
Developed in collaboration with [Claude](https://www.anthropic.com) by Anthropic.
 
---
 
## Licence
 
This project is released into the public domain. Do with it as you will.
 
