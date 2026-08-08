# SillyTavern ⇄ JanitorAI Lorebook Converter

A vibe-coded, browser-based converter for SillyTavern and JanitorAI lorebooks. Fully static, runs entirely client-side — no backend, no accounts, no data leaves your browser.



## Features

- **ST World Info ↔ JAI Lorebook JSON** — full bidirectional conversion
- **JAI `loreEntries` scripts → ST** — with category-based position mapping and automatic scenario/personality splitting
- **JAI dynamic lore scripts → ST** — parses triggered/nested entries, keyword expansion, and probability gates
- **JAI `if-chain` scripts → ST** — handles the common `if (lastMessage.includes(...))` pattern
- **Script repair** — sanitizes broken quotes, unclosed strings, and missing commas in pasted/exported JAI scripts before parsing
- **.docx import** — zero-dependency text extraction, so lorebooks pasted into a Word doc can be dropped straight in
- **Simple / Advanced modes** — Simple mode is drop-file → convert → download; Advanced exposes per-category position mapping, outlet naming, merge behavior, and text prefixes
- **Preview table** — see parsed entries, warnings, and format detection before exporting
- **Drag-and-drop** file import, or paste text directly

**Limits:** ST → JAI export only produces the basic JAI JSON format — it does not generate JAI scripts.



## Usage

- Download the **[latest release](https://github.com/dwenne/LBconverter/releases)** and open it in your browser, or
- Use the hosted version. Mobile-friendly, nothing to install:
  - [GitHub Pages](https://dwenne.github.io/LBconverter)
  - [Neocities](https://drevaine.neocities.org/lbconverter)

Drop a lorebook file (or paste its contents) in, pick Simple or Advanced mode, hit Convert, download the result.



## Related project

**[Character Card Editor](https://github.com/dwenne/CharacterCardEditor)** — a companion tool for building and editing SillyTavern character cards from scratch. Useful once your lorebook's converted and you're ready to attach it to a card.
