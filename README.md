# SillyTavern ⇄ JanitorAI Lorebook Converter
 
A vibe-coded, browser-based converter for SillyTavern and JanitorAI lorebooks. Fully static, runs entirely client-side — no backend, no accounts, no data leaves your browser.
 
## Features
 
- **ST World Info ↔ JAI Lorebook JSON** — full bidirectional conversion
- **JAI `loreEntries` scripts → ST** — with category-based position mapping and automatic scenario/personality splitting
- **JAI dynamic lore scripts → ST** — parses triggered entries, `Shifts`, keyword expansion (`char.entity`), and probability gates. This format supports more logic than ST does natively, so some conditions are approximated rather than reproduced exactly: trigger-tag content is inlined into whichever entry fires it (if two entries share a tag, that content can appear twice in the same turn), `Shifts` become separate entries gated on the parent's keywords, and gate types ST has no equivalent for (emotion detection, tag-gates, `maxMessages`, `notAll`) are dropped. The converter always surfaces a warning listing exactly what was approximated.
- **JAI `if-chain` scripts → ST** — handles the common `if (lastMessage.includes(...))` pattern
- **Script repair** — recovers scripts with common hand-editing typos instead of failing outright: unclosed quotes on `scenario`/`personality` lines, missing commas between array items, single-quoted strings with unescaped apostrophes, and — the one that actually matters most — a stray extra `];` that silently truncates the entry list and orphans everything after it. Repairs only kick in when strict parsing fails, and the UI always tells you exactly what was patched so you can double-check the affected entries.
- **.docx import** — zero-dependency text extraction (hand-parses the `.docx` ZIP container and inflates it with the browser's native `DecompressionStream`, no external libraries), so lorebooks pasted into a Word doc can be dropped straight in
- **Simple / Advanced modes** — Simple mode is drop-file → convert → download; Advanced exposes per-category position mapping, outlet naming, merge behavior, and text prefixes
- **Preview table** — see parsed entries, warnings, and format detection before exporting
- **Drag-and-drop** file import, or paste text directly
**Limits:** ST → JAI export only produces the basic JAI JSON format — it does not generate JAI scripts.
 
## Usage
 
- Download the **[latest release](https://github.com/dwenne/LBconverter/releases)** and open it in your browser, or
- Use the hosted version. Mobile-friendly, nothing to install:
  - [GitHub Pages](https://dwenne.github.io/LBconverter)
  - [Neocities](https://drevaine.neocities.org/lbconverter)
 
## Related project
 
**[Character Card Editor](https://github.com/dwenne/CharacterCardEditor)** — a companion tool for building and editing SillyTavern character cards from scratch. Useful once your lorebook's converted and you're ready to attach it to a card.
