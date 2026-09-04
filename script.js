const POSITIONS = window.POSITIONS;
const POSITIONS_RU = window.POSITIONS_RU;
const I18N = window.I18N;
const BADGE_I18N = window.BADGE_I18N;
const WARN_I18N = window.WARN_I18N;

let LANG = "en";
function getPositions() {
  return LANG === "ru" ? POSITIONS_RU : POSITIONS;
}

function skipStringOrComment(src, i) {
  const c = src[i];
  if (c === "/" && src[i + 1] === "/") {
    const nl = src.indexOf("\n", i);
    return nl === -1 ? src.length : nl + 1;
  }
  if (c === "/" && src[i + 1] === "*") {
    const end = src.indexOf("*/", i + 2);
    return end === -1 ? src.length : end + 2;
  }
  if (c === '"' || c === "'" || c === "`") {
    let j = i + 1;
    while (j < src.length) {
      if (src[j] === "\\") { j += 2; continue; }
      if (src[j] === c) return j + 1;
      j++;
    }
    return j;
  }
  return i;
}

function detectFormat(text) {
  const t = text.trim();
  if (t.startsWith("{") || t.startsWith("[")) {
    try {
      const p = JSON.parse(t);
      if (Array.isArray(p) && p.length && ("keys" in p[0] || "keyword" in p[0] || "keywords" in p[0] || "content" in p[0]))
        return { type: "jai_json", parsed: p };
      if (p && p.entries && typeof p.entries === "object" && !Array.isArray(p.entries))
        return { type: "st_json", parsed: p };
      const ks = Object.keys(p);
      if (ks.length && ks.every((k) => !isNaN(parseInt(k))))
        return { type: "st_json", parsed: { entries: p } };
    } catch (e) {}
  }
  if (/loreEntries\s*=\s*\[/.test(t)) return { type: "jai_script", parsed: null };
  if (detectDynamicLore(t)) return { type: "dynamic_lore", parsed: null };
  const msgVar = detectMsgVar(t);
  const msgVarRe = new RegExp(msgVar.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\.includes\\s*\\(");
  if (msgVarRe.test(t) && /context\.character\.(personality|scenario)\s*(\+=|\.replace\s*\()/.test(t))
    return { type: "if_chain", parsed: null };
  return null;
}

async function extractDocxText(arrayBuffer) {
  const view = new DataView(arrayBuffer);
  const bytes = new Uint8Array(arrayBuffer);
  let eocdOffset = -1;
  for (let i = bytes.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === 0x06054b50) { eocdOffset = i; break; }
  }
  if (eocdOffset === -1) throw new Error("Not a valid ZIP/DOCX file");
  const cdOffset = view.getUint32(eocdOffset + 16, true);
  const cdEntries = view.getUint16(eocdOffset + 10, true);
  let offset = cdOffset, targetOffset = -1, targetCompMethod = -1, targetCompSize = -1;
  for (let i = 0; i < cdEntries; i++) {
    if (view.getUint32(offset, true) !== 0x02014b50) break;
    const compMethod = view.getUint16(offset + 10, true);
    const compSize = view.getUint32(offset + 20, true);
    const nameLen = view.getUint16(offset + 28, true);
    const extraLen = view.getUint16(offset + 30, true);
    const commentLen = view.getUint16(offset + 32, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);
    const nameBytes = bytes.subarray(offset + 46, offset + 46 + nameLen);
    const name = new TextDecoder().decode(nameBytes);
    if (name === "word/document.xml") {
      targetOffset = localHeaderOffset; targetCompMethod = compMethod; targetCompSize = compSize; break;
    }
    offset += 46 + nameLen + extraLen + commentLen;
  }
  if (targetOffset === -1) throw new Error("word/document.xml not found in DOCX");
  if (view.getUint32(targetOffset, true) !== 0x04034b50) throw new Error("Invalid local file header");
  const lfNameLen = view.getUint16(targetOffset + 26, true);
  const lfExtraLen = view.getUint16(targetOffset + 28, true);
  const dataOffset = targetOffset + 30 + lfNameLen + lfExtraLen;
  const compressedData = bytes.subarray(dataOffset, dataOffset + targetCompSize);

  let xmlString = "";
  if (targetCompMethod === 0) {
    xmlString = new TextDecoder().decode(compressedData);
  } else if (targetCompMethod === 8) {
    if (typeof DecompressionStream === "undefined") throw new Error("Browser doesn't support local DOCX extraction. Try copy-pasting instead.");
    const ds = new DecompressionStream("deflate-raw");
    const writer = ds.writable.getWriter();
    writer.write(compressedData); writer.close();
    const reader = ds.readable.getReader();
    const chunks = [];
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }
    const totalLen = chunks.reduce((acc, c) => acc + c.length, 0);
    const decompressed = new Uint8Array(totalLen);
    let ptr = 0;
    for (const c of chunks) { decompressed.set(c, ptr); ptr += c.length; }
    xmlString = new TextDecoder().decode(decompressed);
  } else {
    throw new Error("Unsupported DOCX compression method");
  }

  let paragraphs = xmlString.split("</w:p>");
  let text = paragraphs.map((p) => {
    let cleanP = p.replace(/<w:br[^>]*>/g, "<w:t>\n</w:t>");
    let tMatch = cleanP.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g);
    if (!tMatch) return "";
    let line = tMatch.map((t) => t.replace(/<w:t[^>]*>/, "").replace(/<\/w:t>/, "")).join("");
    return line.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  }).filter((p) => p.trim() !== "").join("\n\n");
  return text.replace(/[\u2018\u2019]/g, "'").replace(/[\u201C\u201D]/g, '"');
}

function sanitizeQuotes(code) {
  let out = "", i = 0;
  while (i < code.length) {
    if (code[i] === "/" && code[i + 1] === "/") {
      const end = code.indexOf("\n", i);
      const e = end === -1 ? code.length : end + 1;
      out += code.slice(i, e); i = e; continue;
    }
    if (code[i] === "/" && code[i + 1] === "*") {
      const end = code.indexOf("*/", i + 2);
      const e = end === -1 ? code.length : end + 2;
      out += code.slice(i, e); i = e; continue;
    }
    if (code[i] === "`") {
      let j = i + 1;
      while (j < code.length) {
        if (code[j] === "\\") { j += 2; continue; }
        if (code[j] === "`") { j++; break; }
        j++;
      }
      out += code.slice(i, j); i = j; continue;
    }
    if (code[i] === '"') {
      let j = i + 1;
      while (j < code.length) {
        if (code[j] === "\\") { j += 2; continue; }
        if (code[j] === '"') { j++; break; }
        j++;
      }
      out += code.slice(i, j); i = j; continue;
    }
    if (code[i] === "'") {
      let j = i + 1, inner = "";
      while (j < code.length) {
        if (code[j] === "\\" && j + 1 < code.length) {
          const nx = code[j + 1];
          if (nx === "'") inner += "'";
          else if (nx === '"') inner += '\\"';
          else inner += code[j] + nx;
          j += 2; continue;
        }
        if (code[j] === "'") { j++; break; }
        if (code[j] === '"') inner += '\\"';
        else inner += code[j];
        j++;
      }
      out += '"' + inner + '"'; i = j; continue;
    }
    out += code[i]; i++;
  }
  return out;
}

let lastScriptRepairStats = null;

// --- Keyword sanitization ---
// If a keyword field is unnaturally big (any item too long or containing a
// newline), the whole field is treated as corrupted — not item by item. The
// entire thing is moved into the entry's content (nothing discarded), `key`
// is left empty, and the entry is flagged so the fix-it popup lets the user
// type real trigger words in by hand.
let lastBadKeyStats = null;
const KEY_MAX_LEN = 80;
function isBadKeyword(s) {
  if (typeof s !== "string") return true;
  if (s.includes("\n")) return true;
  if (s.length > KEY_MAX_LEN) return true;
  return false;
}
function sanitizeKeyList(arr) {
  if (!Array.isArray(arr) || !arr.length) return { keys: arr, movedText: "" };
  if (!arr.some(isBadKeyword)) return { keys: arr, movedText: "" };
  if (!lastBadKeyStats) lastBadKeyStats = { entries: 0, dropped: 0 };
  lastBadKeyStats.entries++;
  lastBadKeyStats.dropped += arr.length;
  return { keys: [], movedText: arr.join(", ") };
}

function fixUnclosedLineStrings(code) {
  const lines = code.split("\n");
  let fixed = 0;
  for (let li = 0; li < lines.length; li++) {
    let line = lines[li];
    const m = line.match(/^(\s*[A-Za-z_$][\w$]*\s*:\s*)(['"])/);
    if (!m) continue;
    const q = m[2];
    const rest = line.slice(m[0].length);
    let closed = false, esc = false;
    for (const ch of rest) {
      if (esc) { esc = false; continue; }
      if (ch === "\\") { esc = true; continue; }
      if (ch === q) { closed = true; break; }
    }
    if (!closed) {
      if (/\[/.test(rest) && !/\]/.test(rest)) continue;
      const trailMatch = line.match(/,\s*$/);
      if (trailMatch) line = line.slice(0, line.length - trailMatch[0].length) + q + ",";
      else line = line.replace(/\s*$/, "") + q;
      lines[li] = line; fixed++;
    }
  }
  return { code: lines.join("\n"), fixed };
}

function fixMissingCommasInLists(code) {
  let fixed = 0;
  const out = code.replace(/\[([^[\]]*)\]/gs, (whole, inner) => {
    const fixedInner = inner.replace(/(['"])(\s+)(?=['"])/g, (m2, q, ws) => { fixed++; return q + "," + ws; });
    return "[" + fixedInner + "]";
  });
  return { code: out, fixed };
}

function repairScriptText(code) {
  const r1 = fixUnclosedLineStrings(code);
  const r2 = fixMissingCommasInLists(r1.code);
  if (!lastScriptRepairStats) lastScriptRepairStats = { unclosedStrings: 0, missingCommas: 0, prematureCloses: 0 };
  lastScriptRepairStats.unclosedStrings += r1.fixed;
  lastScriptRepairStats.missingCommas += r2.fixed;
  return r2.code;
}

function findEntryArrayEnd(code, start) {
  let depth = 0, i = start;
  const falseCloses = [];
  while (i < code.length) {
    const nextI = skipStringOrComment(code, i);
    if (nextI > i) { i = nextI; continue; }
    const c = code[i];
    if (c === "[") { depth++; i++; continue; }
    if (c === "]") {
      depth--;
      if (depth === 0) {
        let j = i + 1;
        while (j < code.length) {
          const skip = skipStringOrComment(code, j);
          if (skip > j) { j = skip; continue; }
          if (/\s/.test(code[j]) || code[j] === ";") { j++; continue; }
          break;
        }
        if (code[j] === "{" && /\b(keywords|tag)\s*:/.test(code.slice(j, j + 400))) {
          falseCloses.push([i, j]); depth = 1; i = j; continue;
        }
        return { end: i, falseCloses };
      }
    }
    i++;
  }
  return { end: -1, falseCloses };
}

function extractRepairedArray(code, start) {
  const { end, falseCloses } = findEntryArrayEnd(code, start);
  if (end === -1) return null;
  let arrStr = code.slice(start, end + 1);
  for (let k = falseCloses.length - 1; k >= 0; k--) {
    const [closeIdx, resumeIdx] = falseCloses[k];
    arrStr = arrStr.slice(0, closeIdx - start) + "," + arrStr.slice(resumeIdx - start);
  }
  if (!lastScriptRepairStats) lastScriptRepairStats = { unclosedStrings: 0, missingCommas: 0, prematureCloses: 0 };
  lastScriptRepairStats.prematureCloses = falseCloses.length;
  return arrStr;
}

function parseScript(code) {
  const startIdx = code.indexOf("loreEntries");
  if (startIdx === -1) throw new Error("loreEntries not found");
  const arrStart = code.indexOf("[", code.indexOf("=", startIdx));
  if (arrStart === -1) throw new Error("Array not found");
  lastScriptRepairStats = null;
  const arrStr = extractRepairedArray(code, arrStart);
  if (arrStr === null) throw new Error("Unclosed bracket");

  try {
    const r = new Function('"use strict";return(' + arrStr + ")")();
    if (Array.isArray(r)) return r;
  } catch (_) {}
  try {
    const r = new Function('"use strict";return(' + sanitizeQuotes(arrStr) + ")")();
    if (Array.isArray(r)) return r;
  } catch (_) {}

  const repairedCode = repairScriptText(code);
  const repairedStart = repairedCode.indexOf("[", repairedCode.indexOf("=", repairedCode.indexOf("loreEntries")));
  const repairedArrStr = repairedStart === -1 ? null : extractRepairedArray(repairedCode, repairedStart);
  if (repairedArrStr === null) throw new Error("Unclosed bracket");
  try {
    const r = new Function('"use strict";return(' + repairedArrStr + ")")();
    if (Array.isArray(r)) return r;
  } catch (_) {}
  try {
    const r = new Function('"use strict";return(' + sanitizeQuotes(repairedArrStr) + ")")();
    if (Array.isArray(r)) return r;
    throw new Error("Result is not an array");
  } catch (e) { throw new Error("Parse failed: " + e.message); }
}

function detectMsgVar(code) {
  const m = code.match(/\b(?:const|let|var)\s+(\w+)\s*=[^;]*context\.chat\.last_messages?\b/);
  return m ? m[1] : "lastMessage";
}

function extractKeywordConstants(code) {
  const map = {};
  const declRe = /\b(?:const|let|var)\s+(\w+)\s*=\s*([(\[])/g;
  let m;
  while ((m = declRe.exec(code)) !== null) {
    const name = m[1];
    const openChar = m[2];
    const closeChar = openChar === "(" ? ")" : "]";
    let i = declRe.lastIndex - 1, depth = 0;
    const bodyStart = i;
    while (i < code.length) {
      const nextI = skipStringOrComment(code, i);
      if (nextI > i) { i = nextI; continue; }
      const c = code[i];
      if (c === openChar) depth++;
      else if (c === closeChar) { depth--; if (depth === 0) { i++; break; } }
      i++;
    }
    const body = code.slice(bodyStart, i);
    const kws = [];
    const strRe = /'([^'\\]*(?:\\.[^'\\]*)*)'|"([^"\\]*(?:\\.[^"\\]*)*)"/g;
    let sm;
    while ((sm = strRe.exec(body)) !== null) kws.push(sm[1] !== undefined ? sm[1] : sm[2]);
    if (kws.length) map[name] = kws;
    declRe.lastIndex = i;
  }
  return map;
}

function parseIfChain(code) {
  const entries = [];
  const msgVar = detectMsgVar(code);
  const kwConsts = extractKeywordConstants(code);

  function getBlock(src, braceIdx) {
    let depth = 0, i = braceIdx;
    while (i < src.length) {
      const nextI = skipStringOrComment(src, i);
      if (nextI > i) { i = nextI; continue; }
      const c = src[i];
      if (c === "{") depth++;
      else if (c === "}") { depth--; if (depth === 0) return { body: src.slice(braceIdx + 1, i), end: i }; }
      i++;
    }
    return null;
  }

  function extractStrEnd(src, from) {
    let i = from;
    while (i < src.length && src[i] !== "'" && src[i] !== '"' && src[i] !== "`") i++;
    if (i >= src.length) return null;
    const q = src[i];
    i++;
    let val = "";
    const ESCAPES = { n: "\n", t: "\t", r: "\r" };
    while (i < src.length) {
      if (src[i] === "\\" && i + 1 < src.length) {
        const esc = src[i + 1]; val += ESCAPES[esc] !== undefined ? ESCAPES[esc] : esc; i += 2; continue;
      }
      if (src[i] === q) { i++; break; }
      val += src[i]; i++;
    }
    return { value: val, end: i };
  }

  function extractStr(src, from) {
    const r = extractStrEnd(src, from);
    return r ? r.value : "";
  }

  function getKws(cond) {
    const msgVarEsc = msgVar.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp(msgVarEsc + "\\.includes\\s*\\(([^)]*)\\)", "g");
    const kws = [];
    let m;
    while ((m = re.exec(cond)) !== null) {
      const parts = m[1].split("||").map((s) => s.trim());
      parts.forEach((p) => {
        const sm = p.match(/^['"](.+)['"]$/);
        if (sm) kws.push(sm[1]);
        else if (kwConsts[p]) kws.push(...kwConsts[p]);
      });
    }
    return [...new Set(kws)];
  }

  function extractReplaceArg(body, callIdx) {
    const openParen = body.indexOf("(", callIdx);
    if (openParen === -1) return "";
    const first = extractStrEnd(body, openParen + 1);
    if (!first) return "";
    const second = extractStrEnd(body, first.end);
    return second ? second.value : "";
  }

  function getContent(body) {
    const pi = body.indexOf("context.character.personality +=");
    const si = body.indexOf("context.character.scenario +=");
    let p = pi !== -1 ? extractStr(body, pi + 32) : "";
    let s = si !== -1 ? extractStr(body, si + 29) : "";
    if (!p) {
      const pr = body.indexOf("context.character.personality.replace(");
      if (pr !== -1) p = extractReplaceArg(body, pr);
    }
    if (!s) {
      const sr = body.indexOf("context.character.scenario.replace(");
      if (sr !== -1) s = extractReplaceArg(body, sr);
    }
    return { p, s };
  }

  function walk(src, minMsg) {
    let i = 0;
    while (i < src.length) {
      const nextI = skipStringOrComment(src, i);
      if (nextI > i) { i = nextI; continue; }
      const prevOk = i === 0 || /\W/.test(src[i - 1]);
      const nextOk = /[\s(]/.test(src[i + 2] || "");
      if (src.slice(i, i + 2) === "if" && prevOk && nextOk) {
        let j = i + 2;
        while (j < src.length && src[j] !== "(") j++;
        if (j >= src.length) { i++; continue; }
        let depth = 0, k = j;
        while (k < src.length) {
          if (src[k] === "(") depth++;
          else if (src[k] === ")") { depth--; if (depth === 0) break; }
          k++;
        }
        const cond = src.slice(j + 1, k);
        let bi = k + 1;
        while (bi < src.length && src[bi] !== "{") bi++;
        if (bi >= src.length) { i++; continue; }
        const block = getBlock(src, bi);
        if (!block) { i++; continue; }

        const mg = cond.match(/context\.chat\.message_count\s*[>>=]+\s*(\d+)/);
        if (mg) {
          walk(block.body, Math.max(minMsg, parseInt(mg[1]) + 1));
          i = block.end + 1; continue;
        }

        const kws = getKws(cond);
        if (kws.length) {
          const { p, s } = getContent(block.body);
          if (p || s) entries.push({ keywords: kws, personality: p, scenario: s, minMessages: minMsg, category: "If-Chain" });
          walk(block.body, minMsg);
          i = block.end + 1; continue;
        }
        i = block.end + 1; continue;
      }
      i++;
    }
  }
  walk(code, 0);
  return entries;
}

let lastDynLoreStats = null;
function dynArr(x) { return Array.isArray(x) ? x : x == null ? [] : [x]; }
function dynProbToPercent(v) {
  if (v == null) return null;
  if (typeof v === "number") return Math.round(v <= 1 ? v * 100 : v);
  const s = String(v).trim();
  const n = parseFloat(s.replace("%", ""));
  if (!isFinite(n)) return null;
  return Math.round(s.indexOf("%") !== -1 ? n : n * 100);
}

function extractTopLevelLiteral(code, varName, openChar) {
  const closeChar = openChar === "[" ? "]" : "}";
  const declRe = new RegExp("\\b(?:const|let|var)\\s+" + varName + "\\s*=\\s*");
  const m = declRe.exec(code);
  if (!m) return null;
  const start = code.indexOf(openChar, m.index + m[0].length);
  if (start === -1) return null;
  let depth = 0, i = start, end = -1;
  while (i < code.length) {
    const nextI = skipStringOrComment(code, i);
    if (nextI > i) { i = nextI; continue; }
    const c = code[i];
    if (c === openChar) depth++;
    else if (c === closeChar) { depth--; if (depth === 0) { end = i; break; } }
    i++;
  }
  if (end === -1) return null;
  return code.slice(start, end + 1);
}

function dynSafeEval(str) {
  if (!str) return null;
  try { return new Function('"use strict";return(' + str + ")")(); } catch (_) {}
  try { return new Function('"use strict";return(' + sanitizeQuotes(str) + ")")(); } catch (_) { return null; }
}

function dynHasUnsupportedGates(e) {
  return !!(e.notAll || e.andAnyTags || e.andAllTags || e.notAnyTags || e.notAllTags || e.requireEmotion || e.blockEmotion || e.maxMessages != null || e["prev.requireAny"] || e["prev.requireAll"] || e["prev.requireNone"] || e["prev.andAny"] || e["prev.andAll"] || e["prev.notAny"]);
}

function dynWordGates(e) {
  const r = e.requires || {};
  return {
    any: [].concat(dynArr(e.requireAny), dynArr(e.andAny), dynArr(r.any)),
    all: [].concat(dynArr(e.requireAll), dynArr(e.andAll), dynArr(r.all)),
    none: [].concat(dynArr(e.requireNone), dynArr(e.notAny), dynArr(e.block), dynArr(r.none)),
  };
}

function detectDynamicLore(code) { return /\bdynamicLore\s*=\s*\[/.test(code); }

function parseDynamicLore(code) {
  lastScriptRepairStats = null;
  let loreStr = extractTopLevelLiteral(code, "dynamicLore", "[");
  let dynamicLore = loreStr && dynSafeEval(loreStr);
  let workingCode = code;
  if (!Array.isArray(dynamicLore)) {
    workingCode = repairScriptText(code);
    loreStr = extractTopLevelLiteral(workingCode, "dynamicLore", "[");
    dynamicLore = loreStr && dynSafeEval(loreStr);
  }
  if (!Array.isArray(dynamicLore)) throw new Error("Could not parse dynamicLore array");

  const entityStr = extractTopLevelLiteral(workingCode, "ENTITY_DB", "{");
  const ENTITY_DB = (entityStr && dynSafeEval(entityStr)) || {};
  const relStr = extractTopLevelLiteral(workingCode, "RELATIONSHIP_DB", "[");
  const RELATIONSHIP_DB = (relStr && dynSafeEval(relStr)) || [];

  let all = dynamicLore.slice();
  for (const name in ENTITY_DB) {
    if (!Object.prototype.hasOwnProperty.call(ENTITY_DB, name)) continue;
    const ent = ENTITY_DB[name];
    if (ent && Array.isArray(ent.lore)) {
      for (const le of ent.lore) all.push(Object.assign({ _entityOwner: name }, le));
    }
  }

  const entityKwRe = /^char\.([a-z0-9_]+)$/i;
  function expandKeywords(list) {
    if (!Array.isArray(list)) return [];
    const out = [];
    for (const kw of list) {
      const m = String(kw).match(entityKwRe);
      if (m && ENTITY_DB[m[1].toLowerCase()]) {
        const ent = ENTITY_DB[m[1].toLowerCase()];
        out.push(m[1].toLowerCase());
        if (Array.isArray(ent.aliases)) out.push(...ent.aliases);
      } else out.push(kw);
    }
    return [...new Set(out)];
  }
  for (const e of all) {
    if (e.keywords) e.keywords = expandKeywords(e.keywords);
    if (Array.isArray(e.Shifts)) {
      for (const sh of e.Shifts) if (sh.keywords) sh.keywords = expandKeywords(sh.keywords);
    }
  }

  const byTag = {};
  for (const e of all) { if (e.tag) (byTag[e.tag] = byTag[e.tag] || []).push(e); }

  const stats = { inlined: 0, shifts: 0, nameBlocked: 0, relationships: 0, unsupportedGates: 0 };
  function collectTriggered(triggers, depth, seen) {
    let p = "", s = "";
    if (!triggers || !triggers.length || depth > 4) return { p, s };
    for (const tag of triggers) {
      const list = byTag[tag];
      if (!list) continue;
      for (const te of list) {
        if (seen.has(te)) continue;
        seen.add(te);
        if (dynHasUnsupportedGates(te)) stats.unsupportedGates++;
        if (te.personality) p += "\n\n" + te.personality;
        if (te.scenario) s += "\n\n" + te.scenario;
        stats.inlined++;
        if (Array.isArray(te.triggers) && te.triggers.length) {
          const nested = collectTriggered(te.triggers, depth + 1, seen);
          p += nested.p; s += nested.s;
        }
      }
    }
    return { p, s };
  }

  function toFlat(e, ownerLabel) {
    if (dynHasUnsupportedGates(e)) stats.unsupportedGates++;
    let keywords = e.keywords && e.keywords.length ? e.keywords : e["prev.keywords"] && e["prev.keywords"].length ? e["prev.keywords"] : [];
    const g = dynWordGates(e);
    const filters = {};
    if (keywords.length) {
      if (g.none.length) filters.notWith = g.none;
      else if (g.all.length) filters.requiresAll = g.all;
      else if (g.any.length) filters.requiresAny = g.any;
    } else if (g.any.length) {
      keywords = g.any.slice();
      if (g.none.length) filters.notWith = g.none;
    } else if (g.all.length) {
      keywords = [g.all[0]];
      if (g.all.length > 1) filters.requiresAll = g.all.slice(1);
      if (g.none.length) filters.notWith = g.none;
    } else if (g.none.length) {
      filters.notWith = g.none;
      stats.unsupportedGates++;
    }
    const { p: trigP, s: trigS } = collectTriggered(e.triggers, 0, new Set());
    const flat = {
      keywords, personality: (e.personality || "") + trigP, scenario: (e.scenario || "") + trigS,
      priority: typeof e.priority === "number" ? e.priority : 3, minMessages: typeof e.minMessages === "number" ? e.minMessages : 0, category: ownerLabel || "Advanced Lore", filters,
    };
    const prob = dynProbToPercent(e.probability);
    if (prob != null) flat.probability = prob;
    if (Array.isArray(e.nameBlock) && e.nameBlock.length) { flat.nameBlock = e.nameBlock.slice(); stats.nameBlocked++; }
    return flat;
  }

  const out = [];
  for (const e of all) {
    if (e.tag && !(e.keywords && e.keywords.length) && !e["prev.keywords"]) continue;
    const ownerLabel = e._entityOwner ? "Character \u2014 " + e._entityOwner : "Advanced Lore";
    out.push(toFlat(e, ownerLabel));
    if (Array.isArray(e.Shifts) && e.Shifts.length) {
      for (const sh of e.Shifts) {
        stats.shifts++;
        if (dynHasUnsupportedGates(sh)) stats.unsupportedGates++;
        const shKeys = sh.keywords && sh.keywords.length ? sh.keywords : e.keywords || [];
        const shFilters = e.keywords && e.keywords.length ? { requiresAll: e.keywords.slice() } : {};
        const shFlat = {
          keywords: shKeys, personality: sh.personality || "", scenario: sh.scenario || "",
          priority: typeof e.priority === "number" ? e.priority : 3, minMessages: typeof e.minMessages === "number" ? e.minMessages : 0, category: ownerLabel + " (Shift)", filters: shFilters,
        };
        const shProb = dynProbToPercent(sh.probability);
        if (shProb != null) shFlat.probability = shProb;
        if (Array.isArray(sh.nameBlock) && sh.nameBlock.length) { shFlat.nameBlock = sh.nameBlock.slice(); stats.nameBlocked++; }
        out.push(shFlat);
      }
    }
  }

  for (const rel of RELATIONSHIP_DB) {
    if (!Array.isArray(rel.pair) || rel.pair.length < 2) continue;
    stats.relationships++;
    const [a, b] = rel.pair;
    const aEnt = ENTITY_DB[a] || {}; const bEnt = ENTITY_DB[b] || {};
    const aKeys = [a].concat(Array.isArray(aEnt.aliases) ? aEnt.aliases : []);
    const bKeys = [b].concat(Array.isArray(bEnt.aliases) ? bEnt.aliases : []);
    out.push({
      keywords: aKeys, personality: rel.injection || "", scenario: "", priority: 3, minMessages: 0, category: "Relationship", filters: { requiresAll: bKeys },
    });
  }
  lastDynLoreStats = stats;
  return out;
}

function smartScenPos(category) {
  const c = (category || "").toLowerCase().replace(/[_-]/g, " ");
  if (/\bnpc\b|character/.test(c)) return "1";
  if (/location|place|dorm|building|region/.test(c)) return "1";
  return "0";
}

function catLabel(category) {
  const c = (category || "").toLowerCase().replace(/[_-]/g, " ").trim();
  if (/^world/.test(c)) return "World";
  if (/hierarchy/.test(c)) return "Social Hierarchy";
  if (/^culture/.test(c)) return "Culture";
  if (/^location/.test(c)) return "Location";
  if (/\bnpc\b|^character/.test(c)) return "NPC";
  if (/underground/.test(c)) return "Underground";
  if (/faction/.test(c)) return "Faction";
  if (/empire/.test(c)) return "Empire";
  if (/species/.test(c)) return "Species";
  if (/creature/.test(c)) return "Creature";
  if (/mechanic/.test(c)) return "Mechanics";
  if (/secret/.test(c)) return "Secret";
  if (/cosm/.test(c)) return "Cosmology";
  if (/if\s*chain/.test(c)) return "If-Chain";
  return c.split(" ").map((w) => (w ? w[0].toUpperCase() + w.slice(1) : "")).join(" ");
}

function entryLabel(entry) {
  const cat = (entry.category || "").toLowerCase();
  if (/\bnpc\b|^character/.test(cat) && entry.personality) {
    const m = entry.personality.match(/knows?\s+([A-Z][a-zA-Záàâäéèêëíìîïóòôöúùûüñç]+(?:\s+[A-Z][a-zA-Záàâäéèêëíìîïóòôöúùûüñç]+)*)\s+(?:very well|:)/i);
    if (m) {
      const pts = m[1].trim().split(/\s+/);
      return pts.length > 2 ? pts[0] + " " + pts[pts.length - 1] : m[1].trim();
    }
  }
  if (!(entry.keywords && entry.keywords[0])) {
    const text = (entry.personality || entry.scenario || "").replace(/\{\{[^}]*\}\}/g, "").trim();
    if (text) {
      const words = text.split(/\s+/).filter(Boolean);
      const snippet = words.slice(0, 5).join(" ").replace(/[.,;:!?]+$/, "");
      if (snippet) return snippet + (words.length > 5 ? "\u2026" : "");
    }
  }
  const kw = entry.keywords && entry.keywords[0] ? entry.keywords[0] : entry.category || "Entry";
  return kw.split(/\s+/).map((w) => (w ? w[0].toUpperCase() + w.slice(1) : "")).join(" ");
}

function cleanPersonality(text) {
  if (!text) return "";
  let t = text.trim().replace(/^,\s*/, "");
  t = t.charAt(0).toUpperCase() + t.slice(1);
  if (!/[.!?]$/.test(t)) t += ".";
  return t;
}
function cleanScenario(text) { return (text || "").trim(); }

function buildCatPanel(entries) {
  const cats = [...new Set(entries.map((e) => e.category || ""))].sort();
  const grid = document.getElementById("catGrid");
  while (grid.children.length > 3) grid.removeChild(grid.lastChild);
  const isMerged = document.getElementById("mergePairsCheck").checked;

  cats.forEach((cat) => {
    const defS = smartScenPos(cat), defP = MODE === "simple" ? "6:0" : "7";
    const nameEl = document.createElement("div");
    nameEl.className = "cat-name";
    nameEl.textContent = cat || "(none)";

    function makeCell(cls, defVal, kind) {
      const cell = document.createElement("div");
      cell.className = "pos-cell";
      const sel = document.createElement("select");
      sel.className = "cat-sel " + cls;
      sel.dataset.cat = cat;
      if (cls.includes("pers") && isMerged) sel.disabled = true;
      Object.entries(getPositions()).forEach(([v, lbl]) => {
        const o = document.createElement("option");
        o.value = v; o.textContent = lbl;
        if (v === defVal) o.selected = true;
        sel.appendChild(o);
      });
      const dep = document.createElement("input");
      dep.type = "number"; dep.className = "depth-inp";
      dep.min = 1; dep.max = 99; dep.value = 4;
      dep.dataset.cat = cat; dep.dataset.kind = kind;
      if (defVal.startsWith("6:")) dep.classList.add("show");
      sel.addEventListener("change", () => {
        dep.classList.toggle("show", sel.value.startsWith("6:"));
        finalizeScript(true);
      });
      dep.addEventListener("input", () => finalizeScript(true));
      cell.appendChild(sel); cell.appendChild(dep);
      return cell;
    }

    grid.appendChild(nameEl);
    grid.appendChild(makeCell("cat-sel-scen", "" + defS, "scen"));
    grid.appendChild(makeCell("cat-sel-pers", defP, "pers"));
  });
  document.getElementById("catPanel").classList.add("visible");
}

function getCatMap() {
  const m = {};
  document.querySelectorAll(".cat-sel-scen").forEach((s) => {
    if (!m[s.dataset.cat]) m[s.dataset.cat] = {};
    m[s.dataset.cat].scen = s.value; m[s.dataset.cat].scenDepth = 4;
  });
  document.querySelectorAll(".cat-sel-pers").forEach((s) => {
    if (!m[s.dataset.cat]) m[s.dataset.cat] = {};
    m[s.dataset.cat].pers = s.value; m[s.dataset.cat].persDepth = 4;
  });
  document.querySelectorAll(".depth-inp").forEach((inp) => {
    const c = inp.dataset.cat, k = inp.dataset.kind, v = parseInt(inp.value) || 4;
    if (m[c]) m[c][k === "scen" ? "scenDepth" : "persDepth"] = v;
  });
  return m;
}

function makeSTEntry(uid, key, keysecondary, selective, selectiveLogic, comment, content, posKey, order, delay, outletName, depthNum, extra) {
  const parts = String(posKey).split(":");
  const position = parseInt(parts[0]) || 0;
  const role = parts.length > 1 ? parseInt(parts[1]) : position === 7 ? null : 0;
  const isOutlet = position === 7;
  const ex = extra || {};
  return {
    uid, key, keysecondary, comment, content,
    constant: !!ex.constant, vectorized: false, selective, selectiveLogic, addMemo: true, order, position,
    disable: false, ignoreBudget: false, excludeRecursion: false, preventRecursion: false, matchPersonaDescription: false,
    matchCharacterDescription: false, matchCharacterPersonality: false, matchCharacterDepthPrompt: false, matchScenario: false,
    matchCreatorNotes: false, delayUntilRecursion: false, probability: ex.probability ?? 100, useProbability: ex.probability != null ? true : false,
    depth: depthNum ?? 4, outletName: isOutlet ? outletName || "" : "", group: "", groupOverride: false, groupWeight: 100,
    scanDepth: null, caseSensitive: null, matchWholeWords: null, useGroupScoring: null, automationId: "", role, sticky: 0, cooldown: 0, delay, triggers: [], displayIndex: uid,
    characterFilter: ex.characterFilter || { isExclude: false, names: [], tags: [] },
  };
}

function scriptToST(entries, catMap, outlet, opts) {
  const { merge, scenPref, persPref } = opts;
  const result = {};
  let uid = 0;
  entries.forEach((e) => {
    const keys = Array.isArray(e.keywords) ? e.keywords : [];
    let sec = [], selective = false, selectiveLogic = 0;
    if (e.filters) {
      if (e.filters.notWith && e.filters.notWith.length) { sec = e.filters.notWith; selectiveLogic = 2; selective = true; }
      else if (e.filters.requiresAll && e.filters.requiresAll.length) { sec = e.filters.requiresAll; selectiveLogic = 1; selective = true; }
      else if (e.filters.requiresAny && e.filters.requiresAny.length) { sec = e.filters.requiresAny; selectiveLogic = 0; selective = true; }
    }
    const order = (e.priority ?? 5) * 10, delay = e.minMessages ?? 0;
    const cmap = catMap[e.category] || { scen: "0", scenDepth: 4, pers: "7", persDepth: 4 };
    let label = entryLabel(e);
    if (e.category && e.category !== "If-Chain") label = catLabel(e.category) + " \u2014 " + label;
    const extra = {};
    if (typeof e.probability === "number") extra.probability = e.probability;
    if (Array.isArray(e.nameBlock) && e.nameBlock.length) { extra.characterFilter = { isExclude: true, names: e.nameBlock.slice(), tags: [] }; }
    if (!keys.length) extra.constant = true;
    let sText = e.scenario ? cleanScenario(e.scenario) : "";
    let pText = e.personality ? cleanPersonality(e.personality) : "";
    if (sText && scenPref) sText = scenPref.replace(/\\n/g, "\n") + sText;
    if (pText && persPref) pText = persPref.replace(/\\n/g, "\n") + pText;

    if (merge) {
      const combined = [sText, pText].filter(Boolean).join("\n\n");
      if (combined) {
        const st = makeSTEntry(uid, keys, sec, selective, selectiveLogic, label, combined, cmap.scen, order, delay, outlet, cmap.scenDepth, extra);
        st._type = "merged"; result[uid] = st; uid++;
      }
    } else {
      if (sText) {
        const st = makeSTEntry(uid, keys, sec, selective, selectiveLogic, label, sText, cmap.scen, order, delay, outlet, cmap.scenDepth, extra);
        st._type = "scenario"; result[uid] = st; uid++;
      }
      if (pText) {
        const st = makeSTEntry(uid, keys, sec, selective, selectiveLogic, label, pText, cmap.pers, order, delay, outlet, cmap.persDepth, extra);
        st._type = "personality"; result[uid] = st; uid++;
      }
    }
  });
  return { entries: result, count: uid };
}

function stToJai(e, uid) {
  const k = Array.isArray(e.key) ? e.key : e.key ? [e.key] : [];
  const jai = {
    id: uid, keyword: k, keys: k, keysecondary: Array.isArray(e.keysecondary) ? e.keysecondary : [],
    comment: e.comment || "", content: e.content || "", enabled: !e.disable, constant: !!e.constant,
    probability: e.useProbability === false ? 100 : (e.probability ?? 100), order: e.order ?? 100, group: e.group || "",
    groupWeight: e.groupWeight ?? 100, groupOverride: !!e.groupOverride, useGroupScoring: e.useGroupScoring ?? false,
    selective: e.selective ?? true, selectiveLogic: e.selectiveLogic ?? 0, minMessages: e.delay ?? 0,
    caseSensitive: e.caseSensitive ?? false, matchWholeWords: e.matchWholeWords ?? true,
  };
  return { ...e, ...jai };
}

function jaiToSt(e, uid) {
  const label = e.comment || e.name || e.title || "";
  let k = e.keys || e.keywords || e.keyword || e.key || [];
  if (typeof k === "string") k = k.split(",").map((s) => s.trim()).filter(Boolean);
  else if (!Array.isArray(k)) k = [k];
  const sanitized = sanitizeKeyList(k);
  k = sanitized.keys;
  let ks = e.keysecondary || e.secondary_keys || [];
  if (typeof ks === "string") ks = ks.split(",").map((s) => s.trim()).filter(Boolean);
  else if (!Array.isArray(ks)) ks = [ks];
  let content = e.content || "";
  if (sanitized.movedText) content = sanitized.movedText + (content ? "\n\n" + content : "");
  const st = makeSTEntry(uid, k, ks, e.selective ?? true, e.selectiveLogic ?? 0, label, content, "0", e.order ?? 100, e.minMessages ?? 0, "", 4);
  st._type = "jai";
  st._flaggedKeys = !!sanitized.movedText;
  return { ...e, ...st };
}

function wt(key) { return (WARN_I18N[LANG] || WARN_I18N.en)[key]; }
function buildWarnings(entries, fmt) {
  const w = [];
  if (fmt === "st") {
    const pos = entries.filter((e) => e.position !== undefined && e.position !== 4);
    if (pos.length) w.push(wt("nonDefaultPos")(pos.length));
    if (entries.filter((e) => e.sticky > 0).length) w.push(wt("stickyDropped"));
    if (entries.filter((e) => e.vectorized).length) w.push(wt("vectorizedDropped"));
  } else if (fmt === "jai") {
    if (lastBadKeyStats && lastBadKeyStats.entries) w.push(wt("badKeywords")(lastBadKeyStats.entries, lastBadKeyStats.dropped));
    w.push(wt("jaiPosSet"));
  } else if (fmt === "jai_script" || fmt === "if_chain" || fmt === "dynamic_lore") {
    if (lastScriptRepairStats && (lastScriptRepairStats.unclosedStrings || lastScriptRepairStats.missingCommas || lastScriptRepairStats.prematureCloses)) {
      w.push(wt("scriptRepaired")(lastScriptRepairStats));
    }
    w.push(wt("scriptSecurity"));
    const isMerged = document.getElementById("mergePairsCheck").checked;
    if (isMerged) w.push(wt("scriptMerged")); else w.push(wt("scriptSplit"));
    w.push(wt("outletReminder"));

    if (fmt === "jai_script") {
      const f = entries.filter((e) => e.filters && Object.keys(e.filters).length > 0);
      if (f.length) w.push(wt("filtersUsed")(f.length));
    }
    if (fmt === "if_chain") w.push(wt("ifChainDup"));
    if (fmt === "dynamic_lore" && lastDynLoreStats) {
      const s = lastDynLoreStats;
      if (s.inlined) w.push(wt("dynLoreInlined")(s.inlined));
      if (s.shifts) w.push(wt("dynLoreShifts")(s.shifts));
      if (s.unsupportedGates) w.push(wt("dynLoreGates")(s.unsupportedGates));
      if (s.nameBlocked) w.push(wt("dynLoreNameBlock")(s.nameBlocked));
      if (s.relationships) w.push(wt("dynLoreRelationships")(s.relationships));
    }
  }
  return w;
}

const modal = document.getElementById("previewModal");
const mTitle = document.getElementById("pmTitle");
const mBody = document.getElementById("pmBody");
const mClose = document.getElementById("pmClose");
function closeModal() {
  modal.classList.remove("visible");
  mBody.classList.remove("editing");
  fixQueue = [];
}
function showPreview(title, content) {
  mBody.classList.remove("editing");
  mTitle.textContent = title; mBody.textContent = content; modal.classList.add("visible");
}
mClose.onclick = closeModal;
modal.onclick = (e) => { if (e.target === modal) closeModal(); };

// --- Fix-it modal: lets the user edit name / keywords / content for
// entries where keyword sanitization stripped or auto-recovered keys. ---
function ft(key) { return (FIX_I18N[LANG] || FIX_I18N.en)[key]; }
let fixQueue = [], fixTotal = 0;
function startFixQueue(uids) {
  fixQueue = uids.slice();
  fixTotal = fixQueue.length;
  if (fixQueue.length) openNextFix();
}
function openNextFix() {
  if (!fixQueue.length) { modal.classList.remove("visible"); mBody.classList.remove("editing"); return; }
  const uid = fixQueue.shift();
  openEntryEditor(uid, fixTotal - fixQueue.length, fixTotal);
}
function openEntryEditor(uid, idx, total) {
  const e = currentSTEntries[uid];
  if (!e) { openNextFix(); return; }
  mBody.classList.add("editing");
  mBody.innerHTML = "";

  const note = document.createElement("div");
  note.className = "edit-note";
  note.innerHTML = ft("fixNote");
  mBody.appendChild(note);

  function field(labelText, inputEl) {
    const wrap = document.createElement("div");
    wrap.className = "edit-field";
    const lbl = document.createElement("label");
    lbl.className = "name-label"; lbl.textContent = labelText;
    wrap.appendChild(lbl); wrap.appendChild(inputEl);
    return wrap;
  }

  const nameInp = document.createElement("input");
  nameInp.className = "name-input"; nameInp.type = "text"; nameInp.value = e.comment || "";
  mBody.appendChild(field(ft("fixNameLabel"), nameInp));

  const kwInp = document.createElement("input");
  kwInp.className = "name-input"; kwInp.type = "text";
  kwInp.value = (e.key || []).join(", ");
  kwInp.placeholder = ft("fixKeywordsPh");
  mBody.appendChild(field(ft("fixKeywordsLabel"), kwInp));

  const contentTa = document.createElement("textarea");
  contentTa.style.minHeight = "220px";
  contentTa.value = e.content || "";
  mBody.appendChild(field(ft("fixContentLabel"), contentTa));

  const countEl = document.createElement("div");
  countEl.className = "edit-count";
  countEl.textContent = ft("fixProgress")(idx, total);
  mBody.appendChild(countEl);

  const btnRow = document.createElement("div");
  btnRow.className = "btn-row";
  const saveBtn = document.createElement("button");
  saveBtn.type = "button"; saveBtn.className = "btn btn-primary";
  saveBtn.textContent = fixQueue.length ? ft("fixSaveNext") : ft("fixSaveClose");
  saveBtn.onclick = () => {
    e.comment = nameInp.value.trim();
    e.key = kwInp.value.split(",").map((s) => s.trim()).filter(Boolean);
    e.content = contentTa.value;
    delete e._flaggedKeys;
    rebuildOutput();
    renderTable(currentSTEntries);
    openNextFix();
  };
  const skipBtn = document.createElement("button");
  skipBtn.type = "button"; skipBtn.className = "btn btn-secondary";
  skipBtn.textContent = ft("fixSkip");
  skipBtn.onclick = () => openNextFix();
  btnRow.appendChild(saveBtn); btnRow.appendChild(skipBtn);
  mBody.appendChild(btnRow);

  mTitle.textContent = ft("fixTitle");
  modal.classList.add("visible");
}

function tagListInner(arr, cls) {
  if (!arr || !arr.length) return '<span style="color:var(--text-muted);font-size:.7rem">—</span>';
  const vis = arr.slice(0, 4), rest = arr.length - vis.length;
  return '<div class="tl">' + vis.map((k) => '<span class="tag ' + cls + '">' + escHtml(k) + "</span>").join("") + (rest > 0 ? '<span class="tag sec">+' + rest + "</span>" : "") + "</div>";
}
function posBadge(pos, role) {
  let key = String(pos);
  if (pos === 6 && role !== null && role !== undefined) key = "6:" + role;
  const cls = [0, 1, 2, 3, 4, 5, 6, 7].includes(pos) ? "p" + pos : "px";
  const lbl = { 0: "0·Before", 1: "1·After", 2: "2·BefEx", 3: "3·AftEx", 4: "4·AN↑", 5: "5·AN↓", "6:0": "6·Sys", "6:1": "6·User", "6:2": "6·Asst", 7: "7·Outlet" }[key] ?? "" + pos;
  return '<span class="pos-badge ' + cls + '">' + lbl + "</span>";
}
function mkTd(html, cls) {
  const td = document.createElement("td");
  if (cls) td.className = cls; td.innerHTML = html; return td;
}

function renderTable(stEntries) {
  currentSTEntries = stEntries;
  const tbody = document.getElementById("tBody");
  tbody.innerHTML = "";
  const entriesArr = Object.entries(stEntries);
  if (entriesArr.length === 0) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center; padding: 24px; color: var(--text-muted); font-style: italic;">No entries generated</td></tr>';
    return;
  }
  const fragment = document.createDocumentFragment();
  entriesArr.forEach(([uid, e], i) => {
    let tb = '<span class="type-badge tb-jai">entry</span>';
    if (e._type === "scenario") tb = '<span class="type-badge tb-scen">scenario</span>';
    else if (e._type === "personality") tb = '<span class="type-badge tb-pers">personality</span>';
    else if (e._type === "merged") tb = '<span class="type-badge tb-merged">merged</span>';

    const tr = document.createElement("tr");
    tr.className = "click-row" + (e._flaggedKeys ? " flagged-row" : "");
    tr.onclick = () => e._flaggedKeys ? openEntryEditor(uid, 1, 1) : showPreview(currentSTEntries[uid].comment || "(unnamed)", e.content);
    tr.appendChild(mkTd(i, "mono"));
    tr.appendChild(mkTd(tb));
    const nameTd = document.createElement("td");
    nameTd.className = "name-cell";
    const nameInp = document.createElement("input");
    nameInp.type = "text"; nameInp.className = "name-inp"; nameInp.value = e.comment || "";
    nameInp.placeholder = "entry name…"; nameInp.title = "Edit entry name";
    nameInp.addEventListener("click", (ev) => ev.stopPropagation());
    nameInp.addEventListener("change", () => { currentSTEntries[uid].comment = nameInp.value; rebuildOutput(); });
    const hint = document.createElement("span");
    hint.className = "tap-hint";
    hint.textContent = e._flaggedKeys ? ft("fixHint") : "Tap row to preview";
    nameTd.appendChild(nameInp); nameTd.appendChild(hint); tr.appendChild(nameTd);

    const kwTd = document.createElement("td");
    kwTd.innerHTML = (e._flaggedKeys ? '<span class="flag-badge" title="' + escHtml(ft("fixHint")) + '">⚠</span>' : "") + tagListInner(e.key, "kw");
    tr.appendChild(kwTd);
    const secTd = document.createElement("td"); secTd.innerHTML = tagListInner(e.keysecondary, "sec"); tr.appendChild(secTd);
    tr.appendChild(mkTd(posBadge(e.position, e.role)));
    tr.appendChild(mkTd(e.order, "mono"));
    tr.appendChild(mkTd(e.delay || 0, "mono"));
    fragment.appendChild(tr);
  });
  tbody.appendChild(fragment);
}

let currentSTEntries = {};
function rebuildOutput() { lastOut = buildExportJson(currentSTEntries); document.getElementById("outputJson").value = lastOut; }
let lastOut = "", pendingEntries = null, pendingFmt = "";
function getOutlet() {
  const v = document.getElementById("outletName").value.trim(), n = document.getElementById("lbName").value.trim();
  if (v) return v;
  if (n) return n.toLowerCase().replace(/[^a-z0-9]/g, "_") + "_knowledge";
  return "lorebook_knowledge";
}
function getFilename() {
  const n = document.getElementById("lbName").value.trim();
  return (n ? n.replace(/[^a-zA-Z0-9 _-]/g, "") : "lorebook") + ".json";
}

document.getElementById("mergePairsCheck").addEventListener("change", (e) => {
  document.querySelectorAll(".cat-sel-pers").forEach((el) => (el.disabled = e.target.checked));
  document.getElementById("persPosHead").style.opacity = e.target.checked ? "0.4" : "1";
  finalizeScript(true);
});
document.getElementById("scenPrefix").addEventListener("input", () => finalizeScript(true));
document.getElementById("persPrefix").addEventListener("input", () => finalizeScript(true));

let lastBadge = null;
function badge(type, key, ...args) {
  lastBadge = { type, key, args };
  const gen = (BADGE_I18N[LANG] || BADGE_I18N.en)[key];
  showBadge(type, gen(...args));
}
let lastLoadedFileName = null;
function bt(key) { return (BADGE_I18N[LANG] || BADGE_I18N.en)[key]; }

function buildExportJson(entriesObj) {
  const exportObj = { entries: {} };
  Object.keys(entriesObj).forEach((k) => {
    exportObj.entries[k] = { ...entriesObj[k] };
    delete exportObj.entries[k]._type;
  });
  return JSON.stringify(exportObj, null, 2);
}

function doConvert(text) {
  lastScriptRepairStats = null;
  lastBadKeyStats = null;
  const det = detectFormat(text);
  if (!det) { badge("error", "cannotDetect"); return; }
  updatePills(det.type);

  if (det.type === "jai_script") {
    let entries;
    try { entries = parseScript(text); } catch (e) { badge("error", "scriptParseFailed", e.message); return; }
    if (!entries.length) { badge("warn", "noEntriesLoreEntries"); return; }
    pendingEntries = entries; pendingFmt = "jai_script";
    buildCatPanel(entries);
    if (MODE === "simple") finalizeScript(false);
    else { badge("ok", "detectedLoreEntries", entries.length); document.getElementById("convertBtn").onclick = () => finalizeScript(false); finalizeScript(true); }
    return;
  }

  if (det.type === "dynamic_lore") {
    let entries;
    try { entries = parseDynamicLore(text); } catch (e) { badge("error", "scriptParseFailed", e.message); return; }
    if (!entries.length) { badge("warn", "noEntriesLoreEntries"); return; }
    pendingEntries = entries; pendingFmt = "dynamic_lore";
    buildCatPanel(entries);
    if (MODE === "simple") finalizeScript(false);
    else { badge("ok", "detectedDynamicLore", entries.length); document.getElementById("convertBtn").onclick = () => finalizeScript(false); finalizeScript(true); }
    return;
  }

  if (det.type === "if_chain") {
    let entries;
    try { entries = parseIfChain(text); } catch (e) { badge("error", "ifChainParseFailed", e.message); return; }
    if (!entries.length) { badge("warn", "noKeywordEntries"); return; }
    pendingEntries = entries; pendingFmt = "if_chain";
    buildCatPanel(entries);
    if (MODE === "simple") finalizeScript(false);
    else { badge("ok", "detectedIfChain", entries.length); document.getElementById("convertBtn").onclick = () => finalizeScript(false); finalizeScript(true); }
    return;
  }

  if (det.type === "st_json") {
    const raw = Object.values(det.parsed.entries);
    if (!raw.length) { badge("warn", "noEntriesFound"); return; }
    lastOut = JSON.stringify(raw.map(stToJai), null, 2);
    showWarnings(buildWarnings(raw, "st"), raw, "st");
    badge("ok", "detectedStWi", raw.length);
    populate(raw.length, raw.length, "ST World Info", "JAI Lorebook JSON", false);
    renderTable(Object.fromEntries(raw.map((e, i) => [i, { ...e, comment: e.comment || "entry " + i, _type: "jai" }])));
    return;
  }

  if (det.type === "jai_json") {
    const raw = det.parsed;
    if (!raw.length) { badge("warn", "noEntriesFound"); return; }
    const obj = { entries: {} };
    raw.forEach((e, i) => { obj.entries[i] = jaiToSt(e, i); });
    lastOut = buildExportJson(obj.entries);
    showWarnings(buildWarnings(raw, "jai"), raw, "jai");
    badge("ok", "detectedJai", raw.length);
    populate(raw.length, Object.keys(obj.entries).length, "JAI Lorebook JSON", "ST World Info", false);
    renderTable(obj.entries);
    const flagged = Object.keys(obj.entries).filter((uid) => obj.entries[uid]._flaggedKeys);
    if (flagged.length) startFixQueue(flagged);
    return;
  }
}

function finalizeScript(isPreview = false) {
  if (!pendingEntries) return;
  const merge = document.getElementById("mergePairsCheck").checked;
  const scenPref = document.getElementById("scenPrefix").value;
  const persPref = document.getElementById("persPrefix").value;
  const { entries, count } = scriptToST(pendingEntries, getCatMap(), getOutlet(), { merge, scenPref, persPref });
  lastOut = buildExportJson(entries);
  if (!isPreview) {
    showWarnings(buildWarnings(pendingEntries, pendingFmt), pendingEntries, pendingFmt);
    badge("ok", "convertedScript", pendingEntries.length, count);
  }
  const sourceName = pendingFmt === "if_chain" ? "if-chain Script" : pendingFmt === "dynamic_lore" ? "Advanced Lore Script" : "loreEntries Script";
  populate(pendingEntries.length, count, sourceName, "ST World Info", isPreview);
  renderTable(entries);
}

function populate(srcCount, stCount, src, tgt, isPreview = false) {
  document.getElementById("rCount").textContent = srcCount;
  document.getElementById("rST").textContent = stCount;
  document.getElementById("rSrc").textContent = src;
  document.getElementById("rTgt").textContent = tgt;
  document.getElementById("outputJson").value = lastOut;
  if (isPreview) {
    document.getElementById("exportSection").classList.remove("visible");
    document.getElementById("previewNote").classList.add("visible");
  } else {
    document.getElementById("exportSection").classList.add("visible");
    document.getElementById("previewNote").classList.remove("visible");
  }
  document.getElementById("resultsSection").classList.add("visible", "fade-in");
}

function showBadge(type, html) {
  document.getElementById("badgeDot").className = "badge-dot dot-" + type;
  document.getElementById("detectionText").innerHTML = html;
  document.getElementById("detectionBadge").classList.add("visible");
}
let lastWarnEntries = null, lastWarnFmt = null;
function showWarnings(warns, entries, fmt) {
  if (entries !== undefined) { lastWarnEntries = entries; lastWarnFmt = fmt; }
  const box = document.getElementById("warningsBox");
  if (warns.length) {
    box.innerHTML = "<strong>" + wt("heading") + "</strong><ul>" + warns.map((w) => "<li>" + w + "</li>").join("") + "</ul>";
    box.classList.add("visible");
  } else box.classList.remove("visible");
}
function updatePills(type) {
  document.querySelectorAll(".fpill").forEach((p) => p.classList.remove("active"));
  const map = { st_json: "pill-st", jai_json: "pill-jai", jai_script: "pill-js", if_chain: "pill-chain" };
  if (map[type]) document.getElementById(map[type]).classList.add("active");
}
function escHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

const pasteEl = document.getElementById("pasteInput");
const dropEl = document.getElementById("dropZone");
const fileEl = document.getElementById("fileInput");

pasteEl.addEventListener("input", () => {
  const v = pasteEl.value.trim();
  if (v) { const d = detectFormat(v); if (d) updatePills(d.type); }
});
pasteEl.addEventListener("keydown", (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
    e.preventDefault(); document.getElementById("convertBtn").click();
  }
});

const PILL_EXAMPLES = {
  "pill-st": '{"entries": {"0": {"key": ["tavern", "inn"], "keysecondary": [], "comment": "The Rusty Tankard", "content": "A dim, smoky tavern near the docks.", "position": 0, "order": 100, "disable": false}}}',
  "pill-jai": '[{"keys": ["tavern", "inn"], "content": "A dim, smoky tavern near the docks.", "enabled": true, "order": 100}]',
  "pill-js": 'const loreEntries = [{ keywords: ["tavern", "inn"], scenario: "A dim, smoky tavern near the docks.", priority: 5, category: "Location" }];',
  "pill-chain": "if (lastMessage.includes('tavern')) { context.character.scenario += \"A dim, smoky tavern near the docks.\"; }",
};
Object.keys(PILL_EXAMPLES).forEach((id) => {
  const el = document.getElementById(id);
  const load = () => { pasteEl.value = PILL_EXAMPLES[id]; pasteEl.focus(); doConvert(PILL_EXAMPLES[id]); };
  el.addEventListener("click", load);
  el.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); load(); } });
});

document.getElementById("lbName").addEventListener("input", () => {
  const outlet = document.getElementById("outletName");
  if (!outlet.dataset.touched) {
    const n = document.getElementById("lbName").value.trim();
    outlet.placeholder = "e.g. " + (n ? n.toLowerCase().replace(/[^a-z0-9]/g, "_") + "_knowledge" : "char_knowledge");
  }
});
document.getElementById("outletName").addEventListener("input", function () { this.dataset.touched = "1"; });

document.getElementById("convertBtn").addEventListener("click", function () {
  const v = pasteEl.value.trim();
  if (!v) { badge("warn", "nothingToConvert"); return; }
  const det = detectFormat(v);
  if (!det || (det.type !== "jai_script" && det.type !== "if_chain")) {
    this.onclick = null; pendingEntries = null; doConvert(v);
  } else if (!pendingEntries) {
    doConvert(v);
  } else {
    finalizeScript(false);
  }
});

document.getElementById("clearBtn").addEventListener("click", () => {
  pasteEl.value = "";
  ["detectionBadge", "warningsBox", "catPanel", "resultsSection"].forEach((id) => document.getElementById(id).classList.remove("visible", "fade-in"));
  document.getElementById("exportSection").classList.remove("visible");
  document.getElementById("previewNote").classList.remove("visible");
  document.querySelectorAll(".fpill").forEach((p) => p.classList.remove("active"));
  document.getElementById("convertBtn").onclick = null;
  document.getElementById("dropLoaded").textContent = "";
  pendingEntries = null; lastOut = ""; lastBadge = null; lastWarnEntries = null; lastWarnFmt = null; lastLoadedFileName = null;
});

dropEl.addEventListener("click", () => fileEl.click());
dropEl.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") fileEl.click(); });
fileEl.addEventListener("change", () => { if (fileEl.files[0]) readFile(fileEl.files[0]); });
dropEl.addEventListener("dragover", (e) => { e.preventDefault(); dropEl.classList.add("drag-over"); });
dropEl.addEventListener("dragleave", () => dropEl.classList.remove("drag-over"));
dropEl.addEventListener("drop", (e) => {
  e.preventDefault(); dropEl.classList.remove("drag-over");
  if (e.dataTransfer.files[0]) readFile(e.dataTransfer.files[0]);
});

function readFile(f) {
  lastLoadedFileName = f.name;
  document.getElementById("dropLoaded").textContent = bt("fileLoaded")(f.name);
  if (f.name.toLowerCase().endsWith(".docx")) {
    const r = new FileReader();
    r.onload = async (ev) => {
      try {
        badge("warn", "extractingDocx");
        const txt = await extractDocxText(ev.target.result);
        pasteEl.value = txt; doConvert(txt);
      } catch (e) { badge("error", "docxReadFailed", e.message); }
    };
    r.readAsArrayBuffer(f);
  } else {
    const r = new FileReader();
    r.onload = (ev) => { pasteEl.value = ev.target.result; doConvert(ev.target.result); };
    r.readAsText(f);
  }
}

document.getElementById("copyBtn").addEventListener("click", () => {
  if (!lastOut) return;
  const outTextarea = document.getElementById("outputJson");
  outTextarea.select();
  navigator.clipboard.writeText(lastOut).then(() => {
    const b = document.getElementById("copyBtn");
    b.textContent = "copied!";
    setTimeout(() => { b.textContent = "copy"; }, 1800);
  });
});
document.getElementById("downloadBtn").addEventListener("click", () => {
  if (!lastOut) return;
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([lastOut], { type: "application/json" }));
  a.download = getFilename();
  a.click();
});

function applyLang(lang) {
  LANG = lang;
  document.querySelectorAll("[data-i18n]").forEach((el) => {
    const key = el.getAttribute("data-i18n");
    if (I18N[lang][key] !== undefined) el.textContent = I18N[lang][key];
  });
  document.querySelectorAll("[data-i18n-ph]").forEach((el) => {
    const key = el.getAttribute("data-i18n-ph");
    if (I18N[lang][key] !== undefined) el.setAttribute("placeholder", I18N[lang][key]);
  });
  document.querySelectorAll("[data-i18n-title]").forEach((el) => {
    const key = el.getAttribute("data-i18n-title");
    if (I18N[lang][key] !== undefined) el.setAttribute("title", I18N[lang][key]);
  });
  document.getElementById("langToggle").textContent = lang === "en" ? "RU" : "EN";
  document.documentElement.lang = lang === "ru" ? "ru" : "en";
  const themeBtn = document.getElementById("themeToggle");
  if (themeBtn) {
    const isLight = document.documentElement.classList.contains("light");
    themeBtn.textContent = isLight ? I18N[lang].themeLight : I18N[lang].themeDark;
  }
  if (typeof pendingEntries !== "undefined" && pendingEntries) buildCatPanel(pendingEntries);
  if (lastWarnEntries) showWarnings(buildWarnings(lastWarnEntries, lastWarnFmt), lastWarnEntries, lastWarnFmt);
  if (lastBadge) {
    const gen = (BADGE_I18N[lang] || BADGE_I18N.en)[lastBadge.key];
    showBadge(lastBadge.type, gen(...lastBadge.args));
  }
  if (lastLoadedFileName) {
    const el = document.getElementById("dropLoaded");
    if (el) el.textContent = (BADGE_I18N[lang] || BADGE_I18N.en).fileLoaded(lastLoadedFileName);
  }
  refreshModeUI();
}

let MODE = "simple";
function refreshModeUI() {
  const dict = I18N[LANG] || I18N.en;
  document.getElementById("modeSimpleBtn").textContent = dict.modeSimple;
  document.getElementById("modeAdvancedBtn").textContent = dict.modeAdvanced;
  document.getElementById("modeHint").textContent = MODE === "simple" ? dict.modeHintSimple : dict.modeHintAdvanced;
}
function applyMode(mode) {
  MODE = mode;
  document.body.classList.toggle("mode-simple", mode === "simple");
  document.getElementById("modeSimpleBtn").classList.toggle("active", mode === "simple");
  document.getElementById("modeAdvancedBtn").classList.toggle("active", mode === "advanced");
  document.getElementById("modeSimpleBtn").setAttribute("aria-selected", mode === "simple");
  document.getElementById("modeAdvancedBtn").setAttribute("aria-selected", mode === "advanced");
  refreshModeUI();
  if (pendingEntries && (pendingFmt === "jai_script" || pendingFmt === "if_chain")) {
    buildCatPanel(pendingEntries);
    if (mode === "simple") finalizeScript(false);
    else finalizeScript(true);
  }
}
document.getElementById("modeSimpleBtn").addEventListener("click", () => applyMode("simple"));
document.getElementById("modeAdvancedBtn").addEventListener("click", () => applyMode("advanced"));
applyMode("simple");

document.getElementById("langToggle").addEventListener("click", () => { applyLang(LANG === "en" ? "ru" : "en"); });
applyLang("en");

(function () {
  const btn = document.getElementById("themeToggle");
  const root = document.documentElement;
  let dark = true;
  if (window.matchMedia && window.matchMedia("(prefers-color-scheme: light)").matches) {
    dark = false; root.classList.add("light"); btn.textContent = I18N[LANG].themeLight;
  }
  btn.addEventListener("click", () => {
    dark = !dark;
    root.classList.add("no-transition");
    root.classList.toggle("light", !dark);
    btn.textContent = dark ? I18N[LANG].themeDark : I18N[LANG].themeLight;
    requestAnimationFrame(() => requestAnimationFrame(() => root.classList.remove("no-transition")));
  });
})();