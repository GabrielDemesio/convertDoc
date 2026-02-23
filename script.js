const $file = document.getElementById("file");
const $run = document.getElementById("run");
const $log = document.getElementById("log");
const $download = document.getElementById("download");
const $downloadXml = document.getElementById("downloadXml");
const $downloadAll = document.getElementById("downloadAll");
const $renamed = document.getElementById("renamed");
const $queue = document.getElementById("queue");
const $bar = document.getElementById("bar");
let renamedObjectUrls = [];
let renamedDownloadsState = [];
let csvObjectUrl = null;
let xmlObjectUrl = null;
const FILE_CONCURRENCY = 2;
let queueItemsState = [];
let queueProgressMode = false;

function log(msg) { $log.textContent += msg + "\n"; }
function clearLog() { $log.textContent = ""; }
function setProgress(p) { $bar.style.width = `${Math.max(0, Math.min(100, p))}%`; }
function setOcrProgress(p) {
    if (queueProgressMode) return;
    setProgress(p);
}

function sanitizeFilePart(v) {
    return String(v || "")
        .replace(/[<>:"/\\|?*\x00-\x1F]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

function normalizeText(t) {
    return (t || "")
        .replace(/\r/g, "")
        .replace(/[ \t]+/g, " ")
        .replace(/[|•·]/g, " ")
        .replace(/(\d)[ \t]+(\d)/g, "$1$2")
        .replace(/\n{3,}/g, "\n\n")
        .trim();
}

function pick(re, text) {
    const m = text.match(re);
    return m ? (m[1] || "").trim() : "";
}

function cleanField(v) {
    return String(v || "")
        .replace(/\s+/g, " ")
        .replace(/^[\s:\-–]+/, "")
        .trim();
}

function normalizeDigits(s) {
    const n = String(s || "").replace(/\D+/g, "");
    if (!n) return "";
    return n.replace(/^0+/, "") || "0";
}

function extractNfFromDocOriginalText(text) {
    const t = String(text || "");
    const m1 = t.match(/\bNF[-\s]?E\b[\s\S]{0,120}?(\d{1,3})\s*[\/\\]\s*(\d{4,9})\b/i);
    if (m1?.[2]) return normalizeDigits(m1[2]);

    const m2 = t.match(/\bNF\b\s*[:\-]?\s*([0-9]{4,9})\s+\bVOL\b/i);
    if (m2?.[1]) return normalizeDigits(m2[1]);

    return "";
}

function extractNfFromSerieDocWords(words, pageWidth, pageHeight) {
    if (!words.length || !pageWidth || !pageHeight) return "";
    const cands = [];

    for (const w of words) {
        const raw = normalizeOcrNumber(String(w.text || ""));
        const m = raw.match(/\b(\d{3})\s*[\/\\]\s*(\d{4,9})\b/);
        if (!m?.[2]) continue;
        // Faixa da tabela "DOCUMENTOS ORIGINÁRIOS" (evita canhoto e datas)
        if (w.yc < pageHeight * 0.42 || w.yc > pageHeight * 0.80) continue;
        if (w.xc < pageWidth * 0.05 || w.xc > pageWidth * 0.78) continue;
        cands.push({ nf: normalizeDigits(m[2]), y: w.yc, x: w.x0 });
    }

    if (!cands.length) return "";
    cands.sort((a, b) => a.y - b.y || a.x - b.x);
    return cands[0].nf;
}

function extractNfFromNfBoxText(text) {
    const t = String(text || "");
    const m =
        t.match(/\bNF[-\s]?E\b[\s\S]{0,100}?\bN[º°O]?\s*([0-9][0-9.\s]{3,10})\s*\bS[ÉE]RIE\b/i) ||
        t.match(/\bN[º°O]\s*([0-9]{1,3}[.\s][0-9]{3,6})\b[\s\S]{0,40}\bS[ÉE]RIE\b/i);
    return m?.[1] ? normalizeDigits(m[1]) : "";
}

function isOnlyConfusableDigitDiff(a, b) {
    if (!a || !b || a.length !== b.length) return false;
    const conf = new Set(["0", "6", "8"]);
    for (let i = 0; i < a.length; i++) {
        if (a[i] === b[i]) continue;
        if (!conf.has(a[i]) || !conf.has(b[i])) return false;
    }
    return true;
}

function pickConsistentNf(cands) {
    const pool = cands.map(normalizeDigits).filter(Boolean);
    if (!pool.length) return "";
    if (pool.length === 1) return pool[0];

    const freq = new Map();
    for (const n of pool) freq.set(n, (freq.get(n) || 0) + 1);
    let max = 0;
    for (const v of freq.values()) if (v > max) max = v;
    const top = [...freq.keys()].filter(k => freq.get(k) === max);
    if (top.length === 1) return top[0];

    const sameLen = top.every(v => v.length === top[0].length);
    if (sameLen) {
        let confusable = false;
        for (let i = 0; i < top.length; i++) {
            for (let j = i + 1; j < top.length; j++) {
                if (isOnlyConfusableDigitDiff(top[i], top[j])) {
                    confusable = true;
                    break;
                }
            }
            if (confusable) break;
        }
        if (confusable) {
            top.sort((a, b) => {
                const zerosA = (a.match(/0/g) || []).length;
                const zerosB = (b.match(/0/g) || []).length;
                if (zerosA !== zerosB) return zerosB - zerosA;
                return Number(a) - Number(b);
            });
            return top[0];
        }
    }

    top.sort((a, b) => a.length - b.length || Number(a) - Number(b));
    return top[0];
}

function extractBestDigits(text, minLen = 4, maxLen = 9) {
    const raw = String(text || "");
    const candidates = raw.match(/\d[\d.\s,-]{2,}\d/g) || [];
    let best = "";
    for (const c of candidates) {
        const d = normalizeDigits(c);
        if (d.length < minLen || d.length > maxLen) continue;
        if (d.length > best.length) best = d;
    }
    if (best) return best;
    // fallback: any digit group
    const simple = raw.match(/\d{4,9}/g) || [];
    for (const s of simple) {
        if (s.length > best.length) best = s;
    }
    return best;
}

function pickNumberAfterLabel(text, labels, minDigits = 4) {
    const labelAlt = labels.map(l => l.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|");
    const re = new RegExp(
        `(?:^|\\b)\\s*(?:${labelAlt})\\b\\s*[:\\-–#°º\\.]?\\s*([\\d\\s.,]{${minDigits},})`,
        "i"
    );
    const m = text.match(re);
    if (!m) return "";
    return normalizeDigits(m[1]);
}

function trimAtLabel(value, labelRe) {
    if (!value) return "";
    const idx = value.search(labelRe);
    if (idx > 0) return cleanField(value.slice(0, idx));
    return cleanField(value);
}

function findBlock(text, labels, stopLabels) {
    const labelAlt = labels.map(l => l.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|");
    const stopAlt = stopLabels.map(l => l.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|");
    const re = new RegExp(
        `(?:^|\\n|\\b)\\s*(?:${labelAlt})\\b\\s*[:\\-–]?\\s*([\\s\\S]*?)(?=\\n\\s*(?:${stopAlt})\\b|$)`,
        "i"
    );
    const m = text.match(re);
    return m ? cleanField(m[1]) : "";
}

function pickClosestMatch(text, re, anchorIdx) {
    let best = "";
    let bestDist = Infinity;
    let m;
    while ((m = re.exec(text)) !== null) {
        const idx = m.index || 0;
        const dist = Math.abs(idx - anchorIdx);
        if (dist < bestDist) {
            bestDist = dist;
            best = m[1] || "";
        }
    }
    return cleanField(best);
}

function extractCepFromBlock(block) {
    if (!block) return "";
    const m = block.match(/\b(\d{5}-\d{3})\b/);
    return m ? m[1] : "";
}

function formatCep(value) {
    const s = String(value || "").trim();
    if (!s) return "";
    if (/\d{5}-\d{3}/.test(s)) return s;
    const digits = s.replace(/\D+/g, "");
    if (digits.length === 8) return digits.slice(0, 5) + "-" + digits.slice(5);
    return s;
}

function normWord(s) {
    return String(s || "")
        .toUpperCase()
        .replace(/[^A-Z0-9]/g, "");
}

function isNumberLike(s) {
    return /[0-9]/.test(s);
}

function normalizeOcrNumber(s) {
    return String(s || "")
        .replace(/[OoQqÐ]/g, "0")
        .replace(/[Il|]/g, "1")
        .replace(/[Zz]/g, "2")
        .replace(/[Ss]/g, "5")
        .replace(/[Bb]/g, "8")
        .replace(/[Gg]/g, "6");
}

function extractNumberFromWord(s) {
    const n = normalizeOcrNumber(s).replace(/\D+/g, "");
    return n;
}

function findLabelAnchor(words, labelTokens) {
    const tokens = labelTokens.map(normWord).filter(Boolean);
    if (!tokens.length) return null;
    for (let i = 0; i < words.length; i++) {
        const w = words[i];
        if (normWord(w.text) !== tokens[0]) continue;
        if (tokens.length === 1) return w;
        let ok = true;
        let last = w;
        for (let j = 1; j < tokens.length; j++) {
            const nxt = words[i + j];
            if (!nxt) { ok = false; break; }
            const sameLine = Math.abs(nxt.yc - w.yc) <= Math.max(6, w.h * 0.8);
            if (!sameLine || normWord(nxt.text) !== tokens[j]) { ok = false; break; }
            last = nxt;
        }
        if (ok) return last;
    }
    return null;
}

function findLabelSpan(words, labelTokens) {
    const tokens = labelTokens.map(normWord).filter(Boolean);
    if (!tokens.length) return null;
    for (let i = 0; i < words.length; i++) {
        const w = words[i];
        if (normWord(w.text) !== tokens[0]) continue;
        let ok = true;
        let last = w;
        let x0 = w.x0, x1 = w.x1, y0 = w.y0, y1 = w.y1;
        if (tokens.length > 1) {
            for (let j = 1; j < tokens.length; j++) {
                const nxt = words[i + j];
                if (!nxt) { ok = false; break; }
                const sameLine = Math.abs(nxt.yc - w.yc) <= Math.max(6, w.h * 0.8);
                if (!sameLine || normWord(nxt.text) !== tokens[j]) { ok = false; break; }
                last = nxt;
                x0 = Math.min(x0, nxt.x0);
                x1 = Math.max(x1, nxt.x1);
                y0 = Math.min(y0, nxt.y0);
                y1 = Math.max(y1, nxt.y1);
            }
        }
        if (ok) {
            return {
                x0, x1, y0, y1,
                xc: (x0 + x1) / 2,
                yc: (y0 + y1) / 2,
                h: y1 - y0
            };
        }
    }
    return null;
}

function findAnchorsByToken(words, tokenNorm) {
    const out = [];
    for (const w of words) {
        if (normWord(w.text) === tokenNorm) out.push(w);
    }
    return out;
}

function median(values) {
    if (!values.length) return 0;
    const a = [...values].sort((x, y) => x - y);
    const mid = Math.floor(a.length / 2);
    return a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}

function groupWordsIntoLines(words) {
    if (!words.length) return { lines: [], lineGap: 0 };
    const heights = words.map(w => w.h).filter(Boolean);
    const medH = median(heights) || 10;
    const lineGap = Math.max(6, medH * 0.8);
    const sorted = [...words].sort((a, b) => a.yc - b.yc);
    const lines = [];
    for (const w of sorted) {
        let line = lines.find(l => Math.abs(l.yc - w.yc) <= lineGap);
        if (!line) {
            line = { yc: w.yc, words: [] };
            lines.push(line);
        }
        line.words.push(w);
        line.yc = (line.yc * (line.words.length - 1) + w.yc) / line.words.length;
    }
    for (const l of lines) {
        l.words.sort((a, b) => a.x0 - b.x0);
        const hs = l.words.map(w => w.h).filter(Boolean);
        l.h = median(hs) || 10;
    }
    lines.sort((a, b) => a.yc - b.yc);
    return { lines, lineGap };
}

function lineText(line, xMin, xMax, labelSet) {
    const parts = [];
    for (const w of line.words) {
        if (w.x1 < xMin || w.x0 > xMax) continue;
        const nw = normWord(w.text);
        if (labelSet && labelSet.has(nw)) continue;
        parts.push(w.text);
    }
    return cleanField(parts.join(" "));
}

function lineTextUntilStop(line, startX, endX, stopSet, labelSet) {
    const parts = [];
    for (const w of line.words) {
        if (w.x1 < startX || w.x0 > endX) continue;
        const nw = normWord(w.text);
        if (labelSet && labelSet.has(nw)) continue;
        if (stopSet && stopSet.has(nw)) break;
        parts.push(w.text);
    }
    return cleanField(parts.join(" "));
}

function collectNumberInLineRange(line, xMin, xMax, opts = {}) {
    const maxGap = opts.maxGap ?? Math.max(12, (line.h || 10) * 1.6);
    const allowNonDigit = opts.allowNonDigit ?? 2;
    const maxDigits = opts.maxDigits ?? 12;
    const maxTokens = opts.maxTokens ?? 4;
    let digits = "";
    let started = false;
    let lastX1 = null;
    let nonDigitLeft = allowNonDigit;
    let tokens = 0;
    for (const w of line.words) {
        if (w.x1 < xMin || w.x0 > xMax) continue;
        if (started && lastX1 != null && w.x0 - lastX1 > maxGap) break;
        const d = extractNumberFromWord(w.text);
        if (d) {
            started = true;
            digits += d;
            tokens++;
            lastX1 = w.x1;
            nonDigitLeft = allowNonDigit;
            if (digits.length >= maxDigits || tokens >= maxTokens) break;
            continue;
        }
        if (started) {
            const wWidth = Math.max(0, w.x1 - w.x0);
            if (nonDigitLeft > 0 && wWidth <= maxGap) {
                nonDigitLeft--;
                lastX1 = w.x1;
                continue;
            }
            break;
        }
    }
    return digits;
}

function collectNumberTokensInRange(line, xMin, xMax) {
    const out = [];
    for (const w of line.words) {
        if (w.x1 < xMin || w.x0 > xMax) continue;
        const d = extractNumberFromWord(w.text);
        if (!d) continue;
        out.push({ digits: d, x0: w.x0, x1: w.x1, yc: w.yc });
    }
    return out.sort((a, b) => a.x0 - b.x0);
}

function joinNumberTokens(tokens, maxGap) {
    const seqs = [];
    let cur = null;
    for (const t of tokens) {
        if (!cur) {
            cur = { digits: t.digits, x1: t.x1 };
            continue;
        }
        if (t.x0 - cur.x1 <= maxGap) {
            cur.digits += t.digits;
            cur.x1 = t.x1;
        } else {
            seqs.push(cur.digits);
            cur = { digits: t.digits, x1: t.x1 };
        }
    }
    if (cur) seqs.push(cur.digits);
    return seqs;
}

function pickLongest(nums, minLen = 1) {
    let best = "";
    for (const n of nums) {
        if (n.length < minLen) continue;
        if (n.length > best.length) best = n;
    }
    return best;
}

function findAllLabelAnchors(words, labelTokens) {
    const tokens = labelTokens.map(normWord).filter(Boolean);
    if (!tokens.length) return [];
    const anchors = [];
    for (let i = 0; i < words.length; i++) {
        const w = words[i];
        if (normWord(w.text) !== tokens[0]) continue;
        if (tokens.length === 1) {
            anchors.push(w);
            continue;
        }
        let ok = true;
        let last = w;
        for (let j = 1; j < tokens.length; j++) {
            const nxt = words[i + j];
            if (!nxt) { ok = false; break; }
            const sameLine = Math.abs(nxt.yc - w.yc) <= Math.max(6, w.h * 0.8);
            if (!sameLine || normWord(nxt.text) !== tokens[j]) { ok = false; break; }
            last = nxt;
        }
        if (ok) anchors.push(last);
    }
    return anchors;
}

function pickBestCandidate(cands) {
    if (!cands.length) return "";
    const freq = new Map();
    for (const c of cands) freq.set(c, (freq.get(c) || 0) + 1);
    let best = cands[0];
    let bestCount = 0;
    for (const [k, v] of freq.entries()) {
        if (v > bestCount) { bestCount = v; best = k; }
    }
    return best;
}

function pickBestCandidateNumber(cands, minLen = 1) {
    const nums = cands
        .map(c => String(c || "").replace(/\D+/g, ""))
        .filter(Boolean);
    if (!nums.length) return "";
    let pool = nums.filter(n => n.length >= minLen);
    if (!pool.length) pool = nums;

    // prefer most frequent across candidates
    const freq = new Map();
    for (const n of pool) freq.set(n, (freq.get(n) || 0) + 1);
    let bestCount = 0;
    for (const [k, v] of freq.entries()) {
        if (v > bestCount) bestCount = v;
    }
    const top = [...freq.keys()].filter(n => freq.get(n) === bestCount);
    if (top.length === 1) return top[0];

    // OCR comum: um "1" grudado no início (ex: 120060 em vez de 20060)
    for (const n of top) {
        if (!n.startsWith("1") || n.length <= minLen) continue;
        const tail = n.slice(1);
        if (!tail || tail.length < minLen) continue;
        if ((freq.get(tail) || 0) >= (freq.get(n) || 0)) return tail;
    }

    // empate: prefere o tamanho mais curto para reduzir ruído de dígito extra
    top.sort((a, b) => a.length - b.length);
    return top[0];
}

function stripAfterMarkers(text, markers) {
    const s = String(text || "");
    const up = s.toUpperCase();
    let idx = -1;
    for (const m of markers) {
        const i = up.indexOf(m);
        if (i >= 0) idx = idx < 0 ? i : Math.min(idx, i);
    }
    return idx >= 0 ? cleanField(s.slice(0, idx)) : cleanField(s);
}

function findLineIndexByY(lines, y, lineGap) {
    let best = -1;
    let bestDist = Infinity;
    for (let i = 0; i < lines.length; i++) {
        const d = Math.abs(lines[i].yc - y);
        if (d < bestDist) {
            bestDist = d;
            best = i;
        }
    }
    if (bestDist <= (lineGap || 8) * 1.5) return best;
    return best;
}

function findNumberNearAnchor(words, anchor, opts = {}) {
    if (!anchor) return "";
    const yTol = opts.yTol ?? Math.max(10, anchor.h * 1.2);
    const minDigits = opts.minDigits ?? 1;
    const maxDigits = opts.maxDigits ?? 20;
    let best = null;
    let bestScore = Infinity;
    for (const w of words) {
        if (!isNumberLike(w.text)) continue;
        const dx = w.x0 - anchor.x1;
        const dy = Math.abs(w.yc - anchor.yc);
        if (dx < -5) continue;
        if (dy > yTol) continue;
        const digits = extractNumberFromWord(w.text);
        if (digits.length < minDigits || digits.length > maxDigits) continue;
        const score = dx + dy * 2.5;
        if (score < bestScore) {
            bestScore = score;
            best = w;
        }
    }
    return best ? extractNumberFromWord(best.text) : "";
}

function findNumberBelowAnchor(words, anchor, opts = {}) {
    if (!anchor) return "";
    const pageWidth = opts.pageWidth ?? Math.max(...words.map(w => w.x1));
    const pageHeight = opts.pageHeight ?? Math.max(...words.map(w => w.y1));
    const dxTol = opts.dxTol ?? Math.max(120, pageWidth * 0.06);
    const minDigits = opts.minDigits ?? 1;
    const maxDigits = opts.maxDigits ?? 20;
    const yMin = anchor.y1 + (opts.yMinGap ?? 2);
    const yMax = anchor.yc + (opts.yMaxGap ?? pageHeight * 0.18);

    let best = null;
    let bestScore = Infinity;
    for (const w of words) {
        if (!isNumberLike(w.text)) continue;
        if (w.yc < yMin || w.yc > yMax) continue;
        const dx = Math.abs(w.xc - anchor.xc);
        if (dx > dxTol) continue;
        const digits = extractNumberFromWord(w.text);
        if (digits.length < minDigits || digits.length > maxDigits) continue;
        const score = (w.yc - anchor.yc) + dx * 0.25;
        if (score < bestScore) {
            bestScore = score;
            best = w;
        }
    }
    return best ? extractNumberFromWord(best.text) : "";
}

function findNumberInColumnBelow(words, anchorSpan, opts = {}) {
    if (!anchorSpan) return "";
    const pageHeight = opts.pageHeight ?? Math.max(...words.map(w => w.y1));
    const minDigits = opts.minDigits ?? 1;
    const maxDigits = opts.maxDigits ?? 12;
    const preferDecimal = opts.preferDecimal ?? false;
    const preferNonZero = opts.preferNonZero ?? false;
    const xPad = opts.xPad ?? Math.max(12, anchorSpan.h * 1.5);
    const yMin = anchorSpan.y1 + (opts.yMinGap ?? 2);
    const yMax = anchorSpan.y1 + (opts.yMaxGap ?? pageHeight * 0.2);

    let best = null;
    let bestScore = Infinity;
    for (const w of words) {
        if (w.yc < yMin || w.yc > yMax) continue;
        if (w.xc < anchorSpan.x0 - xPad || w.xc > anchorSpan.x1 + xPad) continue;
        const raw = String(w.text || "");
        const digits = extractNumberFromWord(raw);
        if (digits.length < minDigits || digits.length > maxDigits) continue;
        const hasDecimal = /[.,]/.test(raw);
        if (preferDecimal && !hasDecimal) continue;
        if (preferNonZero && Number(digits) === 0) continue;
        const score = (w.yc - anchorSpan.y1) + Math.abs(w.xc - anchorSpan.xc) * 0.2;
        if (score < bestScore) {
            bestScore = score;
            best = raw;
        }
    }
    return best ? normalizeOcrNumber(best) : "";
}

function extractNfFromFixedPosition(words, pageWidth, pageHeight) {
    if (!words.length || !pageWidth || !pageHeight) return "";

    // Região fixa do canhoto inferior onde costuma aparecer "NF: 20060"
    const xMin = pageWidth * 0.30;
    const xMax = pageWidth * 0.86;
    const yMin = pageHeight * 0.86;
    const yMax = pageHeight * 0.98;
    const region = words
        .filter(w =>
            w.x0 >= xMin &&
            w.x1 <= xMax &&
            w.yc >= yMin &&
            w.yc <= yMax
        )
        .sort((a, b) => (a.yc - b.yc) || (a.x0 - b.x0));
    if (!region.length) return "";

    // Prioridade: linha inferior com VOL (padrão "NF: 20060 VOL: 1")
    const { lines: regionLines } = groupWordsIntoLines(region);
    const linesDesc = [...regionLines].sort((a, b) => b.yc - a.yc);
    for (const line of linesDesc) {
        const volWord = line.words.find(w => normWord(w.text).startsWith("VOL"));
        if (!volWord) continue;
        const leftNums = line.words
            .filter(w => w.x1 <= volWord.x0 + 2)
            .map(w => ({
                digits: extractNumberFromWord(w.text),
                x1: w.x1
            }))
            .filter(c => c.digits.length >= 4 && c.digits.length <= 9)
            .sort((a, b) => b.x1 - a.x1);
        if (leftNums.length) return leftNums[0].digits;
    }

    // Caso OCR junte em um único token: "NF:20060"
    for (const w of region) {
        const raw = normalizeOcrNumber(String(w.text || ""));
        const m = raw.match(/NF\s*[:\-]?\s*([0-9]{4,9})/i);
        if (m?.[1]) return normalizeDigits(m[1]);
    }

    // Procura âncora "NF" mais baixa da região e pega o número à direita na mesma linha
    const nfAnchors = region
        .filter(w => {
            const n = normWord(w.text);
            return n === "NF" || n.startsWith("NF");
        })
        .sort((a, b) => b.yc - a.yc);

    for (const a of nfAnchors) {
        const yTol = Math.max(6, a.h * 0.9);
        const sameLineRight = region
            .filter(w =>
                w.x0 >= a.x1 - 2 &&
                Math.abs(w.yc - a.yc) <= yTol
            )
            .map(w => ({
                digits: extractNumberFromWord(w.text),
                dx: Math.max(0, w.x0 - a.x1)
            }))
            .filter(c => c.digits.length >= 4 && c.digits.length <= 9)
            .sort((u, v) => u.dx - v.dx);

        if (sameLineRight.length) return sameLineRight[0].digits;
    }

    // Fallback por texto da região
    const regionText = region.map(w => w.text).join(" ");
    const m = normalizeOcrNumber(regionText).match(/\bNF\b\s*[:\-]?\s*([0-9]{4,9})/i);
    return m?.[1] ? normalizeDigits(m[1]) : "";
}

function extractNameByLabel(words, lines, lineGap, labelTokens, columnRange, labelSet, stopSet) {
    const anchor = findLabelAnchor(words, labelTokens);
    if (!anchor) return "";
    const li = findLineIndexByY(lines, anchor.yc, lineGap);
    if (li < 0) return "";
    const line = lines[li];
    // tenta na mesma linha, à direita do rótulo
    let text = lineTextUntilStop(line, anchor.x1 + 3, columnRange.xMax, stopSet, labelSet);
    if (text) return text;
    // se não achou, pega a próxima linha do mesmo bloco/coluna
    for (let i = li + 1; i < lines.length; i++) {
        const l = lines[i];
        if (l.yc - line.yc > lineGap * 6) break;
        const lt = lineTextUntilStop(l, columnRange.xMin, columnRange.xMax, stopSet, labelSet);
        if (!lt) continue;
        if ([...labelSet].some(lbl => lt.toUpperCase().includes(lbl))) break;
        return lt;
    }
    return "";
}

function mapOcrWords(ocrData) {
    if (!Array.isArray(ocrData?.words)) return [];
    return ocrData.words
        .filter(w => w?.bbox)
        .map(w => ({
            text: w.text || "",
            x0: Number(w.bbox.x0) || 0,
            y0: Number(w.bbox.y0) || 0,
            x1: Number(w.bbox.x1) || 0,
            y1: Number(w.bbox.y1) || 0,
            xc: ((Number(w.bbox.x0) || 0) + (Number(w.bbox.x1) || 0)) / 2,
            yc: ((Number(w.bbox.y0) || 0) + (Number(w.bbox.y1) || 0)) / 2,
            h: (Number(w.bbox.y1) || 0) - (Number(w.bbox.y0) || 0)
        }));
}

function extractFields(ocrText, ocrData = null) {
    const t = normalizeText(ocrText);
    const nfFromDocText = extractNfFromDocOriginalText(t);
    const nfFromBoxText = extractNfFromNfBoxText(t);
    const words = mapOcrWords(ocrData);

    const { lines, lineGap } = groupWordsIntoLines(words);
    const maxX = words.length ? Math.max(...words.map(w => w.x1)) : 0;
    const maxY = words.length ? Math.max(...words.map(w => w.y1)) : 0;
    const nfFromSerieDocWords = extractNfFromSerieDocWords(words, maxX, maxY);
    const labelSet = new Set([
        "REM","REMETENTE","DEST","DESTINATARIO","DESTINATÁRIO","CONSIGNATARIO","CONSIGNATÁRIO",
        "CEP","VOLUMES","VOLUME","VOL","PESO","REAL","CTE","CT","NF","NFE","EXPEDIDOR","RECEBEDOR","TOMADOR"
    ].map(normWord));
    const stopSet = new Set([
        "FONE","TELEFONE","ENDERECO","ENDEREÇO","CNPJ","CPF","IE","INSCR","INSCRICAO","INSCRIÇÃO",
        "CEP","CIDADE","UF","PAIS","PAÍS","RUA","AV","AVENIDA","ROD","BR","KM","BAIRRO","CENTRO","NUMERO","Nº","NO","N."
    ].map(normWord));

    const halfX = maxX ? maxX * 0.52 : 0;
    const leftCol = { xMin: 0, xMax: halfX || 10000 };
    const rightCol = { xMin: halfX || 0, xMax: maxX || 10000 };

    let cteByWords = "";
    if (words.length) {
        const anchors = findAllLabelAnchors(words, ["CTE"]);
        const cands = [];
        for (const a of anchors) {
            const li = findLineIndexByY(lines, a.yc, lineGap);
            if (li >= 0) {
                const sameLine = collectNumberInLineRange(lines[li], a.x1 + 2, maxX);
                if (sameLine.length >= 6 && sameLine.length <= 7) cands.push(sameLine);
                // tenta nas linhas abaixo, mesma coluna
                for (let k = li + 1; k < Math.min(lines.length, li + 4); k++) {
                    if (lines[k].yc - lines[li].yc > lineGap * 6) break;
                    const below = collectNumberInLineRange(lines[k], a.x0 - 10, a.x1 + 60);
                    if (below.length >= 6 && below.length <= 7) { cands.push(below); break; }
                }
            }
            const near = findNumberNearAnchor(words, a, { minDigits: 6, maxDigits: 7 });
            if (near) cands.push(near);
            const below2 = findNumberBelowAnchor(words, a, { minDigits: 6, maxDigits: 7, pageWidth: maxX, pageHeight: maxY });
            if (below2) cands.push(below2);
        }
        cteByWords = pickBestCandidate(cands);
    }
    const cte =
        cteByWords ||
        pickNumberAfterLabel(t, ["CT-E", "CTE"], 4) ||
        pick(/\bCT[-\s]?E\b\s*[:\-]?\s*(\d{4,})/i, t) ||
        pick(/\bCTE\b\s*[:\-]?\s*(\d{4,})/i, t) ||
        pick(/\[(\d{6})\]/, t);

    let nfByWords = words.length ? extractNfFromFixedPosition(words, maxX, maxY) : "";
    if (words.length && !nfByWords) {
        const nfSpan = findLabelSpan(words, ["NFE"]) || findLabelSpan(words, ["NF"]);
        if (nfSpan) {
            const li = findLineIndexByY(lines, nfSpan.yc, lineGap);
            if (li >= 0) {
                const line = lines[li];
                const tokens = collectNumberTokensInRange(line, nfSpan.x1 + 2, nfSpan.x1 + 260);
                const seqs = joinNumberTokens(tokens, Math.max(14, (line.h || 10) * 2.2));
                const bestSeq = pickLongest(seqs, 4);
                if (bestSeq) nfByWords = bestSeq;
                if (!nfByWords) {
                    const sameLine = collectNumberInLineRange(
                        line,
                        nfSpan.x1 + 2,
                        nfSpan.x1 + 200,
                        { allowNonDigit: 2, maxDigits: 9, maxTokens: 5 }
                    );
                    if (sameLine.length >= 4) nfByWords = sameLine;
                }
                if (!nfByWords) {
                    for (let k = li + 1; k < Math.min(lines.length, li + 5); k++) {
                        if (lines[k].yc - lines[li].yc > lineGap * 6) break;
                        const l2 = lines[k];
                        const t2 = collectNumberTokensInRange(l2, nfSpan.x0 - 10, nfSpan.x1 + 260);
                        const s2 = joinNumberTokens(t2, Math.max(14, (l2.h || 10) * 2.2));
                        const b2 = pickLongest(s2, 4);
                        if (b2) { nfByWords = b2; break; }
                        const below = collectNumberInLineRange(
                            l2,
                            nfSpan.x0 - 10,
                            nfSpan.x1 + 200,
                            { allowNonDigit: 2, maxDigits: 9, maxTokens: 5 }
                        );
                        if (below.length >= 4) { nfByWords = below; break; }
                    }
                }
            }
            if (!nfByWords) {
                nfByWords = findNumberInColumnBelow(words, nfSpan, {
                    minDigits: 4,
                    maxDigits: 9,
                    preferNonZero: true,
                    pageHeight: maxY
                });
            }
        }
        if (!nfByWords) {
            const anchors = [
                ...findAllLabelAnchors(words, ["NFE"]),
                ...findAllLabelAnchors(words, ["NF"])
            ];
            const cands = [];
            for (const a of anchors) {
                const li = findLineIndexByY(lines, a.yc, lineGap);
                if (li >= 0) {
                    const sameLine = collectNumberInLineRange(lines[li], a.x1 + 2, a.x1 + 140, { allowNonDigit: 2, maxDigits: 9, maxTokens: 4 });
                    if (sameLine.length >= 4 && sameLine.length <= 9) cands.push(sameLine);
                    for (let k = li + 1; k < Math.min(lines.length, li + 4); k++) {
                        if (lines[k].yc - lines[li].yc > lineGap * 6) break;
                        const below = collectNumberInLineRange(lines[k], a.x0 - 10, a.x1 + 140, { allowNonDigit: 2, maxDigits: 9, maxTokens: 4 });
                        if (below.length >= 4 && below.length <= 9) { cands.push(below); break; }
                    }
                }
                const near = findNumberNearAnchor(words, a, { minDigits: 4, maxDigits: 9 });
                if (near) cands.push(near);
                const below2 = findNumberBelowAnchor(words, a, { minDigits: 4, maxDigits: 9, pageWidth: maxX, pageHeight: maxY });
                if (below2) cands.push(below2);
            }
            nfByWords = pickBestCandidateNumber(cands, 4);
        }
    }
    const nfByText =
        pickNumberAfterLabel(t, ["NF-E", "NFE", "NF", "NOTA", "FISCAL", "Nº", "N°", "NO", "N."], 4) ||
        pick(/\bNF(?:-?e)?\b\s*[:\-]?\s*([\d.\s]{4,})/i, t) ||
        pick(/\bNFE\b\s*[:\-]?\s*([\d.\s]{4,})/i, t);
    const nfFromWords = normalizeDigits(nfByWords);
    const nfFromDoc = normalizeDigits(nfFromDocText);
    const nfFromBox = normalizeDigits(nfFromBoxText);
    const nfFromSerie = normalizeDigits(nfFromSerieDocWords);
    const nfFromText = normalizeDigits(nfByText);
    let nf = pickConsistentNf([nfFromSerie, nfFromBox, nfFromDoc, nfFromWords, nfFromText]);

    if (!nfFromSerie && !nfFromBox && !nfFromDoc && nfFromWords && nfFromText && nfFromWords !== nfFromText) {
        const wordHasExtraLeadingOne =
            nfFromWords.length === nfFromText.length + 1 &&
            nfFromWords.startsWith("1") &&
            nfFromWords.endsWith(nfFromText);
        const textHasExtraLeadingOne =
            nfFromText.length === nfFromWords.length + 1 &&
            nfFromText.startsWith("1") &&
            nfFromText.endsWith(nfFromWords);

        if (wordHasExtraLeadingOne) {
            nf = nfFromText;
        } else if (textHasExtraLeadingOne) {
            nf = nfFromWords;
        } else {
            nf = pickBestCandidateNumber([nfFromWords, nfFromText], 4) || nf;
        }
    }

    // Evita confundir NF com CT-e quando OCR aproxima os campos.
    const cteDigits = normalizeDigits(cte);
    if (nf && cteDigits && nf === cteDigits) {
        if (nfFromSerie && nfFromSerie !== cteDigits) {
            nf = nfFromSerie;
        } else if (nfFromBox && nfFromBox !== cteDigits) {
            nf = nfFromBox;
        } else if (nfFromDoc && nfFromDoc !== cteDigits) {
            nf = nfFromDoc;
        } else if (nfFromText && nfFromText !== cteDigits) {
            nf = nfFromText;
        } else if (nfFromWords && nfFromWords !== cteDigits) {
            nf = nfFromWords;
        }
    }

    let rem = extractNameByLabel(words, lines, lineGap, ["REMETENTE", "REM"], leftCol, labelSet, stopSet);
    if (!rem) {
        rem = findBlock(
            t,
            ["REM", "REMETENTE", "EXPEDIDOR"],
            ["DEST", "DESTINAT", "CONSIGNAT", "TOMADOR", "CEP", "VOLUMES", "PESO", "CT-E", "NF"]
        );
        rem = trimAtLabel(rem, /\bDEST\b|\bDESTINAT|\bCONSIGNAT|\bTOMADOR\b/i);
        rem = rem.replace(/\b(FONE|ENDEREÇ?O|CNPJ|CPF|IE|CEP|CIDADE|UF)\b.*$/i, "").trim();
    }
    rem = stripAfterMarkers(rem, [" FONE", " ENDERE", " CNPJ", " CPF", " IE", " CEP", " CIDADE", " UF", " RUA", " AV", " BAIRRO", " CENTRO"]);

    let dest = extractNameByLabel(words, lines, lineGap, ["DESTINATARIO", "DESTINATÁRIO", "DEST"], rightCol, labelSet, stopSet);
    if (!dest) {
        dest = findBlock(
            t,
            ["DEST", "DESTINATÁRIO", "DESTINATARIO", "CONSIGNATÁRIO", "CONSIGNATARIO"],
            ["REM", "REMETENTE", "EXPEDIDOR", "TOMADOR", "CEP", "VOLUMES", "PESO", "CT-E", "NF"]
        );
        dest = trimAtLabel(dest, /\bREM\b|\bREMETENTE\b|\bEXPEDIDOR\b|\bTOMADOR\b/i);
        dest = dest.replace(/\b(FONE|ENDEREÇ?O|CNPJ|CPF|IE|CEP|CIDADE|UF)\b.*$/i, "").trim();
    }
    dest = stripAfterMarkers(dest, [" FONE", " ENDERE", " CNPJ", " CPF", " IE", " CEP", " CIDADE", " UF", " RUA", " AV", " BAIRRO", " CENTRO"]);

    const destLabelIdx = Math.max(
        t.search(/\bDEST\b/i),
        t.search(/\bDESTINAT[ÁA]RIO\b/i)
    );

    let cepByWords = "";
    if (words.length) {
        const destAnchor = findLabelAnchor(words, ["DESTINATARIO", "DESTINATÁRIO", "DEST"]);
        const col = destAnchor && destAnchor.xc > halfX ? rightCol : rightCol;
        const cepAnchors = findAnchorsByToken(words, "CEP")
            .filter(w => w.xc >= col.xMin && w.xc <= col.xMax);
        if (cepAnchors.length) {
            let best = cepAnchors[0];
            if (destAnchor) {
                let bestDist = Infinity;
                for (const a of cepAnchors) {
                    const d = Math.abs(a.yc - destAnchor.yc);
                    if (d < bestDist) { bestDist = d; best = a; }
                }
            }
            cepByWords = findNumberNearAnchor(words, best, { minDigits: 8, yTol: Math.max(10, best.h * 0.8) });
        }
    }

    const destCep = extractCepFromBlock(dest);
    const remCep = extractCepFromBlock(rem);
    const cep =
        formatCep(cepByWords) ||
        formatCep(destCep) ||
        formatCep(remCep) ||
        formatCep(pick(/\bCEP\b\s*[:\-]?\s*(\d{5}-\d{3})/i, t)) ||
        formatCep(pickClosestMatch(t, /(\d{5}-\d{3})/g, destLabelIdx >= 0 ? destLabelIdx : 0));

    const volSpan = words.length
        ? (findLabelSpan(words, ["VOLUMES"]) || findLabelSpan(words, ["VOLUME"]) || findLabelSpan(words, ["VOL"]))
        : null;
    let volumesByWords = "";
    if (words.length && volSpan) {
        volumesByWords =
            findNumberInColumnBelow(words, volSpan, { minDigits: 1, maxDigits: 3, preferNonZero: true, pageHeight: maxY }) ||
            findNumberInColumnBelow(words, volSpan, { minDigits: 1, maxDigits: 3, pageHeight: maxY });
    }
    if (!volumesByWords && words.length) {
        const volAnchor = findLabelAnchor(words, ["VOLUMES"]) || findLabelAnchor(words, ["VOLUME"]) || findLabelAnchor(words, ["VOL"]);
        volumesByWords =
            findNumberNearAnchor(words, volAnchor, { minDigits: 1, maxDigits: 3 }) ||
            findNumberBelowAnchor(words, volAnchor, { minDigits: 1, maxDigits: 3, pageWidth: maxX, pageHeight: maxY });
    }
    const volumes =
        volumesByWords ||
        pick(/\bVOLUMES?\b\s*[:\-]?\s*(\d{1,4})/i, t) ||
        pick(/\bVOL\.?\b\s*[:\-]?\s*(\d{1,4})/i, t) ||
        pick(/\bVOL\s*[:\-]?\s*(\d{1,4})/i, t) ||
        pick(/\bVOLUMES?\b\s*[=:]?\s*(\d{1,4})/i, t);

    const pesoSpan = words.length ? findLabelSpan(words, ["PESO", "REAL"]) : null;
    let pesoByWords = "";
    if (words.length && pesoSpan) {
        pesoByWords =
            findNumberInColumnBelow(words, pesoSpan, { minDigits: 2, maxDigits: 6, preferDecimal: true, pageHeight: maxY }) ||
            findNumberInColumnBelow(words, pesoSpan, { minDigits: 2, maxDigits: 6, pageHeight: maxY });
    }
    if (!pesoByWords && words.length) {
        const pesoAnchor = findLabelAnchor(words, ["PESO", "REAL"]);
        pesoByWords =
            findNumberNearAnchor(words, pesoAnchor, { minDigits: 2, maxDigits: 6 }) ||
            findNumberBelowAnchor(words, pesoAnchor, { minDigits: 2, maxDigits: 6, pageWidth: maxX, pageHeight: maxY });
    }
    const pesoReal =
        pesoByWords ||
        pick(/\bPESO\s+REAL\b\s*[:\-]?\s*([\d.,]+)\s*KG/i, t) ||
        pick(/\bPESO\s+REAL\b\s*[:\-]?\s*([\d.,]+)/i, t) ||
        pick(/\bPESO\b\s*[:\-]?\s*([\d.,]+)\s*KG/i, t);

    return {
        REM: rem,
        DEST: dest,
        CEP: cep,
        VOLUMES: volumes,
        "PESO REAL": pesoReal,
        "CT-e": normalizeDigits(cte),
        NF: nf
    };
}

const OUTPUT_HEADERS = ["REM","DEST","CEP","VOLUMES","PESO REAL","CT-e","NF"];

function toCSV(rows) {
    const headers = OUTPUT_HEADERS;
    const sep = ";";
    const clean = (v) =>
        String(v ?? "")
            .replace(/\r?\n+/g, " ")
            .replace(/[ \t]+/g, " ")
            .trim();
    const esc = (v) => `"${clean(v).replace(/"/g, '""')}"`;

    const lines = [];
    rows.forEach((r, idx) => {
        headers.forEach((h) => {
            lines.push([`"${h}"`, esc(r[h])].join(sep));
        });
        if (idx < rows.length - 1) lines.push("");
    });
    return lines.join("\n");
}

function clearOutputDownloads() {
    if (csvObjectUrl) URL.revokeObjectURL(csvObjectUrl);
    if (xmlObjectUrl) URL.revokeObjectURL(xmlObjectUrl);
    csvObjectUrl = null;
    xmlObjectUrl = null;
    $download.removeAttribute("href");
    $download.classList.add("hidden");
    if ($downloadXml) {
        $downloadXml.removeAttribute("href");
        $downloadXml.classList.add("hidden");
    }
}

function escapeXml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}

function xmlAttr(name, value) {
    if (value === undefined || value === null || value === "") return "";
    return ` ${name}="${escapeXml(value)}"`;
}

function fmtXmlNumber(value) {
    const n = Number(value);
    if (!Number.isFinite(n)) return "";
    return n.toFixed(2);
}

function buildPageXmlSnapshot(pageNumber, fields, ocr, source = "ocr") {
    const words = mapOcrWords(ocr?.data);
    const { lines } = groupWordsIntoLines(words);
    return {
        pageNumber,
        width: Number(ocr?.width) || 0,
        height: Number(ocr?.height) || 0,
        source,
        text: normalizeText(ocr?.text || ""),
        fields: { ...fields },
        lines: lines.map((line, lineIndex) => ({
            index: lineIndex + 1,
            y: line.yc,
            text: cleanField(line.words.map(w => w.text).join(" ")),
            words: line.words.map((w, tokenIndex) => ({
                index: tokenIndex + 1,
                text: w.text || "",
                x0: w.x0,
                y0: w.y0,
                x1: w.x1,
                y1: w.y1
            }))
        }))
    };
}

function toDocLikeXml(documents) {
    const lines = [];
    lines.push('<?xml version="1.0" encoding="UTF-8"?>');
    lines.push(`<doclingLikeExport generatedAt="${escapeXml(new Date().toISOString())}">`);

    for (const doc of documents) {
        lines.push(
            `  <document${xmlAttr("fileName", doc.fileName)}${xmlAttr("sourceType", doc.sourceType)}${xmlAttr("pageCount", doc.pages?.length || 0)}>`
        );

        for (const page of (doc.pages || [])) {
            lines.push(
                `    <page${xmlAttr("number", page.pageNumber)}${xmlAttr("source", page.source)}${xmlAttr("width", fmtXmlNumber(page.width))}${xmlAttr("height", fmtXmlNumber(page.height))}>`
            );

            lines.push("      <fields>");
            for (const field of OUTPUT_HEADERS) {
                lines.push(`        <field name="${escapeXml(field)}">${escapeXml(page.fields?.[field] || "")}</field>`);
            }
            lines.push("      </fields>");

            lines.push(`      <text>${escapeXml(page.text || "")}</text>`);
            lines.push("      <layout>");
            for (const line of (page.lines || [])) {
                lines.push(`        <line${xmlAttr("index", line.index)}${xmlAttr("y", fmtXmlNumber(line.y))}>`);
                lines.push(`          <lineText>${escapeXml(line.text || "")}</lineText>`);
                for (const token of (line.words || [])) {
                    lines.push(
                        `          <token${xmlAttr("index", token.index)}${xmlAttr("x0", fmtXmlNumber(token.x0))}${xmlAttr("y0", fmtXmlNumber(token.y0))}${xmlAttr("x1", fmtXmlNumber(token.x1))}${xmlAttr("y1", fmtXmlNumber(token.y1))}>${escapeXml(token.text || "")}</token>`
                    );
                }
                lines.push("        </line>");
            }
            lines.push("      </layout>");

            lines.push("    </page>");
        }

        lines.push("  </document>");
    }

    lines.push("</doclingLikeExport>");
    return lines.join("\n");
}

function isEmptyField(v) {
    return !cleanField(String(v || ""));
}

function mergeFieldsPreferPrimary(primary, secondary) {
    const out = { ...primary };
    for (const key of OUTPUT_HEADERS) {
        if (isEmptyField(out[key]) && !isEmptyField(secondary?.[key])) {
            out[key] = secondary[key];
        }
    }
    return out;
}

function makeExtractWordsFromLayoutWords(words) {
    return words.map(w => ({
        text: w.text || "",
        bbox: {
            x0: w.x0,
            y0: w.y0,
            x1: w.x1,
            y1: w.y1
        }
    }));
}

async function ocrFromImageBlob(blob, label = "") {
    const { data } = await Tesseract.recognize(blob, "por", {
        logger: m => {
            if (m.status === "recognizing text" && typeof m.progress === "number") {
                setOcrProgress(Math.round(m.progress * 100));
            }
        }
    });
    return data.text;
}

// Renderiza PDF para imagens (dataURL) com PDF.js
async function pdfToPageImages(file) {
    const out = await pdfToPageImagesByNumbers(file);
    return out.images.map(p => p.dataUrl);
}

// Renderiza somente páginas selecionadas (ou todas, se omitido)
async function pdfToPageImagesByNumbers(file, pageNumbers = null) {
    if (!window.pdfjsLib) {
        throw new Error("PDF.js não carregou. Verifique conexão ou o link do PDF.js no HTML.");
    }
    const ab = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data: ab }).promise;

    const images = [];
    const targetPages = Array.isArray(pageNumbers) && pageNumbers.length
        ? [...new Set(pageNumbers)].filter(n => n >= 1 && n <= pdf.numPages).sort((a, b) => a - b)
        : Array.from({ length: pdf.numPages }, (_, i) => i + 1);

    for (const pageNum of targetPages) {
        log(`Render PDF página ${pageNum}/${pdf.numPages}...`);
        const page = await pdf.getPage(pageNum);

        // escala: ajuste se OCR estiver ruim (2.0 a 3.0 costuma ajudar)
        const viewport = page.getViewport({ scale: 2.5 });

        const canvas = document.createElement("canvas");
        const ctx = canvas.getContext("2d", { willReadFrequently: true });
        canvas.width = Math.floor(viewport.width);
        canvas.height = Math.floor(viewport.height);

        await page.render({ canvasContext: ctx, viewport }).promise;

        const dataUrl = canvas.toDataURL("image/jpeg", 0.95);
        images.push({ pageNumber: pageNum, dataUrl });
    }
    return { images, totalPages: pdf.numPages };
}

async function pdfToNativeXmlPages(file) {
    if (!window.pdfjsLib) {
        throw new Error("PDF.js não carregou para extração nativa.");
    }
    const ab = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data: ab }).promise;
    const pages = [];

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
        const page = await pdf.getPage(pageNum);
        const viewport = page.getViewport({ scale: 1 });
        const textContent = await page.getTextContent();
        const words = [];

        for (const item of (textContent.items || [])) {
            const text = String(item.str || "").trim();
            if (!text) continue;

            const x = Number(item.transform?.[4]) || 0;
            const yPdf = Number(item.transform?.[5]) || 0;
            const w = Math.max(Number(item.width) || 0, text.length * 3);
            const h = Math.max(Number(item.height) || 0, 1);
            const yBottom = viewport.height - yPdf;
            const yTop = yBottom - h;

            words.push({
                text,
                x0: x,
                y0: yTop,
                x1: x + w,
                y1: yBottom,
                xc: x + w / 2,
                yc: yTop + h / 2,
                h
            });
        }

        const { lines } = groupWordsIntoLines(words);
        const pageText = lines.map(l => cleanField(l.words.map(w => w.text).join(" "))).filter(Boolean).join("\n");
        const extractWords = makeExtractWordsFromLayoutWords(words);
        const fields = pageText ? extractFields(pageText, { words: extractWords }) : {};
        pages.push({
            pageNumber: pageNum,
            width: viewport.width,
            height: viewport.height,
            source: "pdf-text",
            text: pageText,
            fields,
            lines: lines.map((line, lineIndex) => ({
                index: lineIndex + 1,
                y: line.yc,
                text: cleanField(line.words.map(w => w.text).join(" ")),
                words: line.words.map((w, tokenIndex) => ({
                    index: tokenIndex + 1,
                    text: w.text || "",
                    x0: w.x0,
                    y0: w.y0,
                    x1: w.x1,
                    y1: w.y1
                }))
            }))
        });
    }

    return pages;
}

function dataUrlToBlob(dataUrl) {
    const [meta, b64] = dataUrl.split(",");
    const mime = meta.match(/data:(.*?);base64/)?.[1] || "image/jpeg";
    const bin = atob(b64);
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return new Blob([arr], { type: mime });
}

function buildRenamedPdfName(nfDigits) {
    const nf = normalizeDigits(nfDigits);
    if (!nf) return "NF_NAO_ENCONTRADA.pdf";
    return `${nf}.pdf`;
}

function clearRenamedLinks() {
    for (const u of renamedObjectUrls) URL.revokeObjectURL(u);
    renamedObjectUrls = [];
    renamedDownloadsState = [];
    $renamed.innerHTML = "";
    $renamed.classList.add("hidden");
    if ($downloadAll) {
        $downloadAll.classList.add("hidden");
        $downloadAll.disabled = false;
        $downloadAll.textContent = "Baixar PDFs em massa (ZIP)";
    }
}

function renderRenamedLinks(downloads) {
    renamedDownloadsState = [...downloads];
    if (!downloads.length) return;
    const title = document.createElement("h2");
    title.textContent = "PDFs renomeados";
    $renamed.appendChild(title);

    for (const d of downloads) {
        const a = document.createElement("a");
        a.href = d.url;
        a.download = d.name;
        a.textContent = `Baixar ${d.name}`;
        $renamed.appendChild(a);
    }
    $renamed.classList.remove("hidden");
    if ($downloadAll && downloads.length) {
        $downloadAll.textContent = `Baixar PDFs em massa (ZIP) (${downloads.length})`;
        $downloadAll.classList.remove("hidden");
    }
}

function uniqueZipName(originalName, usedNames) {
    const safe = sanitizeFilePart(originalName || "arquivo.pdf") || "arquivo.pdf";
    if (!usedNames.has(safe)) {
        usedNames.add(safe);
        return safe;
    }
    const dot = safe.lastIndexOf(".");
    const base = dot > 0 ? safe.slice(0, dot) : safe;
    const ext = dot > 0 ? safe.slice(dot) : "";
    let i = 2;
    while (true) {
        const candidate = `${base}_${i}${ext}`;
        if (!usedNames.has(candidate)) {
            usedNames.add(candidate);
            return candidate;
        }
        i += 1;
    }
}

async function downloadRenamedZip(downloads) {
    if (!downloads.length) throw new Error("Nenhum PDF renomeado disponível.");
    if (!window.JSZip) throw new Error("JSZip não carregou. Verifique a conexão.");

    const zip = new window.JSZip();
    const usedNames = new Set();
    for (const d of downloads) {
        const response = await fetch(d.url);
        if (!response.ok) throw new Error(`Falha ao ler ${d.name}.`);
        const blob = await response.blob();
        zip.file(uniqueZipName(d.name, usedNames), blob);
    }

    const zipBlob = await zip.generateAsync({ type: "blob", compression: "DEFLATE" });
    const zipUrl = URL.createObjectURL(zipBlob);
    const a = document.createElement("a");
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    a.href = zipUrl;
    a.download = `pdfs_renomeados_${stamp}.zip`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(zipUrl), 2000);
}

function queueLabel(status, detail = "") {
    const labels = {
        pending: "Na fila",
        processing: "Processando",
        done: "Concluído",
        error: "Erro",
        ignored: "Ignorado"
    };
    const base = labels[status] || "Na fila";
    return detail ? `${base} • ${detail}` : base;
}

function updateQueueSummary() {
    const total = queueItemsState.length;
    const pending = queueItemsState.filter(i => i.status === "pending").length;
    const processing = queueItemsState.filter(i => i.status === "processing").length;
    const done = queueItemsState.filter(i => i.status === "done").length;
    const error = queueItemsState.filter(i => i.status === "error").length;
    const ignored = queueItemsState.filter(i => i.status === "ignored").length;
    const $summary = $queue.querySelector(".queue-summary");
    if (!$summary) return;
    $summary.textContent =
        `Total: ${total} | Na fila: ${pending} | Processando: ${processing} | Concluídos: ${done} | Erros: ${error} | Ignorados: ${ignored}`;
}

function renderQueue(files) {
    queueItemsState = files.map((file, index) => ({
        id: index,
        file,
        name: file.name,
        status: "pending",
        detail: "",
        el: null,
        stateEl: null
    }));

    $queue.innerHTML = "";
    const title = document.createElement("h2");
    title.textContent = "Fila de processamento";
    const summary = document.createElement("div");
    summary.className = "queue-summary";
    const list = document.createElement("ul");
    list.className = "queue-list";

    for (const item of queueItemsState) {
        const li = document.createElement("li");
        li.className = "queue-item pending";
        const name = document.createElement("span");
        name.className = "queue-name";
        name.textContent = item.name;
        const state = document.createElement("span");
        state.className = "queue-state";
        state.textContent = queueLabel("pending");
        li.appendChild(name);
        li.appendChild(state);
        list.appendChild(li);
        item.el = li;
        item.stateEl = state;
    }

    $queue.appendChild(title);
    $queue.appendChild(summary);
    $queue.appendChild(list);
    $queue.classList.remove("hidden");
    updateQueueSummary();
    return queueItemsState;
}

function updateQueueItem(item, status, detail = "") {
    item.status = status;
    item.detail = detail;
    if (item.el) item.el.className = `queue-item ${status}`;
    if (item.stateEl) item.stateEl.textContent = queueLabel(status, detail);
    updateQueueSummary();
}

async function runQueue(items, concurrency, worker) {
    if (!items.length) return;
    const limit = Math.max(1, Math.min(concurrency, items.length));
    let nextIndex = 0;

    async function runner() {
        while (true) {
            const i = nextIndex;
            nextIndex += 1;
            if (i >= items.length) return;
            await worker(items[i], i);
        }
    }

    const runners = Array.from({ length: limit }, () => runner());
    await Promise.all(runners);
}

async function processPdfFile(file, item) {
    log(`\nPDF: ${file.name}`);
    updateQueueItem(item, "processing", "parse PDF");

    let nativePages = [];
    try {
        nativePages = await pdfToNativeXmlPages(file);
        log(`Parse ${file.name}: ${nativePages.length} página(s) lida(s) do PDF.`);
    } catch (e) {
        log(`Parse ${file.name}: falhou (${e.message}). OCR completo será usado.`);
    }

    if (!nativePages.length) {
        updateQueueItem(item, "processing", "OCR completo");
        const pageImages = await pdfToPageImages(file);
        const rows = [];
        const xmlPages = [];
        const nfCandidates = [];

        for (let i = 0; i < pageImages.length; i++) {
            updateQueueItem(item, "processing", `OCR ${i + 1}/${pageImages.length}`);
            log(`OCR ${file.name} página ${i + 1}/${pageImages.length}...`);
            const blob = dataUrlToBlob(pageImages[i]);
            const ocr = await ocrFromImageSmart(blob);
            let fields = extractFields(ocr.text, ocr.data);
            fields = await refineNFWithCrop(fields, ocr);
            rows.push(fields);
            xmlPages.push(buildPageXmlSnapshot(i + 1, fields, ocr, "ocr"));
            if (fields.NF) nfCandidates.push(fields.NF);
            log(`OK ${file.name} pág ${i + 1}: CT-e ${fields["CT-e"]} | NF ${fields["NF"]} | CEP ${fields["CEP"]}`);
        }

        const bestPdfNF = pickBestCandidateNumber(nfCandidates, 4);
        const renamedName = buildRenamedPdfName(bestPdfNF);
        const pdfUrl = URL.createObjectURL(file);
        renamedObjectUrls.push(pdfUrl);
        updateQueueItem(item, "done", `${rows.length} página(s)`);
        log(`Download renomeado: ${renamedName}`);

        return {
            rows,
            xmlDocument: {
                fileName: file.name,
                sourceType: "pdf",
                pages: xmlPages
            },
            renamedDownload: { url: pdfUrl, name: renamedName }
        };
    }

    const rows = nativePages.map(p => ({ ...(p.fields || {}) }));
    let xmlPages = nativePages.map(p => ({ ...p, fields: { ...(p.fields || {}) } }));
    const parsedNfCandidates = rows.map(r => normalizeDigits(r.NF)).filter(Boolean);

    const ocrFallbackPages = [];
    for (let i = 0; i < nativePages.length; i++) {
        const hasNativeText = !isEmptyField(nativePages[i].text);
        const hasNativeNf = !isEmptyField(rows[i]?.NF);
        if (!hasNativeText || !hasNativeNf) {
            ocrFallbackPages.push(i + 1);
        }
    }

    if (ocrFallbackPages.length) {
        updateQueueItem(item, "processing", `OCR fallback (${ocrFallbackPages.length})`);
        const { images: fallbackImages, totalPages } = await pdfToPageImagesByNumbers(file, ocrFallbackPages);
        const imageByPage = new Map(fallbackImages.map(p => [p.pageNumber, p.dataUrl]));

        for (const pageNum of ocrFallbackPages) {
            const img = imageByPage.get(pageNum);
            if (!img) continue;
            updateQueueItem(item, "processing", `OCR pág ${pageNum}/${totalPages}`);
            log(`OCR fallback ${file.name} página ${pageNum}/${totalPages}...`);
            const blob = dataUrlToBlob(img);
            const ocr = await ocrFromImageSmart(blob);
            let ocrFields = extractFields(ocr.text, ocr.data);
            ocrFields = await refineNFWithCrop(ocrFields, ocr);

            const idx = pageNum - 1;
            const merged = mergeFieldsPreferPrimary(rows[idx] || {}, ocrFields);
            rows[idx] = merged;

            if (!isEmptyField(nativePages[idx]?.text)) {
                xmlPages[idx] = {
                    ...xmlPages[idx],
                    source: "pdf-text+ocr",
                    fields: { ...merged }
                };
            } else {
                xmlPages[idx] = buildPageXmlSnapshot(pageNum, merged, ocr, "ocr");
            }
            log(`OK fallback ${file.name} pág ${pageNum}: CT-e ${merged["CT-e"]} | NF ${merged["NF"]} | CEP ${merged["CEP"]}`);
        }
    }

    const bestPdfNF =
        pickBestCandidateNumber(parsedNfCandidates, 4) ||
        pickBestCandidateNumber(rows.map(r => r.NF), 4);
    const renamedName = buildRenamedPdfName(bestPdfNF);
    const pdfUrl = URL.createObjectURL(file);
    renamedObjectUrls.push(pdfUrl);
    updateQueueItem(item, "done", `${rows.length} página(s)`);
    log(`Download renomeado: ${renamedName}`);

    return {
        rows,
        xmlDocument: {
            fileName: file.name,
            sourceType: "pdf",
            pages: xmlPages
        },
        renamedDownload: { url: pdfUrl, name: renamedName }
    };
}

async function processImageFile(file, item) {
    log(`\nIMG: ${file.name}`);
    updateQueueItem(item, "processing", "OCR");
    const ocr = await ocrFromImageSmart(file);
    let fields = extractFields(ocr.text, ocr.data);
    fields = await refineNFWithCrop(fields, ocr);
    log(`OK ${file.name}: CT-e ${fields["CT-e"]} | NF ${fields["NF"]} | CEP ${fields["CEP"]}`);
    updateQueueItem(item, "done", "1 imagem");
    return {
        rows: [fields],
        xmlDocument: {
            fileName: file.name,
            sourceType: "image",
            pages: [buildPageXmlSnapshot(1, fields, ocr, "ocr")]
        },
        renamedDownload: null
    };
}

if ($downloadAll) {
    $downloadAll.addEventListener("click", async () => {
        if (!renamedDownloadsState.length) {
            log("Nenhum PDF renomeado para download em massa.");
            return;
        }
        const originalLabel = $downloadAll.textContent;
        $downloadAll.disabled = true;
        $downloadAll.textContent = "Gerando ZIP...";
        try {
            await downloadRenamedZip(renamedDownloadsState);
            log(`ZIP gerado com ${renamedDownloadsState.length} PDF(s) renomeado(s).`);
        } catch (e) {
            log(`ERRO ZIP: ${e.message}`);
        } finally {
            $downloadAll.disabled = false;
            $downloadAll.textContent = originalLabel;
        }
    });
}

$run.addEventListener("click", async () => {
    clearLog();
    setProgress(0);
    clearOutputDownloads();
    clearRenamedLinks();
    $queue.classList.add("hidden");
    $queue.innerHTML = "";

    const files = [...($file.files || [])];
    if (!files.length) return log("Selecione PDFs e/ou imagens (JPEG/PNG).");

    $run.disabled = true;
    queueProgressMode = true;

    try {
        const rows = [];
        const xmlDocuments = [];
        const renamedDownloads = [];

        const queueItems = renderQueue(files);
        let completed = 0;

        await runQueue(queueItems, FILE_CONCURRENCY, async (item) => {
            const f = item.file;
            const isPdf = f.type === "application/pdf" || f.name.toLowerCase().endsWith(".pdf");
            const isImage = f.type.startsWith("image/");

            try {
                if (isPdf) {
                    const result = await processPdfFile(f, item);
                    rows.push(...result.rows);
                    if (result.xmlDocument) xmlDocuments.push(result.xmlDocument);
                    if (result.renamedDownload) renamedDownloads.push(result.renamedDownload);
                } else if (isImage) {
                    const result = await processImageFile(f, item);
                    rows.push(...result.rows);
                    if (result.xmlDocument) xmlDocuments.push(result.xmlDocument);
                } else {
                    updateQueueItem(item, "ignored", "tipo não suportado");
                    log(`\nIgnorado: ${f.name} (tipo não suportado)`);
                }
            } catch (e) {
                updateQueueItem(item, "error", "falha no processamento");
                log(`ERRO: ${f.name} -> ${e.message}`);
            } finally {
                completed += 1;
                setProgress(Math.round((completed / files.length) * 100));
            }
        });

        if (!rows.length) {
            return log("\nNenhum registro válido gerado.");
        }

        const csv = toCSV(rows);
        const csvBlob = new Blob([csv], { type: "text/csv;charset=utf-8" });
        csvObjectUrl = URL.createObjectURL(csvBlob);

        $download.href = csvObjectUrl;
        $download.download = "dacte_saida.csv";
        $download.textContent = `Baixar CSV (${rows.length} linha(s))`;
        $download.classList.remove("hidden");

        const xml = toDocLikeXml(xmlDocuments);
        const xmlBlob = new Blob([xml], { type: "application/xml;charset=utf-8" });
        xmlObjectUrl = URL.createObjectURL(xmlBlob);
        if ($downloadXml) {
            $downloadXml.href = xmlObjectUrl;
            $downloadXml.download = "dacte_saida.xml";
            $downloadXml.textContent = `Baixar XML (${xmlDocuments.length} doc(s))`;
            $downloadXml.classList.remove("hidden");
        }

        renderRenamedLinks(renamedDownloads);

        log("\nPronto. Fila concluída. Clique em “Baixar CSV”, “Baixar XML”, “Baixar PDFs em massa (ZIP)” e/ou nos PDFs renomeados.");
    } finally {
        queueProgressMode = false;
        $run.disabled = false;
    }
});
async function preprocessToBlob(fileOrBlob, scale = 2.2) {
    const bmp = await createImageBitmap(fileOrBlob);
    const canvas = document.createElement("canvas");
    canvas.width = Math.floor(bmp.width * scale);
    canvas.height = Math.floor(bmp.height * scale);

    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    ctx.drawImage(bmp, 0, 0, canvas.width, canvas.height);

    // grayscale + threshold simples (ajuda MUITO em foto/scan)
    const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const d = img.data;

    for (let i = 0; i < d.length; i += 4) {
        const r = d[i], g = d[i+1], b = d[i+2];
        // luminância
        let v = 0.299*r + 0.587*g + 0.114*b;

        // aumenta contraste leve
        v = (v - 128) * 1.25 + 128;

        // threshold
        const t = v > 165 ? 255 : 0;
        d[i] = d[i+1] = d[i+2] = t;
    }

    ctx.putImageData(img, 0, 0);
    return await new Promise(res => canvas.toBlob(res, "image/png", 1));
}

async function ocrFromImageSmart(fileOrBlob) {
    const pre = await preprocessToBlob(fileOrBlob);
    const bmp = await createImageBitmap(pre);

    const { data } = await Tesseract.recognize(pre, "por", {
        logger: m => {
            if (m.status === "recognizing text" && typeof m.progress === "number") {
                setOcrProgress(Math.round(m.progress * 100));
            }
        }
    });

    bmp.close?.();
    return { text: data.text, data, pre, width: bmp.width, height: bmp.height };
}

async function ocrFromPreCrop(preBlob, rect) {
    const bmp = await createImageBitmap(preBlob);
    const x = Math.max(0, Math.floor(rect.x));
    const y = Math.max(0, Math.floor(rect.y));
    const w = Math.min(bmp.width - x, Math.max(1, Math.floor(rect.w)));
    const h = Math.min(bmp.height - y, Math.max(1, Math.floor(rect.h)));

    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    ctx.drawImage(bmp, x, y, w, h, 0, 0, w, h);
    bmp.close?.();

    const blob = await new Promise(res => canvas.toBlob(res, "image/png", 1));
    const { data } = await Tesseract.recognize(blob, "por", {});
    return data.text || "";
}

function findWordByDigits(words, digits) {
    if (!digits) return null;
    let best = null;
    for (const w of words) {
        const d = extractNumberFromWord(w.text || "");
        if (!d) continue;
        if (d === digits || d.endsWith(digits) || digits.endsWith(d)) {
            if (!best || w.bbox.y1 > best.bbox.y1) best = w;
        }
    }
    return best;
}

async function refineNFWithCrop(fields, ocr) {
    try {
        if (!ocr?.pre || !ocr?.data?.words) return fields;
        if (String(fields.NF || "").length >= 4) return fields;

        const words = ocr.data.words;
        const cteDigits = normalizeDigits(fields["CT-e"]);
        const anchor = findWordByDigits(words, cteDigits);
        if (!anchor) return fields;

        const bbox = anchor.bbox;
        const cellW = Math.max(40, bbox.x1 - bbox.x0);
        const cellH = Math.max(20, bbox.y1 - bbox.y0);
        const rect = {
            x: bbox.x0 - cellW * 0.25,
            y: bbox.y1 + cellH * 0.15,
            w: cellW * 1.8,
            h: cellH * 1.8
        };

        const cropText = await ocrFromPreCrop(ocr.pre, rect);
        const nfDigits = extractBestDigits(cropText, 4, 9);
        if (nfDigits) fields.NF = nfDigits;
    } catch (e) {
        log(`NF crop ERRO: ${e.message}`);
    }
    return fields;
}
