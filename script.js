const $file = document.getElementById("file");
const $run = document.getElementById("run");
const $log = document.getElementById("log");
const $download = document.getElementById("download");
const $bar = document.getElementById("bar");

function log(msg) { $log.textContent += msg + "\n"; }
function clearLog() { $log.textContent = ""; }
function setProgress(p) { $bar.style.width = `${Math.max(0, Math.min(100, p))}%`; }

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
    // prefer longest
    let maxLen = Math.max(...pool.map(n => n.length));
    const longest = pool.filter(n => n.length === maxLen);
    // tie-break by frequency
    const freq = new Map();
    for (const n of longest) freq.set(n, (freq.get(n) || 0) + 1);
    let best = longest[0];
    let bestCount = 0;
    for (const [k, v] of freq.entries()) {
        if (v > bestCount) { bestCount = v; best = k; }
    }
    return best;
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

function extractFields(ocrText, ocrData = null) {
    const t = normalizeText(ocrText);
    const words = Array.isArray(ocrData?.words) ? ocrData.words.map(w => ({
        text: w.text || "",
        x0: w.bbox.x0,
        y0: w.bbox.y0,
        x1: w.bbox.x1,
        y1: w.bbox.y1,
        xc: (w.bbox.x0 + w.bbox.x1) / 2,
        yc: (w.bbox.y0 + w.bbox.y1) / 2,
        h: (w.bbox.y1 - w.bbox.y0)
    })) : [];

    const { lines, lineGap } = groupWordsIntoLines(words);
    const maxX = words.length ? Math.max(...words.map(w => w.x1)) : 0;
    const maxY = words.length ? Math.max(...words.map(w => w.y1)) : 0;
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

    let nfByWords = "";
    if (words.length) {
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
    const nfRaw =
        nfByWords ||
        pickNumberAfterLabel(t, ["NF-E", "NFE", "NF", "NOTA", "FISCAL", "Nº", "N°", "NO", "N."], 4) ||
        pick(/\bNF(?:-?e)?\b\s*[:\-]?\s*([\d.\s]{4,})/i, t) ||
        pick(/\bNFE\b\s*[:\-]?\s*([\d.\s]{4,})/i, t);
    const nf = normalizeDigits(nfRaw);

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

function toCSV(rows) {
    const headers = ["REM","DEST","CEP","VOLUMES","PESO REAL","CT-e","NF"];
    const sep = ";";
    const clean = (v) =>
        String(v ?? "")
            .replace(/\r?\n+/g, " ")
            .replace(/[ \t]+/g, " ")
            .trim();
    const esc = (v) => `"${clean(v).replace(/"/g, '""')}"`;

    // Formato vertical: uma linha por campo (coluna A = campo, coluna B = valor)
    const lines = [];
    rows.forEach((r, idx) => {
        headers.forEach((h) => {
            lines.push([`"${h}"`, esc(r[h])].join(sep));
        });
        // separador entre registros (linha em branco), se houver mais de um
        if (idx < rows.length - 1) lines.push("");
    });
    return lines.join("\n");
}

async function ocrFromImageBlob(blob, label = "") {
    const { data } = await Tesseract.recognize(blob, "por", {
        logger: m => {
            if (m.status === "recognizing text" && typeof m.progress === "number") {
                setProgress(Math.round(m.progress * 100));
            }
        }
    });
    return data.text;
}

// Renderiza PDF para imagens (dataURL) com PDF.js
async function pdfToPageImages(file) {
    if (!window.pdfjsLib) {
        throw new Error("PDF.js não carregou. Verifique conexão ou o link do PDF.js no HTML.");
    }
    const ab = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data: ab }).promise;

    const images = [];
    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
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
        images.push(dataUrl);
    }
    return images;
}

function dataUrlToBlob(dataUrl) {
    const [meta, b64] = dataUrl.split(",");
    const mime = meta.match(/data:(.*?);base64/)?.[1] || "image/jpeg";
    const bin = atob(b64);
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return new Blob([arr], { type: mime });
}

$run.addEventListener("click", async () => {
    clearLog();
    setProgress(0);
    $download.classList.add("hidden");

    const files = [...($file.files || [])];
    if (!files.length) return log("Selecione PDFs e/ou imagens (JPEG/PNG).");

    $run.disabled = true;

    const rows = [];

    for (const f of files) {
        const isPdf = f.type === "application/pdf" || f.name.toLowerCase().endsWith(".pdf");
        const isImage = f.type.startsWith("image/");

        try {
            if (isPdf) {
                log(`\nPDF: ${f.name}`);
                const pageImages = await pdfToPageImages(f);

                // Faz OCR em cada página e extrai 1 linha por página
                for (let i = 0; i < pageImages.length; i++) {
                    log(`OCR página ${i + 1}/${pageImages.length}...`);
                    const blob = dataUrlToBlob(pageImages[i]);
                    const ocr = await ocrFromImageSmart(blob);
                    let fields = extractFields(ocr.text, ocr.data);
                    fields = await refineNFWithCrop(fields, ocr);
                    rows.push(fields);

                    log(`OK pág ${i + 1}: CT-e ${fields["CT-e"]} | NF ${fields["NF"]} | CEP ${fields["CEP"]}`);
                    setProgress(0);
                }
            } else if (isImage) {
                log(`\nIMG: ${f.name}`);
                const ocr = await ocrFromImageSmart(f);
                let fields = extractFields(ocr.text, ocr.data);
                fields = await refineNFWithCrop(fields, ocr);
                rows.push(fields);

                log(`OK: CT-e ${fields["CT-e"]} | NF ${fields["NF"]} | CEP ${fields["CEP"]}`);
                setProgress(0);
            } else {
                log(`\nIgnorado: ${f.name} (tipo não suportado)`);
            }
        } catch (e) {
            log(`ERRO: ${f.name} -> ${e.message}`);
            setProgress(0);
        }
    }

    if (!rows.length) {
        $run.disabled = false;
        return log("\nNenhum registro válido gerado.");
    }

    const csv = toCSV(rows);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);

    $download.href = url;
    $download.download = "dacte_saida.csv";
    $download.textContent = `Baixar CSV (${rows.length} linha(s))`;
    $download.classList.remove("hidden");

    $run.disabled = false;
    log("\nPronto. Clique em “Baixar CSV”.");
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
                setProgress(Math.round(m.progress * 100));
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
