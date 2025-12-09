// === CONFIG ===
// Hard-coded published CSV URL served via your Cloudflare Worker
const CSV_URL = "/sheet";

// Column indices (0-based: A=0, B=1, ..., K=10)
const NAME_COL_INDEX = 1;          // Column B
const SCORE_START_COL_INDEX = 5;   // Column F
const SCORE_END_COL_INDEX = 10;    // Column K (inclusive)

const loadStatus = document.getElementById("load-status");
const debugInfo = document.getElementById("debug-info");
const correlationUI = document.getElementById("correlation-ui");
const personASelect = document.getElementById("person-a");
const personBSelect = document.getElementById("person-b");
const computeBtn = document.getElementById("compute-btn");
const computeStatus = document.getElementById("compute-status");
const resultBox = document.getElementById("result");

// Theme toggle
const themeToggleBtn = document.getElementById("theme-toggle");
const themeIcon = document.getElementById("theme-icon");
const themeLabel = document.getElementById("theme-label");

// NEW: Tabs + overall stats container
const tabCompare = document.getElementById("tab-compare");
const tabOverall = document.getElementById("tab-overall");
const overallStats = document.getElementById("overall-stats");

let people = [];        // { name: string, scores: (number|null)[] }
let headerRow = [];     // first row of the CSV – used for location names

// --- THEME LOGIC ---
function updateThemeToggleUI(currentTheme) {
    if (currentTheme === "dark") {
        // Currently dark → button should say "Light"
        themeIcon.textContent = "☀️";
        themeLabel.textContent = "Light";

        // Force Light to be white, regardless of CSS variables
        themeLabel.style.color = "#f9fafb";
        themeIcon.style.color = "#f9fafb";
    } else {
        // Currently light → button should say "Dark"
        themeIcon.textContent = "🌙";
        themeLabel.textContent = "Dark";

        // Force Dark to be black, regardless of CSS variables
        themeLabel.style.color = "#111827";
        themeIcon.style.color = "#111827";
    }
}

function applyTheme(theme) {
    document.documentElement.setAttribute("data-theme", theme);
    updateThemeToggleUI(theme);
}

function initTheme() {
    const stored = localStorage.getItem("pref-theme");
    if (stored === "light" || stored === "dark") {
        applyTheme(stored);
    } else {
        const prefersDark =
            window.matchMedia &&
            window.matchMedia("(prefers-color-scheme: dark)").matches;
        applyTheme(prefersDark ? "dark" : "light");
    }
}

themeToggleBtn.addEventListener("click", () => {
    const current =
        document.documentElement.getAttribute("data-theme") || "light";
    const next = current === "light" ? "dark" : "light";
    applyTheme(next);
    localStorage.setItem("pref-theme", next);
});

// --- TAB LOGIC ---
function showTab(which) {
    if (!correlationUI || !overallStats) return;

    if (which === "overall") {
        correlationUI.style.display = "none";
        overallStats.style.display = "block";
        tabOverall.classList.add("active");
        tabCompare.classList.remove("active");
    } else {
        correlationUI.style.display = "block";
        overallStats.style.display = "none";
        tabCompare.classList.add("active");
        tabOverall.classList.remove("active");
    }
}

if (tabCompare && tabOverall) {
    tabCompare.addEventListener("click", () => showTab("compare"));
    tabOverall.addEventListener("click", () => showTab("overall"));
}

// --- Robust CSV parser that handles quotes & commas ---
function parseCSV(text) {
    const rows = [];
    let currentRow = [];
    let currentField = "";
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
        const ch = text[i];

        if (ch === '"') {
            // Handle double quotes ("") inside a quoted field
            if (inQuotes && text[i + 1] === '"') {
                currentField += '"';
                i++; // skip the next quote
            } else {
                inQuotes = !inQuotes;
            }
        } else if (ch === "," && !inQuotes) {
            // Field boundary
            currentRow.push(currentField);
            currentField = "";
        } else if ((ch === "\n" || ch === "\r") && !inQuotes) {
            // End of line
            if (ch === "\r" && text[i + 1] === "\n") {
                i++; // handle CRLF
            }
            currentRow.push(currentField);
            rows.push(currentRow);
            currentRow = [];
            currentField = "";
        } else {
            currentField += ch;
        }
    }

    // Last field / row
    if (currentField.length > 0 || currentRow.length > 0) {
        currentRow.push(currentField);
        rows.push(currentRow);
    }

    // Filter out rows that are completely empty
    return rows.filter(row =>
        row.some(cell => String(cell).trim().length > 0)
    );
}

// --- OVERALL STATS HELPERS ---

function computePersonStats(person) {
    const vals = person.scores.filter(v => v !== null);
    if (!vals.length) return null;

    const n = vals.length;
    const mean = vals.reduce((s, v) => s + v, 0) / n;
    const variance = vals.reduce((s, v) => s + (v - mean) * (v - mean), 0) / n;
    const std = Math.sqrt(variance);
    const min = Math.min(...vals);
    const max = Math.max(...vals);

    return { name: person.name, n, mean, std, min, max };
}

function computeGlobalPairExtremes() {
    let bestPair = null;  // { a, b, r, overlap }
    let worstPair = null;

    for (let i = 0; i < people.length; i++) {
        for (let j = i + 1; j < people.length; j++) {
            const pA = people[i];
            const pB = people[j];
            const { xs, ys } = buildOverlap(pA, pB);
            if (xs.length < 2) continue;
            const r = pearsonCorrelation(xs, ys);
            if (!Number.isFinite(r)) continue;

            if (!bestPair || r > bestPair.r) {
                bestPair = { a: pA.name, b: pB.name, r, overlap: xs.length };
            }
            if (!worstPair || r < worstPair.r) {
                worstPair = { a: pA.name, b: pB.name, r, overlap: xs.length };
            }
        }
    }
    return { bestPair, worstPair };
}

function computePolarizationExtremes() {
    if (!people.length) return { most: null, least: null };
    const numLocations = people[0].scores.length;

    let most = null;  // highest std
    let least = null; // lowest std (least polarizing)

    for (let idx = 0; idx < numLocations; idx++) {
        const vals = [];
        for (const p of people) {
            const v = p.scores[idx];
            if (v !== null) vals.push(v);
        }
        if (vals.length < 2) continue;

        const n = vals.length;
        const mean = vals.reduce((s, v) => s + v, 0) / n;
        const variance = vals.reduce((s, v) => s + (v - mean) * (v - mean), 0) / n;
        const std = Math.sqrt(variance);

        const entry = {
            index: idx,
            name: getLocationName(idx),
            std,
            n
        };

        if (!most || std > most.std) {
            most = entry;
        }
        if (!least || std < least.std) {
            least = entry;
        }
    }

    return { most, least };
}

// Signed squared sum = Σ (score * |score|) over everyone for each location
function computeLocationSignedSquaredSums() {
    if (!people.length) return null;
    const numLocations = people[0].scores.length;

    const locStats = [];

    for (let idx = 0; idx < numLocations; idx++) {
        let signedSquaredSum = 0;
        let count = 0;

        for (const p of people) {
            const v = p.scores[idx];
            if (v !== null) {
                signedSquaredSum += v * Math.abs(v);
                count++;
            }
        }

        if (count > 0) {
            locStats.push({
                index: idx,
                name: getLocationName(idx),
                signedSquaredSum,
                count
            });
        }
    }

    if (!locStats.length) return null;

    // Sort descending by signed squared sum
    locStats.sort((a, b) => b.signedSquaredSum - a.signedSquaredSum);

    const best = locStats[0];
    const worst = locStats[locStats.length - 1];

    return { best, worst, ranking: locStats };
}


function computeMostPolarizingLocation() {
    if (!people.length) return null;
    const numLocations = people[0].scores.length;

    let best = null; // { index, name, std, n }

    for (let idx = 0; idx < numLocations; idx++) {
        const vals = [];
        for (const p of people) {
            const v = p.scores[idx];
            if (v !== null) vals.push(v);
        }
        if (vals.length < 2) continue;

        const n = vals.length;
        const mean = vals.reduce((s, v) => s + v, 0) / n;
        const variance = vals.reduce((s, v) => s + (v - mean) * (v - mean), 0) / n;
        const std = Math.sqrt(variance);

        if (!best || std > best.std) {
            best = {
                index: idx,
                name: getLocationName(idx),
                std,
                n
            };
        }
    }

    return best;
}

function renderOverallStats() {
    if (!overallStats || !people.length) return;

    // Global flattened scores
    const allVals = [];
    for (const p of people) {
        for (const v of p.scores) {
            if (v !== null) allVals.push(v);
        }
    }

    const numPeople = people.length;
    const totalRatings = allVals.length;

    let globalMean = 0, globalStd = 0, minScore = null, maxScore = null;
    if (totalRatings > 0) {
        globalMean = allVals.reduce((s, v) => s + v, 0) / totalRatings;
        const variance = allVals.reduce((s, v) => s + (v - globalMean) * (v - globalMean), 0) / totalRatings;
        globalStd = Math.sqrt(variance);
        minScore = Math.min(...allVals);
        maxScore = Math.max(...allVals);
    }

    const personStats = people
        .map(computePersonStats)
        .filter(Boolean)
        .sort((a, b) => b.mean - a.mean);

    const topN = personStats.slice(0, 3);
    const bottomN = personStats.slice(-3);

    const { bestPair, worstPair } = computeGlobalPairExtremes();
    const { most: mostPolar, least: leastPolar } = computePolarizationExtremes();
    const locScores = computeLocationSignedSquaredSums();


    let html = `
      <h2 style="margin-top:0; margin-bottom:0.5rem;">Overall statistics</h2>
      <p class="description" style="margin-top:0;">
        Group-level view of how everyone scores the locations.
      </p>
      <div>
        <strong>Global summary</strong>
        <ul style="margin-top:0.4rem; padding-left:1.2rem;">
          <li>${numPeople} people in the sheet</li>
          <li>${totalRatings} total ratings across all locations</li>
    `;

    if (totalRatings > 0) {
        html += `
          <li>Average score across everyone: ${globalMean.toFixed(2)} (σ ≈ ${globalStd.toFixed(2)})</li>
          <li>Score range used: ${minScore} to ${maxScore}</li>
        `;
    }

    html += `
        </ul>
      </div>
    `;

    if (personStats.length) {
        html += `
          <hr/>
          <div style="margin-top:0.4rem;">
            <strong>People stats</strong>
            <div style="display:flex; gap:2rem; flex-wrap:wrap; margin-top:0.4rem;">
              <div>
                <div style="font-weight:600; margin-bottom:0.25rem;">Highest average ratings</div>
                <ol style="margin:0; padding-left:1.2rem;">
        `;
        for (const p of topN) {
            html += `<li>${p.name}: ${p.mean.toFixed(2)} (σ ≈ ${p.std.toFixed(2)})</li>`;
        }
        html += `
        </ol>
      </div>
      <div>
        <div style="font-weight:600; margin-bottom:0.25rem;">Lowest average ratings</div>`
        const startIndex = Math.max(personStats.length - bottomN.length + 1, 1);
        html += `<ol style="margin:0; padding-left:1.2rem;" start="${startIndex}">`;
        for (const p of bottomN) {
            html += `<li>${p.name}: ${p.mean.toFixed(2)} (σ ≈ ${p.std.toFixed(2)})</li>`;
        }
        html += `
        </ol>
      </div>
            </div>
          </div>
        `;
    }

    if (bestPair || worstPair) {
        html += `
          <hr/>
          <div style="margin-top:0.4rem;">
            <strong>Strongest relationships</strong>
            <ul style="margin-top:0.4rem; padding-left:1.2rem;">
        `;
        if (bestPair) {
            html += `
              <li>
                Most aligned pair:
                <span class="pill pill-positive">
                  ${bestPair.a} &amp; ${bestPair.b}
                  (${bestPair.r.toFixed(3)}, ${bestPair.overlap} locations)
                </span>
              </li>
            `;
        }
        if (worstPair) {
            html += `
              <li>
                Most opposite pair:
                <span class="pill pill-negative">
                  ${worstPair.a} &amp; ${worstPair.b}
                  (${worstPair.r.toFixed(3)}, ${worstPair.overlap} locations)
                </span>
              </li>
            `;
        }
        html += `
            </ul>
          </div>
        `;
    }

    if (mostPolar || leastPolar) {
        html += `
      <hr/>
      <div style="margin-top:0.4rem;">
        <strong>Polarization by location</strong>
        <ul style="margin-top:0.4rem; padding-left:1.2rem;">
    `;
        if (mostPolar) {
            html += `
          <li>
            Most polarizing location:
            ${mostPolar.name} &mdash; people disagree the most here
            (σ ≈ ${mostPolar.std.toFixed(2)} across ${mostPolar.n} ratings).
          </li>
        `;
        }
        if (leastPolar) {
            html += `
          <li>
            Least polarizing location:
            ${leastPolar.name} &mdash; people are most in agreement here
            (σ ≈ ${leastPolar.std.toFixed(2)} across ${leastPolar.n} ratings).
          </li>
        `;
        }
        html += `
        </ul>
      </div>
    `;
    }

    if (locScores) {
        const { best: highLoc, worst: lowLoc, ranking } = locScores;
        html += `
      <hr/>
      <div style="margin-top:0.4rem;">
        <strong>Location score rankings (signed squared sum)</strong>
        <ul style="margin-top:0.4rem; padding-left:1.2rem;">
          <li>
            Highest-scoring location:
            ${highLoc.name} (signed squared sum ≈ ${highLoc.signedSquaredSum.toFixed(2)}).
          </li>
          <li>
            Lowest-scoring location:
            ${lowLoc.name} (signed squared sum ≈ ${lowLoc.signedSquaredSum.toFixed(2)}).
          </li>
        </ul>
        <details style="margin-top:0.6rem;">
            <summary>See full ranking</summary>
            <table style="
                margin-top:0.5rem;
                border-collapse: collapse;
                width: 100%;
                font-size: 0.9rem;
            ">
                <thead>
                <tr>
                    <th style="text-align:left; padding:0.35rem; border-bottom:1px solid var(--card-border);">
                    Rank
                    </th>
                    <th style="text-align:left; padding:0.35rem; border-bottom:1px solid var(--card-border);">
                    Location
                    </th>
                    <th style="text-align:right; padding:0.35rem; border-bottom:1px solid var(--card-border);">
                    Signed squared sum
                    </th>
                </tr>
                </thead>
                <tbody>
        `;
        for (let i = 0; i < ranking.length; i++) {
            const loc = ranking[i];
            html += `
                <tr>
                    <td style="padding:0.35rem; border-bottom:1px solid var(--card-border);">
                    ${i + 1}
                    </td>
                    <td style="padding:0.35rem; border-bottom:1px solid var(--card-border);">
                    ${loc.name}
                    </td>
                    <td style="padding:0.35rem; text-align:right; border-bottom:1px solid var(--card-border);">
                    ${loc.signedSquaredSum.toFixed(2)}
                    </td>
                </tr>
                `;
        }
        html += `
          </tbody>
            </table>
            </details>
      </div>
    `;
    }



    overallStats.innerHTML = html;
}

// --- LOAD SHEET VIA /sheet PROXY ---
function loadSheet() {
    loadStatus.textContent = "Loading sheet…";
    loadStatus.className = "status";

    fetch(CSV_URL)
        .then(resp => {
            if (!resp.ok) {
                throw new Error("Failed to fetch CSV: " + resp.status);
            }
            return resp.text();
        })
        .then(text => {
            const rows = parseCSV(text);

            headerRow = rows[0] || [];          // <-- save header row
            const dataRows = rows.slice(1);     // skip header

            people = [];
            for (const row of dataRows) {
                if (row.length <= NAME_COL_INDEX) continue;

                const rawName = (row[NAME_COL_INDEX] || "").trim();
                if (!rawName) continue;

                const scores = [];
                for (let c = SCORE_START_COL_INDEX; c <= SCORE_END_COL_INDEX; c++) {
                    let val = row[c];

                    if (val === undefined || val === null || String(val).trim() === "") {
                        scores.push(null);
                    } else {
                        // Clean "+3" → "3"
                        const cleaned = String(val).trim().replace(/^\+/, "");
                        const num = Number(cleaned);
                        scores.push(Number.isFinite(num) ? num : null);
                    }
                }

                people.push({ name: rawName, scores });
            }

            if (people.length === 0) {
                throw new Error("No people found in the expected name column.");
            }

            populateSelects();
            loadStatus.textContent = `Loaded ${people.length} people from sheet.`;
            loadStatus.className = "status success";
            correlationUI.style.display = "block";

            // Render overall stats now that we have data
            renderOverallStats();

            // Default to "Compare" tab
            showTab("compare");

            debugInfo.textContent = `Debug: Loaded ${people.length} people; score dimension = ${people[0].scores.length}.`;
        })
        .catch(err => {
            console.error(err);
            loadStatus.textContent =
                "Error loading sheet. (If this is a CORS error, you’ll need a proxy/worker.)\n" +
                err.message;
            loadStatus.className = "status error";
        });
}

function getLocationName(scoreIndex) {
    // scoreIndex is 0-based index into person.scores (0 → column F, 1 → column G, ...)
    const colIndex = SCORE_START_COL_INDEX + scoreIndex;
    const header = headerRow[colIndex] || "";

    // Try to match "Score the following locations: [Location]"
    const m = header.match(/Score the following locations:\s*\[(.+?)\]\s*$/i);
    if (m) {
        return m[1].trim(); // e.g., "Vancouver"
    }

    // Fallback: just use the header text or a generic label
    const trimmed = header.trim();
    return trimmed || `Location ${scoreIndex + 1}`;
}

function populateSelects() {
    personASelect.innerHTML = "";
    personBSelect.innerHTML = "";

    for (const p of people) {
        const optA = document.createElement("option");
        optA.value = p.name;
        optA.textContent = p.name;
        personASelect.appendChild(optA);

        const optB = document.createElement("option");
        optB.value = p.name;
        optB.textContent = p.name;
        personBSelect.appendChild(optB);
    }

    if (people.length >= 2) {
        personASelect.selectedIndex = 0;
        personBSelect.selectedIndex = 1;
    }
}

// --- CORE STATS ---
function pearsonCorrelation(xs, ys) {
    const n = xs.length;
    if (n < 2) return NaN;

    let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0, sumY2 = 0;
    for (let i = 0; i < n; i++) {
        const x = xs[i];
        const y = ys[i];
        sumX += x;
        sumY += y;
        sumXY += x * y;
        sumX2 += x * x;
        sumY2 += y * y;
    }
    const num = n * sumXY - sumX * sumY;
    const den = Math.sqrt(
        (n * sumX2 - sumX * sumX) *
        (n * sumY2 - sumY * sumY)
    );
    if (den === 0) return NaN;
    return num / den;
}

// helper to build overlapping data between two people
function buildOverlap(personA, personB) {
    const xs = [];
    const ys = [];
    const overlaps = []; // { index, a, b }

    for (let i = 0; i < personA.scores.length; i++) {
        const a = personA.scores[i];
        const b = personB.scores[i];
        if (a !== null && b !== null) {
            xs.push(a);
            ys.push(b);
            overlaps.push({ index: i, a, b });
        }
    }

    return { xs, ys, overlaps };
}

function qualitativeFromR(r) {
    const absR = Math.abs(r);
    if (absR < 0.2) return "very weak similarity";
    if (absR < 0.4) return "weak similarity";
    if (absR < 0.6) return "moderate similarity";
    if (absR < 0.8) return "strong similarity";
    return "very strong similarity";
}

function corrClassFromR(r) {
    if (r > 0.25) return "positive";
    if (r < -0.25) return "negative";
    return "neutral";
}

// compute best and worst matches vs a given person
function computeBestAndWorstFor(basePerson) {
    let best = null;  // { name, r, overlap }
    let worst = null; // { name, r, overlap }

    for (const other of people) {
        if (other.name === basePerson.name) continue;

        const { xs, ys } = buildOverlap(basePerson, other);
        if (xs.length < 2) continue;

        const r = pearsonCorrelation(xs, ys);
        if (!Number.isFinite(r)) continue;

        if (best === null || r > best.r) {
            best = { name: other.name, r, overlap: xs.length };
        }
        if (worst === null || r < worst.r) {
            worst = { name: other.name, r, overlap: xs.length };
        }
    }

    return { best, worst };
}

// --- MAIN COMPUTE ---
function computeCorrelation() {
    const nameA = personASelect.value;
    const nameB = personBSelect.value;

    computeStatus.textContent = "";
    resultBox.style.display = "none";

    if (!nameA || !nameB) {
        computeStatus.textContent = "Please choose both people.";
        computeStatus.className = "status error";
        return;
    }
    if (nameA === nameB) {
        computeStatus.textContent = "Please choose two different people.";
        computeStatus.className = "status error";
        return;
    }

    const personA = people.find(p => p.name === nameA);
    const personB = people.find(p => p.name === nameB);
    if (!personA || !personB) {
        computeStatus.textContent = "Could not find one or both people in the dataset.";
        computeStatus.className = "status error";
        return;
    }

    const xs = [];
    const ys = [];

    // For "same score" + disagreement
    let topSameScore = null; // { score, location }
    let maxDisagreement = null; // { diff, location }

    // For explanation stats
    let sumAbsDiff = 0;
    let sumDiff = 0;

    for (let i = 0; i < personA.scores.length; i++) {
        const a = personA.scores[i];
        const b = personB.scores[i];
        if (a === null || b === null) continue;

        xs.push(a);
        ys.push(b);

        // Track largest-magnitude same score
        if (a === b) {
            if (
                !topSameScore ||
                Math.abs(a) > Math.abs(topSameScore.score)
            ) {
                topSameScore = {
                    score: a,
                    location: getLocationName(i),
                };
            }
        }

        // Largest disagreement where signs differ or one is zero
        const signA = Math.sign(a);
        const signB = Math.sign(b);
        const signsOppositeOrZero =
            (signA === 0 || signB === 0 || signA === -signB) && a !== b;

        if (signsOppositeOrZero) {
            const diff = Math.abs(a - b);
            if (!maxDisagreement || diff > maxDisagreement.diff) {
                maxDisagreement = {
                    diff,
                    location: getLocationName(i),
                };
            }
        }

        // For explanation bullets
        sumAbsDiff += Math.abs(a - b);
        sumDiff += (a - b); // positive => A rates higher
    }

    if (xs.length < 2) {
        computeStatus.textContent =
            `Not enough overlapping rated locations between ${nameA} and ${nameB} to compute a correlation (need at least 2).`;
        computeStatus.className = "status error";
        return;
    }

    const r = pearsonCorrelation(xs, ys);
    if (!Number.isFinite(r)) {
        computeStatus.textContent = "Could not compute a valid correlation.";
        computeStatus.className = "status error";
        return;
    }

    const rounded = r.toFixed(3);
    const absR = Math.abs(r);

    let qualitative;
    if (absR < 0.2) qualitative = "very weak";
    else if (absR < 0.4) qualitative = "weak";
    else if (absR < 0.6) qualitative = "moderate";
    else if (absR < 0.8) qualitative = "strong";
    else qualitative = "very strong";

    const signWord = r >= 0 ? "positive" : "negative";
    const correlationPhrase = `${qualitative} ${signWord} correlation`;

    const direction = r >= 0 ? "aligned" : "opposite";

    // --- Explanation stats (relative movement) ---
    const n = xs.length;
    const meanA = xs.reduce((s, v) => s + v, 0) / n;
    const meanB = ys.reduce((s, v) => s + v, 0) / n;

    let sameDirCount = 0;
    let oppDirCount = 0;

    for (let i = 0; i < n; i++) {
        const da = xs[i] - meanA;
        const db = ys[i] - meanB;
        if (da === 0 || db === 0) continue;

        const prod = da * db;
        if (prod > 0) sameDirCount++;
        else if (prod < 0) oppDirCount++;
    }

    const dirPairs = sameDirCount + oppDirCount;
    let movePercent = 0;
    let moveWord = r >= 0 ? "the same direction" : "opposite directions";

    if (dirPairs > 0) {
        const base =
            r >= 0 ? sameDirCount : oppDirCount;
        movePercent = Math.round(100 * (base / dirPairs));
    }

    const avgAbsDiff = sumAbsDiff / n;
    const avgDiff = sumDiff / n; // >0 => A higher, <0 => B higher

    // --- Build trends (using only largest-magnitude same score) ---
    let trendsHtml = "";

    if (topSameScore || maxDisagreement) {
        trendsHtml += `<hr/><div style="margin-top:0.75rem;"><strong>Trends</strong><br/>`;

        if (topSameScore) {
            trendsHtml += `
        <div>You and ${nameB} both put the score ${topSameScore.score} for ${topSameScore.location}.</div>
      `;
        }

        if (maxDisagreement) {
            trendsHtml += `
        <div style="margin-top:0.5rem;">
          You and ${nameB} disagree the most about ${maxDisagreement.location}.
        </div>
      `;
        }

        trendsHtml += `</div>`;
    }

    // --- Explanation block with bullets ---
    let explanationHtml = `
    <div style="margin-top:0.75rem;">
      <strong>Why this correlation?</strong>
      <div>The correlation is a ${correlationPhrase}.</div>
      <ul style="margin-top:0.5rem; padding-left:1.2rem;">
        <li>You score in ${moveWord} on about ${movePercent}% of locations (relative to your own averages).</li>
        <li>Your scores are usually within ~${avgAbsDiff.toFixed(2)} points of each other.</li>
  `;

    if (Math.abs(avgDiff) >= 0.05) {
        if (avgDiff > 0) {
            explanationHtml += `
        <li>${nameA} tends to rate locations higher than ${nameB} (by about ${avgDiff.toFixed(2)} points on average).</li>
      `;
        } else {
            explanationHtml += `
        <li>${nameB} tends to rate locations higher than ${nameA} (by about ${Math.abs(avgDiff).toFixed(2)} points on average).</li>
      `;
        }
    } else {
        explanationHtml += `
      <li>On average, you both rate locations at about the same level.</li>
    `;
    }

    explanationHtml += `
      </ul>
    </div>
  `;

    // Best / worst match for Person A
    const { best, worst } = computeBestAndWorstFor(personA);

    let matchesHtml = "<hr /><div class=\"matches-section\">";
    matchesHtml += `<strong>Matches for ${nameA}</strong><br/>`;

    function matchLine(label, m) {
        if (!m) return `${label}: <span class="pill pill-neutral">no valid match</span><br/>`;
        const cls = corrClassFromR(m.r);
        const pillClass =
            cls === "positive" ? "pill-positive" :
                cls === "negative" ? "pill-negative" :
                    "pill-neutral";
        return `${label}: <span class="pill ${pillClass}">${m.name} (${m.r.toFixed(3)}, ${m.overlap} locations)</span><br/>`;
    }

    matchesHtml += matchLine("Best match", best);
    matchesHtml += matchLine("Worst match", worst);
    matchesHtml += "</div>";

    // Correlation header (number + short summary)
    resultBox.innerHTML = `
    <strong>
      Correlation between ${nameA} and ${nameB}:
      <span class="corr-number ${r > 0.25 ? "corr-positive" :
            r < -0.25 ? "corr-negative" :
                "corr-neutral"
        }">${rounded}</span>
    </strong><br/>
    <br/>
    Compared on <strong>${xs.length}</strong> locations where both provided a score.<br/>
    This indicates <strong>${qualitative}</strong> similarity and your preferences are mostly <strong>${direction}</strong>.
    ${explanationHtml}
    ${trendsHtml}
    ${matchesHtml}
  `;
    resultBox.style.display = "block";
    computeStatus.textContent = "";
}

computeBtn.addEventListener("click", computeCorrelation);

// Init
initTheme();
loadSheet();
