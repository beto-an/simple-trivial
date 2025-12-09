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
const tabMatrix = document.getElementById("tab-matrix");
const matrixBox = document.getElementById("correlation-matrix");
const tabPersonal = document.getElementById("tab-personal");
const personalStatsBox = document.getElementById("personal-stats");
const tabRanking = document.getElementById("tab-ranking");
const customRankingBox = document.getElementById("custom-ranking");
const tabClusters = document.getElementById("tab-clusters");
const clustersBox = document.getElementById("cluster-stats");

let people = [];        // { name: string, scores: (number|null)[] }
let headerRow = [];     // first row of the CSV – used for location names
let matrixOrder = [];   // indices into `people` for the correlation matrix ordering
let cachedClusterAssignments = null;
let cachedClusterK = null;

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
    correlationUI.style.display = "none";
    overallStats.style.display = "none";
    matrixBox.style.display = "none";
    if (personalStatsBox) personalStatsBox.style.display = "none";
    if (customRankingBox) customRankingBox.style.display = "none";
    if (clustersBox) clustersBox.style.display = "none";

    tabCompare.classList.remove("active");
    tabOverall.classList.remove("active");
    tabMatrix.classList.remove("active");
    if (tabPersonal) tabPersonal.classList.remove("active");
    if (tabRanking) tabRanking.classList.remove("active");
    if (tabClusters) tabClusters.classList.remove("active");

    if (which === "overall") {
        overallStats.style.display = "block";
        tabOverall.classList.add("active");
    } else if (which === "matrix") {
        matrixBox.style.display = "block";
        tabMatrix.classList.add("active");
    } else if (which === "personal") {
        if (personalStatsBox) personalStatsBox.style.display = "block";
        if (tabPersonal) tabPersonal.classList.add("active");
    } else if (which === "ranking") {
        if (customRankingBox) customRankingBox.style.display = "block";
        if (tabRanking) tabRanking.classList.add("active");
    } else if (which === "clusters") {
        if (clustersBox) clustersBox.style.display = "block";
        if (tabClusters) tabClusters.classList.add("active");
    } else {
        correlationUI.style.display = "block";
        tabCompare.classList.add("active");
    }
}


let matrixRendered = false;
const MAX_EXPONENT = 8;          // hard-coded max exponent n
let currentExponent = 2;         // default exponent
let customRankingInitialized = false;

if (tabCompare && tabOverall && tabMatrix && tabPersonal && tabRanking && tabClusters) {
    tabCompare.addEventListener("click", () => showTab("compare"));

    tabOverall.addEventListener("click", () => showTab("overall"));

    tabMatrix.addEventListener("click", () => {
        if (!matrixRendered) {
            renderCorrelationMatrix();
            matrixRendered = true;
        }
        showTab("matrix");
    });

    tabPersonal.addEventListener("click", () => {
        renderPersonalStats();   // default to first person
        showTab("personal");
    });

    tabRanking.addEventListener("click", () => {
        renderCustomRanking();   // initializes controls & renders
        showTab("ranking");
    });

    tabClusters.addEventListener("click", () => {
        renderClusters();
        showTab("clusters");
    });
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

function renderCustomRanking() {
    if (!customRankingBox) return;

    if (!people.length) {
        customRankingBox.innerHTML = `
          <h2 style="margin-top:0; margin-bottom:0.5rem;">Custom ranking</h2>
          <p class="description" style="margin-top:0;">
            No data loaded yet. Once the sheet loads, this tab will rank locations
            using a custom exponent.
          </p>
        `;
        return;
    }

    if (!customRankingInitialized) {
        customRankingBox.innerHTML = `
          <h2 style="margin-top:0; margin-bottom:0.5rem;">Custom ranking</h2>
          <p class="description" style="margin-top:0;">
            Each location receives a score based on everyone&apos;s rank values:
            <code>score(location) = Σ sign(rank) · |rank|<sup>exponent</sup></code>.
            Use the slider or numeric input to adjust the exponent and see how
            the ranking and dispersion change.
          </p>

          <div class="rank-controls">
            <label for="rank-exp-slider">
              Exponent
            </label>
            <input
              id="rank-exp-slider"
              type="range"
              min="0"
              max="${MAX_EXPONENT}"
              step="0.1"
            />
            <input
              id="rank-exp-input"
              type="number"
              min="0"
              max="${MAX_EXPONENT}"
              step="0.1"
            />
          </div>

          <div style="margin-top:0.75rem; font-size:0.85rem; color:var(--muted-soft);">
            <ul style="margin:0.25rem 0 0.5rem; padding-left:1.2rem;">
              <li>Exponent 0 treats every non-zero rank as ±1 (only the sign matters).</li>
              <li>Higher exponents emphasize strong opinions (large |rank|) more.</li>
            </ul>
          </div>

          <div style="margin-top:0.75rem; overflow-x:auto;">
            <table style="
              border-collapse: collapse;
              width: 100%;
              font-size: 0.9rem;
              min-width: 480px;
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
                    Score
                  </th>
                  <th style="text-align:right; padding:0.35rem; border-bottom:1px solid var(--card-border);">
                    Std. dev. of contributions
                  </th>
                  <th style="text-align:left; padding:0.35rem; border-bottom:1px solid var(--card-border);">
                    Visual
                  </th>
                </tr>
              </thead>
              <tbody id="custom-ranking-rows">
              </tbody>
            </table>
          </div>
        `;

        const slider = customRankingBox.querySelector("#rank-exp-slider");
        const input = customRankingBox.querySelector("#rank-exp-input");

        if (slider && input) {
            slider.value = String(currentExponent);
            input.value = String(currentExponent);

            // Slider: live updates as you drag
            slider.addEventListener("input", (e) => {
                const val = Number(e.target.value);
                if (!Number.isNaN(val)) {
                    currentExponent = Math.min(Math.max(val, 0), MAX_EXPONENT);
                    input.value = currentExponent.toFixed(1);
                    updateCustomRankingResults();
                }
            });

            // Numeric input: clamp on change
            input.addEventListener("change", (e) => {
                let val = Number(e.target.value);
                if (Number.isNaN(val)) {
                    val = currentExponent;
                }
                val = Math.min(Math.max(val, 0), MAX_EXPONENT);
                currentExponent = val;
                input.value = currentExponent.toFixed(1);
                slider.value = String(currentExponent);
                updateCustomRankingResults();
            });
        }

        customRankingInitialized = true;
    }

    // Always recompute when the tab is opened (in case people/ratings changed)
    updateCustomRankingResults();
}


function updateCustomRankingResults() {
    if (!customRankingBox || !people.length) return;

    const tbody = customRankingBox.querySelector("#custom-ranking-rows");
    if (!tbody) return;

    const exponent = currentExponent;
    const numLocations = people[0].scores.length;

    const locStats = [];
    let maxAbsScore = 0;

    for (let locIdx = 0; locIdx < numLocations; locIdx++) {
        const contributions = [];

        for (const p of people) {
            const v = p.scores[locIdx];
            if (v === null) continue;

            const sign = Math.sign(v);
            const mag = Math.pow(Math.abs(v), exponent);
            const contrib = sign * mag;

            contributions.push(contrib);
        }

        if (!contributions.length) continue;

        const count = contributions.length;
        const sum = contributions.reduce((s, v) => s + v, 0);
        const mean = sum / count;
        const variance = contributions.reduce((s, v) => s + (v - mean) * (v - mean), 0) / count;
        const std = Math.sqrt(Math.max(variance, 0));

        locStats.push({
            index: locIdx,
            name: getLocationName(locIdx),
            score: sum,
            std,
            count
        });

        const absScore = Math.abs(sum);
        if (absScore > maxAbsScore) {
            maxAbsScore = absScore;
        }
    }

    if (!locStats.length) {
        tbody.innerHTML = `
          <tr>
            <td colspan="5" style="padding:0.5rem; text-align:center; color:var(--muted-soft);">
              No locations have any valid scores.
            </td>
          </tr>
        `;
        return;
    }

    // Sort descending by custom score
    locStats.sort((a, b) => b.score - a.score);

    const rows = [];
    for (let rank = 0; rank < locStats.length; rank++) {
        const loc = locStats[rank];
        const widthPct = maxAbsScore > 0
            ? (Math.abs(loc.score) / maxAbsScore) * 100
            : 0;
        const isPositive = loc.score >= 0;
        const barClass = isPositive
            ? "rank-bar-fill-positive"
            : "rank-bar-fill-negative";

        rows.push(`
          <tr>
            <td style="padding:0.35rem; border-bottom:1px solid var(--card-border);">
              ${rank + 1}
            </td>
            <td style="padding:0.35rem; border-bottom:1px solid var(--card-border);">
              ${loc.name}
            </td>
            <td style="padding:0.35rem; text-align:right; border-bottom:1px solid var(--card-border);">
              ${loc.score.toFixed(3)}
            </td>
            <td style="padding:0.35rem; text-align:right; border-bottom:1px solid var(--card-border);">
              ${Number.isFinite(loc.std) ? loc.std.toFixed(3) : "—"}
            </td>
            <td style="padding:0.35rem; border-bottom:1px solid var(--card-border); min-width:120px;">
              <div class="rank-bar-track">
                <div
                  class="${barClass}"
                  style="width:${widthPct}%;"
                  title="${loc.score.toFixed(3)}"
                ></div>
              </div>
            </td>
          </tr>
        `);
    }

    tbody.innerHTML = rows.join("");
}


// --- CLUSTERING HELPERS ---
// Build a normalized score vector for each person (imputing missing with their mean)
function buildNormalizedVectors() {
    if (!people.length) return [];

    const numLocations = people[0].scores.length;
    const vectors = [];

    for (const person of people) {
        const vals = person.scores.slice();

        // compute mean over non-null entries
        let sum = 0;
        let count = 0;
        for (const v of vals) {
            if (v !== null) {
                sum += v;
                count++;
            }
        }
        const mean = count > 0 ? sum / count : 0;

        // impute missing with mean, then normalize
        let sum2 = 0;
        const filled = new Array(numLocations);
        for (let i = 0; i < numLocations; i++) {
            const v = vals[i] === null ? mean : vals[i];
            filled[i] = v;
            sum2 += v;
        }
        const mu = sum2 / numLocations;
        let varAcc = 0;
        for (let i = 0; i < numLocations; i++) {
            const d = filled[i] - mu;
            varAcc += d * d;
        }
        const std = Math.sqrt(varAcc / numLocations) || 1;

        const norm = filled.map(v => (v - mu) / std);
        vectors.push(norm);
    }
    return vectors;
}

// Simple k-means on normalized vectors
function kMeansCluster(vectors, k) {
    const n = vectors.length;
    if (!n || k <= 0) return { assignments: [], centroids: [] };

    const dim = vectors[0].length;
    k = Math.min(k, n);

    // init centroids: first k people
    const centroids = [];
    for (let c = 0; c < k; c++) {
        centroids.push(vectors[c].slice());
    }

    const assignments = new Array(n).fill(0);

    function distance2(a, b) {
        let s = 0;
        for (let i = 0; i < dim; i++) {
            const d = a[i] - b[i];
            s += d * d;
        }
        return s;
    }

    const MAX_ITERS = 25;
    for (let iter = 0; iter < MAX_ITERS; iter++) {
        // assign
        let changed = false;
        for (let i = 0; i < n; i++) {
            let bestC = 0;
            let bestD = Infinity;
            for (let c = 0; c < k; c++) {
                const d = distance2(vectors[i], centroids[c]);
                if (d < bestD) {
                    bestD = d;
                    bestC = c;
                }
            }
            if (assignments[i] !== bestC) {
                assignments[i] = bestC;
                changed = true;
            }
        }

        // recompute centroids
        const sums = Array.from({ length: k }, () => new Array(dim).fill(0));
        const counts = new Array(k).fill(0);
        for (let i = 0; i < n; i++) {
            const c = assignments[i];
            const v = vectors[i];
            counts[c]++;
            for (let d = 0; d < dim; d++) {
                sums[c][d] += v[d];
            }
        }

        for (let c = 0; c < k; c++) {
            if (counts[c] === 0) continue; // leave centroid as is
            for (let d = 0; d < dim; d++) {
                centroids[c][d] = sums[c][d] / counts[c];
            }
        }

        if (!changed) break;
    }

    return { assignments, centroids };
}

// Compute per-cluster favorite & least favorite locations
function computeClusterLocationExtremes(assignments, k) {
    if (!people.length) return [];
    const numLocations = people[0].scores.length;

    const result = [];
    for (let c = 0; c < k; c++) {
        const memberIdxs = [];
        for (let i = 0; i < assignments.length; i++) {
            if (assignments[i] === c) memberIdxs.push(i);
        }
        if (!memberIdxs.length) {
            result.push({
                cluster: c,
                members: [],
                favorite: null,
                leastFavorite: null
            });
            continue;
        }

        // average score per location for this cluster
        const locStats = [];
        for (let loc = 0; loc < numLocations; loc++) {
            let sum = 0;
            let count = 0;
            for (const idx of memberIdxs) {
                const v = people[idx].scores[loc];
                if (v !== null) {
                    sum += v;
                    count++;
                }
            }
            if (count > 0) {
                locStats.push({
                    idx: loc,
                    name: getLocationName(loc),
                    mean: sum / count,
                    count
                });
            }
        }

        if (!locStats.length) {
            result.push({
                cluster: c,
                members: memberIdxs.map(i => people[i].name),
                favorite: null,
                leastFavorite: null
            });
            continue;
        }

        locStats.sort((a, b) => b.mean - a.mean);
        const favorite = locStats[0];
        const leastFavorite = locStats[locStats.length - 1];

        result.push({
            cluster: c,
            members: memberIdxs.map(i => people[i].name),
            favorite,
            leastFavorite
        });
    }

    return result;
}

function computeClustersOnce() {
    if (!people.length) return null;

    if (cachedClusterAssignments) {
        return {
            assignments: cachedClusterAssignments,
            k: cachedClusterK
        };
    }

    const vectors = buildNormalizedVectors();
    const k = 3;
    const { assignments } = kMeansCluster(vectors, k);

    cachedClusterAssignments = assignments;
    cachedClusterK = k;

    return { assignments, k };
}

function getPersonClusterLabel(personIdx) {
    const clusterInfo = computeClustersOnce();
    if (!clusterInfo) return null;

    const c = clusterInfo.assignments[personIdx];
    if (c == null) return null;

    return `${c + 1}`;
}


// Render the Clusters tab
function renderClusters() {
    if (!clustersBox || !people.length) return;

    if (people.length < 2) {
        clustersBox.innerHTML = `
          <h2 style="margin-top:0; margin-bottom:0.5rem;">Clusters</h2>
          <p class="description" style="margin-top:0;">
            Not enough people to form clusters.
          </p>
        `;
        return;
    }

    const clusterInfo = computeClustersOnce();
    if (!clusterInfo) return;

    const { assignments, k } = clusterInfo;

    const clusterSummaries = computeClusterLocationExtremes(assignments, k);

    let html = `
      <h2 style="margin-top:0; margin-bottom:0.5rem;">Clusters</h2>
      <p class="description" style="margin-top:0;">
        People are grouped into clusters based on their overall rating profiles
        (using a simple k-means on normalized scores). For each cluster, we show
        the favorite and least favorite locations based on the cluster's average scores.
      </p>
    `;

    html += `<div style="margin-top:0.5rem;">`;

    clusterSummaries.forEach((cl, idx) => {
        const label = `Cluster ${idx + 1}`;
        const members = cl.members || [];

        html += `
          <div style="margin-top:0.75rem;">
            <strong>${label}</strong>
            <ul style="margin-top:0.25rem; padding-left:1.2rem;">
              <li>Members: ${members.length
                ? members.join(", ")
                : "<em>none assigned</em>"
            }</li>
        `;

        if (cl.favorite) {
            html += `
              <li>
                Favorite location:
                ${cl.favorite.name}
                (average ≈ ${cl.favorite.mean.toFixed(2)}).
              </li>
            `;
        } else {
            html += `<li>Favorite location: <em>not enough ratings</em></li>`;
        }

        if (cl.leastFavorite) {
            html += `
              <li>
                Least favorite location:
                ${cl.leastFavorite.name}
                (average ≈ ${cl.leastFavorite.mean.toFixed(2)}).
              </li>
            `;
        } else {
            html += `<li>Least favorite location: <em>not enough ratings</em></li>`;
        }

        html += `
            </ul>
          </div>
        `;
    });

    html += `</div>`;

    clustersBox.innerHTML = html;
}


// --- OVERALL STATS HELPERS ---


// Compute: for a given person index, which location are they most different
// from the rest of the group (by |score - others' mean|)?
function computeMostDifferentLocation(personIdx) {
    const person = people[personIdx];
    if (!person) return null;

    const numLocations = person.scores.length;
    let best = null; // { index, location, diff, selfScore, groupMean, count }

    for (let idx = 0; idx < numLocations; idx++) {
        const v = person.scores[idx];
        if (v === null) continue;

        let sum = 0;
        let count = 0;

        for (let j = 0; j < people.length; j++) {
            if (j === personIdx) continue;
            const w = people[j].scores[idx];
            if (w !== null) {
                sum += w;
                count++;
            }
        }

        if (count === 0) continue;

        const groupMean = sum / count;
        const diff = Math.abs(v - groupMean);

        if (!best || diff > best.diff) {
            best = {
                index: idx,
                location: getLocationName(idx),
                diff,
                selfScore: v,
                groupMean,
                count
            };
        }
    }

    return best;
}

// Render the Personal stats tab for a given person name.
// If no name is provided, defaults to the first person.
function renderPersonalStats(selectedName) {
    if (!personalStatsBox || !people.length) return;

    const defaultName = selectedName || (people[0] && people[0].name);
    const personName = defaultName || "";
    const personIdx = people.findIndex(p => p.name === personName);
    const clusterLabel = getPersonClusterLabel(personIdx);
    const person = people[personIdx];

    if (!person) {
        personalStatsBox.innerHTML = `
          <h2 style="margin-top:0; margin-bottom:0.5rem;">Personal stats</h2>
          <p class="description" style="margin-top:0;">
            No people found in the dataset.
          </p>
        `;
        return;
    }

    // Basic stats
    const stats = computePersonStats(person);
    const nRated = stats ? stats.n : 0;

    // Favorite / least favorite locations for this person
    let fav = null;   // { name, score }
    let least = null; // { name, score }
    for (let i = 0; i < person.scores.length; i++) {
        const v = person.scores[i];
        if (v === null) continue;
        const locName = getLocationName(i);

        if (!fav || v > fav.score) {
            fav = { name: locName, score: v };
        }
        if (!least || v < least.score) {
            least = { name: locName, score: v };
        }
    }

    // Location where they are most different from the rest of the group
    const mostDiff = computeMostDifferentLocation(personIdx);

    // Best / worst match vs others (moved from Compare tab)
    const { best, worst } = computeBestAndWorstFor(person);

    const meanText = stats ? stats.mean.toFixed(2) : "—";
    const stdText = stats ? stats.std.toFixed(2) : "—";
    const minText = stats ? stats.min : "—";
    const maxText = stats ? stats.max : "—";

    // Build HTML
    let html = `
      <h2 style="margin-top:0; margin-bottom:0.5rem;">Personal stats</h2>
      <p class="description" style="margin-top:0;">
        View how one person compares to the rest of the group.
      </p>

      <label for="personal-person-select">Choose a person</label>
      <select id="personal-person-select" style="margin-top:0.25rem;">
    `;

    for (const p of people) {
        const sel = p.name === person.name ? "selected" : "";
        html += `<option value="${p.name}" ${sel}>${p.name}</option>`;
    }

    html += `
      </select>

      <hr/>

      <div style="margin-top:0.6rem;">
        <strong>Summary for ${person.name}</strong>
        <ul style="margin-top:0.4rem; padding-left:1.2rem;">
          <li>Rated ${nRated} locations.</li>
          <li>Average score: ${meanText} (σ ≈ ${stdText}).</li>
          <li>Score range: ${minText} to ${maxText}.</li>
    `;
    if (clusterLabel) {
        html += `<li>Cluster: ${clusterLabel}</li>`;
    }

    if (fav) {
        html += `
          <li>Favorite location: ${fav.name} (score = ${fav.score}).</li>
        `;
    }
    if (least) {
        html += `
          <li>Least favorite location: ${least.name} (score = ${least.score}).</li>
        `;
    }

    html += `</ul>
      </div>
    `;

    // Most different from group
    if (mostDiff) {
        html += `
          <hr/>
          <div style="margin-top:0.6rem;">
            <strong>Where ${person.name} is most different</strong>
            <p style="margin-top:0.4rem; margin-bottom:0.4rem;">
              Location: <em>${mostDiff.location}</em><br/>
              ${person.name}'s score: ${mostDiff.selfScore}<br/>
              Group average (others only): ${mostDiff.groupMean.toFixed(2)} (based on ${mostDiff.count} others)<br/>
              Absolute difference: ${mostDiff.diff.toFixed(2)} points.
            </p>
          </div>
        `;
    }

    // Matches for this person
    html += `
      <hr/>
      <div style="margin-top:0.6rem;">
        <strong>Matches for ${person.name}</strong>
        <div style="margin-top:0.4rem;">
    `;

    function matchLine(label, m) {
        if (!m) {
            return `${label}: <span class="pill pill-neutral">no valid match</span><br/>`;
        }
        const cls = corrClassFromR(m.r);
        const pillClass =
            cls === "positive" ? "pill-positive" :
                cls === "negative" ? "pill-negative" :
                    "pill-neutral";
        return `${label}: <span class="pill ${pillClass}">${m.name} (${m.r.toFixed(3)}, ${m.overlap} locations)</span><br/>`;
    }

    html += matchLine("Best match", best);
    html += matchLine("Worst match", worst);

    html += `
        </div>
      </div>
    `;

    personalStatsBox.innerHTML = html;

    // Wire up select change → re-render for chosen person
    const selEl = personalStatsBox.querySelector("#personal-person-select");
    if (selEl) {
        selEl.addEventListener("change", (e) => {
            const newName = e.target.value;
            renderPersonalStats(newName);
        });
    }
}


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
// Signed squared sum = Σ (score * |score|) over everyone for each location
function computeLocationSignedSquaredSums() {
    if (!people.length) return null;
    const numLocations = people[0].scores.length;

    const locStats = [];

    for (let idx = 0; idx < numLocations; idx++) {
        let signedSquaredSum = 0;
        let count = 0;
        let sum = 0;
        let sumSq = 0;

        for (const p of people) {
            const v = p.scores[idx];
            if (v !== null) {
                signedSquaredSum += v * Math.abs(v);
                count++;
                sum += v;
                sumSq += v * v;
            }
        }

        if (count > 0) {
            let consensus = null; // we'll use σ as a "consensus" proxy: lower = more agreement
            if (count > 1) {
                const mean = sum / count;
                const variance = sumSq / count - mean * mean;
                consensus = Math.sqrt(Math.max(variance, 0));
            }

            locStats.push({
                index: idx,
                name: getLocationName(idx),
                signedSquaredSum,
                count,
                consensus,   // standard deviation of scores for this location
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

function computeCorrelationMatrix(order = matrixOrder) {
    const n = order.length;
    const matrix = [];

    for (let i = 0; i < n; i++) {
        const idxI = order[i];
        matrix[i] = [];
        for (let j = 0; j < n; j++) {
            const idxJ = order[j];

            if (idxI === idxJ) {
                // diagonal
                matrix[i][j] = {
                    r: 1.0,
                    overlap: people[idxI].scores.length
                };
            } else {
                const { xs, ys } = buildOverlap(people[idxI], people[idxJ]);
                if (xs.length < 2) {
                    matrix[i][j] = { r: null, overlap: xs.length };
                } else {
                    const r = pearsonCorrelation(xs, ys);
                    matrix[i][j] = {
                        r: Number.isFinite(r) ? r : null,
                        overlap: xs.length
                    };
                }
            }
        }
    }
    return matrix;
}

function sortMatrixByPerson(personIdx) {
    const n = people.length;
    const corrs = [];

    for (let j = 0; j < n; j++) {
        if (j === personIdx) {
            corrs.push({ idx: j, r: 1.0 });
            continue;
        }

        const { xs, ys } = buildOverlap(people[personIdx], people[j]);
        if (xs.length < 2) {
            corrs.push({ idx: j, r: null });
        } else {
            const r = pearsonCorrelation(xs, ys);
            corrs.push({
                idx: j,
                r: Number.isFinite(r) ? r : null
            });
        }
    }

    // Sort: clicked person first, then others by descending correlation,
    // with "no data" (null) at the end.
    corrs.sort((a, b) => {
        if (a.idx === personIdx && b.idx === personIdx) return 0;
        if (a.idx === personIdx) return -1;
        if (b.idx === personIdx) return 1;

        const ra = a.r;
        const rb = b.r;

        if (ra === null && rb === null) return 0;
        if (ra === null) return 1;
        if (rb === null) return -1;

        return rb - ra; // descending by r
    });

    matrixOrder = corrs.map(c => c.idx);
    renderCorrelationMatrix();
}

function attachMatrixSortHandlers() {
    if (!matrixBox) return;
    const headers = matrixBox.querySelectorAll("thead th[data-person-idx]");

    headers.forEach(th => {
        const idx = Number(th.getAttribute("data-person-idx"));
        if (Number.isNaN(idx)) return;

        th.style.cursor = "pointer";
        const existingTitle = th.getAttribute("title") || "";
        const hint = existingTitle ? existingTitle + " – " : "";
        th.setAttribute("title", hint + "click header to sort by this person");

        // Click on the header background / name → sort by that person
        th.addEventListener("click", (e) => {
            // If the click came from a reorder button, ignore here
            if (e.target.closest(".reorder-btn")) return;
            sortMatrixByPerson(idx);
        });

        // Wire the tiny left/right buttons for manual reordering
        const leftBtn = th.querySelector(".reorder-left");
        const rightBtn = th.querySelector(".reorder-right");

        if (leftBtn) {
            leftBtn.addEventListener("click", (e) => {
                e.stopPropagation(); // don’t also trigger sort
                movePersonInOrder(idx, -1);
            });
        }

        if (rightBtn) {
            rightBtn.addEventListener("click", (e) => {
                e.stopPropagation();
                movePersonInOrder(idx, +1);
            });
        }
    });
}


function movePersonInOrder(personIdx, direction) {
    // direction: -1 = move left, +1 = move right
    const currPos = matrixOrder.indexOf(personIdx);
    if (currPos === -1) return;

    const newPos = currPos + direction;
    if (newPos < 0 || newPos >= matrixOrder.length) return;

    const tmp = matrixOrder[currPos];
    matrixOrder[currPos] = matrixOrder[newPos];
    matrixOrder[newPos] = tmp;

    // Re-render with updated ordering
    renderCorrelationMatrix();
}


function corrToBgColor(r) {
    if (r === null || !Number.isFinite(r)) return "transparent";

    const alpha = Math.min(Math.abs(r), 1) * 0.35; // light intensity

    if (r > 0) {
        // light green
        return `rgba(34, 197, 94, ${alpha})`;
    } else {
        // light red
        return `rgba(239, 68, 68, ${alpha})`;
    }
}

function renderCorrelationMatrix() {
    if (!matrixBox || !people.length) return;

    const order = (matrixOrder.length === people.length)
        ? matrixOrder
        : people.map((_, idx) => idx);

    const matrix = computeCorrelationMatrix(order);

    let html = `
      <h2 style="margin-top:0; margin-bottom:0.5rem;">Correlation matrix</h2>
      <p class="description" style="margin-top:0;">
        Pairwise Pearson correlations across all shared locations.
        <br/>
        Click a header to sort by that person, or click a cell to jump to the comparison view.
      </p>

      <div style="overflow-x:auto;">
      <table style="
        border-collapse: collapse;
        font-size: 0.85rem;
        min-width: 600px;
      ">
        <thead>
          <tr>
            <th style="padding:0.35rem; border-bottom:1px solid var(--card-border);"></th>
    `;

    // Column headers (clickable for sorting + left/right reordering)
    for (let c = 0; c < order.length; c++) {
        const pIndex = order[c];
        const p = people[pIndex];
        html += `
        <th
            data-person-idx="${pIndex}"
            style="
                padding:0.25rem 0.35rem;
                border-bottom:1px solid var(--card-border);
                text-align:center;
                white-space:nowrap;
            "
        >
            <div style="display:flex; align-items:center; justify-content:center; gap:0.25rem;">
                <button
                    type="button"
                    class="reorder-btn reorder-left"
                    data-person-idx="${pIndex}"
                    style="
                        border:none;
                        background:transparent;
                        color:var(--muted-soft);
                        padding:0;
                        cursor:pointer;
                        font-size:0.7rem;
                    "
                    title="Move left"
                >
                    ◀
                </button>
                <span>${p.name}</span>
                <button
                    type="button"
                    class="reorder-btn reorder-right"
                    data-person-idx="${pIndex}"
                    style="
                        border:none;
                        background:transparent;
                        color:var(--muted-soft);
                        padding:0;
                        cursor:pointer;
                        font-size:0.7rem;
                    "
                    title="Move right"
                >
                    ▶
                </button>
            </div>
        </th>
        `;
    }

    html += `</tr></thead><tbody>`;

    // Rows
    for (let i = 0; i < order.length; i++) {
        const rowIdx = order[i];
        const rowPerson = people[rowIdx];

        html += `
          <tr>
            <th style="padding:0.35rem; border-right:1px solid var(--card-border); text-align:right;">
              ${rowPerson.name}
            </th>
        `;

        for (let j = 0; j < order.length; j++) {
            const cell = matrix[i][j];

            let content = "—";
            let bg = "transparent";

            if (cell.r !== null) {
                content = cell.r.toFixed(2);
                bg = corrToBgColor(cell.r);
            }

            html += `
                <td
                    class="corr-cell"
                    data-row-pos="${i}"
                    data-col-pos="${j}"
                    title="${cell.overlap} shared locations"
                    style="
                        padding: 0.35rem;
                        text-align: center;
                        background-color: ${bg};
                        border: 1px solid var(--card-border);
                    "
                >
                    ${content}
                </td>
            `;
        }

        html += `</tr>`;
    }

    html += `
        </tbody>
      </table>
      </div>
    `;

    matrixBox.innerHTML = html;

    // Make headers sortable & re-orderable
    attachMatrixSortHandlers();
    // NEW: make cells interactive (hover highlight + Matrix → Compare jump)
    attachMatrixHoverHandlers();
    attachMatrixCellHandlers();
}

// --- Hover highlight helpers ---

function clearMatrixHover() {
    if (!matrixBox) return;
    matrixBox
        .querySelectorAll(".matrix-row-hover, .matrix-col-hover")
        .forEach(el => {
            el.classList.remove("matrix-row-hover", "matrix-col-hover");
        });
}

function highlightMatrixHover(rowPos, colPos) {
    if (!matrixBox) return;
    clearMatrixHover();

    const rows = matrixBox.querySelectorAll("tbody tr");
    rows.forEach((tr, i) => {
        if (i === rowPos) {
            tr.classList.add("matrix-row-hover");
        }
        const tds = tr.querySelectorAll("td.corr-cell");
        tds.forEach((td, j) => {
            if (j === colPos) {
                td.classList.add("matrix-col-hover");
            }
        });
    });

    // Highlight corresponding column header
    const headerRow = matrixBox.querySelector("thead tr");
    if (headerRow) {
        const headerCells = headerRow.querySelectorAll("th[data-person-idx]");
        if (headerCells[colPos]) {
            headerCells[colPos].classList.add("matrix-col-hover");
        }
    }
}

function attachMatrixHoverHandlers() {
    if (!matrixBox) return;
    const cells = matrixBox.querySelectorAll("td.corr-cell");

    cells.forEach(cell => {
        const rowPos = Number(cell.getAttribute("data-row-pos"));
        const colPos = Number(cell.getAttribute("data-col-pos"));

        cell.addEventListener("mouseenter", () => {
            if (!Number.isNaN(rowPos) && !Number.isNaN(colPos)) {
                highlightMatrixHover(rowPos, colPos);
            }
        });

        cell.addEventListener("mouseleave", () => {
            clearMatrixHover();
        });
    });
}

// --- Matrix → Compare jump (click a cell to go to Compare tab) ---

function attachMatrixCellHandlers() {
    if (!matrixBox) return;
    const cells = matrixBox.querySelectorAll("td.corr-cell");

    cells.forEach(cell => {
        cell.addEventListener("click", () => {
            const rowPos = Number(cell.getAttribute("data-row-pos"));
            const colPos = Number(cell.getAttribute("data-col-pos"));
            if (
                Number.isNaN(rowPos) ||
                Number.isNaN(colPos) ||
                rowPos === colPos ||
                !matrixOrder.length
            ) {
                return;
            }

            // Map matrix positions → actual people indices
            const idxA = matrixOrder[colPos];
            const idxB = matrixOrder[rowPos];
            const personA = people[idxA];
            const personB = people[idxB];
            if (!personA || !personB) return;

            // Set dropdowns in Compare UI
            personASelect.value = personA.name;
            personBSelect.value = personB.name;

            // Jump to Compare tab and immediately run the comparison
            showTab("compare");
            computeCorrelation();
        });
    });
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
            (σ ≈ ${mostPolar.std.toFixed(2)}).
          </li>
        `;
        }
        if (leastPolar) {
            html += `
          <li>
            Least polarizing location:
            ${leastPolar.name} &mdash; people are most in agreement here
            (σ ≈ ${leastPolar.std.toFixed(2)}).
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
            ${highLoc.name} (${highLoc.signedSquaredSum.toFixed(2)}).
          </li>
          <li>
            Lowest-scoring location:
            ${lowLoc.name} (${lowLoc.signedSquaredSum.toFixed(2)}).
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
                    <th style="text-align:right; padding:0.35rem; border-bottom:1px solid var(--card-border);">
                    Consensus (σ; lower = more agreement)
                    </th>
                </tr>
                </thead>
                <tbody>
        `;
        for (let i = 0; i < ranking.length; i++) {
            const loc = ranking[i];
            const consensusText =
                loc.consensus != null ? loc.consensus.toFixed(2) : "—";
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
                    <td style="padding:0.35rem; text-align:right; border-bottom:1px solid var(--card-border);">
                      ${consensusText}
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

            // NEW: reset matrix ordering to the natural order
            matrixOrder = people.map((_, idx) => idx);
            matrixRendered = false;

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
    let topSameScore = null;      // { score, location }
    let maxDisagreement = null;   // { diff, location }

    // For explanation stats
    let sumAbsDiff = 0;
    let sumDiff = 0;

    // For agreement strength breakdown
    let exactSameCount = 0;
    let offBy1Count = 0;
    let offBy2Count = 0;
    let diff3PlusCount = 0;

    // For representative / compromise logic – keep metadata aligned with xs/ys
    const overlapDetails = [];    // [{ index, a, b, location }]

    // First pass: gather overlaps & simple stats
    for (let i = 0; i < personA.scores.length; i++) {
        const a = personA.scores[i];
        const b = personB.scores[i];
        if (a === null || b === null) continue;

        const locationName = getLocationName(i);

        xs.push(a);
        ys.push(b);
        overlapDetails.push({ index: i, a, b, location: locationName });

        const absDiff = Math.abs(a - b);

        // Track largest-magnitude same score
        if (a === b) {
            if (!topSameScore || Math.abs(a) > Math.abs(topSameScore.score)) {
                topSameScore = {
                    score: a,
                    location: locationName,
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
                    location: locationName,
                };
            }
        }

        // For explanation bullets
        sumAbsDiff += absDiff;
        sumDiff += (a - b); // positive => A rates higher

        // Agreement-strength buckets
        if (absDiff === 0) {
            exactSameCount++;
        } else if (absDiff === 1) {
            offBy1Count++;
        } else if (absDiff === 2) {
            offBy2Count++;
        } else if (absDiff >= 3) {
            diff3PlusCount++;
        }
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

    // For "Most representative location" (largest contribution to the correlation)
    let mostRep = null; // { location, a, b, contribution }

    for (let i = 0; i < n; i++) {
        const da = xs[i] - meanA;
        const db = ys[i] - meanB;
        const prod = da * db;

        if (prod > 0) sameDirCount++;
        else if (prod < 0) oppDirCount++;

        // Signed contribution aligned with overall r
        const signedContribution = r >= 0 ? prod : -prod;
        const meta = overlapDetails[i];

        if (!mostRep || signedContribution > mostRep.contribution) {
            mostRep = {
                contribution: signedContribution,
                location: meta?.location ?? `Location ${i + 1}`,
                a: meta?.a ?? xs[i],
                b: meta?.b ?? ys[i],
            };
        }
    }

    const dirPairs = sameDirCount + oppDirCount;
    let movePercent = 0;
    let moveWord = r >= 0 ? "the same direction" : "opposite directions";

    if (dirPairs > 0) {
        const base = r >= 0 ? sameDirCount : oppDirCount;
        movePercent = Math.round(100 * (base / dirPairs));
    }

    const avgAbsDiff = sumAbsDiff / n;
    const avgDiff = sumDiff / n; // >0 => A higher, <0 => B higher

    // --- Polarization index (share of locations where you move in opposite directions) ---
    const oppShare = dirPairs > 0 ? (oppDirCount / dirPairs) : 0;
    const polarizationPct = Math.round(oppShare * 100);
    let polarizationLabel = "low";
    if (oppShare >= 0.5) polarizationLabel = "high";
    else if (oppShare >= 0.25) polarizationLabel = "medium";

    // --- Agreement-strength percentages ---
    const exactPct = Math.round((exactSameCount / n) * 100);
    const offBy1Pct = Math.round((offBy1Count / n) * 100);
    const offBy2Pct = Math.round((offBy2Count / n) * 100);
    const diff3PlusPct = Math.round((diff3PlusCount / n) * 100);

    // --- "If you had to compromise…" recommendation ---
    let compromise = null; // { location, a, b, absDiff, avg }
    for (const meta of overlapDetails) {
        const absDiff = Math.abs(meta.a - meta.b);
        const avgScore = (meta.a + meta.b) / 2;

        if (
            !compromise ||
            absDiff < compromise.absDiff - 1e-9 ||
            (Math.abs(absDiff - compromise.absDiff) < 1e-9 && avgScore > compromise.avg)
        ) {
            compromise = {
                location: meta.location,
                a: meta.a,
                b: meta.b,
                absDiff,
                avg: avgScore,
            };
        }
    }

    // --- Build Trends (including new items) ---
    let trendsHtml = "";

    if (n > 0) {
        trendsHtml += `<hr/><div style="margin-top:0.75rem;"><strong>Trends</strong><br/>`;

        // Largest-magnitude same score
        if (topSameScore) {
            trendsHtml += `
        <div>You and ${nameB} both put the score ${topSameScore.score} for ${topSameScore.location}.</div>
      `;
        }

        // Largest disagreement
        if (maxDisagreement) {
            trendsHtml += `
        <div style="margin-top:0.5rem;">
          You and ${nameB} disagree the most about ${maxDisagreement.location}.
        </div>
      `;
        }

        // Polarization index
        trendsHtml += `
      <div style="margin-top:0.5rem;">
        <strong>Polarization index:</strong> ${polarizationPct}% (${polarizationLabel} – share of locations where you move in opposite directions relative to your usual scores).
      </div>
    `;

        // Agreement strength breakdown
        trendsHtml += `
      <div style="margin-top:0.5rem;">
        <strong>Agreement strength breakdown</strong>
        <ul style="margin-top:0.25rem; padding-left:1.2rem;">
          <li>${exactPct}% of locations: exact same score</li>
          <li>${offBy1Pct}%: differ by 1 point</li>
          <li>${offBy2Pct}%: differ by 2 points</li>
          <li>${diff3PlusPct}%: differ by 3+ points</li>
        </ul>
      </div>
    `;

        // Most representative location
        if (mostRep) {
            trendsHtml += `
        <div style="margin-top:0.5rem;">
          <strong>Most representative location:</strong> ${mostRep.location}
          (you rated it ${mostRep.a}, and ${nameB} rated it ${mostRep.b}, contributing strongly to this ${signWord} correlation).
        </div>
      `;
        }

        // If you had to compromise…
        if (compromise) {
            trendsHtml += `
        <div style="margin-top:0.5rem;">
          <strong>If you had to compromise…</strong>
          A fair pick might be <em>${compromise.location}</em>, where your scores are closest
          (average ≈ ${compromise.avg.toFixed(2)} across both of you).
        </div>
      `;
        }

        trendsHtml += `</div>`;
    }

    // --- Explanation block with bullets (unchanged structure, but now coexists with new trends) ---
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
    Together, your scoring exhibits <strong>${qualitative}</strong> similarity and your preferences are mostly <strong>${direction}</strong>.
    ${explanationHtml}
    ${trendsHtml}
  `;
    resultBox.style.display = "block";
    computeStatus.textContent = "";
}

computeBtn.addEventListener("click", computeCorrelation);

// Init
initTheme();
loadSheet();
