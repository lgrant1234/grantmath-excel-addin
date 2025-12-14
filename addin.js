/**
 * GrantMath Excel Add-in
 * PhD-level computational reasoning for Excel
 * © 2025 SolonAI
 */

// API Configuration
const GRANTMATH_API_URL = "https://api.grantmath.com/nlp-analysis";
const LOCAL_DEV_URL = "http://localhost:8000/nlp-analysis"; // For development

// Use local API for development, production API for deployed version
const API_URL = window.location.hostname === "localhost" ? LOCAL_DEV_URL : GRANTMATH_API_URL;

// Initialize Office
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("GrantMath Excel Add-in loaded");

        // Set up event listeners
        document.getElementById("save-key-btn").onclick = saveApiKey;
        document.getElementById("solve-btn").onclick = solveCell;
        document.getElementById("api-key").addEventListener("input", updateKeyIndicator);

        // Load saved API key
        loadApiKey();

        // Enable solve button if API key exists
        updateSolveButtonState();
    }
});

/**
 * Load API key from localStorage
 */
function loadApiKey() {
    const apiKey = localStorage.getItem("grantmath_api_key");
    if (apiKey) {
        document.getElementById("api-key").value = apiKey;
        updateKeyIndicator();
    }
}

/**
 * Save API key to localStorage
 */
function saveApiKey() {
    const apiKey = document.getElementById("api-key").value.trim();

    if (!apiKey) {
        showStatus("Please enter an API key", "error");
        return;
    }

    localStorage.setItem("grantmath_api_key", apiKey);
    updateKeyIndicator();
    updateSolveButtonState();
    showStatus("API key saved successfully", "success");

    // Clear success message after 3 seconds
    setTimeout(() => hideStatus(), 3000);
}

/**
 * Update key indicator visual state
 */
function updateKeyIndicator() {
    const apiKey = document.getElementById("api-key").value.trim();
    const indicator = document.getElementById("key-indicator");

    if (apiKey && apiKey.length > 0) {
        indicator.classList.add("active");
    } else {
        indicator.classList.remove("active");
    }

    updateSolveButtonState();
}

/**
 * Enable/disable solve button based on API key presence
 */
function updateSolveButtonState() {
    const solveBtn = document.getElementById("solve-btn");
    // Always enable the solve button (API is currently public)
    solveBtn.disabled = false;
}

/**
 * Solve mathematical question from selected cell
 */
async function solveCell() {
    const apiKey = localStorage.getItem("grantmath_api_key") || "";

    try {
        await Excel.run(async (context) => {
            // Get selected range
            const range = context.workbook.getSelectedRange();
            range.load("values, address");
            await context.sync();

            // Get question from selected cell
            const question = range.values[0][0];

            if (!question || question.toString().trim() === "") {
                showStatus("Selected cell is empty. Please select a cell with a question.", "error");
                return;
            }

            // Show loading state
            showStatus("Computing solution...", "loading");
            document.getElementById("solve-btn").disabled = true;

            // Build headers
            const headers = {
                "Content-Type": "application/json"
            };

            // Add authorization if API key exists
            if (apiKey) {
                headers["Authorization"] = `Bearer ${apiKey}`;
            }

            // Call GrantMath API
            const response = await fetch(API_URL, {
                method: "POST",
                headers: headers,
                body: JSON.stringify({
                    question: question.toString()
                })
            });

            if (!response.ok) {
                if (response.status === 401) {
                    throw new Error("Invalid API key. Please check your API key.");
                } else if (response.status === 429) {
                    throw new Error("Rate limit exceeded. Please wait or upgrade to Pro tier.");
                } else if (response.status === 400) {
                    const errorData = await response.json();
                    throw new Error(errorData.error || "Invalid question format");
                } else {
                    throw new Error(`API error: ${response.status} ${response.statusText}`);
                }
            }

            const result = await response.json();

            // Check if request was successful
            if (!result.success) {
                const error = result.error || "Unknown error occurred";
                showStatus(error, "error");
                document.getElementById("solve-btn").disabled = false;
                return;
            }

            // Get the answer (strip HTML for Excel)
            const answer = stripHtmlTags(result.analysis || result.formatted_answer || "No answer returned");

            // Write result to cell to the right of selected cell
            const outputCell = range.getOffsetRange(0, 1);
            outputCell.values = [[answer]];

            // Format the output cell
            outputCell.format.font.color = "#166534";
            outputCell.format.wrapText = true;
            outputCell.format.verticalAlignment = "Top";

            await context.sync();

            // Show success
            showStatus("Solution computed successfully!", "success");

            // Re-enable button and clear success message after 3 seconds
            setTimeout(() => {
                document.getElementById("solve-btn").disabled = false;
                hideStatus();
            }, 3000);
        });

    } catch (error) {
        console.error("Error solving question:", error);
        const errorMessage = error.message || "Failed to compute solution. Please contact support@solonai.com";
        showStatus(errorMessage, "error");
        document.getElementById("solve-btn").disabled = false;
    }
}

/**
 * Strip HTML tags from text (for Excel display)
 */
function stripHtmlTags(html) {
    if (!html) return "";

    // Remove script and style tags
    let text = html.replace(/<script[^>]*>.*?<\/script>/gi, "");
    text = text.replace(/<style[^>]*>.*?<\/style>/gi, "");

    // Convert headers to text with separators (more compact)
    text = text.replace(/<h[1-6][^>]*>(.*?)<\/h[1-6]>/gi, "\n$1\n" + "=".repeat(40) + "\n");

    // Convert paragraphs and divs to single newlines
    text = text.replace(/<\/?p[^>]*>/gi, "\n");
    text = text.replace(/<\/?div[^>]*>/gi, "\n");

    // Convert line breaks
    text = text.replace(/<br\s*\/?>/gi, "\n");

    // Convert bold/strong to uppercase markers
    text = text.replace(/<strong[^>]*>(.*?)<\/strong>/gi, "**$1**");
    text = text.replace(/<b[^>]*>(.*?)<\/b>/gi, "**$1**");

    // Remove all remaining HTML tags
    text = text.replace(/<[^>]+>/g, "");

    // Decode HTML entities
    const textarea = document.createElement("textarea");
    textarea.innerHTML = text;
    text = textarea.value;

    // Clean up whitespace (more aggressive)
    text = text.replace(/ +/g, " "); // Multiple spaces to single space
    text = text.replace(/\n\n+/g, "\n"); // Multiple newlines to single newline
    text = text.split("\n").map(line => line.trim()).filter(line => line.length > 0).join("\n"); // Trim and remove empty lines

    // Add blank line before disclaimer for visual separation
    text = text.replace(/(\*\*Disclaimer:\*\*)/g, "\n$1");

    return text.trim();
}

/**
 * Show status message
 */
function showStatus(message, type) {
    const statusEl = document.getElementById("status");
    statusEl.textContent = message;
    statusEl.className = `status ${type}`;
    statusEl.style.display = "block";
}

/**
 * Hide status message
 */
function hideStatus() {
    const statusEl = document.getElementById("status");
    statusEl.style.display = "none";
}

