/* global Word, Office */
/* Version 1.1 - Fixed code counting and removed generatedCount */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("generateBtn").onclick = generateUniqueCode;
        document.getElementById("insertBtn").onclick = insertCode;
    }
});

let currentCode = null;

/**
 * Generate a random code in format fc-XXX-XXX
 */
function generateCode() {
    const part1 = Math.floor(Math.random() * 900) + 100; // 100-999
    const part2 = Math.floor(Math.random() * 900) + 100; // 100-999
    return `fc-${part1}-${part2}`;
}

/**
 * Search the entire document for existing codes
 */
async function getAllCodesInDocument() {
    return await Word.run(async (context) => {
        const body = context.document.body;
        // Word wildcard pattern for fc-XXX-XXX: fc- followed by 3 digits, hyphen, 3 digits
        const searchResults = body.search("fc-[0-9]{3}-[0-9]{3}", { matchWildcards: true });
        
        searchResults.load("text");
        await context.sync();
        
        const codes = new Set();
        for (let i = 0; i < searchResults.items.length; i++) {
            codes.add(searchResults.items[i].text);
        }
        
        return codes;
    });
}

/**
 * Generate a unique code that doesn't exist in the document
 */
async function generateUniqueCode() {
    try {
        // showStatus("Generating unique code...", "info");
        disableButtons(true);
        
        // Get all existing codes
        const existingCodes = await getAllCodesInDocument();
        
        // Maximum possible unique codes: 900 * 900 = 810,000
        const MAX_UNIQUE_CODES = 810000;
        
        // Check if we've reached the maximum
        if (existingCodes.size >= MAX_UNIQUE_CODES) {
            throw new Error("You have reached the maximum quantity of unique numbers.");
        }
        
        // Warn if getting close to maximum (within 1000 codes)
        if (existingCodes.size >= MAX_UNIQUE_CODES - 1000) {
            console.warn(`Warning: ${existingCodes.size} codes used out of ${MAX_UNIQUE_CODES} possible`);
        }
        
        // Generate a new code
        let newCode;
        let attempts = 0;
        const maxAttempts = 10000; // Increased from 1000
        
        do {
            newCode = generateCode();
            attempts++;
            
            if (attempts > maxAttempts) {
                throw new Error("Unable to generate unique code. You may have reached the maximum quantity of unique numbers.");
            }
        } while (existingCodes.has(newCode));
        
        // Display the code
        currentCode = newCode;
        document.getElementById("codeDisplay").textContent = newCode;
        document.getElementById("insertBtn").disabled = false;
        
        // showStatus(`✓ Generated unique code: ${newCode}`, "success");
        disableButtons(false);
        
    } catch (error) {
        showStatus(`Error: ${error.message}`, "error");
        disableButtons(false);
        console.error(error);
    }
}

/**
 * Insert the generated code at the cursor position with existing formatting
 */
async function insertCode() {
    if (!currentCode) {
        showStatus("Please generate a code first", "error");
        return;
    }
    
    try {
        // showStatus("Inserting code...", "info");
        disableButtons(true);
        
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            
            // Insert the code
            const insertedRange = selection.insertText(currentCode, Word.InsertLocation.end);
            
            // Apply fixed formatting: Arial, 8pt, light gray, plain style
            insertedRange.font.name = "Arial";
            insertedRange.font.size = 8;
            insertedRange.font.color = "#BFBFBF"; // Light gray color
            insertedRange.font.bold = false; // Ensure not bold
            insertedRange.font.italic = false; // Ensure not italic
            insertedRange.font.underline = "None"; // Ensure not underlined
            insertedRange.font.strikeThrough = false; // Ensure not strikethrough
            insertedRange.font.superscript = false; // Ensure not superscript
            insertedRange.font.subscript = false; // Ensure not subscript
            
            // Don't move cursor - prevents page jumping
            // insertedRange.select(Word.SelectionMode.end);
            
            await context.sync();
        });
        
        // showStatus(`✓ Code ${currentCode} inserted successfully!`, "success");
        
        // Reset for next generation
        currentCode = null;
        document.getElementById("codeDisplay").textContent = "fc-000-000";
        document.getElementById("insertBtn").disabled = true;
        
        disableButtons(false);
        
    } catch (error) {
        // Only show error if the insertion actually failed
        if (!error.message.includes("InvalidArgument")) {
            showStatus(`Error inserting code: ${error.message}`, "error");
        }
        disableButtons(false);
        console.error(error);
    }
}

/**
 * Show status message
 */
function showStatus(message, type) {
    const statusDiv = document.getElementById("status");
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
    
    if (type === "success") {
        setTimeout(() => {
            statusDiv.style.display = "none";
        }, 3000);
    }
}

/**
 * Disable/enable buttons during operations
 */
function disableButtons(disabled) {
    document.getElementById("generateBtn").disabled = disabled;
    
    // Only disable insert button if we're disabling, or if there's no current code
    if (disabled) {
        document.getElementById("insertBtn").disabled = true;
    } else if (currentCode) {
        document.getElementById("insertBtn").disabled = false;
    }
}
