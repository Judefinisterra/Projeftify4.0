/**
 * CodeCollection.js
 * Functions for processing and managing code collections
 */

/**
 * Parses code strings and creates a code collection
 * @param {string} inputText - The input text containing code strings
 * @returns {Array} - An array of code objects with type and parameters
 */
export function populateCodeCollection(inputText) {
    try {
        console.log("Processing input text for code collection");
        
        // Initialize an empty code collection
        const codeCollection = [];
        
        // Split the input text by newlines to process each line
        const lines = inputText.split('\n');
        
        for (const line of lines) {
            // Skip empty lines
            if (!line.trim()) continue;
            
            // Extract the code type and parameters
            const codeMatch = line.match(/<([^;>]+);(.*?)>/);
            if (!codeMatch) continue;
            
            const codeType = codeMatch[1].trim();
            const paramsString = codeMatch[2].trim();
            
            // Parse parameters
            const params = {};
            const paramMatches = paramsString.matchAll(/(\w+)\s*=\s*"([^"]*)"/g);
            
            for (const match of paramMatches) {
                const paramName = match[1].trim();
                const paramValue = match[2].trim();
                params[paramName] = paramValue;
            }
            
            // Add the code to the collection
            codeCollection.push({
                type: codeType,
                params: params
            });
        }
        
        console.log(`Processed ${codeCollection.length} codes`);
        return codeCollection;
    } catch (error) {
        console.error("Error in populateCodeCollection:", error);
        throw error;
    }
}

/**
 * Exports a code collection to text format
 * @param {Array} codeCollection - The code collection to export
 * @returns {string} - A formatted text representation of the code collection
 */
export function exportCodeCollectionToText(codeCollection) {
    try {
        if (!codeCollection || !Array.isArray(codeCollection)) {
            throw new Error("Invalid code collection");
        }
        
        let result = "Code Collection:\n";
        result += "================\n\n";
        
        codeCollection.forEach((code, index) => {
            result += `Code ${index + 1}: ${code.type}\n`;
            result += "Parameters:\n";
            
            for (const [key, value] of Object.entries(code.params)) {
                result += `  ${key}: ${value}\n`;
            }
            
            result += "\n";
        });
        
        return result;
    } catch (error) {
        console.error("Error in exportCodeCollectionToText:", error);
        throw error;
    }
} 

/**
 * Processes a code collection and performs operations based on code types
 * @param {Array} codeCollection - The code collection to process
 * @returns {Object} - Results of processing the code collection
 */
export async function runCodes(codeCollection) {
    try {
        console.log("Running code collection processing");
        
        if (!codeCollection || !Array.isArray(codeCollection)) {
            throw new Error("Invalid code collection");
        }
        
        // Initialize result object
        const result = {
            processedCodes: 0,
            createdTabs: [],
            errors: []
        };
        
        // Initialize state variables (similar to VBA variables)
        let currentWorksheetName = null;
        const assumptionTabs = [];
        
        // Process each code in the collection
        for (let i = 0; i < codeCollection.length; i++) {
            const code = codeCollection[i];
            const codeType = code.type;
            
            try {
                // Handle MODEL code type
                if (codeType === "MODEL") {
                    // Skip for now as mentioned in the original VBA code
                    console.log("MODEL code type encountered - skipping for now");
                    continue;
                }
                
                // Handle TAB code type
                if (codeType === "TAB") {
                    // Accept both label1 and Label1 for backward compatibility
                    const tabName = code.params.label1 || code.params.Label1 || `Tab_${i}`;
                    
                    // Check if worksheet exists and delete it
                    await Excel.run(async (context) => {
                        try {
                            // Get all worksheets
                            const sheets = context.workbook.worksheets;
                            sheets.load("items/name");
                            console.log("sheets", sheets);
                            await context.sync();
                            
                            // Check if worksheet exists
                            const existingSheet = sheets.items.find(sheet => sheet.name === tabName);
                            console.log("existingSheet", existingSheet);
                            if (existingSheet) {
                                // Delete the worksheet if it exists
                                existingSheet.delete();
                                await context.sync();
                            }
                            console.log("existingSheet deleted");
                            
                            // Get the Calcs worksheet
                            const sourceCalcsWS = context.workbook.worksheets.getItem("Calcs");
                            console.log("sourceCalcsWS", sourceCalcsWS);
                            
                            // Create a new worksheet by copying the Calcs worksheet
                            const newSheet = sourceCalcsWS.copy();
                            console.log("newSheet created by copying Calcs worksheet");
                            
                            // Rename it
                            newSheet.name = tabName;
                            console.log("newSheet renamed to", tabName);
                            
                            // Set the first row
                            const firstRow = 9; // Equivalent to calcsfirstrow in VBA
                            console.log("firstRow", firstRow);
                            
                            // Clear all cells including and below the first row
                            const usedRange = newSheet.getUsedRange();
                            usedRange.load("rowCount");
                            await context.sync();
                            
                            if (usedRange.rowCount >= firstRow) {
                                const clearRange = newSheet.getRange(`${firstRow}:${usedRange.rowCount}`);
                                clearRange.clear();
                                console.log(`Cleared all cells from row ${firstRow} to ${usedRange.rowCount}`);
                            }
                            
                            // Add to assumption tabs collection
                            assumptionTabs.push({
                                name: tabName,
                                worksheet: newSheet
                            });
                            
                            // Set the current worksheet name
                            currentWorksheetName = tabName;
                            
                            await context.sync();
                            
                            result.createdTabs.push(tabName);
                            console.log("Tab created successfully:", tabName);
                        } catch (error) {
                            console.error("Detailed error in TAB processing:", error);
                            throw error;
                        }
                    }).catch(error => {
                        console.error(`Error processing TAB code: ${error.message}`);
                        result.errors.push({
                            codeIndex: i,
                            codeType: codeType,
                            error: error.message
                        });
                    });
                    
                    continue;
                }
                
                // Handle non-TAB codes
                if (codeType !== "TAB") {
                    await Excel.run(async (context) => {
                        try {
                            // Get the Codes worksheet
                            const codesWS = context.workbook.worksheets.getItem("Codes");
                            console.log("Got Codes worksheet");
                            
                            // Get the used range of the Codes worksheet
                            const usedRange = codesWS.getUsedRange();
                            usedRange.load("rowCount");
                            usedRange.load("columnCount");
                            await context.sync();
                            console.log(`Used range: ${usedRange.rowCount} rows x ${usedRange.columnCount} columns`);
                            
                            // Get the current worksheet
                            const currentWS = context.workbook.worksheets.getItem(currentWorksheetName);
                            console.log("Got current worksheet:", currentWorksheetName);
                            
                            // Get the last row in the current worksheet
                            const lastUsedRow = currentWS.getUsedRange().getLastRow();
                            lastUsedRow.load("rowIndex");
                            await context.sync();
                            const pasteRow = lastUsedRow.rowIndex + 1;
                            console.log("Paste row:", pasteRow);
                            
                            // Search for the code type in column D (index 3)
                            let firstRow = -1;
                            let lastRow = -1;
                            
                            // Load the values of column D
                            const columnD = codesWS.getRange(`D1:D${usedRange.rowCount}`);
                            columnD.load("values");
                            await context.sync();
                            
                            console.log("Loaded column D values");
                            
                            // Check if values are loaded properly
                            if (!columnD.values) {
                                console.error("columnD.values is null or undefined");
                                throw new Error(`Failed to load values from column D in Codes worksheet`);
                            }
                            
                            console.log(`columnD.values length: ${columnD.values.length}`);
                            
                            // Debug print the first few values in column D
                            console.log("First 10 values in column D:");
                            for (let i = 0; i < Math.min(10, columnD.values.length); i++) {
                                console.log(`Row ${i+1}: ${columnD.values[i][0]}`);
                            }
                            
                            // Find the first and last row with the code
                            for (let row = 0; row < columnD.values.length; row++) {
                                if (columnD.values[row][0] === codeType) {
                                    if (firstRow === -1) {
                                        firstRow = row + 1; // Excel rows are 1-indexed
                                    }
                                    lastRow = row + 1;
                                }
                            }
                            
                            if (firstRow === -1 || lastRow === -1) {
                                throw new Error(`Code type ${codeType} not found in Codes worksheet`);
                            }
                            
                            console.log(`Found code type ${codeType} in rows ${firstRow} to ${lastRow}`);
                            
                            // Try the suggested approach to copy the range with all properties
                            await Excel.run(async (context) => {
                                // Get the source range
                                const sourceRange = context.workbook.worksheets.getItem("Codes").getRange(`A${firstRow}:CX${lastRow}`);
                                
                                // Get the destination range
                                const destinationRange = context.workbook.worksheets.getItem(currentWorksheetName).getRange(`A${pasteRow}`);
                                
                                // Copy the range with all properties
                                destinationRange.copyFrom(sourceRange, Excel.RangeCopyType.all);
                                
                                await context.sync();
                            });
                            
                            await context.sync();
                            
                            // Process driver and assumption inputs
                            await processDriverAndAssumptionInputs(currentWS, pasteRow, i, codeCollection[i]);
                            
                            result.processedCodes++;
                            console.log(`Processed code type: ${codeType}`);
                        } catch (error) {
                            console.error(`Error processing code ${codeType}:`, error);
                            throw error;
                        }
                    }).catch(error => {
                        console.error(`Error processing code ${codeType}: ${error.message}`);
                        result.errors.push({
                            codeIndex: i,
                            codeType: codeType,
                            error: error.message
                        });
                    });
                }
            } catch (error) {
                console.error(`Error processing code ${i}:`, error);
                result.errors.push({
                    codeIndex: i,
                    codeType: codeType,
                    error: error.message
                });
            }
        }
        
        // COMMENTED OUT: Final operations (similar to the end of the VBA sub)
        /*
        await Excel.run(async (context) => {
            // Recalculate all formulas
            context.workbook.application.calculate("Full");
            
            // Copy Financials column B
            const financialsWS = context.workbook.worksheets.getItem("Financials");
            const financialsColB = financialsWS.getRange("B:B");
            financialsColB.copy();
            
            // Paste to the same column
            financialsColB.paste();
            
            await context.sync();
        }).catch(error => {
            console.error(`Error in final operations: ${error.message}`);
            result.errors.push({
                phase: "final",
                error: error.message
            });
        });
        */
        
        // Return the result
        return result;
    } catch (error) {
        console.error("Error in runCodes:", error);
        throw error;
    }
}

/**
 * Process driver and assumption inputs for a code
 * @param {Excel.Worksheet} worksheet - The worksheet to process
 * @param {number} startRow - The starting row
 * @param {number} codeIndex - The index of the code in the collection
 * @param {Object} code - The code object
 */
async function processDriverAndAssumptionInputs(worksheet, startRow, codeIndex, code) {
    return Excel.run(async (context) => {
        // Get the code parameters
        const params = code.params;
        
        // Process driver inputs
        if (params.Driver) {
            const driverValue = params.Driver;
            const driverCell = worksheet.getRange(`D${startRow}`);
            driverCell.values = [[driverValue]];
        }
        
        // Process assumption inputs
        if (params.Assumptions) {
            const assumptions = params.Assumptions.split(',');
            
            for (let i = 0; i < assumptions.length; i++) {
                const assumption = assumptions[i].trim();
                const assumptionCell = worksheet.getRange(`E${startRow + i}`);
                assumptionCell.values = [[assumption]];
            }
        }
        
        await context.sync();
    });
}