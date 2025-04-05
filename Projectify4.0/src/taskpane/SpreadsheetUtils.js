// SpreadsheetUtils.js - Utility functions for spreadsheet operations

/**
 * Inserts worksheets from a base64-encoded Excel file into the current workbook
 * @param {string} base64String - Base64-encoded string of the source Excel file
 * @param {string[]} [sheetNames] - Optional array of sheet names to insert. If not provided, all sheets will be inserted.
 * @returns {Promise<void>}
 */
export async function handleInsertWorksheetsFromBase64(base64String, sheetNames = null) {
    try {
        // Validate base64 string
        if (!base64String || typeof base64String !== 'string') {
            throw new Error("Invalid base64 string provided");
        }

        // Validate base64 format
        if (!/^[A-Za-z0-9+/]*={0,2}$/.test(base64String)) {
            throw new Error("Invalid base64 format");
        }

        await Excel.run(async (context) => {
            const workbook = context.workbook;
            
            // Check if we have the required API version
            if (!workbook.insertWorksheetsFromBase64) {
                throw new Error("This feature requires Excel API requirement set 1.13 or later");
            }
            
            // Insert the worksheets with error handling
            try {
                await workbook.insertWorksheetsFromBase64(base64String, {
                    sheetNames: sheetNames
                });
                
                await context.sync();
                console.log("Worksheets inserted successfully");
            } catch (error) {
                console.error("Error during worksheet insertion:", error);
                throw new Error(`Failed to insert worksheets: ${error.message}`);
            }
        });
    } catch (error) {
        console.error("Error inserting worksheets from base64:", error);
        throw error;
    }
} 