// Remove all these imports as they're no longer needed
// import { promises as fs } from 'fs';
// import { join } from 'path';
// import { fileURLToPath } from 'url';
// import { dirname } from 'path';

async function validateCodeStrings(inputCodeStrings) {
    const errors = [];
    const tabLabels = new Set();
    const rowValues = new Set();
    const codeTypes = new Set();
    
    // Track codes by their suffixes
    const vvCodes = new Set();  // codes ending in -VV
    const vrCodes = new Set();  // codes ending in -VR
    const rrCodes = new Set();  // codes ending in -RR
    const rvCodes = new Set();  // codes ending in -RV
    const evCodes = new Set();  // codes ending in -EV
    const erCodes = new Set();  // codes ending in -ER

    // Load valid codes from the correct location
    let validCodes = new Set();
    try {
        const response = await fetch('../prompts/Codes.txt');
        if (!response.ok) {
            throw new Error('Failed to load Codes.txt file');
        }
        const fileContent = await response.text();
        validCodes = new Set(fileContent.split('\n')
            .map(line => line.trim())
            .filter(line => line.length > 0));
    } catch (error) {
        errors.push(`Error reading Codes.txt file: ${error.message}`);
        return errors;
    }

    // First pass: collect all row values, code types, and suffixes
    for (const codeString of inputCodeStrings) {
        if (!codeString.startsWith('<') || !codeString.endsWith('>')) {
            errors.push(`Invalid code string format: ${codeString}`);
            continue;
        }
        // Extract and store row IDs
        if (codeString.startsWith('<BR>')) {
            // Skip extraction for BR tags
            continue;
        }

        const rowMatches = codeString.match(/row\d+\s*=\s*"([^"]*)"/g);
        if (rowMatches) {
            rowMatches.forEach(match => {
                const rowContent = match.match(/row\d+\s*=\s*"([^"]*)"/)[1];
                // Handle spaces before/after the pipe delimiter
                const parts = rowContent.split('|');
                if (parts.length > 0) {
                    const rowId = parts[0].trim();
                    if (rowId) {
                        rowValues.add(rowId);
                    }
                }
            });
        }

        
        // console.log(rowMatches);
        // Extract code type and suffix
        const codeMatch = codeString.match(/<([^;]+);/);
        if (codeMatch) {
            const codeType = codeMatch[1].trim();
            codeTypes.add(codeType);

            // Check for suffixes and store them
            if (codeType.endsWith('-VV')) vvCodes.add(codeType);
            if (codeType.endsWith('-VR')) vrCodes.add(codeType);
            if (codeType.endsWith('-RR')) rrCodes.add(codeType);
            if (codeType.endsWith('-RV')) rvCodes.add(codeType);
            if (codeType.endsWith('-EV')) evCodes.add(codeType);
            if (codeType.endsWith('-ER')) erCodes.add(codeType);
        }
    }

    // Validate suffix relationships
    // Rule 1: -VV or -VR must have -EV or -RV
    for (const code of [...vvCodes, ...vrCodes]) {
        if (evCodes.size === 0 && rvCodes.size === 0) {
            errors.push(`Code ${code} requires another code with suffix -EV or -RV. Fix by adding another code with the correct suffix or by change to the EV/ER version of the same code.`);
        }
    }

    // Rule 2: -RR or -RV must have -ER or -VR
    for (const code of [...rrCodes, ...rvCodes]) {
        if (erCodes.size === 0 && vrCodes.size === 0) {
            errors.push(`Code ${code} requires another code with suffix -ER or -VR`);
        }
    }

// Second pass: detailed validation
for (const codeString of inputCodeStrings) {
    // Skip BR tags completely
    if (codeString === '<BR>') {
        continue;
    }
    
    const codeMatch = codeString.match(/<([^;]+);/);
    if (!codeMatch) {
        errors.push(`Cannot extract code type from: ${codeString}`);
        continue;
    }

    const codeType = codeMatch[1].trim();
    
    // Validate code exists in description file
    if (!validCodes.has(codeType)) {
        errors.push(`Invalid code type: "${codeType}" not found in valid codes list`);
    }



        // Validate TAB labels
        if (codeType === 'TAB') {
            const labelMatch = codeString.match(/label\d+="([^"]*)"/);
            if (!labelMatch) {
                errors.push('TAB code missing label parameter');
                continue;
            }

            const label = labelMatch[1];
            
            if (label.length > 30) {
                errors.push(`Tab label too long (max 30 chars): "${label}"`);
            }
            
            if (/[&,":;]/.test(label)) {
                errors.push(`Tab label contains illegal characters (&,":;): "${label}"`);
            }
            
            if (tabLabels.has(label)) {
                errors.push(`Duplicate tab label: "${label}"`);
            }
            tabLabels.add(label);
        }

        // Validate row format
        const rowMatches = codeString.match(/row\d+="([^"]*)"/g);
        if (rowMatches) {
            rowMatches.forEach(match => {
                const rowContent = match.match(/row\d+="([^"]*)"/)[1];
                const parts = rowContent.split('|');
                if (parts.length < 2) {
                    errors.push(`Invalid row format (missing required fields): "${rowContent}"`);
                }
            });
        }
    } // <-- This is the end of the second pass loop

    // Third pass: validate driver references after all row IDs are collected
    for (const codeString of inputCodeStrings) {
        const driverMatches = codeString.match(/driver\d+\s*=\s*"([^"]*)"/g);
        if (driverMatches) {
            driverMatches.forEach(match => {
                const driverValue = match.match(/driver\d+\s*=\s*"([^"]*)"/)[1].trim();
                // console.log(driverValue);
                if (!rowValues.has(driverValue)) {
                    errors.push(`Driver value "${driverValue}" not found in any row`);
                }
            });
        }
    }

    return errors;
} // <-- This is the end of the validateCodeStrings function

// Modified export function
export async function runValidation(inputStrings) {
    try {
        const cleanedStrings = inputStrings
            .map(line => line.trim())
            .filter(line => line.length > 0)
            .map(line => line.replace(/^'|',$|,$|^"|"$/g, ''));

        const codeStrings = cleanedStrings.join(' ').match(/<[^>]+>/g) || [];

        if (codeStrings.length === 0) {
            console.error('Validation run with no codestrings - input is empty');
            return ['No code strings found to validate'];
        }

        const validationErrors = await validateCodeStrings(codeStrings);
        return validationErrors;
    } catch (error) {
        console.error('Validation failed:', error);
        return [`Validation failed: ${error.message}`];
    }
}


