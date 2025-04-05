class Code {
    constructor() {
        this.code = '';
        this.codeValue = '';
        this.codestringvalue = '';
        this.timeseries = '';
        this.sections = '';
        this.beginningmonth = '';
        this.Delete = false;
        this.financialsdriver = '';
        this.labelrow = '';
        this.columnlabels = '';
        
        // Initialize arrays for labels, fincodes, drivers, and rows
        for (let i = 1; i <= 9; i++) {
            this[`label${i}`] = '';
            this[`fincode${i}`] = '';
            this[`driver${i}`] = '';
        }
        
        // Initialize rows (up to 200)
        for (let i = 1; i <= 200; i++) {
            this[`row${i}`] = '';
        }
    }
}

function populateCodeCollection(inputText) {
    const codeCollection = new Map();
    
    // Clean up input to extract only content within angle brackets
    let cleanedInput = '';
    let currentPos = 0;
    
    while (true) {
        const startBracket = inputText.indexOf('<', currentPos);
        if (startBracket === -1) break;
        
        const endBracket = inputText.indexOf('>', startBracket);
        if (endBracket === -1) break;
        
        cleanedInput += inputText.substring(startBracket, endBracket + 1);
        currentPos = endBracket + 1;
    }
    
    // Split the text into individual codes using the "<>" format
    const codeStrings = cleanedInput.split('>').filter(code => code.trim() !== '');
    
    for (const codeString of codeStrings) {
        const newCode = new Code();
        const codePart = codeString.replace('<', '').trim();
        const paramParts = codePart.split(';');
        
        // Assign unique key and add to codeCollection
        const baseCode = paramParts[0];
        let incrementedCode = baseCode;
        let counter = 1;
        
        // Check if the baseCode is already in the collection
        while (codeCollection.has(incrementedCode)) {
            incrementedCode = baseCode + counter;
            counter++;
        }
        
        newCode.code = incrementedCode;
        newCode.codeValue = paramParts[0];
        newCode.codestringvalue = codeString;
        
        for (const param of paramParts) {
            const processParam = (labelName) => {
                if (param.toLowerCase().includes(`${labelName.toLowerCase()}=`)) {
                    const startPos = param.indexOf('"');
                    const endPos = param.indexOf('"', startPos + 1);
                    if (startPos > -1 && endPos > -1) {
                        newCode[labelName.toLowerCase()] = param.substring(startPos + 1, endPos);
                    }
                }
            };
            
            // Process special parameters
            processParam('timeseries');
            processParam('sections');
            processParam('beginningmonth');
            processParam('financialsdriver');
            processParam('labelrow');
            processParam('columnlabels');
            
            // Process Delete parameter (special case)
            if (param.includes('Delete')) {
                newCode.Delete = true;
            }
            
            // Process labels, fincodes, and drivers (1-9)
            for (let i = 1; i <= 9; i++) {
                processParam(`label${i}`);
                processParam(`fincode${i}`);
                processParam(`driver${i}`);
            }
            
            // Process rows (1-200)
            for (let i = 1; i <= 200; i++) {
                const rowParam = `row${i}`;
                if (param.toLowerCase().includes(`${rowParam.toLowerCase()}=`) || 
                    param.toLowerCase().includes(`${rowParam.toLowerCase()} =`)) {
                    const startPos = param.indexOf('"');
                    const endPos = param.indexOf('"', startPos + 1);
                    if (startPos > -1 && endPos > -1) {
                        newCode[rowParam] = param.substring(startPos + 1, endPos);
                    }
                }
            }
        }
        
        codeCollection.set(incrementedCode, newCode);
    }
    
    return codeCollection;
}

// Function to export code collection to text
function exportCodeCollectionToText(codeCollection, filename = 'code_collection_results.txt') {
    let resultText = 'Code Collection Test Results\n';
    resultText += '===========================\n\n';
    
    // Iterate through the collection and format each code
    codeCollection.forEach((code, key) => {
        resultText += `Code: ${key}\n`;
        resultText += '-------------------\n';
        
        // Add all properties to the text
        for (const prop in code) {
            if (code[prop] !== '' && code[prop] !== false) {
                resultText += `${prop}: ${code[prop]}\n`;
            }
        }
        
        resultText += '\n';
    });
    
    return resultText;
}

// Export the functions and class
module.exports = {
    Code,
    populateCodeCollection,
    exportCodeCollectionToText
}; 