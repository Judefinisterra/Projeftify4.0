/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Remove imports from Langchain to avoid ESM module issues
// Using direct fetch calls instead
// Add this test function
import { validateCodeStrings } from './Validation.js';
// Mock fs module for browser environment
const fs = {
    writeFileSync: (path, content) => {
        console.log(`Mock writeFileSync called with path: ${path}`);
        // In browser, we'll just log the content instead of writing to file
        console.log(`Content would be written to ${path}:`, content.substring(0, 100) + '...');
    }
};

//*********Setup*********
// Start the timer
const startTime = performance.now();

//Debugging Toggle
const DEBUG = true; 

// API keys storage
let API_KEYS = {
  OPENAI_API_KEY: "",
  PINECONE_API_KEY: ""
};

const srcPaths = [
  'https://localhost:3002/src/prompts/Encoder_System.txt',
  'https://localhost:3002/src/prompts/Encoder_Main.txt',
  'https://localhost:3002/src/prompts/Followup_System.txt',
  'https://localhost:3002/src/prompts/Structure_System.txt',
  'https://localhost:3002/src/prompts/Validation_System.txt',
  'https://localhost:3002/src/prompts/Validation_Main.txt'
];

// Function to load API keys from a config file
// This allows the keys to be stored in a separate file that's .gitignored
async function initializeAPIKeys() {
  try {
    console.log("Initializing API keys...");
    
    // Try to load config.js which is .gitignored
    try {
      const configResponse = await fetch('https://localhost:3002/config.js');
      if (configResponse.ok) {
        const configText = await configResponse.text();
        // Extract keys from the config text using regex
        const openaiKeyMatch = configText.match(/OPENAI_API_KEY\s*=\s*["']([^"']+)["']/);
        const pineconeKeyMatch = configText.match(/PINECONE_API_KEY\s*=\s*["']([^"']+)["']/);
        
        if (openaiKeyMatch && openaiKeyMatch[1]) {
          API_KEYS.OPENAI_API_KEY = openaiKeyMatch[1];
          console.log("OpenAI API key loaded from config.js");
        }
        
        if (pineconeKeyMatch && pineconeKeyMatch[1]) {
          API_KEYS.PINECONE_API_KEY = pineconeKeyMatch[1];
          console.log("Pinecone API key loaded from config.js");
        }
      }
    } catch (error) {
      console.warn("Could not load config.js, will use empty API keys:", error);
    }
    
    // Add debug logging with secure masking of keys
    console.log("OPENAI_API_KEY:", API_KEYS.OPENAI_API_KEY ? 
      `${API_KEYS.OPENAI_API_KEY.substring(0, 3)}...${API_KEYS.OPENAI_API_KEY.substring(API_KEYS.OPENAI_API_KEY.length - 3)}` : 
      "Not found");
    console.log("PINECONE_API_KEY:", API_KEYS.PINECONE_API_KEY ? 
      `${API_KEYS.PINECONE_API_KEY.substring(0, 3)}...${API_KEYS.PINECONE_API_KEY.substring(API_KEYS.PINECONE_API_KEY.length - 3)}` : 
      "Not found");
    
    return API_KEYS.OPENAI_API_KEY && API_KEYS.PINECONE_API_KEY;
  } catch (error) {
    console.error("Error initializing API keys:", error);
    return false;
  }
}

// Update Pinecone configuration to handle multiple indexes
const PINECONE_ENVIRONMENT = "gcp-starter";

// Define configurations for each index
const PINECONE_INDEXES = {
    codes: {
        name: "codes",
        apiEndpoint: "https://codes-zmg9zog.svc.aped-4627-b74a.pinecone.io"
    },
    call2trainingdata: {
        name: "call2trainingdata",
        apiEndpoint: "https://call2trainingdata-zmg9zog.svc.aped-4627-b74a.pinecone.io"
    },
    call2context: {
        name: "call2context",
        apiEndpoint: "https://call2context-zmg9zog.svc.aped-4627-b74a.pinecone.io"
    },
    call1context: {
        name: "call1context",
        apiEndpoint: "https://call1context-zmg9zog.svc.aped-4627-b74a.pinecone.io"
    }
};

//Models
const GPT4O_MINI = "gpt-4o-mini"
const GPT4O = "gpt-4o"
const GPT45_TURBO = "gpt-4.5-turbo"
const GPT35_TURBO = "gpt-3.5-turbo"
const GPT4_TURBO = "gpt-4-turbo"
const GPTFT1 =  "ft:gpt-3.5-turbo-1106:orsi-advisors:cohcolumnsgpt35:B6Wlrql1"

// Conversation history storage
let conversationHistory = [];

// Functions to save and load conversation history
function saveConversationHistory(history) {
    try {
        localStorage.setItem('conversationHistory', JSON.stringify(history));
        if (DEBUG) console.log('Conversation history saved to localStorage');
    } catch (error) {
        console.error("Error saving conversation history:", error);
    }
}

function loadConversationHistory() {
    try {
        const history = localStorage.getItem('conversationHistory');
        if (history) {
            if (DEBUG) console.log('Loaded conversation history from localStorage');
            const parsedHistory = JSON.parse(history);
            
            if (!Array.isArray(parsedHistory)) {
                console.error("Invalid history format, expected array");
                return [];
            }
            
            return parsedHistory;
        }
        if (DEBUG) console.log("No conversation history found in localStorage");
        return [];
    } catch (error) {
        console.error("Error loading conversation history:", error);
        return [];
    }
}

// Direct OpenAI API call function (replaces LangChain)
async function callOpenAI(messages, model = GPT4O, temperature = 0.7) {
  try {
    console.log(`Calling OpenAI API with model: ${model}`);
    
    // Check for API key
    if (!API_KEYS.OPENAI_API_KEY) {
      throw new Error("OpenAI API key not found. Please check your API keys.");
    }
    
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${API_KEYS.OPENAI_API_KEY}`
      },
      body: JSON.stringify({
        model: model,
        messages: messages,
        temperature: temperature
      })
    });
    
    if (!response.ok) {
      const errorData = await response.json().catch(() => null);
      console.error("OpenAI API error response:", errorData);
      throw new Error(`OpenAI API error: ${response.status} ${response.statusText}`);
    }
    
    const data = await response.json();
    console.log("OpenAI API response received");
    
    return data.choices[0].message.content;
  } catch (error) {
    console.error("Error calling OpenAI API:", error);
    throw error;
  }
}

// OpenAI embeddings function (replaces LangChain)
async function createEmbedding(text) {
  try {
    console.log("Creating embedding for text");
    
    // Check for API key
    if (!API_KEYS.OPENAI_API_KEY) {
      throw new Error("OpenAI API key not found. Please check your API keys.");
    }
    
    const response = await fetch('https://api.openai.com/v1/embeddings', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${API_KEYS.OPENAI_API_KEY}`
      },
      body: JSON.stringify({
        model: "text-embedding-3-large",
        input: text
      })
    });
    
    if (!response.ok) {
      const errorData = await response.json().catch(() => null);
      console.error("OpenAI Embeddings API error response:", errorData);
      throw new Error(`OpenAI Embeddings API error: ${response.status} ${response.statusText}`);
    }
    
    const data = await response.json();
    console.log("OpenAI Embeddings API response received");
    
    return data.data[0].embedding;
  } catch (error) {
    console.error("Error creating embedding:", error);
    throw error;
  }
}

// Remove the PROMPTS object and add a function to load prompts
async function loadPromptFromFile(promptKey) {
  try {
    // Use a simplified path approach that works with dev server with correct port
    const paths = [
      `https://localhost:3002/prompts/${promptKey}.txt`,
    ];
    
    // Combine all paths to try
    paths.push(...srcPaths);
 
    // Try each path until one works
    let response = null;
    for (const path of paths) {
      console.log(`Attempting to load prompt from: ${path}`);
      try {
        response = await fetch(path);
        if (response.ok) {
          console.log(`Successfully loaded prompt from: ${path}`);
          break;
        }
      } catch (err) {
        console.log(`Path ${path} failed: ${err.message}`);
      }
    }
    
    if (!response || !response.ok) {
      throw new Error(`Failed to load prompt: ${promptKey} (Could not find file in any location)`);
    }
    
    return await response.text();
  } catch (error) {
    console.error(`Error loading prompt ${promptKey}:`, error);
    throw error; // Re-throw the error to be handled by the caller
  }
}

// Update the getSystemPromptFromFile function
const getSystemPromptFromFile = async (promptKey) => {
  try {
    const prompt = await loadPromptFromFile(promptKey);
    if (!prompt) {
      throw new Error(`Prompt key "${promptKey}" not found`);
    }
    return prompt;
  } catch (error) {
    console.error(`Error getting prompt for key ${promptKey}:`, error);
    return null;
  }
};

//************Functions************
// Function 1: OpenAI Call with conversation history support
async function processPrompt({ userInput, systemPrompt, model, temperature, history = [] }) {
    console.log("API Key being used:", API_KEYS.OPENAI_API_KEY ? `${API_KEYS.OPENAI_API_KEY.substring(0, 3)}...` : "None");
    
    // Format messages in the way OpenAI expects
    const messages = [
        { role: "system", content: systemPrompt }
    ];
    
    // Add conversation history
    if (history.length > 0) {
        history.forEach(message => {
            messages.push({ 
                role: message[0] === "human" ? "user" : "assistant", 
                content: message[1] 
            });
        });
    }
    
    // Add current user input
    messages.push({ role: "user", content: userInput });
    
    try {
        // Call OpenAI API directly
        const responseContent = await callOpenAI(messages, model, temperature);
        
        // Try to parse JSON response if applicable
        try {
            const parsed = JSON.parse(responseContent);
            if (Array.isArray(parsed)) {
                return parsed;
            }
            return responseContent.split('\n').filter(line => line.trim());
        } catch (e) {
            // If not JSON, treat as text and split by lines
            return responseContent.split('\n').filter(line => line.trim());
        }
    } catch (error) {
        console.error("Error in processPrompt:", error);
        throw error;
    }
}

// Function 3: Query Vector Database using Pinecone REST API
async function queryVectorDB({ queryPrompt, indexName = 'codes', numResults = 10, similarityThreshold = null }) {
    try {
        console.log("Generating embeddings for query:", queryPrompt);
        
        // Generate embeddings using our direct API call
        const embedding = await createEmbedding(queryPrompt);
        console.log("Embeddings generated successfully");
        
        // Get the correct endpoint for the specified index
        const indexConfig = PINECONE_INDEXES[indexName];
        if (!indexConfig) {
            throw new Error(`Invalid index name: ${indexName}`);
        }
        
        const url = `${indexConfig.apiEndpoint}/query`;
        console.log("Making Pinecone API request to:", url);
        
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'api-key': API_KEYS.PINECONE_API_KEY,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                vector: embedding,
                topK: numResults,
                includeMetadata: true,
                namespace: "ns1"
            })
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.error("Pinecone API error details:", {
                status: response.status,
                statusText: response.statusText,
                error: errorText
            });
            throw new Error(`Pinecone API error: ${response.status} ${response.statusText} - ${errorText}`);
        }

        const data = await response.json();
        console.log("Pinecone API response received");
        
        let matches = data.matches || [];

        if (similarityThreshold !== null) {
            matches = matches.filter(match => match.score >= similarityThreshold);
        }

        matches = matches.slice(0, numResults);

        matches = matches.map(match => {
            try {
                if (match.metadata && match.metadata.text) {
                    return {
                        ...match,
                        text: match.metadata.text
                    };
                }
                return match;
            } catch (error) {
                console.error("Error processing match:", error);
                return match;
            }
        });

        if (DEBUG) {
            const matchesDescription = matches
                .map((match, i) => `Match ${i + 1} (score: ${match.score.toFixed(4)}): ${match.text || JSON.stringify(match.metadata)}`)
                .join('\n');
            console.log(matchesDescription);
        }

        const cleanMatches = matches.map(match => extractTextFromJson(match));
        return cleanMatches.filter(text => text !== "");

    } catch (error) {
        console.error("Error during vector database query:", error);
        throw error;
    }
}

function extractTextFromJson(jsonInput) {
   try {
       const jsonData = typeof jsonInput === 'string' ? JSON.parse(jsonInput) : jsonInput;
       
       if (Array.isArray(jsonData)) {
           for (const item of jsonData) {
               if (item.metadata && item.metadata.text) {
                   return item.metadata.text;
               }
           }
           throw new Error("No text field found in the JSON array");
       } 
       else if (jsonData.metadata && jsonData.metadata.text) {
           return jsonData.metadata.text;
       } 
       else {
           throw new Error("Invalid JSON structure: missing metadata.text field");
       }
   } catch (error) {
       console.error(`Error processing JSON: ${error.message}`);
       return "";
   }
}

function safeJsonForPrompt(obj, readable = true) {
    if (!readable) {
        let jsonString = JSON.stringify(obj);
        jsonString = jsonString.replace(/"values":\s*\[\],\s*"metadata":/g, '');
        return jsonString
            .replace(/{/g, '\\u007B')
            .replace(/}/g, '\\u007D');
    }
    
    if (Array.isArray(obj)) {
        return obj.map(item => {
            if (item.metadata && item.metadata.text) {
                const text = item.metadata.text.replace(/~/g, ',');
                const parts = text.split(';');
                
                let result = '';
                if (parts.length >= 1) result += parts[0].trim();
                if (parts.length >= 2) result += '\n' + parts[1].trim();
                if (parts.length >= 3) result += '\n' + parts[2].trim();
                
                if (item.score) {
                    result += `\nSimilarity Score: ${item.score.toFixed(4)}`;
                }
                
                return result;
            }
            return JSON.stringify(item).replace(/~/g, ',');
        }).join('\n\n');
    }
    
    const jsonString = JSON.stringify(obj, null, 2).replace(/~/g, ',');
    return jsonString
        .replace(/{/g, '\\u007B')
        .replace(/}/g, '\\u007D');
}

async function handleFollowUpConversation(clientprompt) {
    if (DEBUG) console.log("Processing follow-up question. Loading conversation history...");
    conversationHistory = loadConversationHistory();
    
    if (conversationHistory.length > 0) {
        if (DEBUG) console.log("Processing follow-up question:", clientprompt);
        if (DEBUG) console.log("Loaded conversation history:", JSON.stringify(conversationHistory, null, 2));
        
        const systemPrompt = await getSystemPromptFromFile('Followup_System');
        const MainPrompt = await getSystemPromptFromFile('main');
        
        const trainingdataCall2 = await queryVectorDB({
            queryPrompt: clientprompt,
            similarityThreshold: .4,
            indexName: 'call2trainingdata',
            numResults: 3
        });

        const call2context = await queryVectorDB({
            queryPrompt: clientprompt + trainingdataCall2,
            similarityThreshold: .3,
            indexName: 'call2context',
            numResults: 5
        });

        const call1context = await queryVectorDB({
            queryPrompt: clientprompt + trainingdataCall2,
            similarityThreshold: .3,
            indexName: 'call1context',
            numResults: 5
        });

        const codeOptions = await queryVectorDB({
            queryPrompt: clientprompt + trainingdataCall2 + call1context,
            indexName: 'codes',
            numResults: 10,
            similarityThreshold: .1
        });
        
        const followUpPrompt = "Client request: " + clientprompt + "\n" +
                       "Main Prompt: " + MainPrompt + "\n" +
                       "Training Data: " + safeJsonForPrompt(trainingdataCall2).replace(/~/g, ',') + "\n" +
                       "Code choosing context: " + safeJsonForPrompt(call1context) + "\n" +
                       "Code editing Context: " + safeJsonForPrompt(call2context) + "\n" +
                       "Code descriptions: " + safeJsonForPrompt(codeOptions);
        
        const response = await processPrompt({
            userInput: followUpPrompt,
            systemPrompt: systemPrompt,
            model: GPT4O,
            temperature: 1,
            history: conversationHistory
        });
        
        conversationHistory.push(["human", clientprompt]);
        conversationHistory.push(["assistant", response.join("\n")]);
        
        saveConversationHistory(conversationHistory);
        
        if (DEBUG) console.log("Updated conversation history:", JSON.stringify(conversationHistory, null, 2));
        
        savePromptAnalysis(clientprompt, systemPrompt, MainPrompt, call2context, call1context, trainingdataCall2, codeOptions, response);
        saveTrainingData(clientprompt, response);
        
        return response;
    } else {
        if (DEBUG) console.log("No conversation history found. Treating as initial question.");
        return handleInitialConversation(clientprompt);
    }
}

async function handleConversation(clientprompt, isFollowUp = false) {
    try {
        if (isFollowUp) {
            return await handleFollowUpConversation(clientprompt);
        } else {
            return await handleInitialConversation(clientprompt);
        }
    } catch (error) {
        console.error("Error in conversation handling:", error);
        return ["Error processing your request: " + error.message];
    }
}

async function handleInitialConversation(clientprompt) {
    if (DEBUG) console.log("Processing initial question:", clientprompt);
    
    const systemPrompt = await getSystemPromptFromFile('Encoder_System');
    console.log("SYSTEM PROMPT: ", systemPrompt);
    const MainPrompt = await getSystemPromptFromFile('Encoder_Main');
    console.log("MAIN PROMPT: ", MainPrompt);


    const Call2prompt = "Client request: " + clientprompt + "\n" +
                       "Main Prompt: " + MainPrompt;
    
    const outputArray2 = await processPrompt({
        userInput: Call2prompt,
        systemPrompt: systemPrompt,
        model: GPT4O,
        temperature: 1 
    });
    
    conversationHistory = [
        ["human", clientprompt],
        ["assistant", outputArray2.join("\n")]
    ];
    
    saveConversationHistory(conversationHistory);
    
    savePromptAnalysis(clientprompt, systemPrompt, MainPrompt, [], [], [], [], outputArray2);
    saveTrainingData(clientprompt, outputArray2);
    
    return outputArray2;
}

async function structureDatabasequeries(clientprompt) {
    if (DEBUG) console.log("Processing structured database queries:", clientprompt);

    try {
        console.log("Getting structure system prompt");
        const systemStructurePrompt = await getSystemPromptFromFile('Structure_System');
        
        if (!systemStructurePrompt) {
            throw new Error("Failed to load structure system prompt");
        }

        console.log("Got system prompt, processing query strings");
        const queryStrings = await processPrompt({
            userInput: clientprompt,
            systemPrompt: systemStructurePrompt,
            model: GPT4O,
            temperature: 1
        });

        if (!queryStrings || !Array.isArray(queryStrings)) {
            throw new Error("Failed to get valid query strings");
        }

        console.log("Got query strings:", queryStrings);
        const results = [];

        for (const queryString of queryStrings) {
            console.log("Processing query:", queryString);
            try {
                const queryResults = {
                    query: queryString,
                    trainingData: await queryVectorDB({
                        queryPrompt: queryString,
                        similarityThreshold: .4,
                        indexName: 'call2trainingdata',
                        numResults: 3
                    }),
                    call2Context: await queryVectorDB({
                        queryPrompt: queryString,
                        similarityThreshold: .2,
                        indexName: 'call2context',
                        numResults: 5
                    }),
                    call1Context: await queryVectorDB({
                        queryPrompt: queryString,
                        similarityThreshold: .2,
                        indexName: 'call1context',
                        numResults: 5
                    }),
                    codeOptions: await queryVectorDB({
                        queryPrompt: queryString,
                        indexName: 'codes',
                        numResults: 3,
                        similarityThreshold: .1
                    })
                };

                results.push(queryResults);
                console.log("Successfully processed query:", queryString);
            } catch (error) {
                console.error(`Error processing query "${queryString}":`, error);
                // Continue with next query instead of failing completely
                continue;
            }
        }

        if (results.length === 0) {
            throw new Error("No valid results were obtained from any queries");
        }

        return results;
    } catch (error) {
        console.error("Error in structureDatabasequeries:", error);
        throw error;
    }
}

function savePromptAnalysis(clientprompt, systemPrompt, MainPrompt, validationSystemPrompt, validationMainPrompt, validationResults, call2context, call1context, trainingdataCall2, codeOptions, outputArray2) {
    try {
        const analysisData = {
            clientRequest: clientprompt,
            systemPrompt,
            mainPrompt: MainPrompt,
            validationSystemPrompt,
            validationMainPrompt,
            validationResults,
            call2context,
            call1context,
            trainingdataCall2,
            codeOptions,
            outputArray2
        };
        
        localStorage.setItem('promptAnalysis', JSON.stringify(analysisData));
        if (DEBUG) console.log('Prompt analysis saved to localStorage');
    } catch (error) {
        console.error("Error saving prompt analysis:", error);
    }
}

function saveTrainingData(clientprompt, outputArray2) {
    try {
        function cleanText(text) {
            if (!text) return '';
            return text.toString()
                .replace(/\r?\n|\r/g, ' ')
                .trim();
        }
        
        const trainingData = {
            prompt: cleanText(clientprompt),
            response: cleanText(JSON.stringify(outputArray2))
        };
        
        localStorage.setItem('trainingData', JSON.stringify(trainingData));
        if (DEBUG) console.log('Training data saved to localStorage');
    } catch (error) {
        console.error("Error saving training data:", error);
    }
}

async function validationCorrection(clientprompt, initialResponse, validationResults) {
    try {
        const conversationHistory = loadConversationHistory();
        
        const trainingData = localStorage.getItem('trainingData') || "";
        const codeDescriptions = localStorage.getItem('codeDescriptions') || "";
        const lastCallContext = localStorage.getItem('lastCallContext') || "";
        
        const validationSystemPrompt = await getSystemPromptFromFile('Validation_System');
        const validationMainPrompt = await getSystemPromptFromFile('Validation_Main');
        
        if (!validationSystemPrompt) {
            throw new Error("Failed to load validation system prompt");
        }
        
        const responseString = Array.isArray(initialResponse) ? initialResponse.join("\n") : String(initialResponse);
        
        const correctionPrompt = 
            "Main Prompt: " + validationMainPrompt + "\n\n" +
            "Original User Input: " + clientprompt + "\n\n" +
            "Initial Response: " + responseString + "\n\n" +
            "Validation Results: " + validationResults + "\n\n" +
            "Training Data: " + trainingData + "\n\n" +
            "Code Descriptions: " + codeDescriptions + "\n\n" +
            "Context from Last Call: " + lastCallContext;
        
        if (DEBUG) {
            console.log("====== VALIDATION CORRECTION INPUT ======");
            console.log(correctionPrompt.substring(0, 500) + "...(truncated)");
            console.log("=========================================");
        }
        
        const correctedResponse = await processPrompt({
            userInput: correctionPrompt,
            systemPrompt: validationSystemPrompt,
            model: GPT4O,
            temperature: 0.7
        });
        
        const correctionOutputPath = "C:\\Users\\joeor\\Dropbox\\B - Freelance\\C_Projectify\\VanPC\\Training Data\\Main Script Training and Context Data\\validation_correction_output.txt";
        fs.writeFileSync(correctionOutputPath, Array.isArray(correctedResponse) ? correctedResponse.join("\n") : correctedResponse);
        
        if (DEBUG) console.log(`Validation correction saved to ${correctionOutputPath}`);
        
        return correctedResponse;
    } catch (error) {
        console.error("Error in validation correction:", error);
        console.error(error.stack);
        return ["Error during validation correction: " + error.message];
    }
}

// Add this function at the top level
function showError(message) {
    const errorDiv = document.createElement('div');
    errorDiv.style.color = 'red';
    errorDiv.style.padding = '10px';
    errorDiv.style.margin = '10px';
    errorDiv.style.border = '1px solid red';
    errorDiv.style.borderRadius = '4px';
    errorDiv.textContent = `Error: ${message}`;
    
    const appBody = document.getElementById('app-body');
    appBody.insertBefore(errorDiv, appBody.firstChild);
    
    // Remove the error message after 5 seconds
    setTimeout(() => {
        errorDiv.remove();
    }, 5000);
}

// Add this function at the top level
function setButtonLoading(isLoading) {
    const runButton = document.getElementById('run');
    if (runButton) {
        if (isLoading) {
            runButton.disabled = true;
            runButton.innerHTML = '<span class="ms-Button-label">Processing...</span>';
        } else {
            runButton.disabled = false;
            runButton.innerHTML = '<span class="ms-Button-label">Run</span>';
        }
    }
}

// Move the run function declaration to module level
export async function run() {
  console.log("Run function started");
  setButtonLoading(true);
  try {
    await Excel.run(async (context) => {
      console.log("Excel.run started");
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.load("values");
      await context.sync();
      
      console.log("Selected range:", range.address);
      const selectedText = range.values[0][0];
      console.log("Selected text:", selectedText);
      
      if (!selectedText) {
        throw new Error("No text selected in the range");
      }
      
      // Process the text through the main function
      console.log("Starting structureDatabasequeries");
      const dbResults = await structureDatabasequeries(selectedText);
      console.log("Database queries completed");
      
      if (!dbResults || !Array.isArray(dbResults)) {
        console.error("Invalid database results:", dbResults);
        throw new Error("Failed to get valid database results");
      }
      
      // Format the database results into a string
      const plainTextResults = dbResults.map(result => {
        if (!result) {
          console.error("Invalid result in dbResults:", result);
          return "No results found";
        }
        
        return `Query: ${result.query || 'No query'}\n` +
               `Training Data:\n${(result.trainingData || []).join('\n')}\n` +
               `Code Options:\n${(result.codeOptions || []).join('\n')}\n` +
               `Code Choosing Context:\n${(result.call1Context || []).join('\n')}\n` +
               `Code Editing Context:\n${(result.call2Context || []).join('\n')}\n` +
               `---\n`;
      }).join('\n');

      // Create an enhanced prompt that includes the database results
      const enhancedPrompt = `Client Request: ${selectedText}\n\nDatabase Results:\n${plainTextResults}`;
      console.log("Enhanced prompt created");

      // Process the conversation with the enhanced prompt
      console.log("Starting handleConversation");
      const response = await handleConversation(enhancedPrompt, false);
      console.log("Conversation completed");

      if (!response || !Array.isArray(response)) {
        console.error("Invalid response:", response);
        throw new Error("Failed to get valid response from conversation");
      }

      // Run validation and correction
      console.log("Starting validation");
      const validationResults = await validateCodeStrings(response);
      console.log("Validation completed:", validationResults);

      let finalResponse;
      if (validationResults && validationResults.includes("Validation successful - no errors found")) {
        finalResponse = response;
      } else {
        console.log("Starting validation correction");
        finalResponse = await validationCorrection(selectedText, response, validationResults);
        console.log("Validation correction completed");
      }
      
      if (!finalResponse || !Array.isArray(finalResponse)) {
        console.error("Invalid final response:", finalResponse);
        throw new Error("Failed to get valid final response");
      }
      
      // Write the final response back to Excel
      console.log("Writing response to Excel");
      range.values = [[finalResponse.join('\n')]];
      await context.sync();
      console.log("Response written to Excel");
    });
  } catch (error) {
    console.error("Error in run function:", error);
    console.error("Error stack:", error.stack);
    showError(error.message);
  } finally {
    setButtonLoading(false);
  }
}

// Update the Office.onReady callback to reference the run function
Office.onReady(() => {
  console.log("Office.onReady called");
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  
  // Initialize API keys before doing anything else
  initializeAPIKeys().then(success => {
    if (success) {
      console.log("API keys initialized successfully");
    } else {
      console.error("Failed to initialize API keys");
      showError("Failed to initialize API keys. Some features may not work correctly.");
    }

    // Add click handler with visual feedback
    const runButton = document.getElementById("run");
    if (runButton) {
      runButton.onclick = () => {
        console.log("Run button clicked");
        runButton.style.backgroundColor = "#0078d4"; // Visual feedback
        setTimeout(() => {
          runButton.style.backgroundColor = ""; // Reset color
        }, 200);
        run();
      };
      console.log("Run button click handler attached");
    } else {
      console.error("Run button not found in DOM");
    }
  });
});



async function testValidation() {
    // Test cases
    const testCases = [
        // Valid case
        [
            '<TAB; label1="Test Tab">',
            '<CODE-VV; row1="123|Description">',
            '<CODE-EV; row2="456|Another Description">'
        ],
        // Invalid case (missing EV/RV code)
        [
            '<TAB; label1="Test Tab">',
            '<CODE-VV; row1="123|Description">'
        ],
        // Invalid format case
        [
            '<INVALID; row1="123">'
        ]
    ];

    console.log("Starting validation tests...");

    for (let i = 0; i < testCases.length; i++) {
        console.log(`\nTest Case ${i + 1}:`);
        console.log("Input:", testCases[i]);
        
        const errors = await validateCodeStrings(testCases[i]);
        
        if (errors && errors.length > 0) {
            console.log("Validation Errors:");
            errors.forEach(error => console.log(`- ${error}`));
        } else {
            console.log("Validation Successful - No errors");
        }
    }
}

// Add this to your Office.onReady handler if not already present
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("run-test").onclick = testValidation;
    }
});


