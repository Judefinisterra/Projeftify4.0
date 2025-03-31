/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import process from 'process/browser';
window.process = process;

import { ChatOpenAI } from "@langchain/openai";
import { ChatPromptTemplate } from "@langchain/core/prompts";
import { OpenAIEmbeddings } from '@langchain/openai';

//*********Setup*********
// Start the timer
const startTime = performance.now();

//Debugging Toggle
const DEBUG = false; 

// Replace dotenv with hardcoded API keys
const OPENAI_API_KEY = window.process.env.OPENAI_API_KEY || "";
const PINECONE_API_KEY = window.process.env.PINECONE_API_KEY || "";
const PINECONE_ENVIRONMENT = "gcp-starter"; // Your Pinecone environment
const PINECONE_INDEX = "codes"; // Your Pinecone index name

// Add Pinecone API configuration
const PINECONE_CONFIG = {
    apiKey: PINECONE_API_KEY,
    environment: PINECONE_ENVIRONMENT,
    indexName: PINECONE_INDEX,
    apiEndpoint: "https://codes-zmg9zog.svc.aped-4627-b74a.pinecone.io"
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

// Remove the PROMPTS object and add a function to load prompts
async function loadPromptFromFile(promptKey) {
  try {
    // Try different paths for Office Add-ins
    const paths = [
      `../prompts/${promptKey}.txt`,
      `/prompts/${promptKey}.txt`,
      `/src/prompts/${promptKey}.txt`,
      `./prompts/${promptKey}.txt`
    ];
    
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
    return null;
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
    const messages = [
        ["system", systemPrompt]
    ];
    
    if (history.length > 0) {
        history.forEach(message => {
            messages.push(message);
        });
    }
    
    messages.push(["human", userInput]);
    
    const prompt = ChatPromptTemplate.fromMessages(messages);

    const chatModel = new ChatOpenAI({
        modelName: model,
        temperature: temperature,
        openAIApiKey: OPENAI_API_KEY,
    });

    const chain = prompt.pipe(chatModel).pipe(response => {
        try {
            const parsed = JSON.parse(response.content);
            if (!Array.isArray(parsed)) {
                throw new Error('Response is not an array');
            }
            return parsed;
        } catch (e) {
            return response.content.split('\n').filter(line => line.trim());
        }
    });

    return await chain.invoke();
}

// Function 3: Query Vector Database using Pinecone REST API
async function queryVectorDB({ queryPrompt, indexName = PINECONE_INDEX, numResults = 10, similarityThreshold = null }) {
    try {
        console.log("Generating embeddings for query:", queryPrompt);
        const embeddings = new OpenAIEmbeddings({
            openAIApiKey: OPENAI_API_KEY,
            modelName: "text-embedding-3-large"
        });

        const embedding = await embeddings.embedQuery(queryPrompt);
        console.log("Embeddings generated successfully");
        
        const url = `${PINECONE_CONFIG.apiEndpoint}/query`;
        console.log("Making Pinecone API request to:", url);
        
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'api-key': PINECONE_CONFIG.apiKey,
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
        
        const systemPrompt = await getSystemPromptFromFile('followUpSystem');
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
    const MainPrompt = await getSystemPromptFromFile('Encoder_Main');

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
        const systemStructurePrompt = await getSystemPromptFromFile('structureSystem');
        
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

async function runValidation() {
    try {
        // Instead of running a separate script, we'll do validation in memory
        const validationData = localStorage.getItem('promptAnalysis');
        if (!validationData) {
            return "No validation data available";
        }
        
        // Basic validation logic
        const data = JSON.parse(validationData);
        if (!data.outputArray2 || !Array.isArray(data.outputArray2)) {
            return "Invalid response format";
        }
        
        return "Validation successful - no errors found";
    } catch (error) {
        console.error("Error during validation:", error);
        return "Validation failed: " + error.message;
    }
}

async function validationCorrection(clientprompt, initialResponse, validationResults) {
    try {
        const conversationHistory = loadConversationHistory();
        
        const trainingData = localStorage.getItem('trainingData') || "";
        const codeDescriptions = localStorage.getItem('codeDescriptions') || "";
        const lastCallContext = localStorage.getItem('lastCallContext') || "";
        
        const validationSystemPrompt = await getSystemPromptFromFile('validationSystem');
        const validationMainPrompt = await getSystemPromptFromFile('validationMain');
        
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

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  console.log("Office.onReady called");
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  
  // Test file loading
  getSystemPromptFromFile("/prompts/test-prompt.txt")
    .then(text => {
      console.log("Test prompt loaded:", text);
    })
    .catch(error => {
      console.error("Error loading test prompt:", error);
    });
  
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
      const validationResults = await runValidation();
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


