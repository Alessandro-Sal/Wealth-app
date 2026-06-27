/**
 * CONFIGURATION & SECRETS FILE
 * This file is tracked by Git. 
 * DO NOT HARDCODE REAL API KEYS OR PERSONAL DATA HERE.
 */

// Global API keys accessible from all other files in the project.
// Keys are safely fetched from Google Apps Script Properties.
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY") || "";
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY") || "";
const ANTHROPIC_API_KEY = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY") || "";
const OPENROUTER_API_KEY = PropertiesService.getScriptProperties().getProperty("OPENROUTER_API_KEY") || "";

/**
 * --- UTILITY: SET SCRIPT PROPERTIES PROGRAMMATICALLY ---
 * If your Apps Script Project Settings UI is locked because you have >50 properties (due to cache),
 * you can use this function to set your API keys without using the visual interface.
 * 
 * 1. Replace the empty strings below with your actual API keys.
 * 2. Select `setupMyApiKeys` from the run dropdown and click "Run".
 * 3. WARNING: Do not commit this file with your real keys to GitHub. Undo your changes after running.
 */
function setupMyApiKeys() {
  const props = PropertiesService.getScriptProperties();
  
  props.setProperties({
    "GEMINI_API_KEY": "", // <-- Paste your key here
    "OPENAI_API_KEY": "", // <-- Paste your key here
    "ANTHROPIC_API_KEY": "", // <-- Paste your key here
    "OPENROUTER_API_KEY": "" // <-- Paste your key here
  });
  
  console.log("API Keys updated successfully in PropertiesService!");
}