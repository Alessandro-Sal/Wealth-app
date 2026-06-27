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