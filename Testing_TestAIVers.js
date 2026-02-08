/**
 * ======================================================================================
 * MODULE: AI MODEL DEBUGGER
 * Helper script to list all available Gemini models for your API Key.
 * ======================================================================================
 */

function debugAvailableModels() {
  // ⚠️ SECURITY WARNING: Do not hardcode API keys in production or public repos.
  // Use PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') instead.
  const API_KEY = "YOUR_API_KEY_HERE"; 
  
  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${API_KEY}`;

  console.log("🔍 Contacting Google...");

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const text = response.getContentText();
    const json = JSON.parse(text);

    if (code !== 200) {
      console.error("❌ KEY/API ERROR (" + code + "):");
      console.error(json.error ? json.error.message : text);
      return;
    }

    console.log("✅ SUCCESS! Here are the exact models your key can access:");
    console.log("---------------------------------------------------");

    let foundFlash = false;
    
    if (json.models) {
      json.models.forEach(m => {
        // Filter only models that support text generation (generateContent)
        if (m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent")) {
          // Clean the name (remove "models/" prefix)
          const name = m.name.replace("models/", "");
          console.log("🔹 " + name);
          
          if(name.includes("flash")) foundFlash = true;
        }
      });
    }

    console.log("---------------------------------------------------");
    if(foundFlash) {
      console.log("💡 TIP: Use 'gemini-1.5-flash' (or the exact name listed above).");
    } else {
      console.log("⚠️ WARNING: No 'flash' model found. You may need to use 'gemini-pro'.");
    }

  } catch (e) {
    console.error("❌ CONNECTION ERROR: " + e.toString());
  }
}