/**
 * Ultimate Universal AI router.
 * Features cascading model fallbacks for OpenRouter, Anthropic, and Gemini.
 * ORDERED BY INTELLIGENCE: Heaviest/Smartest models first, Lite models last.
 * * @param {string} prompt - The user prompt.
 * @param {string} provider - 'OPENROUTER', 'ANTHROPIC', 'GEMINI', or 'OPENAI'.
 * @param {boolean} useWebSearch - Web search capability (currently Gemini only).
 */
function fetchUniversalAI(prompt, provider = 'GEMINI', useWebSearch = false) {
  let url, payload, headers;

  // ==========================================
  // 🟢 OPENROUTER (The Ultimate Aggregator)
  // ==========================================
  if (provider === 'OPENROUTER') {
    const apiKey = typeof OPENROUTER_API_KEY !== 'undefined' ? OPENROUTER_API_KEY : null;
    if (!apiKey) { console.warn("Missing OPENROUTER_API_KEY"); return null; }
    
    url = 'https://openrouter.ai/api/v1/chat/completions';
    headers = { 
      'Authorization': 'Bearer ' + apiKey,
      'HTTP-Referer': 'https://github.com/alessandro-sal', 
      'X-Title': 'WealthApp' 
    };

    // Cascade: SMARTEST FIRST
    const orModels = [
      "anthropic/claude-3.5-sonnet",       // 1. Top Tier: Best reasoning
      "openai/gpt-4o",                     // 2. Top Tier: OpenAI Flagship
      "deepseek/deepseek-chat",            // 3. Top Tier: DeepSeek V3
      "meta-llama/llama-3.3-70b-instruct", // 4. Mid Tier: Best Open Source
      "openai/gpt-4o-mini"                 // 5. Fallback: Fast and cheap
    ];

    for (let i = 0; i < orModels.length; i++) {
      payload = {
        model: orModels[i],
        messages: [{ role: 'user', content: prompt }],
        temperature: 0.7,
        max_tokens: 1500 // Prevents Error 402 on low balance accounts
      };
      
      console.log(`Attempting OpenRouter: ${orModels[i]}...`);
      const response = executeFetch(url, headers, payload, 'OPENROUTER');
      if (response) return response;
    }
    return null;
  } 

  // ==========================================
  // 🟣 ANTHROPIC (Native API)
  // ==========================================
  else if (provider === 'ANTHROPIC') {
    const apiKey = typeof ANTHROPIC_API_KEY !== 'undefined' ? ANTHROPIC_API_KEY : null;
    if (!apiKey) { console.warn("Missing ANTHROPIC_API_KEY"); return null; }
    
    url = 'https://api.anthropic.com/v1/messages';
    headers = { 
      'x-api-key': apiKey, 
      'anthropic-version': '2023-06-01' 
    };

    // Cascade: SMARTEST FIRST
    const claudeModels = [
      "claude-3-5-sonnet-20241022", // Heavy
      "claude-3-5-haiku-20241022"   // Light
    ];

    for (let i = 0; i < claudeModels.length; i++) {
      payload = {
        model: claudeModels[i],
        max_tokens: 1500,
        messages: [{ role: 'user', content: prompt }]
      };
      
      console.log(`Attempting Anthropic: ${claudeModels[i]}...`);
      const response = executeFetch(url, headers, payload, 'ANTHROPIC');
      if (response) return response;
    }
    return null;
  } 

  // ==========================================
  // 🔵 GEMINI (Native Google API)
  // ==========================================
  else if (provider === 'GEMINI') {
    const apiKey = typeof GEMINI_API_KEY !== 'undefined' ? GEMINI_API_KEY : null;
    if (!apiKey) { console.warn("Missing GEMINI_API_KEY"); return null; }

    // --- EXHAUSTIVE CASCADE ORDERED BY INTELLIGENCE ---
    const geminiModels = [
      // 1. THE HEAVYWEIGHTS (Pro Models for Deep Reasoning)
      "gemini-3.1-pro-preview",
      "gemini-3-pro-preview",
      "gemini-2.5-pro",
      "gemini-pro-latest",
      
      // 2. THE SMART SPRINTERS (Flash Models)
      "gemini-2.5-flash",
      "gemini-3-flash-preview",
      "gemini-2.0-flash",
      "gemini-flash-latest",
      
      // 3. THE LIGHTWEIGHTS (Lite Models for fast fallback)
      "gemini-2.5-flash-lite",
      "gemini-2.0-flash-lite",
      "gemini-flash-lite-latest",
      
      // 4. EXPERIMENTAL & ALIASES
      "gemini-3.1-pro-preview-customtools",
      "gemini-2.5-flash-lite-preview-09-2025",
      "nano-banana-pro-preview",
      "deep-research-pro-preview-12-2025",
      "gemini-2.0-flash-001",
      "gemini-2.0-flash-lite-001",
      
      // 5. OPEN WEIGHTS (Ultimate Fallbacks)
      "gemma-3-27b-it",
      "gemma-3-12b-it",
      "gemma-3-4b-it",
      "gemma-3-1b-it",
      "gemma-3n-e4b-it",
      "gemma-3n-e2b-it"
    ];

    for (let i = 0; i < geminiModels.length; i++) {
      url = `https://generativelanguage.googleapis.com/v1beta/models/${geminiModels[i]}:generateContent?key=${apiKey}`;
      payload = { contents: [{ parts: [{ text: prompt }] }] };
      if (useWebSearch) payload.tools = [{ googleSearch: {} }];

      console.log(`Attempting Gemini: ${geminiModels[i]}...`);
      const response = executeFetch(url, {}, payload, 'GEMINI');
      if (response) {
          console.log(`✅ Success with: ${geminiModels[i]}`);
          return response;
      }
    }
    return null;
  }
}

/**
 * Executes HTTP requests and parses the response based on the provider's specific JSON structure.
 */
function executeFetch(url, headers, payload, providerName) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const res = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    const data = JSON.parse(res.getContentText());

    if (code === 200) {
      // Parse output based on provider standard
      if (providerName === 'OPENROUTER' || providerName === 'OPENAI') {
        return data.choices[0].message.content;
      } 
      else if (providerName === 'ANTHROPIC') {
        return data.content[0].text;
      } 
      else if (providerName === 'GEMINI') {
        return data.candidates[0].content.parts[0].text;
      }
    } else {
      console.warn(`[${providerName}] Error ${code}: ${res.getContentText()}`);
      if (code === 429) Utilities.sleep(2000); // Back-off for rate limits
      return null;
    }
  } catch (e) {
    console.error(`Fetch Exception on ${providerName}:`, e);
    return null;
  }
}