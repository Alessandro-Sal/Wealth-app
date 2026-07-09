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

    // FIX (1.7): cascata ridotta ai SOLI modelli reali ed esistenti.
    // Prima erano ~22 modelli (molti inesistenti: gemini-3.x-preview, nano-banana-pro,
    // deep-research-pro...), provati in serie con fetch bloccante -> nel caso peggiore
    // 60-150s e rischio di superare il limite di 6 minuti di esecuzione su analyzeAsset.
    const geminiModels = [
      "gemini-2.0-flash",       // latest stable
      "gemini-1.5-flash",       // fallback stabile
      "gemini-1.5-flash-8b",    // faster fallback
      "gemini-1.5-pro"          // heavy fallback
    ];

    // Deadline di sicurezza: non superare ~90s complessivi nella cascata Gemini.
    const geminiDeadline = Date.now() + 90000;

    for (let i = 0; i < geminiModels.length; i++) {
      if (Date.now() > geminiDeadline) {
        console.warn("Gemini cascade: deadline 90s raggiunta, interrompo per evitare il timeout.");
        break;
      }
      url = `https://generativelanguage.googleapis.com/v1beta/models/${geminiModels[i]}:generateContent?key=${apiKey}`;
      // FIX (1.8 parziale): limita l'output anche per Gemini (prima solo OpenRouter).
      payload = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { maxOutputTokens: 1500 } };
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
      let errStr = data.error ? (data.error.message || JSON.stringify(data.error)) : res.getContentText();
      console.warn(`[${providerName}] Error ${code}: ${errStr}`);
      if (code === 429) Utilities.sleep(2000); // Back-off for rate limits
      return `DEBUG_ERROR: [${providerName} ${code}] ${errStr}`;
    }
  } catch (e) {
    console.error(`Fetch Exception on ${providerName}:`, e);
    return `DEBUG_ERROR: [EXCEPTION] ${e.toString()}`;
  }
}