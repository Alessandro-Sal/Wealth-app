/**
 * Verifies the user's PIN against the secure Script Property.
 * Fetches 'APP_PIN' from the script properties service.
 * * @param {string} inputPin - The PIN entered by the user.
 * @return {boolean} True if the PIN matches, otherwise False.
 */
function verifyUserPin(inputPin) {
  const secretPin = PropertiesService.getScriptProperties().getProperty('APP_PIN');
  // If no PIN is set in properties, use emergency default "0000" or return false
  if (!secretPin) return inputPin === "0000"; 
  return inputPin === secretPin;
}

/**
 * SETUP UTILITY: Run this function ONCE manually from the Editor to set your secret PIN.
 * SECURITY WARNING: Never commit this file with your actual PIN inside the code.
 * Execute it once, then clear the value or delete the function.
 
function setupMyPin() {
  PropertiesService.getScriptProperties().setProperty('APP_PIN', 'YOUR_REAL_PIN_HERE'); // Replace with your actual PIN
  Logger.log("PIN saved securely!");
}*/