/**
 * Entry point for the Web App (HTTP GET request).
 * Evaluates the 'Html_Index' template, sets metadata (title, viewport),
 * and serves the final HTML output.
 */
function doGet() {
  // 1. Create the TEMPLATE object
  var template = HtmlService.createTemplateFromFile('Html_Index');
  
  // 2. Evaluate the template (executes scriptlets <?!= ?>) to get the output
  var output = template.evaluate();
  
  // 3. Configure final output settings
  output
    .setTitle('Wealth Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  
    // Note: XFrameOptions.ALLOWALL is often blocked by modern browsers; kept commented out for security.
    // output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return output;
}
/** Helper function to import content from other files (e.g., CSS, JS) into the HTML template.*/
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


