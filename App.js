/*
 * Wealth-app
 * Copyright (C) 2026 Alessandro Saladino
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 *
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


