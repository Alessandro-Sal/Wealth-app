/**
 * --- TRIP DATA & MAPS MODULE ---
 * Manages travel history and geocoding for map visualizations.
 */

/**
 * Retrieves trip details for a specific year from the "Trip" sheet.
 * Parses currency strings, handles different number formats (EU/US), 
 * and calculates the "Other" costs (Total - Transport - Accommodation).
 * * @param {number|string} year - The year to filter by.
 * @return {Array<Object>} List of trip objects including costs and location.
 */
function getTripData(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Ensure the sheet is named exactly "Trip" (singular)
  const sheet = ss.getSheetByName("Trip");
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
  const targetYear = String(year).trim(); 

  return data
    .filter(row => String(row[0]).trim() === targetYear)
    .map(row => {
      // Helper function to sanitize numbers (handles currencies, dots, and commas)
      const cleanAndPos = (val) => {
        if (!val) return 0;
        let str = String(val).replace(/[€$£]/g, '').trim();
        
        // Handle "1.000,00" format
        if (str.includes(',') && str.includes('.')) {
             str = str.replace(/\./g, ''); 
             str = str.replace(',', '.');  
        } 
        // Handle "100,00" format
        else if (str.includes(',')) {
             str = str.replace(',', '.');
        } 
        // Handle "1.000" format (ambiguous, treated as thousands separator removal)
        else if (str.split('.').length > 1) {
             str = str.replace(/\./g, '');
        }
        
        let num = parseFloat(str);
        return isNaN(num) ? 0 : Math.abs(num);
      };

      const total = cleanAndPos(row[4]);
      const transport = cleanAndPos(row[5]);
      const accom = cleanAndPos(row[6]);
      let other = total - transport - accom;

      return {
        year: row[0],
        days: row[1],
        dest: row[2],
        city: row[3],
        total: total,
        transport: transport,
        accom: accom,
        other: other > 0 ? other : 0, 
        type: row[7]
      };
    });
}

/**
 * Fetches Geolocation (Lat/Lng) for trips using the Google Maps Geocoding API.
 * WARNING: This consumes Google Maps Quota. Use sparingly or cache results.
 * * @param {number|string} year - The year to filter by.
 * @return {Array<Object>} List of trips with added 'lat' and 'lng' properties.
 */
function getTripsWithCoords(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Trip");
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
  const targetYear = String(year).trim();
  let tripData = [];
  
  // Filter by year
  const filtered = data.filter(row => String(row[0]).trim() === targetYear);

  filtered.forEach(row => {
    let city = row[3]; 
    let dest = row[2];
    let fullLocation = city + ", " + dest; 
    let lat = 0, lng = 0;
    
    // Call Google Maps Geocoder
    try {
      let geocoder = Maps.newGeocoder().geocode(fullLocation);
      if (geocoder.status === 'OK') {
        lat = geocoder.results[0].geometry.location.lat;
        lng = geocoder.results[0].geometry.location.lng;
      }
    } catch (e) { Logger.log("Geocoding error for " + fullLocation); }

    // Only push if coordinates were found
    if (lat !== 0) {
      tripData.push({
        city: city,
        dest: dest,
        lat: lat,
        lng: lng,
        total: row[4]
      });
    }
  });
  return tripData;
}

/**
 * --- UNIFIED YEAR MANAGEMENT ---
 * Aggregates all available years from both "Expenses Tracker" sheets 
 * and the "Trip" log to populate UI dropdowns dynamically.
 * * @return {Array<string>} Unique list of years sorted descending (e.g., [2026, 2025]).
 */
function getAvailableYears() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let yearsSet = new Set();
  
  // 1. Scan "Expenses Tracker YYYY" sheets
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith("Expenses Tracker ")) {
      let y = name.replace("Expenses Tracker ", "").trim();
      if (!isNaN(parseFloat(y))) yearsSet.add(y);
    }
  });

  // 2. Scan the "Trip" sheet (Column A)
  const tripSheet = ss.getSheetByName("Trip");
  if (tripSheet && tripSheet.getLastRow() >= 2) {
     const tripYears = tripSheet.getRange(2, 1, tripSheet.getLastRow() - 1, 1).getValues().flat();
     tripYears.forEach(y => {
        if(y) yearsSet.add(String(y).trim());
     });
  }

  // Fallback: If no data found, ensure at least the current year is present
  if (yearsSet.size === 0) yearsSet.add(String(new Date().getFullYear()));

  // Return sorted descending array
  return Array.from(yearsSet).sort().reverse();
}