const SPREADSHEET_NAME = "V1_GHG EMISSION INVENTORY FORM 2025"; // <<<--- Confirmed Google Sheet Name


/**
 * Maps column letters to their 0-indexed positions for easier and more robust access.
 * Excel A=1, B=2, ... Z=26, AA=27, etc.
 * Apps Script arrays are 0-indexed, so A=0, B=1, etc.
 *
 * IMPORTANT: These indices are based on the user's provided 1-indexed column numbers.
 */
const COLUMNS = {
  PLANT: 1, // B (2nd column) - from GHG Calculation
  MONTH_COVERAGE: 2, // C (3rd column) - from GHG Calculation

  // Section 1: Total Annual GHG Emission
  TOTAL_LAST_YEAR: 114, // DK (115th column)
  TOTAL_ACTUAL: 91, // CN (92nd column)
  TOTAL_TARGET: 118, // DO (119th column)

  // Section 2: Scope 1
  SCOPE1_SCORECARD_ACTUAL: 80, // CC (81st column)
  SCOPE1_SCORECARD_TARGET: 115, // DL (116th column)

  // Scope 1 Category Pie Chart (data for slices)
  SCOPE1_PIE_MOBILE_ON_ROAD: 79, // CB (80th column)
  SCOPE1_PIE_MOBILE_NON_ROAD: 52, // BA (53rd column)
  SCOPE1_PIE_STATIONARY: 27, // AB (28th column)

  // Scope 1 Subsections Breakdown Cards (values)
  STATIONARY_TOTAL_SUB: 27, // AB (28th column) - This is the TCO2eq total for Stationary
  STATIONARY_DIESEL: 3, // D (4th column) - Raw Diesel consumption
  STATIONARY_LPG: 15, // P (16th column) - Raw LPG consumption

  MOBILE_NON_ROAD_TOTAL_SUB: 52, // BA (53rd column) - This is the TCO2eq total for Mobile Non-Road
  MOBILE_NON_ROAD_DIESEL: 28, // AC (29th column) - Raw Diesel consumption
  MOBILE_NON_ROAD_GASOLINE: 40, // AO (41st column) - Raw Gasoline consumption

  MOBILE_ON_ROAD_TOTAL_SUB: 79, // CB (80th column) - This is the TCO2eq total for Mobile On-Road
  MOBILE_ON_ROAD_DIESEL: 53, // BB (54th column) - Raw Diesel consumption
  MOBILE_ON_ROAD_GASOLINE: 66, // BO (67th column) - Raw Gasoline consumption

  // Monthly & Plant Bar Charts (Scope 1)
  SCOPE1_BAR_LAST_YEAR: 111, // DH (112th column) - This is generally 'Last Year' value used for bar charts
  SCOPE1_BAR_ACTUAL: 80, // CC (81st column) - Same as SCOPE1_SCORECARD_ACTUAL
  SCOPE1_BAR_TARGET: 115, // DL (116th column) - Same as SCOPE1_SCORECARD_TARGET

  // Section 3: Scope 2 (New)
  SCOPE2_ACTUAL: 90, // CM (91st column) - Used for scorecard, energy mix pie, and bar charts
  SCOPE2_TARGET: 116, // DM (117th column) - Used for scorecard performance and bar charts
  SCOPE2_LAST_YEAR: 112, // DI (113th column) - Used for bar charts (Last Year)

  // Scope 2 Energy Mix (TCO2eq for Pie Chart)
  SCOPE2_PIE_NON_RENEWABLE_TCO2EQ: 87, // CJ (88th column)
  SCOPE2_PIE_RENEWABLE_TCO2EQ: 89, // CL (90th column)

  // Scope 2 Subsection Cards (kWh values)
  SCOPE2_KWH_NON_RENEWABLE: 86, // CI (87th column)
  SCOPE2_KWH_RENEWABLE: 88, // CK (89th column)

  // NEW: Section 4: Scope 3
  SCOPE3_ACTUAL: 94, // CQ (95th column) - For Scope 3 Scorecard

  // NEW: Water Data
  TOTAL_WATER_WITHDRAWN: 93, // CP (94th column)
  TOTAL_WATER_TREATMENT: 96, // CS (97th column)

  // NEW: Scope 3 Waste Category 1 (for subsection cards)
  MIXED_SOLID_WASTE_KG: 97, // CT (98th column)
  PAPER_KG: 98, // CU (99th column)
  PLASTIC_KG: 99, // CV (100th column)

  // NEW: Scope 3 Waste Category 2 (for subsection cards)
  ASSORTED_METAL_KG: 100, // CW (101st column)
  FABRIC_KG: 102, // CY (103rd column)
  FOAM_SCRAPS_KG: 103, // CZ (104th column)

  // NEW: Scope 3 Waste Category 3 (for subsection cards)
  USED_OILS_LITERS: 104, // DA (105th column)
  CHEMICAL_FLASHING_MT: 105, // DB (106th column)
  OIL_CONTAMINATED_RAGS_KG: 107 // DD (108th column)
};


/**
 * Serves the HTML page for the web dashboard.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('GHG Emission Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * Includes external HTML files (CSS, JS) into the main HTML file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}


/**
 * Helper function to parse month and year from a "Month Coverage" string.
 * @param {string} monthStr - The month string, e.g., "January 1-31, 2025".
 * @returns {{month: string, year: number|null}} An object containing the month name and year (as number), or null if year not found.
*/
function parseMonthYearString(monthStr) {
  // Regex to extract month name and year from formats like "Month 1-31, YYYY" or "Month D-D YYYY"
  const match = monthStr.match(/([A-Za-z]+)\s+\d{1,2}-\d{1,2},?\s*(\d{4})/);
  if (match) {
    return { month: match[1], year: parseInt(match[2]) };
  }
  // Fallback for just "Month" or other formats if needed
  const parts = monthStr.split(/[\s,-]+/);
  const month = parts[0];
  const year = parseInt(parts[parts.length - 1]);
  return { month: month, year: isNaN(year) ? null : year }; // Default year to null if not found
}


/**
 * Fetches unique Plant and Month Coverage data for filters.
 * Assumes 'GHG Calculation' sheet has 'Plant' in Column B and 'Month Coverage' in Column C, starting from row 10.
 * @returns {Object} An object containing arrays of unique plants and months.
 */
function getFilterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GHG Calculation');
  if (!sheet) {
    throw new Error('Sheet "GHG Calculation" not found. Please ensure the sheet name is correct.');
  }

  // Get data from Column B (Plant) and C (Month Coverage) starting row 10
  const dataRange = sheet.getRange(10, COLUMNS.PLANT + 1, sheet.getLastRow() - 9, 2); // Start at B10, 2 columns wide
  const data = dataRange.getValues();

  const plants = new Set();
  const months = new Set();
  const years = new Set(); // New Set for years

  data.forEach(row => {
    if (row[0]) plants.add(row[0].toString().trim()); // Column B is index 0 in the fetched range (within dataRange)
    if (row[1]) {
      const monthStr = row[1].toString().trim();
      months.add(monthStr); // Column C is index 1 in the fetched range (within dataRange)
      const parsed = parseMonthYearString(monthStr);
      if (parsed.year) years.add(parsed.year.toString()); // Add year as string
    }
  });

  return {
    plants: Array.from(plants).filter(String).sort(), // Filter out empty strings and sort alphabetically
    months: Array.from(months).filter(String).sort(sortByMonth), // Filter out empty strings and sort chronologically
    years: Array.from(years).filter(String).sort((a,b) => parseInt(a) - parseInt(b)) // Sort years numerically
  };
}


/**
 * Custom sort function for months (e.g., January, February, March).
 * Handles different year formats if present.
 */
function sortByMonth(a, b) {
  const monthOrder = {
    "January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
    "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12
  };

  const parsedA = parseMonthYearString(a);
  const parsedB = parseMonthYearString(b);

  const yearA = parsedA.year || 0; // Treat null year as 0 for sorting
  const yearB = parsedB.year || 0; // Treat null year as 0 for sorting

  if (yearA !== yearB) {
    return yearA - yearB;
  }

  return (monthOrder[parsedA.month] || 0) - (monthOrder[parsedB.month] || 0);
}


/**
 * Fetches and processes dashboard data based on selected filters.
 * @param {string} selectedMonth - The month to filter by (e.g., "January 1-31, 2025").
 * @param {string} selectedPlant - The plant to filter by (e.g., "Polyfoam Valenzuela").
 * @param {string} selectedYear - The year to filter by (e.g., "2025").
 * @returns {Object} Processed data for dashboard components.
*/
function getDashboardData(selectedMonth, selectedPlant, selectedYear) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ghgSheet = ss.getSheetByName('GHG Calculation');

  if (!ghgSheet) {
    throw new Error('Sheet "GHG Calculation" not found. Please ensure the sheet name is correct.');
  }

  // Determine the last column index needed for data fetching from 'GHG Calculation'
  const ghgMaxColIndex = Math.max(
    ...Object.values(COLUMNS)
  );

  // Fetch only the necessary columns from 'GHG Calculation' starting from row 10
  const rawGhgData = ghgSheet.getRange(10, 1, ghgSheet.getLastRow() - 9, ghgMaxColIndex + 1).getValues();

  // Apply filters to the raw data first
  const filteredGhgData = rawGhgData.filter(row => {
    // Ensure row length is sufficient before accessing indices
    if (row.length <= ghgMaxColIndex) { // Check against 0-indexed max column
        return false; // Skip malformed rows
    }
    const plant = String(row[COLUMNS.PLANT]).trim();
    const month = String(row[COLUMNS.MONTH_COVERAGE]).trim();
    
    // Parse month string for its year
    const parsedMonth = parseMonthYearString(month);
    const rowYear = parsedMonth.year ? parsedMonth.year.toString() : null; // Ensure rowYear is string or null

    const matchesMonth = (selectedMonth === 'All Months' || month === selectedMonth);
    const matchesPlant = (selectedPlant === 'All Plants' || plant === selectedPlant);
    const matchesYear = (selectedYear === 'All Years' || rowYear === selectedYear); // Compare string with string

    return matchesMonth && matchesPlant && matchesYear;
  });

  // --- Initialize data structures for ALL sections ---
  // Total Annual GHG Emission
  let totalGHGEmission = 0;
  let totalTargetLimit = 0;

  // Scope 1 Section
  let scope1ActualTotal = 0;
  let scope1TargetTotal = 0;
  let scope1MobileOnRoadTotalPie = 0; // for pie
  let scope1MobileNonRoadTotalPie = 0; // for pie
  let scope1StationaryTotalPie = 0; // for pie
  let scope1StationaryTotalSub = 0; // for subsection (TCO2eq)
  let scope1StationaryDieselTotal = 0; // for subsection (Liters) - aggregated from column D
  let scope1StationaryLPGTotal = 0; // for subsection (KG) - aggregated from column P
  let mobileNonRoadTotalSub = 0; // for 2nd subsection card (TCO2eq)
  let mobileNonRoadDiesel = 0; // aggregated from column AC
  let mobileNonRoadGasoline = 0; // aggregated from column AO
  let mobileOnRoadTotalSub = 0; // for 3rd subsection card (TCO2eq)
  let mobileOnRoadDiesel = 0; // aggregated from column BB
  let mobileOnRoadGasoline = 0; // aggregated from column BO
  
  // Data for main pie chart (Scope 1 and Scope 2 totals from GHG Calculation)
  let mainPieChartScope1Total = 0;
  let mainPieChartScope2Total = 0;

  // Monthly & Plant Bar Charts Data Maps
  const monthlyTotalGhgDataMap = {};
  const plantTotalGhgDataMap = {};
  const monthlyScope1DataMap = {};
  const plantScope1DataMap = {};

  // NEW: Scope 2 Section
  let scope2ActualTotal = 0;
  let scope2TargetTotal = 0;
  let scope2LastYearTotal = 0;
  let scope2NonRenewableEnergyPie = 0;
  let scope2RenewableEnergyPie = 0;
  let scope2KWHNonRenewable = 0;
  let scope2KWHRenewable = 0;

  const monthlyScope2DataMap = {};
  const plantScope2DataMap = {};

  // NEW: Scope 3 Section
  let scope3ActualTotal = 0;

  let totalWaterWithdrawn = 0;
  let totalWaterTreatment = 0;

  let mixedSolidWasteKG = 0;
  let paperKG = 0;
  let plasticKG = 0;
  let assortedMetalKG = 0;
  let fabricKG = 0;
  let foamScrapsKG = 0;
  let usedOilsLiters = 0;
  let chemicalFlashingMT = 0;
  let oilContaminatedRagsKG = 0;


  filteredGhgData.forEach(row => {
    const plant = String(row[COLUMNS.PLANT]).trim();
    const month = String(row[COLUMNS.MONTH_COVERAGE]).trim();
    
    // --- Parse values for Total Annual GHG Emission Section ---
    const totalActual = Number(row[COLUMNS.TOTAL_ACTUAL]) || 0;
    const totalLastYear = Number(row[COLUMNS.TOTAL_LAST_YEAR]) || 0;
    const totalTarget = Number(row[COLUMNS.TOTAL_TARGET]) || 0;

    // --- Parse values for Scope 1 Bar Charts & Scorecard ---
    const s1ActualBar = Number(row[COLUMNS.SCOPE1_BAR_ACTUAL]) || 0;
    const s1LastYearBar = Number(row[COLUMNS.SCOPE1_BAR_LAST_YEAR]) || 0;
    const s1TargetBar = Number(row[COLUMNS.SCOPE1_BAR_TARGET]) || 0;

    // Scope 1 Scorecard values (using CC & DL)
    const s1ScorecardActual = Number(row[COLUMNS.SCOPE1_SCORECARD_ACTUAL]) || 0;
    const s1ScorecardTarget = Number(row[COLUMNS.SCOPE1_SCORECARD_TARGET]) || 0;

    // --- Parse values for Scope 1 Category Breakdown data (for Scope 1 Pie Chart) ---
    const s1MobileOnRoadPieVal = Number(row[COLUMNS.SCOPE1_PIE_MOBILE_ON_ROAD]) || 0;
    const s1MobileNonRoadPieVal = Number(row[COLUMNS.SCOPE1_PIE_MOBILE_NON_ROAD]) || 0;
    const s1StationaryTotalPieVal = Number(row[COLUMNS.SCOPE1_PIE_STATIONARY]) || 0;

    // --- Parse values for Scope 1 Stationary Combustion Breakdown data (for 1st subsection cards) ---
    const stationaryTotalSubVal = Number(row[COLUMNS.STATIONARY_TOTAL_SUB]) || 0;
    const stationaryDiesel = Number(row[COLUMNS.STATIONARY_DIESEL]) || 0;
    const stationaryLPG = Number(row[COLUMNS.STATIONARY_LPG]) || 0;

    // --- Parse values for Scope 1 Mobile Combustion Non-Road Breakdown data (for 2nd subsection cards) ---
    const mobileNonRoadTotalVal = Number(row[COLUMNS.MOBILE_NON_ROAD_TOTAL_SUB]) || 0;
    const mobileNonRoadDieselVal = Number(row[COLUMNS.MOBILE_NON_ROAD_DIESEL]) || 0;
    const mobileNonRoadGasolineVal = Number(row[COLUMNS.MOBILE_NON_ROAD_GASOLINE]) || 0;

    // --- Parse values for Scope 1 Mobile Combustion On-Road Breakdown data (for 3rd subsection cards) ---
    const mobileOnRoadTotalVal = Number(row[COLUMNS.MOBILE_ON_ROAD_TOTAL_SUB]) || 0;
    const mobileOnRoadDieselVal = Number(row[COLUMNS.MOBILE_ON_ROAD_DIESEL]) || 0;
    const mobileOnRoadGasolineVal = Number(row[COLUMNS.MOBILE_ON_ROAD_GASOLINE]) || 0;
    
    // --- Parse value for Main GHG Emission Scope Category Contribution Pie Chart (Scope 2 slice) ---
    const s2TotalForMainPie = Number(row[COLUMNS.SCOPE2_ACTUAL]) || 0; // Use SCOPE2_ACTUAL for the main pie

    // --- NEW: Parse values for Scope 2 Section ---
    const s2Actual = Number(row[COLUMNS.SCOPE2_ACTUAL]) || 0;
    const s2Target = Number(row[COLUMNS.SCOPE2_TARGET]) || 0;
    const s2LastYear = Number(row[COLUMNS.SCOPE2_LAST_YEAR]) || 0;
    const s2NonRenewableTCO2eq = Number(row[COLUMNS.SCOPE2_PIE_NON_RENEWABLE_TCO2EQ]) || 0;
    const s2RenewableTCO2eq = Number(row[COLUMNS.SCOPE2_PIE_RENEWABLE_TCO2EQ]) || 0;
    const s2KWHNonRenewableVal = Number(row[COLUMNS.SCOPE2_KWH_NON_RENEWABLE]) || 0;
    const s2KWHRenewableVal = Number(row[COLUMNS.SCOPE2_KWH_RENEWABLE]) || 0;

    // NEW: Parse values for Scope 3 Section
    const s3Actual = Number(row[COLUMNS.SCOPE3_ACTUAL]) || 0;

    const waterWithdrawn = Number(row[COLUMNS.TOTAL_WATER_WITHDRAWN]) || 0;
    const waterTreatment = Number(row[COLUMNS.TOTAL_WATER_TREATMENT]) || 0;

    const msWasteKG = Number(row[COLUMNS.MIXED_SOLID_WASTE_KG]) || 0;
    const paperValKG = Number(row[COLUMNS.PAPER_KG]) || 0;
    const plasticValKG = Number(row[COLUMNS.PLASTIC_KG]) || 0;
    const assortedMetalValKG = Number(row[COLUMNS.ASSORTED_METAL_KG]) || 0;
    const fabricValKG = Number(row[COLUMNS.FABRIC_KG]) || 0;
    const foamScrapsValKG = Number(row[COLUMNS.FOAM_SCRAPS_KG]) || 0;
    const usedOilsValLiters = Number(row[COLUMNS.USED_OILS_LITERS]) || 0;
    const chemicalFlashingValMT = Number(row[COLUMNS.CHEMICAL_FLASHING_MT]) || 0;
    const oilContaminatedRagsValKG = Number(row[COLUMNS.OIL_CONTAMINATED_RAGS_KG]) || 0;


    // --- Aggregate for Total Annual GHG Emission Section ---
    totalGHGEmission += totalActual;
    totalTargetLimit += totalTarget;

    if (!monthlyTotalGhgDataMap[month]) {
      monthlyTotalGhgDataMap[month] = { lastYear: 0, actual: 0, target: 0 };
    }
    monthlyTotalGhgDataMap[month].lastYear += totalLastYear;
    monthlyTotalGhgDataMap[month].actual += totalActual;
    monthlyTotalGhgDataMap[month].target += totalTarget;

    if (!plantTotalGhgDataMap[plant]) {
      plantTotalGhgDataMap[plant] = { lastYear: 0, actual: 0, target: 0 };
    }
    plantTotalGhgDataMap[plant].lastYear += totalLastYear;
    plantTotalGhgDataMap[plant].actual += totalActual;
    plantTotalGhgDataMap[plant].target += totalTarget;


    // --- Aggregate for Scope 1 Section ---
    scope1ActualTotal += s1ScorecardActual; // For scorecard
    scope1TargetTotal += s1ScorecardTarget; // For scorecard

    // For Scope 1 Category Pie Chart
    scope1MobileOnRoadTotalPie += s1MobileOnRoadPieVal;
    scope1MobileNonRoadTotalPie += s1MobileNonRoadPieVal;
    scope1StationaryTotalPie += s1StationaryTotalPieVal;

    // For 1st Subsection (Stationary Combustion)
    scope1StationaryTotalSub += stationaryTotalSubVal;
    scope1StationaryDieselTotal += stationaryDiesel;
    scope1StationaryLPGTotal += stationaryLPG;
    
    // For 2nd Subsection (Mobile Combustion Non-Road)
    mobileNonRoadTotalSub += mobileNonRoadTotalVal;
    mobileNonRoadDiesel += mobileNonRoadDieselVal;
    mobileNonRoadGasoline += mobileNonRoadGasolineVal;

    // For 3rd Subsection (Mobile Combustion On-Road)
    mobileOnRoadTotalSub += mobileOnRoadTotalVal;
    mobileOnRoadDiesel += mobileOnRoadDieselVal;
    mobileOnRoadGasoline += mobileOnRoadGasolineVal;

    // For Main GHG Emission Scope Category Contribution Pie Chart
    mainPieChartScope1Total += s1ScorecardActual; // Scope 1 Actual (CC) for main pie
    mainPieChartScope2Total += s2TotalForMainPie; // Scope 2 Total (CM) for main pie

    // Aggregate for Monthly Scope 1 Chart
    if (!monthlyScope1DataMap[month]) {
        monthlyScope1DataMap[month] = { lastYear: 0, actual: 0, target: 0 };
    }
    monthlyScope1DataMap[month].lastYear += s1LastYearBar; // DH
    monthlyScope1DataMap[month].actual += s1ActualBar; // CC
    monthlyScope1DataMap[month].target += s1TargetBar; // DL

    // Aggregate for Plant Scope 1 Chart (Now consistent with Monthly Scope 1 Chart)
    if (!plantScope1DataMap[plant]) {
        plantScope1DataMap[plant] = { lastYear: 0, actual: 0, target: 0 };
    }
    plantScope1DataMap[plant].lastYear += s1LastYearBar; // DH
    plantScope1DataMap[plant].actual += s1ActualBar; // CC
    plantScope1DataMap[plant].target += s1TargetBar; // DL

    // --- NEW: Aggregate for Scope 2 Section ---
    scope2ActualTotal += s2Actual;
    scope2TargetTotal += s2Target;
    scope2LastYearTotal += s2LastYear;
    scope2NonRenewableEnergyPie += s2NonRenewableTCO2eq;
    scope2RenewableEnergyPie += s2RenewableTCO2eq;
    scope2KWHNonRenewable += s2KWHNonRenewableVal;
    scope2KWHRenewable += s2KWHRenewableVal;

    if (!monthlyScope2DataMap[month]) {
      monthlyScope2DataMap[month] = { lastYear: 0, actual: 0, target: 0 };
    }
    monthlyScope2DataMap[month].lastYear += s2LastYear;
    monthlyScope2DataMap[month].actual += s2Actual;
    monthlyScope2DataMap[month].target += s2Target;

    if (!plantScope2DataMap[plant]) {
      plantScope2DataMap[plant] = { lastYear: 0, actual: 0, target: 0 };
    }
    plantScope2DataMap[plant].lastYear += s2LastYear;
    plantScope2DataMap[plant].actual += s2Actual;
    plantScope2DataMap[plant].target += s2Target;

    // NEW: Aggregate for Scope 3 Section
    scope3ActualTotal += s3Actual;

    totalWaterWithdrawn += waterWithdrawn;
    totalWaterTreatment += waterTreatment;

    mixedSolidWasteKG += msWasteKG;
    paperKG += paperValKG;
    plasticKG += plasticValKG; // CORRECTED LINE
    assortedMetalKG += assortedMetalValKG;
    fabricKG += fabricValKG;
    foamScrapsKG += foamScrapsValKG;
    usedOilsLiters += usedOilsValLiters;
    chemicalFlashingMT += chemicalFlashingValMT;
    oilContaminatedRagsKG += oilContaminatedRagsValKG;
  });

  // --- Processed Data for Total Annual GHG Emission Section ---
  const totalGhgPerformancePercentage = (totalTargetLimit !== 0) ? ((totalGHGEmission - totalTargetLimit) / totalTargetLimit) * 100 : 0;
  const totalGhgPerformanceText = Math.abs(totalGhgPerformancePercentage).toFixed(2) + "%";
  const totalGhgPerformanceColor = (totalGhgPerformancePercentage > 0) ? "red" : "green";
  let totalGhgArrowDirection = "&#9650;"; // Default to up arrow
  if (totalGhgPerformancePercentage < 0) totalGhgArrowDirection = "&#9660;"; // Down arrow if lower
  if (totalGhgPerformancePercentage === 0 && totalTargetLimit !== 0) totalGhgArrowDirection = "&#9679;"; // Dot for equal
  if (totalTargetLimit === 0) { // If no target, remove percentage and arrow
    totalGhgPerformanceText = "N/A";
    totalGhgPerformanceColor = "gray";
    totalGhgArrowDirection = "";
  }

  const sortedMonthlyTotalGhgData = Object.keys(monthlyTotalGhgDataMap).sort(sortByMonth).map(month => ({
    month: month,
    lastYear: monthlyTotalGhgDataMap[month].lastYear,
    actual: monthlyTotalGhgDataMap[month].actual,
    target: monthlyTotalGhgDataMap[month].target
  }));

  const sortedPlantTotalGhgData = Object.keys(plantTotalGhgDataMap).sort((a, b) => a.localeCompare(b)).map(plant => ({
    plant: plant,
    lastYear: plantTotalGhgDataMap[plant].lastYear,
    actual: plantTotalGhgDataMap[plant].actual,
    target: plantTotalGhgDataMap[plant].target
  }));
  
  // Calculate max values for Y-axis dynamic scaling for Total GHG Charts
  let maxMonthlyTotalGhgValue = 0;
  sortedMonthlyTotalGhgData.forEach(item => {
      maxMonthlyTotalGhgValue = Math.max(maxMonthlyTotalGhgValue, item.lastYear, item.actual, item.target);
  });

  let maxPlantTotalGhgValue = 0;
  sortedPlantTotalGhgData.forEach(item => {
      maxPlantTotalGhgValue = Math.max(maxPlantTotalGhgValue, item.lastYear, item.actual, item.target);
  });


  // --- Processed Data for Scope 1 Section ---
  const scope1PerformancePercentage = (scope1TargetTotal !== 0) ? ((scope1ActualTotal - scope1TargetTotal) / scope1TargetTotal) * 100 : 0;
  const scope1PerformanceText = Math.abs(scope1PerformancePercentage).toFixed(2) + "%";
  const scope1PerformanceColor = (scope1PerformancePercentage > 0) ? "red" : "green";
  let scope1ArrowDirection = "&#9650;"; // Default to up arrow
  if (scope1PerformancePercentage < 0) scope1ArrowDirection = "&#9660;"; // Down arrow if lower
  if (scope1PerformancePercentage === 0 && scope1TargetTotal !== 0) scope1ArrowDirection = "&#9679;"; // Dot for equal
  if (scope1TargetTotal === 0) { // If no target, remove percentage and arrow
    scope1PerformanceText = "N/A";
    scope1PerformanceColor = "gray";
    scope1ArrowDirection = "";
  }

  // Main Scope Category Contribution (for Section 1) - now from GHG Calculation sums
  const mainScopeCategoryContribution = [
    ['Scope', 'TCO2eq'],
    ['Scope 1', mainPieChartScope1Total],
    ['Scope 2', mainPieChartScope2Total],
    ['Scope 3', scope3ActualTotal],
  ];


  // Scope 1 Breakdown Pie Chart Data (for Scope 1 Section)
  const scope1BreakdownPieData = [
    ['Category', 'TCO2eq'],
    ['Mobile On-Road', scope1MobileOnRoadTotalPie],
    ['Mobile Non-Road', scope1MobileNonRoadTotalPie],
    ['Stationary', scope1StationaryTotalPie],
  ];

  // Monthly Scope 1 Chart Data
  const sortedMonthlyScope1Data = Object.keys(monthlyScope1DataMap).sort(sortByMonth).map(month => ({
    month: month,
    lastYear: monthlyScope1DataMap[month].lastYear,
    actual: monthlyScope1DataMap[month].actual,
    target: monthlyScope1DataMap[month].target
  }));

  // Plant Scope 1 Chart Data
  const sortedPlantScope1Data = Object.keys(plantScope1DataMap).sort((a, b) => a.localeCompare(b)).map(plant => ({
    plant: plant,
    lastYear: plantScope1DataMap[plant].lastYear,
    actual: plantScope1DataMap[plant].actual,
    target: plantScope1DataMap[plant].target
  }));

  // Calculate max values for Y-axis dynamic scaling for Scope 1 Charts
  let maxMonthlyScope1Value = 0;
  sortedMonthlyScope1Data.forEach(item => {
      maxMonthlyScope1Value = Math.max(maxMonthlyScope1Value, item.lastYear, item.actual, item.target);
  });

  let maxPlantScope1Value = 0;
  sortedPlantScope1Data.forEach(item => {
      maxPlantScope1Value = Math.max(maxPlantScope1Value, item.lastYear, item.actual, item.target);
  });

  // --- NEW: Processed Data for Scope 2 Section ---
  const scope2PerformancePercentage = (scope2TargetTotal !== 0) ? ((scope2ActualTotal - scope2TargetTotal) / scope2TargetTotal) * 100 : 0;
  const scope2PerformanceText = Math.abs(scope2PerformancePercentage).toFixed(2) + "%";
  const scope2PerformanceColor = (scope2PerformancePercentage > 0) ? "red" : "green"; // Red is worse (above target), Green is better (below target)
  let scope2ArrowDirection = "&#9650;"; // Default to up arrow
  if (scope2PerformancePercentage < 0) scope2ArrowDirection = "&#9660;"; // Down arrow if lower
  if (scope2PerformancePercentage === 0 && scope2TargetTotal !== 0) scope2ArrowDirection = "&#9679;"; // Dot for equal
  if (scope2TargetTotal === 0) { // If no target, remove percentage and arrow
    scope2PerformanceText = "N/A";
    scope2PerformanceColor = "gray";
    scope2ArrowDirection = "";
  }

  const scope2EnergyMixPieData = [
    ['Energy Type', 'TCO2eq'],
    ['Non-Renewable Energy', scope2NonRenewableEnergyPie],
    ['Renewable Energy', scope2RenewableEnergyPie]
  ];

  // Monthly Scope 2 Chart Data
  const sortedMonthlyScope2Data = Object.keys(monthlyScope2DataMap).sort(sortByMonth).map(month => ({
    month: month,
    lastYear: monthlyScope2DataMap[month].lastYear,
    actual: monthlyScope2DataMap[month].actual,
    target: monthlyScope2DataMap[month].target
  }));

  // Plant Scope 2 Chart Data
  const sortedPlantScope2Data = Object.keys(plantScope2DataMap).sort((a, b) => a.localeCompare(b)).map(plant => ({
    plant: plant,
    lastYear: plantScope2DataMap[plant].lastYear,
    actual: plantScope2DataMap[plant].actual,
    target: plantScope2DataMap[plant].target
  }));

  // Calculate max values for Y-axis dynamic scaling for Scope 2 Charts
  let maxMonthlyScope2Value = 0;
  sortedMonthlyScope2Data.forEach(item => {
      maxMonthlyScope2Value = Math.max(maxMonthlyScope2Value, item.lastYear, item.actual, item.target);
  });

  let maxPlantScope2Value = 0;
  sortedPlantScope2Data.forEach(item => {
      maxPlantScope2Value = Math.max(maxPlantScope2Value, item.lastYear, item.actual, item.target);
  });

  // NEW: Processed Data for Scope 3 Section
  // No performance percentage for Scope 3 as requested.
  const scope3PerformanceText = ""; // Empty string now
  const scope3PerformanceColor = "gray";
  const scope3ArrowDirection = "";


  return {
    // Total Annual GHG Emission Section Data
    totalGHGEmission: totalGHGEmission,
    totalGhgPerformanceText: totalGhgPerformanceText,
    totalGhgPerformanceColor: totalGhgPerformanceColor,
    totalGhgArrowDirection: totalGhgArrowDirection,
    mainScopeCategoryContribution: mainScopeCategoryContribution,
    monthlyTotalGhgData: sortedMonthlyTotalGhgData,
    maxMonthlyTotalGhgValue: maxMonthlyTotalGhgValue,
    plantTotalGhgData: sortedPlantTotalGhgData,
    maxPlantTotalGhgValue: maxPlantTotalGhgValue,

    // Scope 1 Section Data
    scope1ActualTotal: scope1ActualTotal,
    scope1TargetTotal: scope1TargetTotal,
    scope1PerformanceText: scope1PerformanceText,
    scope1PerformanceColor: scope1PerformanceColor,
    scope1ArrowDirection: scope1ArrowDirection,
    scope1BreakdownPieData: scope1BreakdownPieData,
    scope1StationaryTotalSub: scope1StationaryTotalSub,
    scope1StationaryDieselTotal: scope1StationaryDieselTotal,
    scope1StationaryLPGTotal: scope1StationaryLPGTotal,
    mobileNonRoadTotalSub: mobileNonRoadTotalSub,
    mobileNonRoadDiesel: mobileNonRoadDiesel,
    mobileNonRoadGasoline: mobileNonRoadGasoline,
    mobileOnRoadTotalSub: mobileOnRoadTotalSub,
    mobileOnRoadDiesel: mobileOnRoadDiesel,
    mobileOnRoadGasoline: mobileOnRoadGasoline,
    monthlyScope1Data: sortedMonthlyScope1Data,
    maxMonthlyScope1Value: maxMonthlyScope1Value,
    plantScope1Data: sortedPlantScope1Data,
    maxPlantScope1Value: maxPlantScope1Value,

    // NEW: Scope 2 Section Data
    scope2ActualTotal: scope2ActualTotal,
    scope2TargetTotal: scope2TargetTotal,
    scope2LastYearTotal: scope2LastYearTotal,
    scope2PerformanceText: scope2PerformanceText,
    scope2PerformanceColor: scope2PerformanceColor,
    scope2ArrowDirection: scope2ArrowDirection,
    scope2EnergyMixPieData: scope2EnergyMixPieData,
    scope2KWHNonRenewable: scope2KWHNonRenewable,
    scope2KWHRenewable: scope2KWHRenewable,
    monthlyScope2Data: sortedMonthlyScope2Data,
    maxMonthlyScope2Value: maxMonthlyScope2Value,
    plantScope2Data: sortedPlantScope2Data,
    maxPlantScope2Value: maxPlantScope2Value,

    // NEW: Scope 3 Section Data
    scope3ActualTotal: scope3ActualTotal,
    scope3PerformanceText: scope3PerformanceText, // Empty string now
    scope3PerformanceColor: scope3PerformanceColor, // Always gray as per request
    scope3ArrowDirection: scope3ArrowDirection, // Always empty as per request
    totalWaterWithdrawn: totalWaterWithdrawn,
    totalWaterTreatment: totalWaterTreatment,
    mixedSolidWasteKG: mixedSolidWasteKG,
    paperKG: paperKG,
    plasticKG: plasticKG,
    assortedMetalKG: assortedMetalKG,
    fabricKG: fabricKG,
    foamScrapsKG: foamScrapsKG,
    usedOilsLiters: usedOilsLiters,
    chemicalFlashingMT: chemicalFlashingMT,
    oilContaminatedRagsKG: oilContaminatedRagsKG
  };
}
