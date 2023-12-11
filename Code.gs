// Global variables
const sheetNameOverview = "Overview";
const sheetNameIndividualExpenses = "User Expenses";
const sheetNameReceipts = "Receipts";


// Expense categories and colors
const categories = ["Donations", "Operational", "Events"];
const colors = {
  Donations: "#43A047",
  Operational: "#F0E68C",
  Events: "#E91E63",
};

// Minimum and maximum amount for random data
const minAmount = 5;
const maxAmount = 100;
const MIN_DATE = new Date(2023, 11, 1);

// Function to generate random date for an expense
function getRandomDate(maxDate = new Date(2023, 11, 31)) {
  const minDateMs = MIN_DATE.getTime();
  const maxDateMs = maxDate.getTime();
  const randomDateMs = Math.floor(Math.random() * (maxDateMs - minDateMs + 1)) + minDateMs;
  return new Date(randomDateMs);
}

// Function to generate random expense data for individual sheets
function generateRandomExpenseData(numTransactions) {
  const expenseData = [];
  for (let i = 0; i < numTransactions; i++) {
    const category = categories[Math.floor(Math.random() * categories.length)];
    expenseData.push({
      date: getRandomDate(new Date(2023, 11, 1), new Date(2023, 11, 31)),
      description: "Transaction " + (i + 1),
      amount: Math.floor(Math.random() * (maxAmount - minAmount + 1)) + minAmount,
      category,
      paymentMethodType: "Cash",  // Placeholder for future implementation
    });
  }
  return expenseData;
}

// Function to calculate total expenses per category
function getTotalExpensesPerCategory(expenses) {
  const totalExpensesPerCategory = {};
  for (const category of categories) {
    totalExpensesPerCategory[category] = 0;
  }
  for (const expense of expenses) {
    // Assuming the 4th element in expense array is the category, and the 3rd is the amount
    const category = expense[3]; // Category
    const amount = expense[2]; // Amount
    if (totalExpensesPerCategory.hasOwnProperty(category)) {
      totalExpensesPerCategory[category] += amount;
    }
  }
  return totalExpensesPerCategory;
}

// Function to generate random receipt data
function generateRandomReceiptData(numReceipts) {
  const receiptData = [];
  for (let i = 0; i < numReceipts; i++) {
    receiptData.push({
      date: getRandomDate(new Date(2023, 11, 1), new Date(2023, 11, 31)),
      category: categories[Math.floor(Math.random() * categories.length)],
      amount: Math.floor(Math.random() * (maxAmount - minAmount + 1)) + minAmount,
      imageLink: "https://placehold.it/100x100",  // Placeholder for uploaded image
    });
  }
  return receiptData;
}

/*
// Create overview sheet if it doesn't exist
function createOverviewSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi(); // Get the user interface object

    if (!spreadsheet.getSheetByName(sheetNameOverview)) {
      const sheet = spreadsheet.insertSheet(sheetNameOverview);

      // Set header row and format
      const headerRow = sheet.getRange(1, 1, 1, 4);
      headerRow.setBackground("#F2F2F2");
      headerRow.setValue("Vision:");
      headerRow.offset(0, 1).setValue("Goals:");
      headerRow.offset(0, 2).setValue("To-Dos:");

      // Prompt for user input
      const visionResponse = ui.prompt('Enter your vision statement:');
      const goalsResponse = ui.prompt('Enter your key financial goals:');
      const todosResponse = ui.prompt('Enter your top 3 to-dos (separated by commas):');

      if (visionResponse.getSelectedButton() == ui.Button.OK) {
          sheet.getRange(2, 1).setValue(visionResponse.getResponseText());
      }
      if (goalsResponse.getSelectedButton() == ui.Button.OK) {
          sheet.getRange(4, 1).setValue(goalsResponse.getResponseText());
      }
      if (todosResponse.getSelectedButton() == ui.Button.OK) {
          sheet.getRange(6, 1).setValue("To-Dos:");
          sheet.getRange(7, 1).setValue(todosResponse.getResponseText().split(",").join("\n"));
      }
    }
  } catch (error) {
    console.error(error.message);
  }
}
*/

/////////////////////////////////////////////////////////
function createOverviewSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let overviewSheet = spreadsheet.getSheetByName(sheetNameOverview);

  if (!overviewSheet) {
    // Create the Overview sheet if it doesn't exist
    overviewSheet = spreadsheet.insertSheet(sheetNameOverview);
    initializeOverviewSheet(overviewSheet);
  }

  // Additional setup or updates can be added here
  // For now, this is minimal to avoid breaking existing structure
}

function initializeOverviewSheet(sheet) {
  // Initial setup for the Overview sheet
  // Set up headers, layout, and initial content
  sheet.getRange('A1').setValue('Foundation Vision');
  sheet.getRange('B1').setValue('Temple Goals');
  sheet.getRange('C1').setValue('To-Do Items');

  // More detailed setup can be added here in future iterations
}
///////////////////////////////////////////////////////////

function createIndividualExpenseSheets() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    console.log("Accessing the active spreadsheet for creating individual expense sheets.");

    let numSheets = 3; 
    console.log("Number of individual expense sheets to create: " + numSheets);

    for (let i = 1; i <= numSheets; i++) {
      let sheetName = sheetNameIndividualExpenses + " " + i;
      console.log("Checking for the existence of sheet: " + sheetName);

      if (!spreadsheet.getSheetByName(sheetName)) {
        console.log("Creating new sheet: " + sheetName);
        const sheet = spreadsheet.insertSheet(sheetName);

        console.log("Setting header for " + sheetName);
        const headerRow = sheet.getRange(1, 1, 1, 5);
        headerRow.setBackground("#F2F2F2");
        headerRow.setValues([["Date", "Description", "Amount", "Category", "Payment Method"]]);

        console.log("Populating " + sheetName + " with random data.");
        const randomData = generateRandomExpenseData(10).map(row => [row.date, row.description, row.amount, row.category, row.paymentMethodType]);
        sheet.getRange(2, 1, randomData.length, 5).setValues(randomData);

        console.log("Applying color formatting based on category in " + sheetName);
        for (let row = 2; row <= randomData.length + 1; row++) {
          const categoryCell = sheet.getRange(row, 4);
          const category = categoryCell.getValue();
          if (category in colors) {
            console.log("Applying color for category: " + category + " in row: " + row);
            categoryCell.setBackground(colors[category]);
          } else {
            console.log("Category not found for color formatting in row " + row + ": " + category);
          }
        }
      } else {
        console.log("Sheet already exists: " + sheetName);
      }
    }

    console.log("Finished creating and populating individual expense sheets.");

  } catch (error) {
    console.error("Error in createIndividualExpenseSheets: " + error.message);
  }
}


// Create individual expense sheets if they don't exist
function createIndividualExpenseSheetsOld() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const existingSheets = spreadsheet.getSheets().filter(sheet => sheet.getName().includes(sheetNameIndividualExpenses));

    
    let numSheets = 3; // // For example, create 3 fixed number of sheets for testing purpose.
    
    // ####### Ideally:Generate random number of sheets if none exist (Useless in current scenario, numSheet fixed above) #######
    if (!existingSheets.length) {
      numSheets = Math.floor(Math.random() * 3) + 1;
    }

    // Create and populate individual expense sheets
    for (let i = 1; i <= numSheets; i++) {
      let sheetName = sheetNameIndividualExpenses + " " + i;
      if (!existingSheets.some(sheet => sheet.getName() === sheetName)) {
        const sheet = spreadsheet.insertSheet(sheetName);

        // Set header row and format
        const headerRow = sheet.getRange(1, 1, 1, 5);
        headerRow.setBackground("#F2F2F2");
        headerRow.setValue("Date").setFontColor("#000000");
        headerRow.offset(0, 1).setValue("Description").setFontColor("#000000");
        headerRow.offset(0, 2).setValue("Amount").setFontColor("#000000");
        headerRow.offset(0, 3).setValue("Category").setFontColor("#000000");
        headerRow.offset(0, 4).setValue("Payment Method").setFontColor("#000000");

        // Generate and populate sheet with random data
        const randomData = generateRandomExpenseData(10);
        const dataRange = sheet.getRange(2, 1, randomData.length, 5);
        dataRange.setValues(randomData.map(data => [data.date, data.description, data.amount, data.category, data.paymentMethodType]));

        // Apply color formatting based on category
        for (let row = 1; row <= dataRange.getNumRows() + 1; row++) {
          const categoryCell = dataRange.getCell(row, 4);
          const category = categoryCell.getValue();
          if (category in colors) {
            categoryCell.setBackground(colors[category]);
          }
        }
      }
    }
  } catch (error) {
    console.error(error.message);
  }
}

// Create the "Receipts" sheet if it doesn't exist
function createReceiptsSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet.getSheetByName(sheetNameReceipts)) {
      spreadsheet.insertSheet(sheetNameReceipts);
    }
  } catch (error) {
    console.error(error.message);
  }
}

//////////////////////////////////////////////
function updateOverviewSheet() {
  // Existing logic for updating the Overview sheet
  // This includes processing data from individual expense sheets and updating the overview

  // Additional logic for updating vision, goals, etc., can be incorporated here
  // For now, keeping it minimal to ensure stability

  // Placeholder for future enhancements:
  // updateOverviewContent();
}

// Placeholder for a future function to update content like vision, goals, and to-dos
function updateOverviewContent() {
  // Future logic to dynamically update content based on new data or inputs
}

//////////////////////////////////////////////

/*
function updateOverviewSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    console.log("Accessing the active spreadsheet.");

    let overviewSheet = spreadsheet.getSheetByName(sheetNameOverview);

    // Check and initialize Overview sheet if it does not exist
    if (!overviewSheet) {
      console.log("Overview sheet does not exist. Creating a new one.");
      overviewSheet = spreadsheet.insertSheet(sheetNameOverview);
      overviewSheet.getRange(1, 1, 1, 4).setValues([["Vision:", "Goals:", "To-Dos:", ""]]);
      overviewSheet.getRange(11, 1, 1, 2).setValues([["Category", "Total Expenses"]]);
    }

    // Initialize Overview sheet if it's blank
    if (overviewSheet.getLastRow() < 11) {
      console.log("Initializing blank Overview sheet.");
      overviewSheet.getRange(1, 1, 1, 4).setBackground("#F2F2F2").setValues([["Vision:", "Goals:", "To-Dos:", ""]]);
      overviewSheet.getRange(11, 1, 1, 2).setValues([["Category", "Total Expenses"]]);
    }

    // Collect data from individual expense sheets
    const expenses = [];
    const individualSheets = spreadsheet.getSheets().filter(sheet => sheet.getName().includes(sheetNameIndividualExpenses));
    console.log("Found " + individualSheets.length + " individual expense sheets.");

    individualSheets.forEach(sheet => {
      if (sheet.getLastRow() > 1) {
        console.log("Processing sheet: " + sheet.getName());
        const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5);
        expenses.push(...dataRange.getValues());
      } else {
        console.log("Skipping empty sheet: " + sheet.getName());
      }
    });

    // Calculate total expenses per category
    console.log("Calculating total expenses per category.");
    const totalExpensesPerCategory = getTotalExpensesPerCategory(expenses);
    console.log("Total expenses per category: " + JSON.stringify(totalExpensesPerCategory));

    // Update total expenses per category on the Overview sheet
    console.log("Updating total expenses per category on the Overview sheet.");
    for (let i = 0; i < categories.length; i++) {
      const row = 12 + i;
      overviewSheet.getRange(row, 1).setValue(categories[i]);
      overviewSheet.getRange(row, 2).setValue(totalExpensesPerCategory[categories[i]] || 0);
    }

    // Creating or updating the pie chart
    console.log("Creating or updating the pie chart.");
    const chartRange = overviewSheet.getRange(11, 1, categories.length + 1, 2);
    const charts = overviewSheet.getCharts();
    if (charts.length > 0) {
      console.log("Updating existing chart.");
      const existingChart = charts[0];
      const modifiedChart = existingChart.modify()
        .setOption('title', 'Total Expenses by Category')
        .setPosition(2, 6, 0, 0)
        .asPieChart()
        .addRange(chartRange)
        .build();
      overviewSheet.updateChart(modifiedChart);
    } else {
      console.log("Creating new chart.");
      const pieChart = overviewSheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(chartRange)
        .setPosition(2, 6, 0, 0)
        .setOption('title', 'Total Expenses by Category')
        .build();
      overviewSheet.insertChart(pieChart);
    }

    console.log("Finished updating the Overview sheet.");

  } catch (error) {
    console.error("Error in updateOverviewSheet: " + error.message);
  }
}
*/

// Calculate and update total expenses on overview sheet
function updateOverviewSheetOld() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const overviewSheet = spreadsheet.getSheetByName(sheetNameOverview);

    // Validate existence of the Overview sheet
    if (!overviewSheet) {
      console.error("Overview sheet does not exist.");
      return;
    }

    // Check if the Overview sheet is blank and initialize it if necessary
    if (overviewSheet.getLastRow() < 11) {
      // Initialize the Overview sheet if it's blank
      const headerRow = overviewSheet.getRange(1, 1, 1, 4);
      headerRow.setBackground("#F2F2F2");
      headerRow.setValues([["Vision:", "Goals:", "To-Dos:", ""]]);
      overviewSheet.getRange(11, 1, 1, 2).setValues([["Category", "Total Expenses"]]);
    }

    // Get data from individual expense sheets
    const expenses = [];
    const individualSheets = spreadsheet.getSheets().filter(sheet => sheet.getName().includes(sheetNameIndividualExpenses));
    console.log("Number of individual sheets: " + individualSheets.length);
   
    // Check if there are individual sheets with data
    if (individualSheets.length === 0 || !individualSheets.some(sheet => sheet.getLastRow() > 1)) {
      console.error("No data in individual expense sheets.");
      return; // Exit the function if there are no expense sheets with data
    }
   
    for (const sheet of individualSheets) {
      console.log("Processing sheet: " + sheet.getName() + " with rows: " + sheet.getLastRow());
      if (sheet.getLastRow() > 1) {
        const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5);
        const sheetValues = dataRange.getValues();
        if (sheetValues.length > 0) {
          expenses.push(...sheetValues);
        }
      }
    }

    // Calculate total expenses per category
    const totalExpensesPerCategory = getTotalExpensesPerCategory(expenses);

    // Clear and set the header for the categories and expenses
    const categoryHeaderRange = overviewSheet.getRange(11, 1, 1, 2);
    categoryHeaderRange.setValues([["Category", "Total Expenses"]]);

    // Update total expenses per category
    for (let i = 0; i < categories.length; i++) {
      const row = 12 + i; // Start from row 12 to preserve the header
      overviewSheet.getRange(row, 1).setValue(categories[i]);
      overviewSheet.getRange(row, 2).setValue(totalExpensesPerCategory[categories[i]] || 0); // Default to 0 if no expense
    }

    // Check if there's enough data to create/update a chart
    if (individualSheets.length > 0 && overviewSheet.getLastRow() >= 12) {
      // Define the chart range
      const chartRange = overviewSheet.getRange(11, 1, categories.length + 1, 2);
      const charts = overviewSheet.getCharts();

      // Create or update the pie chart
      if (charts.length > 0) {
        const existingChart = charts[0];
        const modifiedChart = existingChart.modify()
          .setOption('title', 'Total Expenses by Category')
          .setPosition(2, 6, 0, 0)
          .asPieChart()
          .addRange(chartRange)
          .build();
        overviewSheet.updateChart(modifiedChart);
      } else {
        const pieChart = overviewSheet.newChart()
          .setChartType(Charts.ChartType.PIE)
          .addRange(chartRange)
          .setPosition(2, 6, 0, 0)
          .setOption('title', 'Total Expenses by Category')
          .build();
        overviewSheet.insertChart(pieChart);
      }
    } else {
      console.error("Not enough data to create a chart.");
    }

    // Add additional visualizations and summaries as needed
    // Placeholder for future implementation of additional charts and data summaries

  } catch (error) {
    console.error("Error in updateOverviewSheet: " + error.message);
  }
}


// Function to create a menu for user interaction
function createMenu() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Temple Budget Management');

  // Add menu items here as needed, for example:
  menu.addItem('Update Overview', 'updateOverviewSheet');
  menu.addItem('Manage Receipts', 'createReceiptsSheet');

  // Adding the menu to the UI
  menu.addToUi();
}

// Placeholder for future implementation of user authentication and access control
// ...

// Budget planning functionality
function createBudgetPlan() {
  // Placeholder for future implementation of budget planning features
  // ...
}

// Reporting functionality
function generateReport() {
  // Placeholder for future implementation of report generation features
  // ...
}

// Data export functionality
function exportData() {
  // Placeholder for future implementation of data export features
  // ...
}

// Placeholder for future preparation for major events
// ...

// Placeholder for future implementation of GPT integration
// ...

// Placeholder for future implementation of data security measures and system maintenance features
// ...

// Run script automatically when spreadsheet is opened
function onOpen() {
  createOverviewSheet();
  createIndividualExpenseSheets();
  createReceiptsSheet();
  updateOverviewSheet();
  createMenu(); // Ensure this is called to create the custom menu
  // Placeholder for future integration of additional features
}





