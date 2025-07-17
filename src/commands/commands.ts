/* global console, Excel, Office */

import { copyOperatingExpenses, showNotification as serviceShowNotification } from '../services/operatingExpensesService';

// Declare global types for Office.js
declare global {
  interface Window {
    officeReady: boolean;
  }
}

// Initialize Office.js when ready
function initializeOffice() {
  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady(() => {
      console.log('Office.js is ready for commands');
      // Office is ready - commands will be available
    });
  } else {
    console.log('Office.js not ready for commands, retrying...');
    setTimeout(initializeOffice, 100);
  }
}

// Start initialization
initializeOffice();

async function runCalculation(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      console.log("Starting calculation from ribbon button...");

      context.application.calculationMode = Excel.CalculationMode.automatic;
      context.application.calculate(Excel.CalculationType.full);
      await context.sync();

      await new Promise(resolve => setTimeout(resolve, 500));

      await copyOperatingExpenses(context);

      await serviceShowNotification("Calculation Complete", "Operating expenses copied successfully!");

      event.completed();
      
    });
  } catch (error) {
    console.error("Error in runCalculation:", error);
    const errorMessage = error instanceof Error ? error.message : String(error);
    await serviceShowNotification("Calculation Error", `Error: ${errorMessage}`);
    event.completed();
  }
}

(global as any).runCalculation = runCalculation;
