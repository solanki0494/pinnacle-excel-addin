/* global console, Excel, Office */

import { copyOperatingExpenses, showNotification as serviceShowNotification } from '../services/operatingExpensesService';

async function runCalculation(event: Office.AddinCommands.Event) {
  try {
    console.log("Starting calculation from ribbon button...");

    await Excel.run(async (context) => {
      console.log("Excel.run context established");

      // Force automatic calculation mode
      context.application.calculationMode = Excel.CalculationMode.automatic;
      context.application.calculate(Excel.CalculationType.full);
      await context.sync();
      console.log("Calculation mode set and full calculation completed");

      // Execute the main copy operation
      await copyOperatingExpenses(context);
      console.log("Copy operation completed successfully");

      await serviceShowNotification("Calculation Complete", "Operating expenses copied successfully!");
    });

    // Signal completion to Office
    event.completed();
    console.log("Event completed successfully");

  } catch (error) {
    console.error("Error in runCalculation:", error);
    const errorMessage = error instanceof Error ? error.message : String(error);

    try {
      await serviceShowNotification("Calculation Error", `Error: ${errorMessage}`);
    } catch (notificationError) {
      console.error("Failed to show notification:", notificationError);
    }

    // Always complete the event, even on error
    event.completed();
  }
}

// Ensure the function is available globally for Office.js
(global as any).runCalculation = runCalculation;

// Also register it on window for desktop compatibility
if (typeof window !== 'undefined') {
  (window as any).runCalculation = runCalculation;
}

// Initialize when Office is ready
Office.onReady(() => {
  console.log("Office.js ready - commands.ts loaded");
  console.log("runCalculation function registered:", typeof runCalculation);
});
