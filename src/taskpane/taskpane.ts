/* global console, document, Excel, Office */

import { copyOperatingExpenses } from '../services/operatingExpensesService';

// Declare global types for Office.js
declare global {
  interface Window {
    officeReady: boolean;
  }
}

// Wait for both DOM and Office.js to be ready
function initializeApp() {
  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        console.log('Office.js and Excel are ready');
        document.getElementById("runCalculation")!.onclick = runTaskpaneCalculation;
      }
    });
  } else {
    console.log('Office.js not ready, retrying...');
    setTimeout(initializeApp, 100);
  }
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initializeApp);
} else {
  initializeApp();
}

async function runTaskpaneCalculation() {
  try {
    await Excel.run(async (context) => {
      showStatus("Running calculation...", "info");

      context.application.calculationMode = Excel.CalculationMode.automatic;
      context.application.calculate(Excel.CalculationType.full);

      await copyOperatingExpenses(context);

      showStatus("Calculation completed successfully!", "success");

    });
  } catch (error) {
    console.error(error);
    const errorMessage = error instanceof Error ? error.message : String(error);
    showStatus(`Error: ${errorMessage}`, "error");
  }
}

function showStatus(message: string, type: "success" | "error" | "info") {
  const statusDiv = document.getElementById("status")!;
  statusDiv.textContent = message;
  statusDiv.className = `status ${type}`;
  statusDiv.style.display = "block";
  
  if (type !== "error") {
    setTimeout(() => {
      statusDiv.style.display = "none";
    }, 5000);
  }
}

