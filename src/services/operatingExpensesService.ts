export async function copyOperatingExpenses(context: Excel.RequestContext): Promise<void> {
  try {
    console.log("Starting template copy from Outputs to Software Engineer Cash Flow...");

    // First, let's check what sheets exist in the workbook
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    console.log("Available sheets:", worksheets.items.map(sheet => sheet.name));

    // Check if Outputs sheet exists
    const outputsSheetExists = worksheets.items.some(sheet => sheet.name === "Outputs");
    if (!outputsSheetExists) {
      throw new Error("The 'Outputs' sheet was not found. Please make sure the sheet exists and is named exactly 'Outputs'.");
    }

    // Get the template sheet (Outputs)
    const templateSheet = context.workbook.worksheets.getItem("Outputs");

    // Check if Software Engineer Cash Flow sheet exists (we'll delete and recreate for clean copy)
    const engineeringSheetExists = worksheets.items.some(sheet => sheet.name === "Software Engineer Cash Flow");
    let sheetExists = false;

    if (engineeringSheetExists) {
      console.log("Software Engineer Cash Flow sheet exists, will be replaced with fresh copy");
      sheetExists = true;
    } else {
      console.log("Software Engineer Cash Flow sheet does not exist, will create fresh copy");
      sheetExists = false;
    }

    // Simple approach: Copy the entire Outputs sheet and then convert formulas to values
    console.log("Using sheet copy approach for perfect formatting preservation");

    // Check if the Outputs sheet is hidden before copying
    templateSheet.load("visibility");
    await context.sync();

    const isOutputsHidden = templateSheet.visibility === Excel.SheetVisibility.hidden;

    // If the target sheet already exists, delete it first
    if (sheetExists) {
      console.log("Deleting existing Software Engineer Cash Flow sheet");
      const existingSheet = context.workbook.worksheets.getItem("Software Engineer Cash Flow");
      existingSheet.delete();
      await context.sync();
    }
    console.log(`Outputs sheet visibility: ${templateSheet.visibility} (hidden: ${isOutputsHidden})`);

    // Copy the entire Outputs sheet
    console.log("Copying Outputs sheet...");
    const copiedSheet = templateSheet.copy(Excel.WorksheetPositionType.after, templateSheet);
    copiedSheet.name = "Software Engineer Cash Flow";

    // If the source sheet was hidden, make sure the copied sheet is visible
    if (isOutputsHidden) {
      console.log("Source sheet was hidden, ensuring Software Engineer Cash Flow sheet is visible");
      copiedSheet.visibility = Excel.SheetVisibility.visible;
    }

    await context.sync();

    console.log("Sheet copied successfully, now converting formulas to values...");

    // Get the used range from the copied sheet and convert formulas to values
    const copiedUsedRange = copiedSheet.getUsedRange();
    if (copiedUsedRange) {
      copiedUsedRange.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      console.log(`Converting formulas to values for range: ${copiedUsedRange.rowCount} rows x ${copiedUsedRange.columnCount} columns`);

      // This converts all formulas to their calculated values
      copiedUsedRange.copyFrom(copiedUsedRange, Excel.RangeCopyType.values, false, false);
      await context.sync();

      console.log("Successfully converted all formulas to values");
    } else {
      console.warn("No used range found in copied sheet");
    }

    // Final verification - check that the sheet exists in the workbook
    const finalWorksheets = context.workbook.worksheets;
    finalWorksheets.load("items/name");
    await context.sync();

    const finalSheetExists = finalWorksheets.items.some(sheet => sheet.name === "Software Engineer Cash Flow");
    if (!finalSheetExists) {
      throw new Error("Software Engineer Cash Flow sheet was not found after copying - this should not happen");
    }

    // Activate the newly created Software Engineer Cash Flow sheet
    const finalSheet = context.workbook.worksheets.getItem("Software Engineer Cash Flow");
    finalSheet.activate();
    await context.sync();

    console.log("✅ Successfully copied Outputs sheet to Software Engineer Cash Flow with perfect formatting");
    console.log("✅ All formulas converted to calculated values");
    console.log("✅ Colors, borders, fonts, and layout preserved exactly");
    console.log("✅ Software Engineer Cash Flow sheet is now active and visible");
    if (isOutputsHidden) {
      console.log("✅ Source sheet was hidden, but result sheet is kept visible for user access");
    }

  } catch (error) {
    console.error("Error copying template:", error);

    // Provide more helpful error messages
    if (error instanceof Error) {
      if (error.message.includes("doesn't exist")) {
        throw new Error("Sheet not found. Please ensure both 'Outputs' sheet exists in your workbook.");
      } else if (error.message.includes("empty")) {
        throw error; // Re-throw our custom empty sheet message
      } else {
        throw new Error(`Template copy failed: ${error.message}`);
      }
    } else {
      throw new Error("An unexpected error occurred while copying the template.");
    }
  }
}



export async function showNotification(title: string, message: string): Promise<void> {
  try {
    console.log(`${title}: ${message}`);
    // For now, just log to console. In a full implementation,
    // you could show Excel notifications or task pane messages
  } catch (error) {
    console.log(`${title}: ${message}`);
  }
}
