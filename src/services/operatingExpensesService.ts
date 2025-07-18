export async function copyOperatingExpenses(context: Excel.RequestContext): Promise<void> {
  try {
    console.log("Starting template copy from Outputs to Software Engineer Cash Flow...");

    // Get the template sheet (Outputs)
    const templateSheet = context.workbook.worksheets.getItem("Outputs");

    // Check if Software Engineer Cash Flow sheet exists, if not create it
    let targetSheet: Excel.Worksheet;
    try {
      targetSheet = context.workbook.worksheets.getItem("Software Engineer Cash Flow");
      console.log("Software Engineer Cash Flow sheet exists, will clear and override");
    } catch (error) {
      console.log("Creating new Software Engineer Cash Flow sheet");
      targetSheet = context.workbook.worksheets.add("Software Engineer Cash Flow");
    }

    await context.sync();

    // Clear the target sheet if it exists
    const usedRange = targetSheet.getUsedRange();
    if (usedRange) {
      usedRange.clear();
      await context.sync();
      console.log("Cleared existing content from Software Engineer Cash Flow sheet");
    }

    // Get the used range from template sheet to copy structure and values
    const templateUsedRange = templateSheet.getUsedRange();
    templateUsedRange.load(["values", "formulas", "format", "rowCount", "columnCount"]);
    await context.sync();

    console.log(`Template range size: ${templateUsedRange.rowCount} rows x ${templateUsedRange.columnCount} columns`);

    // Create target range of same size starting from A1
    const targetRange = targetSheet.getRangeByIndexes(0, 0, templateUsedRange.rowCount, templateUsedRange.columnCount);

    // Copy calculated values (not formulas) from template
    targetRange.values = templateUsedRange.values;

    await context.sync();

    console.log("Successfully copied template structure and calculated values to Software Engineer Cash Flow sheet");

  } catch (error) {
    console.error("Error copying template:", error);
    throw error;
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
