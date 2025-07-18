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

    // Check if Software Engineer Cash Flow sheet exists, if not create it
    let targetSheet: Excel.Worksheet;
    let sheetExists = false;

    const engineeringSheetExists = worksheets.items.some(sheet => sheet.name === "Software Engineer Cash Flow");

    if (engineeringSheetExists) {
      targetSheet = context.workbook.worksheets.getItem("Software Engineer Cash Flow");
      sheetExists = true;
      console.log("Software Engineer Cash Flow sheet exists, will clear and override");
    } else {
      console.log("Creating new Software Engineer Cash Flow sheet");
      targetSheet = context.workbook.worksheets.add("Software Engineer Cash Flow");
      sheetExists = false;
    }

    await context.sync();

    // Clear the target sheet if it exists
    if (sheetExists) {
      const usedRange = targetSheet.getUsedRange();
      if (usedRange) {
        usedRange.clear();
        await context.sync();
        console.log("Cleared existing content from Software Engineer Cash Flow sheet");
      }
    }

    // Get the used range from template sheet to copy structure and values
    const templateUsedRange = templateSheet.getUsedRange();

    // Check if the template sheet has any content
    if (!templateUsedRange) {
      throw new Error("The 'Outputs' sheet appears to be empty. Please add some content to copy.");
    }

    templateUsedRange.load(["values", "formulas", "format", "rowCount", "columnCount"]);
    await context.sync();

    console.log(`Template range size: ${templateUsedRange.rowCount} rows x ${templateUsedRange.columnCount} columns`);

    // Create target range of same size starting from A1
    const targetRange = targetSheet.getRangeByIndexes(0, 0, templateUsedRange.rowCount, templateUsedRange.columnCount);

    try {
      // First copy all formatting from template to maintain appearance
      await copyFormatting(templateUsedRange, targetRange, context);

      // Then copy calculated values (not formulas) from template
      targetRange.values = templateUsedRange.values;

      // Copy column widths and row heights for exact layout match
      await copyDimensions(templateSheet, targetSheet, templateUsedRange, context);

      await context.sync();
    } catch (rangeError) {
      console.error("Error during range operations:", rangeError);
      throw new Error(`Failed to copy data: ${rangeError instanceof Error ? rangeError.message : String(rangeError)}`);
    }

    console.log("Successfully copied template structure, formatting, and calculated values to Software Engineer Cash Flow sheet");

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

async function copyDimensions(sourceSheet: Excel.Worksheet, targetSheet: Excel.Worksheet, usedRange: Excel.Range, context: Excel.RequestContext): Promise<void> {
  try {
    // Copy column widths
    for (let col = 0; col < usedRange.columnCount; col++) {
      const sourceColumn = sourceSheet.getRange(`${String.fromCharCode(65 + col)}:${String.fromCharCode(65 + col)}`);
      const targetColumn = targetSheet.getRange(`${String.fromCharCode(65 + col)}:${String.fromCharCode(65 + col)}`);

      sourceColumn.format.load("columnWidth");
      await context.sync();

      targetColumn.format.columnWidth = sourceColumn.format.columnWidth;
    }

    // Copy row heights
    for (let row = 0; row < usedRange.rowCount; row++) {
      const sourceRow = sourceSheet.getRange(`${row + 1}:${row + 1}`);
      const targetRow = targetSheet.getRange(`${row + 1}:${row + 1}`);

      sourceRow.format.load("rowHeight");
      await context.sync();

      targetRow.format.rowHeight = sourceRow.format.rowHeight;
    }

    await context.sync();
    console.log("Column widths and row heights copied successfully");

  } catch (error) {
    console.warn("Could not copy all dimensions, continuing:", error);
  }
}

async function copyFormatting(sourceRange: Excel.Range, targetRange: Excel.Range, context: Excel.RequestContext): Promise<void> {
  try {
    // Load formatting properties from source
    sourceRange.format.load([
      "columnWidth", "rowHeight", "horizontalAlignment", "verticalAlignment",
      "wrapText", "textOrientation", "shrinkToFit", "readingOrder",
      "borders", "fill", "font", "protection"
    ]);

    await context.sync();

    // Copy basic formatting properties
    targetRange.format.horizontalAlignment = sourceRange.format.horizontalAlignment;
    targetRange.format.verticalAlignment = sourceRange.format.verticalAlignment;
    targetRange.format.wrapText = sourceRange.format.wrapText;
    targetRange.format.textOrientation = sourceRange.format.textOrientation;
    targetRange.format.shrinkToFit = sourceRange.format.shrinkToFit;
    targetRange.format.readingOrder = sourceRange.format.readingOrder;

    // Copy font formatting
    sourceRange.format.font.load(["name", "size", "bold", "italic", "underline", "color"]);
    await context.sync();

    targetRange.format.font.name = sourceRange.format.font.name;
    targetRange.format.font.size = sourceRange.format.font.size;
    targetRange.format.font.bold = sourceRange.format.font.bold;
    targetRange.format.font.italic = sourceRange.format.font.italic;
    targetRange.format.font.underline = sourceRange.format.font.underline;
    targetRange.format.font.color = sourceRange.format.font.color;

    // Copy fill formatting
    sourceRange.format.fill.load(["color", "pattern"]);
    await context.sync();

    targetRange.format.fill.color = sourceRange.format.fill.color;
    targetRange.format.fill.pattern = sourceRange.format.fill.pattern;

    await context.sync();
    console.log("Formatting copied successfully");

  } catch (error) {
    console.warn("Could not copy all formatting, continuing with values:", error);
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
