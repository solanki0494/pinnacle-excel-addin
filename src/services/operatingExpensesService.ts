export async function copyOperatingExpenses(context: Excel.RequestContext): Promise<void> {
  try {
    console.log("Starting operating expenses copy...");

    const outputsSheet = context.workbook.worksheets.getItem("Outputs");
    const engineeringSheet = context.workbook.worksheets.getItem("Software Engineer Cash Flow");

    const sourceRange = outputsSheet.getRange("O32:CI35");
    const targetRange = engineeringSheet.getRange("O32:CI35");
    
    sourceRange.load("values");
    await context.sync();
    
    targetRange.values = sourceRange.values;
    await context.sync();
    
    console.log("Operating expenses copied successfully from Outputs to Engineering sheet");

  } catch (error) {
    console.error("Error copying operating expenses:", error);
    throw error;
  }
}

export async function showNotification(title: string, message: string): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const notification = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: message,
        icon: "icon1",
        persistent: false
      };
      
      console.log(`${title}: ${message}`);
    });
  } catch (error) {
    console.log(`${title}: ${message}`);
  }
}
