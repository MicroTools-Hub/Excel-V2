Office.onReady(() => {
  Office.actions.associate("openSmartExcelCopilot", () => {
    return Office.addin.showAsTaskpane();
  });
});
