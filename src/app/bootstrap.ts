Office.onReady(async (info) => {
    if (info.host === Office.HostType.PowerPoint) {
        const taskpane = await import("./taskpane");
        taskpane.initializeTaskPaneListener();
    }
});