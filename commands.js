Office.onReady(() => {
    console.log("Office Ready");
});

function openHQSite(event) {
    console.log("Popup Test clicked");

    const dialogUrl = "https://kondo0401.github.io/OutlookPopupTest/redirect.html";

    try {
        Office.context.ui.displayDialogAsync(
            dialogUrl,
            {
                height: 80,
                width: 80,
                displayInIframe: false
            },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error(
                        "Dialog open failed",
                        result.error.code,
                        result.error.message
                    );
                } else {
                    console.log("Office dialog opened");

                    const dialog = result.value;
                    dialog.addEventHandler(
                        Office.EventType.DialogEventReceived,
                        (args) => {
                            console.warn("Dialog event received", args.error);
                        }
                    );
                }
            }
        );
    } catch (error) {
        console.error("Unexpected dialog error", error);
    } finally {
        event.completed();
    }
}
