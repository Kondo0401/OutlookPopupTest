Office.onReady(() => {
    console.log("Office Ready - window.open rollback");
});

function openHQSite(event) {
    console.log("Button Click");

    try {
        const popup = window.open(
            "https://www.yahoo.co.jp",
            "_blank"
        );

        if (popup === null) {
            console.warn("Popup blocked");
        } else {
            console.log("Popup opened");
        }
    } catch (error) {
        console.error("Failed to open popup", error);
    } finally {
        event.completed();
    }
}
