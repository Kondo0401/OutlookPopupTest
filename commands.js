Office.onReady(() => {
    console.log("Office Ready");
});

function openHQSite(event) {
    console.log("Button Click");

    const popup = window.open(
        "https://www.yahoo.co.jp",
        "_blank"
    );

    if (popup === null) {
        console.warn("Popup blocked");
    } else {
        console.log("Popup opened");
    }

    event.completed();
}
