Office.onReady(() => {
    console.log("Office Ready");
});

function openHQSite(event) {
    console.log("Button Click");

    window.open(
        "https://www.yahoo.co.jp",
        "_blank"
    );

    event.completed();
}
