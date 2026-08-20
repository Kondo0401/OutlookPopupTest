Office.onReady(() => {
});

/*
  Outlookボタン押下時
*/
function openHQSite(event) {

    console.log("Button Click");

    try {

        window.open(
            "https://www.microsoft.com",
            "_blank"
        );

        console.log("window.open executed");

    } catch (e) {

        console.error(e);

    }

    event.completed();
}