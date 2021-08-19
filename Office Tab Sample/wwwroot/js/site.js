// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.

let loginDialog = null;


const onReadyTimer = new Date().getTime();

Office.onReady(async () => {

    let timerResult = new Date().getTime() - onReadyTimer;

    $("#timings").html("<div>Office.OnReady = " + timerResult + "ms</div>");

    const getAccessTokenTimer = new Date().getTime();

    try {

        let token = await OfficeRuntime.auth.getAccessToken({ allowConsentPrompt: true, allowSignInPrompt: true }); //Time this...

        doSomethingWithToken("SSO", token);
    }
    catch (ex) {
        if (ex.code == 13003) {
            $("#tokenSpan").html("Cannot SSO with an on-prem domain account...");

            $("#loginButton").show();

        }
        else {
            $("#tokenSpan").html("An exception occured: Code - " + ex.code);
		}
    }

    timerResult = new Date().getTime() - getAccessTokenTimer;

    $("#timings").html($("#timings").html() + "<div>OfficeRuntime.auth.getAccessToken (+ HTML rendering) = " + timerResult + "ms</div>");
});

function doSomethingWithToken(method, token) {
    $("#tokenSpan").html("Method: <strong>" + method + "</strong>, Token: <strong>" + token + "</strong>");
}

function processMessage(arg) {

    let message = JSON.parse(arg.message);

    if (message.status == "success") {
        loginDialog.close();
        doSomethingWithToken("Dialog", message.result);
    }
    else {
        login.close();
        $("#tokenSpan").html("Something went wrong with during login");
	}
}

function showLoginPopup() {

    var url = location.protocol + "//" + location.hostname + "/auth-start.html";

    Office.context.ui.displayDialogAsync(url, { height: 60, width: 30 }, result => {
        loginDialog = result.value;
        loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    });
}