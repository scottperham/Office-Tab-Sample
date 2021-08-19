// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.

$("#officeLoadSpan").html("We are loading...");

let loginDialog = null;

Office.onReady(async () => {

    $("#officeLoadSpan").html("Loaded");
    $("#pageLoadSpan").html("Loaded");

    try {

        let token = await OfficeRuntime.auth.getAccessToken({ allowConsentPrompt: true, allowSignInPrompt: true });

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
});

function doSomethingWithToken(method, token) {
    $("#tokenSpan").html("Method: " + method + ", Token: " + token);
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

    console.log("Showing login popup");

    var url = location.protocol + "//" + location.hostname + "/auth-start.html";

    Office.context.ui.displayDialogAsync(url, { height: 60, width: 30 }, result => {
        loginDialog = result.value;
        loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    });
}