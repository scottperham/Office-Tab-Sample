// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.

let loginDialog = null;

//capture date/time in variable before Office.onReady function is executed
const onReadyTimer = new Date().getTime();

Office.onReady(async () => {

    //capture date/time after Office.onReady function is executed. Perform some math to store the time taken to execture Office.onReady in the timerResult variable
    let timerResult = new Date().getTime() - onReadyTimer;

    //print on the index page timerResult for Office.OnReady
    $("#timings").html("<div>Office.OnReady = " + timerResult + "ms</div>");

    //capture date/time in variable before OfficeRuntime.auth.getAccessToken function is executed
    const getAccessTokenTimer = new Date().getTime();

    try {

        //use getAccessToken from the Office SDK to try and perform an SSO, if Consent is required, a pop-up window will be displayed to the end-user
        let token = await OfficeRuntime.auth.getAccessToken({ allowConsentPrompt: true, allowSignInPrompt: true });

        //if getAccessToken is successful, pass "SSO" as the method and the Access Token to the doSomethingWithToken function and call it
        doSomethingWithToken("SSO", token);
    }

    //if getAccessToken is unsuccessful, then handle the exceptions. 13003 with show the LoginButton on the index page. Clicking this button will execute the showLoginPopup function.
    
    catch (ex) {
        if (ex.code == 13003) {
            $("#tokenSpan").html("Cannot SSO with an on-prem domain account...");

            $("#loginButton").show();

        }
        //if any other exception code is thrown by getAccessToken then the code will be printed on the index page
        else {
            $("#tokenSpan").html("An exception occured: Code - " + ex.code);
		}
    }
    //capture date/time after OfficeRuntime.auth.getAccessToken function is executed. Perform some math to store the time taken to execture OfficeRuntime.auth.getAccessToken in the timerResult variable
    timerResult = new Date().getTime() - getAccessTokenTimer;

    //print on the index page timerResult for OfficeRuntime.auth.getAccessToken
    //NOTE: this time can include the consent and HTML rendering. To get an accurate time for how long this takes, refresh the page once consent has been granted
    $("#timings").html($("#timings").html() + "<div>OfficeRuntime.auth.getAccessToken (+ HTML rendering & Consent) = " + timerResult + "ms</div>");
});

//this function prints the method used (SSO or dialog) to get the access token, and the access token on the index page
function doSomethingWithToken(method, token) {
    $("#tokenSpan").html("Method: <strong>" + method + "</strong>, Token: <strong>" + token + "</strong>");
}

//this function is the fallback method for sign-in. It used the Office SDK to open the /auth-start.html page in a pop-up window.
//This is required, as Azure AD (and many other auth providers) do not support iFraming. Office.context.ui.displayDialogAsync opens a pop-up, not an iFrame.
function showLoginPopup() {

    var url = location.protocol + "//" + location.hostname + "/auth-start.html";

    Office.context.ui.displayDialogAsync(url, { height: 60, width: 30 }, result => {
        loginDialog = result.value;
        loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    });
}

//if dialog is required to get access token, this function is called to process the result.
//If succesful, doSomethingWithToken is called with "Dialog" set for method, and the access token is passed to the function
//if unsuccessful, the message 'Something went wrong during login is displayed'
function processMessage(arg) {

    let message = JSON.parse(arg.message);

    if (message.status == "success") {
        loginDialog.close();
        doSomethingWithToken("Dialog", message.result);
    }
    else {
        login.close();
        $("#tokenSpan").html("Something went wrong during login");
    }
}