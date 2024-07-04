var _clickEvent;

function CompleteEvent() {
    if (_clickEvent !== undefined) {
      _clickEvent.completed();
    }
  }

function cmdPhishing(event) {
    CompleteEvent();
    _clickEvent = event;
    try {
        reportPhishing(); 
    } catch (e) {
        var msg = JSON.stringify(e);
        console.log("Phishing command ended, with exception: " + msg);
        CompleteEvent();
    } finally {
        console.log("Phishing command ended.");
    }
}

function reportPhishing() {
    getPolicy()
        .then(policy => {
            var action = "Phish";
            var junkOption = mergeUserSettingWithPolicy(JUNK_OPTION_MAP.ASK, policy);
            console.log("ReportPhishing junkOption", junkOption);
            if (junkOption === JUNK_OPTION_MAP.ASK) {
                reportWithDialog(action, policy);
            } else {
                reportToMicrosoft(action, junkOption === JUNK_OPTION_MAP.AUTO)
                .then((action) => {
                        console.log('Report Success!');
                        return showPostNotification(action);
                    })
                    .catch((e) => {
                        console.error('Report Error!')
                        console.error(e)
                    });
            }
        });
}

function reportWithDialog(action, policy) {
    var dialog;
    var dialogUrl;
    var dlgName = "https://github.com/enersecy/reporter/blob/main/src/ReportPhishMobile.html";
    if (window.location === undefined || window.location.origin === undefined || window.location.pathname === undefined) {
        // Fallback
        dialogUrl = "https://github.com/enersecy/reporter/blob/main/src/FunctionFileMobile.html/" + dlgName;
    } else {
        dialogUrl = window.location.origin + window.location.pathname + "/" + dlgName;
    }

    console.log("Open dialog: " + dialogUrl);

    localStorage.setItem('action', action);
    localStorage.setItem('policy', JSON.stringify(policy));
    localStorage.setItem('locale', Office.context.displayLanguage);

    Office.context.ui.displayDialogAsync(dialogUrl, { height: 30, width: 20, displayInIframe: true },
        function (asyncResult) {
            console.log("open dialog...asyncResult",asyncResult);
            dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
            console.log("open dialog...done, dialog", dialog);
        }
    );

    function processMessage(arg) {
        dialog.close();

        var messageFromDialog = JSON.parse(arg.message);
        var isReport = messageFromDialog.isReport;
        if (isReport) {
            reportCore(action);
        } else {
            cancel();
        }
    }
}

function reportCore(action) {
    reportToMicrosoft(action)
        .then(() => {
            return showPostNotification();
        })
        .finally(() => {
            Office.context.ui.closeContainer();
        });
}

function cancel() {
    Office.context.ui.closeContainer();
}
