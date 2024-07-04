const LOCAL_MAP = {
"cs": "cs-cz",
"da": "da-dk",
"de": "de-de",
"en": "en-us",
"en-gb": "en-gb",
"es": "es-es",
"fi": "fi-fi",
"fr": "fr-fr",
"hu": "hu-hu",
"it": "it-it",
"ja": "ja-jp",
"ko": "ko-kr",
"nb": "nb-no",
"nl": "nl-nl",
"pl": "pl-pl",
"pt": "pt-br",
"ru": "ru-ru",
"sv": "sv-se",
"tr": "tr-tr",
"zh-hans": "zh-cn",
"zh-hant": "zh-tw"
};

const SETTING_MAP = {
    ReportMessageApiAvailable: {
        key: 'ReportMessageApiAvailable',
        value: undefined,
    },
    RestReportMessageApiAvailable: {
        key: 'RestReportMessageApiAvailable',
        value: undefined,
    },
    SpecificUserConfigurationApiAvailable: {
        key: 'SpecificUserConfigurationApiAvailable',
        value: undefined,
    },

    AdminReportJunkEmailEnabled: {
        key: 'AdminReportJunkEmailEnabled',
        value: undefined,
    },
    CheckForReportJunkDialog: {
        key: 'CheckForReportJunkDialog',
        value: undefined,
    },
    ReportJunkSelected: {
        key: 'ReportJunkSelected',
        value: undefined,
    }
};

let POLICY = null;

const QueryString = (function (a) {
    if (a === "") return {};
    const b = {};
    for (var i = 0; i < a.length; ++i) {
        var p = a[i].split('=');
        if (p.length !== 2) continue;
        b[p[0]] = decodeURIComponent(p[1].replace(/\+/g, " "));
    }
    return b;
})(window.location.search.substr(1).split('&'));

function getLocaleStrings(strings, locale = 'en-us') {
    const finalLocale = LOCAL_MAP[locale.toLowerCase()] || locale.toLowerCase();
    let result = strings[finalLocale];

    if (!result) {
        result = strings['en-us'];
    }
    return result;
}

function loadTranslation() {
    function ViewModel() {
        this.Strings = getLocaleStrings(ReportMessageStrings, Office.context.displayLanguage);
    }

    ko.applyBindings(new ViewModel());
}

function getSettings(force = false) {
    Object.keys(SETTING_MAP).forEach(key => {
        if (SETTING_MAP[key].value === undefined || force) {
            SETTING_MAP[key] = $.extend({}, SETTING_MAP[key], Office.context.roamingSettings.get(key));
        }
    });
    return Promise.resolve(SETTING_MAP);
}

function getSetting(key) {
    if (!SETTING_MAP[key] || SETTING_MAP[key].value === undefined) {
        SETTING_MAP[key] = {
            key,
            value: Office.context.roamingSettings.get(key),
        };
    }
    return Promise.resolve(SETTING_MAP[key]);
}

function setSetting(key, value) {
    Office.context.roamingSettings.set(key, value);
    return new Promise((resolve, reject) => {
        Office.context.roamingSettings.saveAsync(asyncResult => {
            if (asyncResult.status === 'failed') {
                reject(asyncResult);
            } else {
                SETTING_MAP[key] = {
                    key,
                    value,
                };
                resolve(SETTING_MAP[key]);
            }
        });
    });
}

function setSettings(settingsMap) {
    Object.keys(settingsMap).forEach(key => {
        Office.context.roamingSettings.set(key, settingsMap[key]);
    });
    return new Promise((resolve, reject) => {
        Office.context.roamingSettings.saveAsync(asyncResult => {
            if (asyncResult.status === 'failed') {
                reject(asyncResult);
            } else {
                resolve();
            }
        });
    });
}

const USER_OPTION = {
    ASK: 1,
    AUTO: 2,
    NEVER: 4
};

const JUNK_OPTION_MAP = {
    ASK: 'ask',
    AUTO: 'auto',
    NEVER: 'never',
};

const CONFIRMATION_ACTION_MAP = {
    JUNK: 'junk',
    NOT_JUNK: 'notjunk',
    PHISH: 'phish',
};

function getReportJunkOption(settingMap) {
    let reportJunkToMicrosoft = JUNK_OPTION_MAP.ASK;  // "ask", "auto", "never"

    if (settingMap.CheckForReportJunkDialog.value === true && settingMap.ReportJunkSelected.value === true) {
        reportJunkToMicrosoft = JUNK_OPTION_MAP.AUTO;
    } else if (settingMap.CheckForReportJunkDialog.value === true && settingMap.ReportJunkSelected.value === false) {
        reportJunkToMicrosoft = JUNK_OPTION_MAP.NEVER;
    }
    return reportJunkToMicrosoft;
}

function getSettingCode(userSetting) {
    switch (userSetting) {
      case JUNK_OPTION_MAP.ASK:
        return USER_OPTION.ASK;
      case JUNK_OPTION_MAP.AUTO:
        return USER_OPTION.AUTO;
      case JUNK_OPTION_MAP.NEVER:
        return USER_OPTION.NEVER;
    }
  }

/**
 * merge report action when users' setting conflict with policy
 * @param {*} userReportSetting "ask" "auto" "never"
 * @param {*} policySetting report submission policy setting
 */
function mergeUserSettingWithPolicy(userReportSetting, policySetting) {
    if(userReportSetting === JUNK_OPTION_MAP.NEVER){
      // Back compatible for users who select 'Never' before because the 'Never' option is removed.
      userReportSetting = JUNK_OPTION_MAP.ASK;
    }
    var preSubmitMsgEnabled = policySetting && policySetting.CSME;
  
    if(preSubmitMsgEnabled && userReportSetting === JUNK_OPTION_MAP.ASK){
      return JUNK_OPTION_MAP.ASK;
    }
    else{
      return JUNK_OPTION_MAP.AUTO;
    }
}

function getDeviceClient() {
    if (QueryString.platform) {
        return QueryString.platform;
    }
    const userAgent = navigator.userAgent;
    if (/android/i.test(userAgent)) {
        return "Android";
    }

    if (/iPad|iPhone|iPod/.test(userAgent) && !window.MSStream) {
        return "iOS";
    }


    return "iOS";
}

function loadCss() {
    const platform = getDeviceClient();

    let cls = 'ios';

    if (platform === 'Android') {
        cls = 'android';
    }
    $('.mobile').addClass(cls);
    return Promise.resolve();
}

function hideSpinner() {
    $('.mobile-spinner').hide();
}
function showSpinner() {
    $('.mobile-spinner').show();
}

function withSpinner(promise) {
    showSpinner();
    return promise.then(result => {
        hideSpinner();
        return Promise.resolve(result);
    }).catch(result => {
        hideSpinner();
        return Promise.reject(result);
    });
}

function getAccessToken() {
    const option = {
        isRest: true,
    };

    return new Promise((resolve, reject) => {
        Office.context.mailbox.getCallbackTokenAsync(option, (asyncResult) => {
            if (asyncResult.status === 'succeeded') {
                resolve(asyncResult.value);
            } else {
                reject();
            }
        });
    })
}

/** Convert the Office.context.platform value to submission constants. */
function convertPlatformValue(platform) {
    switch (platform) {
      case "OutlookAndroid":
        return 'Android';
      case "OutlookIOS":
        return 'iOS';
      case "OutlookWebApp":
        return 'OfficeOnline';
      case "Outlook":
        return 'Desktop';
      default:
        return 'Unknown';
    }
  }

function report (accessToken, action, reportOrNot, endPoint) {
    const itemId = Office.context.mailbox.convertToRestId(
        Office.context.mailbox.item.itemId,
        Office.MailboxEnums.RestVersion.v2_0
    );
    const url = (endPoint || Office.context.mailbox.restUrl) + '/beta/me/messages/' + encodeURIComponent(itemId) + '/ReportMessage';
    const clientPlatform = convertPlatformValue(Office.context.mailbox.diagnostics.hostName);
    return new Promise((resolve, reject) => {
        $.ajax({
            type: "POST",
            url: url,
            contentType: "application/json; charset=utf-8",
            headers: { "Authorization": "Bearer " + accessToken },
            data: JSON.stringify({
                "ReportAction": action,
                "SubmitToMicrosoft": reportOrNot,
                "ClientPlatform": clientPlatform
            })
        }).done(resolve).fail(reject);
    });
}

function reportToMicrosoft(action, reportOrNot = true) {
    return getAccessToken().then(accessToken => {
        return report(accessToken, action, reportOrNot, Office.context.mailbox.restUrl)
            .catch((e) => {
                try {
                    const decoded = jwt_decode(accessToken);
                    if (e.status === 401 && decoded && decoded.aud && (!Office.context.mailbox.restUrl.startsWith(decoded.aud))) {
                        return report(accessToken, action, reportOrNot, decoded.aud + '/api');
                    } else {
                        return Promise.reject(e);
                    }
                } catch (error) {
                    return Promise.reject(e);
                }
            });
    });
}

function setJunkOption(selectedOption) {
    let CheckForReportJunkDialog;
    let ReportJunkSelected;
    const _TTL = 1000 * 3600 * 24 * 7;
    switch (selectedOption) {
        case JUNK_OPTION_MAP.AUTO: {
            CheckForReportJunkDialog = {
                value: true,
                ttl: Number(new Date()) + _TTL,
            };
            ReportJunkSelected = {
                value: true,
                ttl: Number(new Date()) + _TTL,
            };
            break;
        }
        case JUNK_OPTION_MAP.NEVER: {
            CheckForReportJunkDialog = {
                value: true,
                ttl: Number(new Date()) + _TTL,
            };
            ReportJunkSelected = {
                value: false,
                ttl: Number(new Date()) + _TTL,
            };
            break;
        }
        case JUNK_OPTION_MAP.ASK:
        default: {
            CheckForReportJunkDialog = {
                value: false,
                ttl: Number(new Date()) + _TTL,
            };
            ReportJunkSelected = {
                value: false,
                ttl: Number(new Date()) + _TTL,
            };
            break;
        }
    }

    return setSettings({
        CheckForReportJunkDialog,
        ReportJunkSelected,
    });
}

function getType(reportAction) {
    var strings = getLocaleStrings(ManifestStrings, Office.context.displayLanguage);
    if (reportAction === "Junk") {
        return strings.ReportAsJunkLabel;
    } else if (reportAction === "Phish") {
        return strings.ReportAsPhishingLabel;
    } else if (reportAction === "NotJunk") {
        return strings.ReportAsNotJunkLabel;
    } else if (reportAction === "FocusedFeedback") {
        return strings.ReportFocusedFeedbackLabel;
    }
}

function setPostCustomizedMessage(reportAction) {
    if (POLICY && POLICY.PSME && (POLICY.PST || POLICY.PSM) && (!POLICY.OPD || reportAction === "Phish")) {
        $('#post-title').html((POLICY.PST || '').replace(/%type%/g, getType(reportAction)));
        $('#post-content').html((POLICY.PSM || '').replace(/%type%/g, getType(reportAction)));
        $('#post-report').css('display', 'flex');
    }
}

function getPolicy() {
    if (POLICY) {
        return Promise.resolve(POLICY);
    } else {
        return reportToMicrosoft('policy')
            .then(result => {
                if (result && result.Properties && result.Properties[0]) {
                    POLICY = JSON.parse(result.Properties[0].Value || '{}');
                }
                return POLICY;
            }).catch(() => {
                POLICY = {};
                return Promise.resolve(POLICY);
            });
    }
}

function showPostCustomizedMessage(reportAction) {
    return Promise.resolve();
}

function showPostNotification() {
    return new Promise((resolve, reject) => {
        // Adds an informational notification to the mail item.
        const id = "ReportPhishingPostMessage";
        const details =
        {
            type: "informationalMessage",
            message: "Report phishing success!",
            icon: "icon16Phish",
            persistent: false
        };
        Office.context.mailbox.item.notificationMessages.addAsync(id, details, handleResult);
        resolve();
    });
}

function handleResult(result) {
    // Helper method to display the result of an asynchronous call.
    console.log("The notification result: ", result);
}

function remove() {
    // Removes a notification message from the current mail item.
    const id = "ReportPhishingPostMessage"; // notifications
    Office.context.mailbox.item.notificationMessages.removeAsync(id, handleResult);
}

function getAllNotification() {
    // Gets all the notification messages and their keys for the current mail item.
    Office.context.mailbox.item.notificationMessages.getAllAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
      }
  
      console.log(asyncResult.value);
    });
  }