"use strict";
var Sdk = window.Sdk || {};

/**
 * @function getClientUrl
 * @description Get the client URL.
 * @returns {string} The client URL.
 */

Sdk.getDependecySubgrid = function () {
    return true;
};

Sdk.getClientUrl = function () {
    
    var globalContext = Xrm.Utility.getGlobalContext();
    return globalContext.getClientUrl();

};

/**
 * An object instantiated to manage detecting the
 * Web API version in conjunction with the 
 * Sdk.retrieveVersion function
 */
Sdk.versionManager = new function () {
    //Start with base version
    var _webAPIMajorVersion = 9;
    var _webAPIMinorVersion = 0;
    //Use properties to increment version and provide WebAPIPath string used by Sdk.request;
    Object.defineProperties(this, {
        "WebAPIMajorVersion": {
            get: function () {
                return _webAPIMajorVersion;
            },
            set: function (value) {
                if (typeof value != "number") {
                    throw new Error("Sdk.versionManager.WebAPIMajorVersion property must be a number.")
                }
                _webAPIMajorVersion = parseInt(value, 10);
            }
        },
        "WebAPIMinorVersion": {
            get: function () {
                return _webAPIMinorVersion;
            },
            set: function (value) {
                if (isNaN(value)) {
                    throw new Error("Sdk.versionManager._webAPIMinorVersion property must be a number.")
                }
                _webAPIMinorVersion = parseInt(value, 10);
            }
        },
        "WebAPIPath": {
            get: function () {
                return "/api/data/v" + _webAPIMajorVersion + "." + _webAPIMinorVersion;
            }
        }
    })

}

Sdk.retrieveVersion = function () {
    return new Promise(function (resolve, reject) {
        Sdk.requestV9("GET", "/RetrieveVersion")
            .then(function (request) {
                try {
                    var RetrieveVersionResponse = JSON.parse(request.response);
                    var fullVersion = RetrieveVersionResponse.Version;
                    var versionData = fullVersion.split(".");
                    Sdk.versionManager.WebAPIMajorVersion = parseInt(versionData[0], 10);
                    Sdk.versionManager.WebAPIMinorVersion = parseInt(versionData[1], 10);
                    resolve();
                } catch (err) {
                    reject(new Error("Error processing version: " + err.message))
                }
            })
            .catch(function (err) {
                reject(new Error("Error retrieving version: " + err.message))
            })
    });
};

/**
 * @function request
 * @description Generic helper function to handle basic XMLHttpRequest calls.
 * @param {string} action - The request action. String is case-sensitive.
 * @param {string} uri - An absolute or relative URI. Relative URI starts with a "/".
 * @param {bool} async - asyncronous.
 * @param {object} data - An object representing an entity. Required for create and update actions.
 * @param {object} addHeader - An object with header and value properties to add to the request
 * @param {function} resolve - successcallback
 * @param {function} reject - errorcallback
 */
Sdk.requestV9 = function (action, uri, async, data, addHeader, resolve, reject, useAdmin) {
    if (!RegExp(action, "g").test("POST PATCH PUT GET DELETE")) { // Expected action verbs.
        throw new Error("Sdk.request: action parameter must be one of the following: " +
            "POST, PATCH, PUT, GET, or DELETE.");
    }
    if (!typeof uri === "string") {
        throw new Error("Sdk.request: uri parameter must be a string.");
    }
    if ((RegExp(action, "g").test("POST PATCH PUT")) && (!data)) {
        throw new Error("Sdk.request: data parameter must not be null for operations that create or modify data.");
    }
    if (addHeader) {
        if (typeof addHeader.header != "string" || typeof addHeader.value != "string") {
            throw new Error("Sdk.request: addHeader parameter must have header and value properties that are strings.");
        }
    }

    // Construct a fully qualified URI if a relative URI is passed in.
    if (uri.charAt(0) === "/") {
        //This sample will try to use the latest version of the web API as detected by the 
        // Sdk.retrieveVersion function.
        uri = Sdk.getClientUrl() + Sdk.versionManager.WebAPIPath + uri;
    }

    var request = new XMLHttpRequest();
    request.open(action, encodeURI(uri), async);
    request.setRequestHeader("OData-MaxVersion", "4.0");
    request.setRequestHeader("OData-Version", "4.0");
    request.setRequestHeader("Accept", "application/json");
    request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    request.setRequestHeader("Prefer", "odata.include-annotations=OData.Community.Display.V1.FormattedValue")
    if (useAdmin) {
        request.setRequestHeader("MSCRMCallerID", Sdk.getAdminId());
    }
    if (addHeader) {
        request.setRequestHeader(addHeader.header, addHeader.value);
    }
    request.onreadystatechange = function () {
        if (this.readyState === 4) {
            request.onreadystatechange = null;
            switch (this.status) {
                case 200: // Operation success with content returned in response body.
                case 201: // Create success. 
                case 204: // Operation success with no content returned in response body.
                case 1223: // Operation success with no content returned in response body.
                    resolve(this);
                    break;
                default: // All other statuses are unexpected so are treated like errors.
                    var error;
                    try {
                        error = JSON.parse(request.response).error;
                    } catch (e) {
                        error = new Error("Unexpected Error");
                    }
                    reject(error);
                    break;
            }
        }
    };
    request.send(JSON.stringify(data));
};

Sdk.showErrorMessage = function (message) {
    if (typeof (formContext) != "undefined" && formContext.ui != "undefined") {
        formContext.ui.setFormNotification("Errore nell'esecuzione del Processo: " + message, 'ERROR', "errorSDK");
    }
    else if (typeof (Xrm) != "undefined" && typeof (Xrm.Page) != "undefined" && Xrm.Page.ui != null) {
        Xrm.Page.ui.setFormNotification("Errore nell'esecuzione del Processo: " + message, 'ERROR', "errorSDK");
    }
    else if (typeof (parent.Xrm) != "undefined" && typeof (parent.Xrm.Page) != "undefined" && parent.Xrm.Page.ui != null) {
        parent.Xrm.Page.ui.setFormNotification("Errore nell'esecuzione del Processo: " + message, 'ERROR', "errorSDK");
    }
    else {
        var alertStrings = { text: "Errore nell'esecuzione del Processo: " + message };
        var alertOptions = { height: 120, width: 520 };
        parent.Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
    }
};

Sdk.clearErrorMessage = function ()
{
    if (typeof (formContext) != "undefined" && formContext.ui != "undefined") {
        formContext.ui.clearFormNotification("errorSDK");
    }
    else if (typeof (Xrm) != "undefined" && typeof (Xrm.Page) != "undefined" && Xrm.Page.ui != null) {
        Xrm.Page.ui.clearFormNotification("errorSDK");
    }
    else if (typeof (parent.Xrm) != "undefined" && typeof (parent.Xrm.Page) != "undefined" && parent.Xrm.Page.ui != null) {
        parent.Xrm.Page.ui.clearFormNotification("errorSDK");
    }
    else {
        var alertStrings = { text: "Errore nell'esecuzione del Processo"};
        var alertOptions = { height: 120, width: 520 };
        parent.Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
    }
};

Sdk.executeFetch = function (entity, fetchXML, hideError) {
    var res;

    var url = Sdk.getClientUrl() + "/api/data/v9.0/" + entity + "?fetchXml=" + fetchXML;
    Sdk.requestV9("GET", url, false, null, null, function (result) {
        var returned = JSON.parse(result.responseText);
        res = returned.value;
        res.executionSucceded = true;
    }, function (result) {
        if (!hideError) {
            Sdk.showErrorMessage(result.message);
        }
        result.executionSucceded = false;
        res = result;
    });
    return res;
};
Sdk.executeFetchAdmin = function (entity, fetchXML, hideError) {
    var res;

    var url = Sdk.getClientUrl() + "/api/data/v9.0/" + entity + "?fetchXml=" + fetchXML;
    Sdk.requestV9("GET", url, false, null, null, function (result) {
        var returned = JSON.parse(result.responseText);
        res = returned.value;
        res.executionSucceded = true;
    }, function (result) {
        if (!hideError) {
            Sdk.showErrorMessage(result.message);
        }
        result.executionSucceded = false;
        res = result;
    }, true);
    return res;
};
Sdk.executeRetrieveByGuid = function (entity, guid, options, hideError) {
    var res;

    var url = Sdk.getClientUrl() + "/api/data/v9.0/" + entity + "(" + guid.replace('{', '').replace('}', '') + ")" + options;
    Sdk.requestV9("GET", url, false, null, null, function (result) {
        res = JSON.parse(result.responseText);
        res.executionSucceded = true;
    }, function (result) {
        if (!hideError) {
            Sdk.showErrorMessage(result.message);
        }
        result.executionSucceded = false;
        res = result;
    });
    return res;
};
Sdk.executeUpdate = function (entity, guid, data, hideError) {
    var res = {};
    var url = Sdk.getClientUrl() + "/api/data/v9.0/" + entity + "(" + guid.replace('{', '').replace('}', '') + ")";
    Sdk.requestV9("PATCH", url, false, data, null, function (result) {
        //res = JSON.parse(result.responseText);
        res.executionSucceded = true;
    }, function (result) {
        if (!hideError) {
            Sdk.showErrorMessage(result.message);
        }
        result.executionSucceded = false;
        res = result;
    });
    return res;
};
Sdk.executeUpdateAdmin = function (entity, guid, data, hideError) {
    var res = {};
    var url = Sdk.getClientUrl() + "/api/data/v9.0/" + entity + "(" + guid.replace('{', '').replace('}', '') + ")";
    Sdk.requestV9("PATCH", url, false, data, null, function (result) {
        //res = JSON.parse(result.responseText);
        res.executionSucceded = true;
    }, function (result) {
        if (!hideError) {
            Sdk.showErrorMessage(result.message);
        }
        result.executionSucceded = false;
        res = result;
    }, true);
    return res;
};
Sdk.executeCreate = function (entity, data, hideError) {
    var res = {};
    var url = Sdk.getClientUrl() + "/api/data/v9.0/" + entity;
    Sdk.requestV9("POST", url, false, data, null, function (result) {
        //res = JSON.parse(result.responseText);
        res.executionSucceded = true;
    }, function (result) {
        if (!hideError) {
            Sdk.showErrorMessage(result.message);
        }
        result.executionSucceded = false;
        res = result;
    });
    return res;
};
Sdk.executeCreateAdmin = function (entity, data, hideError) {
    var res = {};
    var url = Sdk.getClientUrl() + "/api/data/v9.0/" + entity;
    Sdk.requestV9("POST", url, false, data, null, function (result) {
        //res = JSON.parse(result.responseText);
        res.executionSucceded = true;
    }, function (result) {
        if (!hideError) {
            Sdk.showErrorMessage(result.message);
        }
        result.executionSucceded = false;
        res = result;
    }, true);
    return res;
};
Sdk.executeRetrieve = function (entity, options, hideError) {
    var res;

    var url = Sdk.getClientUrl() + "/api/data/v9.0/" + entity + options;
    Sdk.requestV9("GET", url, false, null, null, function (result) {
        var returned = JSON.parse(result.responseText);
        res = returned.value;
        res.executionSucceded = true;
    }, function (result) {
        if (!hideError) {
            Sdk.showErrorMessage(result.message);
        }
        result.executionSucceded = false;
        res = result;
    });
    return res;
};
Sdk.executeRetrieveAdmin = function (entity, options, hideError) {
    var res;

    var url = Sdk.getClientUrl() + "/api/data/v9.0/" + entity + options;
    Sdk.requestV9("GET", url, false, null, null, function (result) {
        var returned = JSON.parse(result.responseText);
        res = returned.value;
        res.executionSucceded = true;
    }, function (result) {
        if (!hideError) {
            Sdk.showErrorMessage(result.message);
        }
        result.executionSucceded = false;
        res = result;
    },true);
    return res;
};
Sdk.executeBoundAction = function (entityName, entityId, actionName, data, hideError, useAdmin) {
        var res;
        Sdk.requestV9("POST", "/" + entityName + "(" + entityId.replace('{', '').replace('}', '') + ")/Microsoft.Dynamics.CRM." + actionName, false, data, null,
            function (result) {
                var returned = {};
                if (result.response != "") {
                    returned = JSON.parse(result.response);
                }
                returned.executionSucceded = true;
                res = returned;
            }, function (result) {
                if (!hideError) {
                    Sdk.showErrorMessage(result.message);
                }
                result.executionSucceded = false;
                res = result;
            },
            useAdmin
        );
        return res;
    }
Sdk.executeUnboundAction = function (actionName, data, hideError) {
    var res;
    Sdk.requestV9("POST", "/" + actionName, false, data, null,
        function (result) {
            var returned = {};
            if (result.response != "") {
                returned = JSON.parse(result.response);
            }
            returned.executionSucceded = true;
            res = returned;
        }, function (result) {
            if (!hideError) {
                Sdk.showErrorMessage(result.message);
            }
            result.executionSucceded = false;
            res = result;
        });
    return res;
}
Sdk.executeDisassociate = function (entity1, entity1Id, relationshipName, entity2Id, hideError) {
    var res;
    Sdk.requestV9("DELETE", "/" + entity1 + "(" + entity1Id + ")/" + relationshipName + "(" + entity2Id + ")/$ref", false, data, null,
        function (result) {
            var returned = {};
            if (result.response != "") {
                returned = JSON.parse(result.response);
            }
            returned.executionSucceded = true;
            res = returned;
        }, function (result) {
            if (!hideError) {
                Sdk.showErrorMessage(result.message);
            }
            result.executionSucceded = false;
            res = result;
        });
    return res;
}

Sdk.executeAssassociate = function (entity1, entity1Id, relationshipName, entity2Id) {
   var res;
   Sdk.request("PATCH", "/" + entity1 + "(" + entity1Id + ")/" + relationshipName + "(" + entity2Id + ")/$ref", false, data, null,
       function (result) {
           var returned = {};
           if (result.response != "") {
               returned = JSON.parse(result.response);
           }
           returned.executionSucceded = true;
           res = returned;
       }, function (result) {
           Sdk.showErrorMessage(result.message);
           result.executionSucceded = false;
           res = result;
       });
   return res;
}