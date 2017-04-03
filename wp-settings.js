// save SharePoint custom Web Part settings to SPList via REST api call
// wpPageURL + wpTitle are primary key
// last updated 03-21-2017

// execute REST with callback handler.  small version of jQuery $.ajax()
function wpsAjax(verb, url, data, callback, update, callback2) {
    // create XHR
    var rel = _spPageContextInfo.webServerRelativeUrl;
    var jsonHeader = 'application/json;odata=verbose';
    var req = new XMLHttpRequest();
    req.open(verb, rel + url);
    // build headers
    req.setRequestHeader('Accept', jsonHeader);
    req.setRequestHeader('Content-Type', jsonHeader);
    if (window.wpDigest) {
        req.setRequestHeader('X-RequestDigest', window.wpDigest);
    }
    if (update) {
        req.setRequestHeader('X-HTTP-Method', 'MERGE');
        req.setRequestHeader('If-Match', '*');
    }
    // define handler
    req.onload = function () {
        if (url.indexOf('contextinfo') > 0) {
            // Form security digest
            var obj = JSON.parse(this.response);
            window.wpDigest = obj.d.GetContextWebInformation.FormDigestValue;
        }
        if (callback) {
            // execute
            callback(this, callback2);
        }
    }
    // execute request
    if (verb == 'POST') {
        req.send(JSON.stringify(data));
    } else {
        req.send();
    }
}

// create SharePoint custom list 'wpSetting'
function wpsEnsureList(callback) {
    var data = {
        '__metadata': {
            'type': 'SP.List'
        },
        'AllowContentTypes': true,
        'BaseTemplate': 100,
        'ContentTypesEnabled': true,
        'Description': '',
        'Title': 'wpSetting'
    };
    wpsAjax('POST', '/_api/web/lists', data, function (resp) {
        if (resp.status == 201) {
            wpsEnsureFields(['wpTitle', 'wpPageURL'], callback)
        } else {
            callback();
        }
    });
}

// create SharePoint field 'wpPageURL'
function wpsEnsureFields(fields, callback) {
    var data = {
        '__metadata': {
            'type': 'SP.Field'
        },
        'FieldTypeKind': 2,
        'Title': fields[0]
    };
    if (Array.isArray(fields) && fields.length > 0) {
        // more fields to create
        fields.shift();
        wpsAjax('POST', '/_api/web/lists/getbytitle(\'wpSetting\')/fields', data, function () {
            wpsEnsureFields(fields, callback);
        });
    } else {
        // finalize
        callback();
    }
}

// locate web part title
function wpsGetTitle(css) {
    var obj = document.querySelector(css);
    do {
        if (obj.className.indexOf('ms-WPBody') > -1) {
            return obj.id;
        }
        obj = obj.parentNode;
    } while (obj);
}

// execute REST write 
function wpsWriteItem(css, setting, callback) {
    var url = _spPageContextInfo.webServerRelativeUrl + _spPageContextInfo.serverRequestPath;
    if (css) {
        var title = wpsGetTitle(css);
    } else {
        var title = 'fullpage';
    }
    var data = {
        '__metadata': {
            'type': 'SP.ListItem'
        },
        'Title': JSON.stringify(setting),
        'wpPageURL': url,
        'wpTitle': title
    };

    // READ
    wpsAjax('GET', '/_api/web/lists/getbytitle(\'wpSetting\')/items?$select=Id,Title&$filter=(wpPageURL+eq+\'' + url + '\')+and+(wpTitle+eq+\'' + title + '\')', null, function (resp) {
        var obj = JSON.parse(resp.response);
        if (obj.d.results.length) {
            // parse ID for matching row
            var id = obj.d.results[0].ID;
        }
        if (id) {
            // UPDATE
            data.Id = id;
            data.ID = id;
            wpsAjax('POST', '/_api/web/lists/getbytitle(\'wpSetting\')/items(' + id + ')', data, function (resp) {
                if (callback) {
                    callback(id);
                }
            }, true);
        } else {
            // INSERT
            wpsAjax('POST', '/_api/web/lists/getbytitle(\'wpSetting\')/items', data, function (resp) {
                var obj = JSON.parse(resp.response);
                if (callback) {
                    callback(obj.d.ID);
                }
            });
        }
    });
}

// button click - Write
function wpsWrite(css, setting, save) {
    wpsAjax('POST', '/_api/contextinfo', {}, function () {
        wpsEnsureList(function () {
            wpsWriteItem(css, setting, save);
        });
    });
}

// button click - Read
function wpsRead(css, callback) {
    var url = _spPageContextInfo.webServerRelativeUrl + _spPageContextInfo.serverRequestPath;
    if (css) {
        var title = wpsGetTitle(css);
    } else {
        var title = 'fullpage';
    }
    wpsAjax('GET', '/_api/web/lists/getbytitle(\'wpSetting\')/items?$select=Id,Title&$filter=(wpPageURL+eq+\'' + url + '\')+and+(wpTitle+eq+\'' + title + '\')', null, function (resp) {
        var obj = JSON.parse(resp.response);
        if (callback) {
            var title = null;
            if (obj.d) {
                if (obj.d.results[0]) {
                    title = obj.d.results[0].Title
                }
            }
            callback(JSON.parse(title));
        }
    });
}
