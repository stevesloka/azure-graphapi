//-----------------------------------------------------------------------------
// Azure Active Directory (AAD) Graph API
//
// Provides HTTP method access to the AAD Graph API. This module gets the
// authentication token for the app specified in the configuration file and
// then uses that to make the API calls. If a call fails due to a 401 error,
// a new token is obtained and the request is retried.
//-----------------------------------------------------------------------------

var https = require('https'),
    querystring = require('querystring'),
    AAD_LOGIN_HOSTNAME = 'login.windows.net',
    GRAPH_API_HOSTNAME = 'graph.windows.net',
    DEFAULT_API_VERSION = '1.5';

//-----------------------------------------------------------------------------
// PUBLIC
//-----------------------------------------------------------------------------

// Constructor
function GraphAPI(tenant, clientId, clientSecret, apiVersion) {
    if (!(this instanceof arguments.callee)) {
        throw new Error("Constructor called as a function");
    }
    this.tenant = tenant;
    this.clientId = clientId;
    this.clientSecret = clientSecret;
    this.apiVersion = apiVersion || DEFAULT_API_VERSION;
    this.accessToken = null;
}

// HTTPS GET
GraphAPI.prototype.get = function (path, callback) {
    this.request('GET', path, null, callback);
}

// HTTPS POST
GraphAPI.prototype.post = function (path, data, callback) {
    this.request('POST', path, data, callback);
}

// HTTPS PUT
GraphAPI.prototype.put = function (path, data, callback) {
    this.request('PUT', path, data, callback);
}

// HTTPS PATCH
GraphAPI.prototype.patch = function (path, data, callback) {
    this.request('PATCH', path, data, callback);
}

// HTTPS DELETE
GraphAPI.prototype.delete = function (path, callback) {
    this.request('DELETE', path, null, callback);
}

// HTTPS GET with odata.nextList recursive call
GraphAPI.prototype.getObjects = function (uri, objectType, callback) {
    this._getObjects(uri, [], objectType, callback);
}

// Generic HTTPS request
GraphAPI.prototype.request = function (method, uri, data, callback) {
    var self = this;
    if (self.accessToken) {
        self._request(method, uri, data, false, callback);
    } else {
        self._requestAccessToken(function (err, token) {
            if (err) {
                callback(err);
            } else {
                self.accessToken = token;
                self._request(method, uri, data, false, callback);
            }
        });
    }
}

//-----------------------------------------------------------------------------
// PRIVATE
//-----------------------------------------------------------------------------

// Recursive method that follows the odata.nextLink.
GraphAPI.prototype._getObjects = function (uri, objects, objectType, callback) {
    var self = this;
    get(uri, function (err, response) {
        if (err) return callback(err);
        var value = response.value;
        for (var i = 0, n = value.length; i < n; i++) {
            if (value[i].objectType === objectType) {
                objects.push(value[i]);
            }
        }
        var nextLink = response['odata.nextLink'];
        if (nextLink) {
            self._getObjects(nextLink, objects, objectType, callback);
        } else {
            callback(null, objects);
        }
    });
}

// Performs the HTTPS request and tries again on a 401 error
// by getting another access token and repeating the request.
GraphAPI.prototype._request = function (method, uri, data, secondAttempt, callback) {
    var self = this;
    var path = ['/'];
    path.push(self.tenant);
    path.push('/');
    path.push(uri);
    if (uri.indexOf('?') < 0) {
        path.push('?');
    }
    path.push('api-version=');
    path.push(self.apiVersion);
    var options = {
        hostname: GRAPH_API_HOSTNAME,
        path: path.join(''),
        method: method,
        headers: {
            'Authorization': 'Bearer ' + self.accessToken
        }
    };
    httpsRequest(options, data, function (err, response) {
        if (err) {
            if (err.statusCode === 401 && !secondAttempt) {
                self._requestAccessToken(function (err, token) {
                    if (err) {
                        callback(err);
                    } else {
                        accessToken = token;
                        self._request(method, uri, data, true, callback);
                    }
                });
            } else {
                log.error(err);
                callback(err);
            }
        } else {
            if (typeof response === 'undefined') {
                // Handle 204 responses by not adding a second argument.
                callback(null);
            } else {
                callback(null, response);
            }
        }
    });
}

// Gets an access token using the client id and secret.
GraphAPI.prototype._requestAccessToken = function (callback) {
    var query = {
        client_id: this.clientId,
        client_secret: this.clientSecret,
        grant_type: 'client_credentials',
        resource: 'https://' + GRAPH_API_HOSTNAME
    };
    var content = querystring.stringify(query);
    var options = {
        hostname: AAD_LOGIN_HOSTNAME,
        path: '/' + this.tenant + '/oauth2/token',
        method: 'POST'
    };
    httpsRequest(options, content, function (err, response) {
        if (err) {
            callback(err);
        } else {
            callback(null, response.access_token);
        }
    });
}

// Our own wrapper around the https.request method.
function httpsRequest(options, content, callback) {
    options.headers = options.headers || {};
    options.headers['Accept'] = 'application/json';
    if (!callback) {
        callback = content;
        content = null;
    } else if (typeof content === 'string') {
        options.headers['Content-Type'] = 'application/x-www-form-urlencoded';
        options.headers['Content-Length'] = content.length;
    } else if (content !== null && typeof content === 'object') {
        content = JSON.stringify(content);
        options.headers['Content-Type'] = 'application/json';
        options.headers['Content-Length'] = content.length;
    } else {
        content = null;
    }
    var req = https.request(options, function (res) {
        res.setEncoding('utf8');
        var buf = [];
        res.on('data', function (data) {
            buf.push(data);
        });
        res.on('end', function () {
            var data = buf.join('');
            if (data.length > 0) {
                data = JSON.parse(data);
            } else {
                data = null;
            }
            if (res.statusCode === 204) {
                callback(null); // no data
            } else if (res.statusCode >= 200 && res.statusCode <= 299) {
                callback(null, data); // success
            } else {
                if (data && data.error_description) {
                    data = data.error_description.split(/[\r\n]/)[0];
                } else if (data && data['odata.error']) {
                    data = data['odata.error'].message.value;
                } else {
                    data = https.STATUS_CODES[status];
                }
                callback(new Error(errmsg(res.statusCode, data)));
            }
        });
    });
    req.on('error', function (err) {
        callback(err);
    })
    if (content) {
        req.write(content);
    }
    req.end();
}

// Creates an exception error message.
function errmsg(status, message) {
    return 'Graph API Error: (' + status + ') ' + message;
}

//-----------------------------------------------------------------------------
// EXPORTS
//-----------------------------------------------------------------------------

module.exports = GraphAPI; // export the constructor

