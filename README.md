# azure-graphapi

Node.js package for making Azure Active Directory Graph API calls

## Installation

```
npm install azure-graphapi --save
```

## Usage

```javascript
var GraphAPI = require('azure-graphapi');

var graph = new GraphAPI(tenant, clientId, clientSecret);
// The tenant, clientId, and clientSecret are usually in a configuration file.

graph.get('users/a8272675-dc21-4ff4-bc8d-8647830fa7db', function(err, user) {
    if (!err) {
        console.dir(user);
    }
}
```

## Details

This package provides an HTTPS interface to the [Azure Active Directory Graph API](https://msdn.microsoft.com/en-us/library/azure/hh974476.aspx). You will need the tenant (i.e., domain) of your Azure AD instance as well as an application within that AD instance that has permissions to access your directory. This application is identified by a `clientId` and authenticated using a `clientSecret`. The `clientSecret` is also called the application key.

The typical verbs are supported (GET, POST, PUT, PATCH, and DELETE). The generic `request` method supports arbitrary requests. The `getObjects` method is useful for reading a large number of objects. Azure AD limits each response to 100 objects. The `getObject` method follows the `odata.nextLink` and accumulates all objects of a specific object type.

## Interface

### Constructor

```javascript
GraphAPI(tenant, clientId, clientSecret, [apiVersion])
```

[Constructor] Creates a new `GraphAPI` instance. If the `apiVersion` is not specified, it defaults to version 1.5.

### HTTPS GET

```javascript
get(uri, callback)
```

[Method] Performs an HTTPS GET request. The `uri` must *not* begin with a slash. The callback signature is `callback(err, response)`.

### HTTPS POST

```javascript
post(uri, data, callback)
```

[Method] Performs an HTTPS POST request. The `uri` must *not* begin with a slash. The `data` is the request object. The callback signature is `callback(err, response)`.

### HTTPS PUT
```javascript
put(uri, data, callback)
```

[Method] Performs an HTTPS PUT request. The `uri` must *not* begin with a slash. The `data` is the request object. The callback signature is `callback(err, response)`.

### HTTPS PATCH

```javascript
patch(uri, data, callback)
```

[Method] Performs an HTTPS PATCH request. The `uri` must *not* begin with a slash. The `data` is the request object. The callback signature is `callback(err, response)`.

### HTTPS DELETE

```javascript
delete(uri, callback)
```

[Method] Performs an HTTPS DELETE request. The `uri` must *not* begin with a slash. The callback signature is `callback(err)`.

### HTTPS (generic)

```javascript
request(method, uri, data, callback)
```

[Method] Performs the HTTPS request specified by the `method`. The `uri` must *not* begin with a slash. The `data` is the request object and can be `null`. The callback signature is `callback(err)` for 204 (No Content) responses and `callback(err, response)` for all other success status codes.

### HTTPS GET (multiple objects)

```javascript
getObjects(uri, objectType, callback)
```

[Method] Performs an HTTPS GET request and accumulates all objects having the specified `objectType` (e.g., "User"). The `uri` must *not* begin with a slash. The callback signature is `callback(err, response)`. This method follows the `odata.nextLink` property in the response and continues until no more batches of objects are available.

## Notes

1. The HTTPS request logic parses out the `error_description` and `odata.error` messages from JSON respones to unsuccessful requests. These become part of the error message in the `err` object provided to the callback method.
2. If the request data is a string, instead of a JavaScript object, it is assumed to be form data and is sent as `application/x-www-form-urlencoded` content instead of `application/json`.

## License

(The MIT License)

Copyright (c) 2015 Frank Hellwig

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


