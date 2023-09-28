A simple and easy-to-use HTTP client for VBA based on WinHttpRequest.

Donate:

- 1F1wLQEWfw7QCqFkwrkLcf2SZnMb6Wj8Bj (Bitcoin network BTC);
- 0x6d3e448bbdf46b2fe8f9d4d92fcee7379597950e (Ethereum network ERC20).

# 1. Quick Start

Just add the **Http.cls** file to your VBA project. Create an instance of the Http class, call the **init()** method by passing it the URL (not to be confused with URI), and call the **methodGet()** method by passing it the URN.

Example:

```vb
Private Sub CommandButton1_Click()
  Dim myHttp As New Http
  myHttp.init "https://api.publicapis.org"
  Set result = myHttp.methodGet("entries")
  Debug.Print result.ResponseText
End Sub
```

where:

- https://api.publicais.org is the URL
- entries is the – URN
- https://api.publicapis.org/entries is the URI

# 2. Method GET

To send a GET request, you need to call the **methodGet()** class method. This method takes two optional parameters:

- **urn** ("/", by default, the request will be sent to the URL specified when the init method is called);
- **params** - a collection of GET parameters for the request. For each GET parameter, you need to add an array with a key/value pair to the collection.

Example:

```vb
Dim params As New Collection
params.Add Array("param1", "value1")
params.Add Array("param2", "value2")
Set result = myHttp.methodGet("entries", params)
```

The method will return an instance of a **WinHttpRequest** object (see Section 6).

# 3. Method POST

To send a POST request, you need to call the **methodPost()** class method. This method takes two optional parameters:

- **urn** (if not specified, the request will be sent to the URL specified when the init method is called);
- **payload** - the request payload. It can be passed in the form of:
  - **string** - a JSON string;
  - **collection** - a collection of parameters, each in the form of a key-value pair;
  - **byte** - the request form in binary form. To create the form, use the **createForm()** method (see Section 4).

Depending on the type of payload passed, the appropriate Content-Type header for the request will be generated automatically.

Example 1:

```vb
Dim params As New Collection
params.Add Array("param1", "value1")
params.Add Array("param2", "value2")
Set result = myHttp.methodPost("entries", params)
```

Example 2:

```vb
Dim strJson as String
strJson = "[""Sunday"", ""Monday"", ""Tuesday"", ""Wednesday"",""Thursday"", ""Friday"", ""Saturday""]"
Set result = myHttp.methodPost("entries", strJson)
```

# 4. Creating and sending a form

If you need to send a form using the POST method (including sending files to the server), you need to create the form using the **createForm()** method. This method will return the form in binary format, after which it can be passed as payload to the **methodPost()** method (see Section 3). The **createForm()** method accepts two optional parameters (at least one of them is requiredmust be):

- **params** - a collection of parameters, each a key/value pair;
- **files** - a collection of files, each a key/path to the file.

Example 1:

```vb
Dim files As New Collection
files.Add Array("file1", "d:\file1.txt")
files.Add Array("file2", "d:\file2.txt")
Dim params As New Collection
params.Add Array("date", "2023-10-02")
bytForm = myHttp.createForm(params, files)
Set result = myHttp.methodPost("api/upload", bytForm)
```

Example 2:

```vb
Dim params As New Collection
params.Add Array("date", "2023-10-02")
bytForm = myHttp.createForm(params)
Set result = myHttp.methodPost("api/upload", bytForm)
```

# 5. Working with request headers

There are a number of methods available for working with request headers.

## 5.1. Method setHeader

This method can be used to add a custom request header. The method takes a header parameter, which can be specified either as a string like "Content-Type: application/x-www-form-urlencoded", or as an array with two values, the header name and its body.

Example 1:

```vb
myHttp.setHeader "Content-Type: application/x-www-form-urlencoded"
```

Example 2:

```vb
myHttp.setHeader Array("Content-Type", "application/x-www-form-urlencoded")
```

## 5.2. Method removeHeader

This method allows you to remove a header from the request. It takes one string parameter - the name of the header to be removed.

Example:

```vb
myHttp.removeHeader "Content-Type"
```

## 5.3. Method removeAllHeaders

This method allows you to remove all headers from the request. It takes no parameters.

Example:

```vb
myHttp.removeAllHeaders
```

## 5.4. Method setBaseHeaders

When the Http class is initialized, a list of basic headers is created, which includes the following headers:

- Accept: \*/\*
- Accept-Language: en-US
- User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)
- Content-Type: text/html;charset=utf-8
- Host, *userHost*

These headers can also be set using the **setBaseHeaders()** method, in which case you need to re-create the list of headers for the next request.

Example:

```vb
myHttp.setBaseHeaders
```

# 6. Working with the response

The result of calling the **methodGet()** and **methodPost()** methods is an instance of a **WinHttpRequest** object. The main methods and properties of this object for obtaining a response are:

- **ResponseText** - property returns the body of the response (the response data is decoded according to the Content-Type header).
- **Status** - property contains the server response code.
- **StatusText** - property contains the text description of the server response.
- **GetResponseHeader()** - method takes the name of the response header and returns the value of this header.
- **GetAllResponseHeaders()** - method returns all response headers as a string.



# P.S.

If you want to enable the automatic logon policy, for example, with ActiveDirectory Windows, set the **autoLogonPolicy** property to true.

To work with JSON format, it is recommended to use this [module](https://github.com/VBA-tools/VBA-JSON).

