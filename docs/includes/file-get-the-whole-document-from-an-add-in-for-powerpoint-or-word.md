您可以创建一个只需一次单击即可将 Word 2013  或 PowerPoint 2013 文档发送到远程位置的 Office 外接程序。本文说明如何构建一个简单的 PowerPoint 2013 任务窗格外接程序，以便以数据对象的形式获取所有演示文稿并将相关数据通过 HTTP 请求发送到 Web 服务器。

## <a name="prerequisites-for-creating-an-add-in-for-powerpoint-or-word"></a>创建 PowerPoint 或 Word 外接程序的先决条件

本文假定您使用文本编辑器创建 PowerPoint 或 Word 任务窗格外接程序。 若要创建任务窗格加载项，必须创建以下文件。

- 在共享网络文件夹或 Web 服务器上，需要以下文件。

  - 一个 HTML 文件 (GetDoc_App.html) ，其中包含用户界面以及指向 JavaScript 文件的链接 (包括特定于office.js和应用程序的.js文件) 和级联样式表 (CSS) 文件。

  - 要包含外接程序编程逻辑的 JavaScript 文件 (GetDoc_App.js)。

  - 一个要包含外接程序的样式和格式的 CSS 文件 (Program.css)。

- 共享网络文件夹或外接程序目录中提供的外接程序的 XML 清单文件 (GetDoc_App.xml)。该清单文件必须指向前面提到的 HTML 文件的位置。

还可以使用 [Visual Studio](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio) 或适用于 Office 外接程序的 [Yeoman 生成器](../quickstarts/powerpoint-quickstart.md?tabs=yeomangenerator) ，或使用 Office 外接程序的 [Visual Studio](../quickstarts/word-quickstart.md?tabs=visualstudio) 或 [Yeoman 生成器为](../quickstarts/word-quickstart.md?tabs=yeomangenerator) Word 创建适用于 PowerPoint 的外接程序。

### <a name="core-concepts-to-know-for-creating-a-task-pane-add-in"></a>创建任务窗格加载项需要了解的核心概念

在开始创建 PowerPoint 或 Word 的此外接程序之前，您应知道如何构建 Office 外接程序和使用 HTTP 请求。 本文不讨论如何从 Web 服务器上的 HTTP 请求解码 Base64 编码的文本。

## <a name="create-the-manifest-for-the-add-in"></a>为外接程序创建清单

PowerPoint 外接程序的 XML 清单文件提供有关外接程序的重要信息：可以托管它的应用程序、HTML 文件的位置、外接程序标题和说明以及许多其他特征。

1. 在文本编辑器中，将以下代码添加到清单文件中。

    ```xml  
    <?xml version="1.0" encoding="utf-8" ?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xsi:type="TaskPaneApp">
        <Id>[Replace_With_Your_GUID]</Id>
        <Version>1.0</Version>
        <ProviderName>[Provider Name]</ProviderName>
        <DefaultLocale>EN-US</DefaultLocale>
        <DisplayName DefaultValue="Get Doc add-in" />
        <Description DefaultValue="My get PowerPoint or Word document add-in." />
        <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
        <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
        <Hosts>
        <Host Name="Document" />
        <Host Name="Presentation" />
        </Hosts>
        <DefaultSettings>
        <SourceLocation DefaultValue="[Network location of app]/GetDoc_App.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

2. 使用 UTF-8 编码将文件以 GetDoc_App.xml 形式保存到网络位置或外接程序目录。

## <a name="create-the-user-interface-for-the-add-in"></a>为外接程序创建用户界面

要为外接程序创建用户界面，可使用直接写入 GetDoc_App.html 文件的 HTML。外接程序的编程逻辑和功能必须包含在 JavaScript 文件（如 GetDoc_App.js）中。

使用以下过程可为该外接程序创建一个包含标题和单个按钮的简单用户界面。

1. 在文本编辑器的新文件中，添加以下 HTML。

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
            <title>Publish presentation</title>
            <link rel="stylesheet" type="text/css" href="Program.css" />
            <script src="https://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js" type="text/javascript"></script>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
            <script src="GetDoc_App.js"></script>
        </head>
        <body>
        <form>
            <h1>Publish presentation</h1>
            <br />
            <div><input id='submit' type="button" value="Submit" /></div>
            <br />
            <div><h2>Status</h2> 
                <div id="status"></div>
            </div>
        </form>
        </body>
    </html>
    ```

2. 使用 UTF-8 编码将文件以 GetDoc_App.html 形式保存到网络位置或 Web 服务器。

    > [!NOTE]
    > 请确保加载项的 **head** 标记包含 **script** 标记，其中包含 office.js 文件的有效链接。

    我们将使用一些 CSS 为外接程序提供一个简洁、现代且具专业水准的外观。使用以下 CSS 可定义外接程序的样式。

3. 在文本编辑器的新文件中，添加以下 CSS。

    ```css  
    body
    {
        font-family: "Segoe UI Light","Segoe UI",Tahoma,sans-serif;
    }
    h1,h2
    {
        text-decoration-color:#4ec724;
    }
    input [type="submit"], input[type="button"]
    {
        height:24px;
        padding-left:1em;
        padding-right:1em;
        background-color:white;
        border:1px solid grey;
        border-color: #dedfe0 #b9b9b9 #b9b9b9 #dedfe0;
        cursor:pointer;
    }
    ```

4. 使用 UTF-8 编码将该文件以 Program.css 形式保存到网络位置，或保存到 GetDoc_App.html 文件所在的 Web 服务器。

## <a name="add-the-javascript-to-get-the-document"></a>添加 JavaScript 以获取文档

在外接程序的代码中，[Office.initialize](/javascript/api/office) 事件的处理程序会向表单上 **提交** 按钮的 Click 事件中添加处理程序，并告知用户外接程序准备就绪。

下面的代码示例演示事件的事件处理程序 `Office.initialize` 以及帮助程序函数， `updateStatus`用于写入状态 div。

```js
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {

    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {

        // Execute sendFile when submit is clicked.
        $('#submit').click(function () {
            sendFile();
        });

        // Update status
        updateStatus("Ready to send file.");
    });
}

// Create a function for writing to the status div.
function updateStatus(message) {
    var statusInfo = $('#status');
    statusInfo[0].innerHTML += message + "<br/>";
}
```

在 UI 中选择 **“提交** ”按钮时，外接程序将调用 `sendFile` 函数，该函数包含对 [Document.getFileAsync 方法的](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) 调用。 该 `getFileAsync` 方法使用异步模式，类似于适用于 Office 的 JavaScript API 中的其他方法。 It has one required parameter, _fileType_, and two optional parameters,  _options_ and _callback_.

_fileType_ 参数需要 [FileType](/javascript/api/office/office.filetype) 枚举中的三个常量之一： `Office.FileType.Compressed` (“压缩”) 、 **Office.FileType.PDF** (“pdf”) 或 **Office.FileType.Text** (“text”) 。 [Document.getFileType](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) 注释下列出了每个平台的当前文件类型支持。 为 _fileType_ 参数传入 **压缩** 时，`getFileAsync`该方法通过在本地计算机上创建文件的临时副本，将文档作为 PowerPoint 2013 演示文稿文件 (*.pptx) 或 Word 2013 文档文件 (*.docx) 。

该 `getFileAsync` 方法以 [File](/javascript/api/office/office.file) 对象形式返回对文件的引用。 该 `File` 对象公开四个成员： [size](/javascript/api/office/office.file#office-office-file-size-member) 属性、 [sliceCount](/javascript/api/office/office.file#office-office-file-slicecount-member) 属性、 [getSliceAsync](/javascript/api/office/office.file#office-office-file-getsliceasync-member(1)) 方法和 [closeAsync](/javascript/api/office/office.file#office-office-file-closeasync-member(1)) 方法。 该 `size` 属性返回文件中的字节数。 返回 `sliceCount` 本文后面 (文件中) 讨论的 [Slice](/javascript/api/office/office.slice) 对象数。

使用以下代码使用该方法将 PowerPoint 或 Word 文档作为 `File` 对象 `Document.getFileAsync` 获取，然后调用本地定义 `getSlice` 的函数。 请注意，在 `File` 匿名对象的调用 `getSlice` 中，将传递对象、计数器变量和文件中切片总数。

```js
// Get all of the content from a PowerPoint or Word document in 100-KB chunks of text.
function sendFile() {
    Office.context.document.getFileAsync("compressed",
        { sliceSize: 100000 },
        function (result) {

            if (result.status == Office.AsyncResultStatus.Succeeded) {

                // Get the File object from the result.
                var myFile = result.value;
                var state = {
                    file: myFile,
                    counter: 0,
                    sliceCount: myFile.sliceCount
                };

                updateStatus("Getting file of " + myFile.size + " bytes");
                getSlice(state);
            }
            else {
                updateStatus(result.status);
            }
        });
}
```

本地函数 `getSlice` 调 `File.getSliceAsync` 用该方法以从 `File` 对象中检索切片。 该 `getSliceAsync` 方法从切片集合中返回一个 `Slice` 对象。 它具有两个必需参数： _sliceIndex_ 和 _callback_。 _sliceIndex_ 参数将整数作为切块集合中的索引器。 与 JavaScript API for Office 中的其他方法一样，该 `getSliceAsync` 方法还采用回调函数作为参数来处理方法调用的结果。
ion `getSlice` 调用 **File.getSliceAsync** 方法以从 **File** 对象中检索切片。 **getSliceAsync** 方法返回切片集合中的 **Slice** 对象。 它具有两个必需参数： _sliceIndex_ 和 _callback_。 _sliceIndex_ 参数将整数作为切块集合中的索引器。 与 Office JavaScript API 中的其他方法一样， **getSliceAsync** 方法还采用回调函数作为参数来处理方法调用的结果。

该 `Slice` 对象允许你访问文件中包含的数据。 除非方法的 _options_ 参数 `getFileAsync` 中另有指定，否则 `Slice` 该对象的大小为 4 MB。 该 `Slice` 对象公开三个属性： [大小](/javascript/api/office/office.slice#office-office-slice-size-member)、 [数据](/javascript/api/office/office.slice#office-office-slice-data-member)和 [索引](/javascript/api/office/office.slice#office-office-slice-index-member)。 该 `size` 属性获取切片的大小（以字节为单位）。 该 `index` 属性获取一个整数，该整数表示切片在切片集合中的位置。

```js
// Get a slice from the file and then call sendSlice.
function getSlice(state) {
    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {
            updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
            sendSlice(result.value, state);
        }
        else {
            updateStatus(result.status);
        }
    });
}
```

该 `Slice.data` 属性将文件的原始数据作为字节数组返回。 如果数据采用文本格式（即 XML 或纯文本），则切片包含原始文本。 如果传入 **Office.FileType.Compressed** 作为 _fileType_ 参数 `Document.getFileAsync`，则切片包含文件的二进制数据作为字节数组。 对于 PowerPoint 或 Word 文件，切片包含字节数组。

您必须实施自己的函数（或使用可用库），将字节数组数据转换为 Base64 编码的字符串。有关使用 JavaScript 进行 Base64 编码的信息，请参阅 [Base64 编码和解码](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding)。

将数据转换为 Base64 后，即可通过多种方法（包括作为 HTTP POST 请求的正文）将其传输到 Web 服务器。

添加以下代码，将切片发送到 Web 服务。

> [!NOTE]
> 此代码将 PowerPoint 或 Word 文件发送到多个切片中的 Web 服务器。 Web 服务器或服务必须将每个单个切片追加到单个文件中，然后将其另存为.pptx或.docx文件，然后才能对其执行任何操作。

```js
function sendSlice(slice, state) {
    var data = slice.data;

    // If the slice contains data, create an HTTP request.
    if (data) {

        // Encode the slice data, a byte array, as a Base64 string.
        // NOTE: The implementation of myEncodeBase64(input) function isn't
        // included with this example. For information about Base64 encoding with
        // JavaScript, see https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding.
        var fileData = myEncodeBase64(data);

        // Create a new HTTP request. You need to send the request
        // to a webpage that can receive a post.
        var request = new XMLHttpRequest();

        // Create a handler function to update the status
        // when the request has been sent.
        request.onreadystatechange = function () {
            if (request.readyState == 4) {

                updateStatus("Sent " + slice.size + " bytes.");
                state.counter++;

                if (state.counter < state.sliceCount) {
                    getSlice(state);
                }
                else {
                    closeFile(state);
                }
            }
        }

        request.open("POST", "[Your receiving page or service]");
        request.setRequestHeader("Slice-Number", slice.index);

        // Send the file as the body of an HTTP POST
        // request to the web server.
        request.send(fileData);
    }
}
```

顾名思义， `File.closeAsync` 该方法关闭与文档的连接并释放资源。 虽然 Office 外接程序沙盒垃圾可回收对文件的范围外引用，但在使用这些文件完成您的代码后，最好显式关闭它们。 该 `closeAsync` 方法有一个参数， _即回调_，该参数指定在调用完成时要调用的函数。

```js
function closeFile(state) {
    // Close the file when you're done with it.
    state.file.closeAsync(function (result) {

        // If the result returns as a success, the
        // file has been successfully closed.
        if (result.status == "succeeded") {
            updateStatus("File closed.");
        }
        else {
            updateStatus("File couldn't be closed.");
        }
    });
}
```