# <a name="get-the-whole-document-from-an-add-in-for-powerpoint-or-word"></a><span data-ttu-id="0a7ec-101">从 PowerPoint 或 Word 相关外接程序获取整个文档</span><span class="sxs-lookup"><span data-stu-id="0a7ec-101">Get the whole document from an add-in for PowerPoint or Word</span></span>

<span data-ttu-id="0a7ec-p101">您可以创建一个只需一次单击即可将 Word 2013  或 PowerPoint 2013 文档发送到远程位置的 Office 外接程序。本文说明如何构建一个简单的 PowerPoint 2013 任务窗格外接程序，以便以数据对象的形式获取所有演示文稿并将相关数据通过 HTTP 请求发送到 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p101">You can create an Office Add-in to provide one-click sending or publishing of a Word 2013 or PowerPoint 2013 document to a remote location. This article demonstrates how to build a simple task pane add-in for PowerPoint 2013 that gets all of the presentation as a data object and sends that data to a web server via an HTTP request.</span></span>

## <a name="prerequisites-for-creating-an-add-in-for-powerpoint-or-word"></a><span data-ttu-id="0a7ec-104">创建 PowerPoint 或 Word 外接程序的先决条件</span><span class="sxs-lookup"><span data-stu-id="0a7ec-104">Prerequisites for creating an add-in for PowerPoint or Word</span></span>

<span data-ttu-id="0a7ec-p102">本文假定您使用文本编辑器创建 PowerPoint 或 Word 任务窗格外接程序。若要创建任务窗格外接程序，您必须创建以下文件：</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p102">This article assumes that you are using a text editor to create the task pane add-in for PowerPoint or Word. To create the task pane add-in, you must create the following files:</span></span>

- <span data-ttu-id="0a7ec-107">在共享网络文件夹或 Web 服务器上，您需要以下文件：</span><span class="sxs-lookup"><span data-stu-id="0a7ec-107">On a shared network folder or on a web server, you need the following files:</span></span>
    
    - <span data-ttu-id="0a7ec-108">HTML 文件 (GetDoc_App.html)，其中包含用户界面、指向 JavaScript 文件（包括 office.js 和主机特定的 .js 文件）的链接和级联样式表 (CSS) 文件。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-108">An HTML file (GetDoc_App.html) that contains the user interface plus links to the JavaScript files (including office.js and host-specific .js files) and Cascading Style Sheet (CSS) files.</span></span>
           
    - <span data-ttu-id="0a7ec-109">要包含外接程序编程逻辑的 JavaScript 文件 (GetDoc_App.js)。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-109">A JavaScript file (GetDoc_App.js) to contain the programming logic of the add-in.</span></span>
    
    - <span data-ttu-id="0a7ec-110">一个要包含外接程序的样式和格式的 CSS 文件 (Program.css)。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-110">A CSS file (Program.css) to contain the styles and formatting for the add-in.</span></span>
    
- <span data-ttu-id="0a7ec-p103">共享网络文件夹或外接程序目录中提供的外接程序的 XML 清单文件 (GetDoc_App.xml)。该清单文件必须指向前面提到的 HTML 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p103">An XML manifest file (GetDoc_App.xml) for the add-in, available on a shared network folder or add-in catalog. The manifest file must point to the location of the HTML file mentioned previously.</span></span>
    
<span data-ttu-id="0a7ec-113">也可以使用 [Visual Studio](../quickstarts/powerpoint-quickstart.md?tabs=visual-studio) 或[任意编辑器](../quickstarts/powerpoint-quickstart.md?tabs=visual-studio-code)创建 PowerPoint 加载项，或使用 [Visual Studio](../quickstarts/word-quickstart.md?tabs=visual-studio) 或[任意编辑器](../quickstarts/word-quickstart.md?tabs=visual-studio-code)创建 Word 加载项。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-113">You can also create an add-in for PowerPoint by using [Visual Studio](../quickstarts/powerpoint-quickstart.md?tabs=visual-studio) or [any editor](../quickstarts/powerpoint-quickstart.md?tabs=visual-studio-code) or for Word by using [Visual Studio](../quickstarts/word-quickstart.md?tabs=visual-studio) or [any editor](../quickstarts/word-quickstart.md?tabs=visual-studio-code).</span></span> 

### <a name="core-concepts-to-know-for-creating-a-task-pane-add-in"></a><span data-ttu-id="0a7ec-114">创建任务窗格加载项需要了解的核心概念</span><span class="sxs-lookup"><span data-stu-id="0a7ec-114">Core concepts to know for creating a task pane add-in</span></span>

<span data-ttu-id="0a7ec-p104">在开始创建 PowerPoint 或 Word 的此外接程序之前，您应知道如何构建 Office 外接程序和使用 HTTP 请求。本文不讨论如何解码 Web 服务器上 HTTP 请求中 Base64 编码的文本。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p104">Before you begin creating this add-in for PowerPoint or Word, you should be familiar with building Office Add-ins and working with HTTP requests. This article does not discuss how to decode Base64-encoded text from an HTTP request on a web server.</span></span> 

## <a name="create-the-manifest-for-the-add-in"></a><span data-ttu-id="0a7ec-117">为外接程序创建清单</span><span class="sxs-lookup"><span data-stu-id="0a7ec-117">Create the manifest for the add-in</span></span>


<span data-ttu-id="0a7ec-118">PowerPoint 外接程序的 XML 清单文件提供有关外接程序的重要信息：可以托管它的应用程序、HTML 文件的位置、外接程序标题和说明以及许多其他特征。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-118">The XML manifest file for the add-in for PowerPoint provides important information about the add-in: what applications can host it, the location of the HTML file, the add-in title and description, and many other characteristics.</span></span>

1. <span data-ttu-id="0a7ec-119">在文本编辑器中，将以下代码添加到清单文件中。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-119">In a text editor, add the following code to the manifest file.</span></span>
    
    ```xml  
    <?xml version="1.0" encoding="utf-8" ?> 
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
    xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
    xsi:type="TaskPaneApp">
        <Id>[Replace_With_Your_GUID]</Id> 
        <Version>1.0</Version> 
        <ProviderName>[Provider Name]</ProviderName> 
        <DefaultLocale>EN-US</DefaultLocale> 
        <DisplayName DefaultValue="Get Doc add-in" /> 
        <Description DefaultValue="My get PowerPoint or Word document add-in." /> 
        <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" /> 
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

2. <span data-ttu-id="0a7ec-120">使用 UTF-8 编码将文件以 GetDoc_App.xml 形式保存到网络位置或外接程序目录。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-120">Save the file as GetDoc_App.xml using UTF-8 encoding to a network location or to an add-in catalog.</span></span>
    
## <a name="create-the-user-interface-for-the-add-in"></a><span data-ttu-id="0a7ec-121">为外接程序创建用户界面</span><span class="sxs-lookup"><span data-stu-id="0a7ec-121">Create the user interface for the add-in</span></span>

<span data-ttu-id="0a7ec-p105">要为外接程序创建用户界面，可使用直接写入 GetDoc_App.html 文件的 HTML。外接程序的编程逻辑和功能必须包含在 JavaScript 文件（如 GetDoc_App.js）中。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p105">For the user interface of the add-in, you can use HTML, written directly into the GetDoc_App.html file. The programming logic and functionality of the add-in must be contained in a JavaScript file (for example, GetDoc_App.js).</span></span>

<span data-ttu-id="0a7ec-124">使用以下过程可为该外接程序创建一个包含标题和单个按钮的简单用户界面。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-124">Use the following procedure to create a simple user interface for the add-in that includes a heading and a single button.</span></span>

1. <span data-ttu-id="0a7ec-125">在文本编辑器的新文件中，添加以下 HTML。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-125">In a new file in the text editor, add the following HTML.</span></span>
        
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

2. <span data-ttu-id="0a7ec-126">使用 UTF-8 编码将文件以 GetDoc_App.html 形式保存到网络位置或 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-126">Save the file as GetDoc_App.html using UTF-8 encoding to a network location or to a web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="0a7ec-127">请确保加载项的 **head** 标记包含 **script** 标记，其中包含 office.js 文件的有效链接。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-127">Be sure that the **head** tags of the add-in contains a **script** tag with a valid link to the office.js file.</span></span> 

    <span data-ttu-id="0a7ec-p106">我们将使用一些 CSS 为外接程序提供一个简洁、现代且具专业水准的外观。使用以下 CSS 可定义外接程序的样式。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p106">We'll use some CSS to give the add-in a simple, yet modern and professional appearance. Use the following CSS to define the style of the add-in.</span></span>

3. <span data-ttu-id="0a7ec-130">在文本编辑器的新文件中，添加以下 CSS。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-130">In a new file in the text editor, add the following CSS.</span></span>
        
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

4. <span data-ttu-id="0a7ec-131">使用 UTF-8 编码将该文件以 Program.css 形式保存到网络位置，或保存到 GetDoc_App.html 文件所在的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-131">Save the file as Program.css using UTF-8 encoding to the network location or to the web server where the GetDoc_App.html file is located.</span></span>
    
## <a name="add-the-javascript-to-get-the-document"></a><span data-ttu-id="0a7ec-132">添加 JavaScript 以获取文档</span><span class="sxs-lookup"><span data-stu-id="0a7ec-132">Add the JavaScript to get the document</span></span>

<span data-ttu-id="0a7ec-133">在外接程序的代码中，[Office.initialize](https://docs.microsoft.com/javascript/api/office) 事件的处理程序会向表单上**提交**按钮的 Click 事件中添加处理程序，并告知用户外接程序准备就绪。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-133">In the code for the add-in, a handler to the [Office.initialize](https://docs.microsoft.com/javascript/api/office) event adds a handler to the click event of the **Submit** button on the form and informs the user that the add-in is ready.</span></span>

<span data-ttu-id="0a7ec-134">以下代码示例演示  **Office.initialize** 事件的事件处理程序，以及用于写入状态 div 的 Helper 函数 `updateStatus`。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-134">The following code example shows the event handler for the  **Office.initialize** event along with a helper function, `updateStatus`, for writing to the status div.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {

        // Execute sendFile when submit is clicked 
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
    statusInfo.innerHTML += message + "<br/>";
}
```

<span data-ttu-id="0a7ec-p107">当您选择 UI 中的**提交**按钮时，外接程序会调用 `sendFile` 函数（包含对 [Document.getFileAsync](https://docs.microsoft.com/javascript/api/office/office.document#getfileasync-filetype--options--callback-) 方法的调用）。**getFileAsync** 方法使用异步模式，这与 JavaScript API for Office 中的其他方法类似。它包含一个必需参数 _fileType_ 以及两个可选参数 _options_ 和 _callback_。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p107">When you choose the  **Submit** button in the UI, the add-in calls the `sendFile` function, which contains a call to the [Document.getFileAsync](https://docs.microsoft.com/javascript/api/office/office.document#getfileasync-filetype--options--callback-) method. The **getFileAsync** method uses the asynchronous pattern, similar to other methods in the JavaScript API for Office. It has one required parameter, _fileType_, and two optional parameters,  _options_ and _callback_.</span></span> 

<span data-ttu-id="0a7ec-p108">_fileType_ 形参需要 [FileType](https://docs.microsoft.com/javascript/api/office/office.filetype) 枚举中三个常量中的一个：**Office.FileType.Compressed** ("compressed")、**Office.FileType.PDF** ("pdf") 或 **Office.FileType.Text** ("text")。PowerPoint 仅支持将 **Compressed** 作为实参；Word 支持这三者。当您为 **fileType** 形参传入 _Compressed_ 时，**getFileAsync** 方法将通过在本地计算机上创建文件的临时副本，来将文档作为 PowerPoint 2013 演示文稿文件 (*.pptx) 或 Word 2013 文档文件 (*.docx) 返回。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p108">The  _fileType_ parameter expects one of three constants from the [FileType](https://docs.microsoft.com/javascript/api/office/office.filetype) enumeration: **Office.FileType.Compressed** ("compressed"), **Office.FileType.PDF** ("pdf"), or **Office.FileType.Text** ("text"). PowerPoint supports only **Compressed** as an argument; Word supports all three. When you pass in **Compressed** for the _fileType_ parameter, the **getFileAsync** method returns the document as a PowerPoint 2013 presentation file (*.pptx) or Word 2013 document file (*.docx) by creating a temporary copy of the file on the local computer.</span></span>

<span data-ttu-id="0a7ec-p109">**getFileAsync** 方法将对文件的引用作为 [File](https://docs.microsoft.com/javascript/api/office/office.file) 对象返回。**File** 对象公开四个成员：[size](https://docs.microsoft.com/javascript/api/office/office.file#size) 属性、[sliceCount](https://docs.microsoft.com/javascript/api/office/office.file#slicecount) 属性、[getSliceAsync](https://docs.microsoft.com/javascript/api/office/office.file#getsliceasync-sliceindex--callback-) 方法和 [closeAsync](https://docs.microsoft.com/javascript/api/office/office.file#closeasync-callback-) 方法。**size** 属性返回文件中的字节数。**sliceCount** 返回文件中 [Slice](https://docs.microsoft.com/javascript/api/office/office.slice) 对象（在下文中讨论）的数目。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p109">The  **getFileAsync** method returns a reference to the file as a [File](https://docs.microsoft.com/javascript/api/office/office.file) object. The **File** object exposes four members: the [size](https://docs.microsoft.com/javascript/api/office/office.file#size) property, [sliceCount](https://docs.microsoft.com/javascript/api/office/office.file#slicecount) property, [getSliceAsync](https://docs.microsoft.com/javascript/api/office/office.file#getsliceasync-sliceindex--callback-) method, and [closeAsync](https://docs.microsoft.com/javascript/api/office/office.file#closeasync-callback-) method. The **size** property returns the number of bytes in the file. The **sliceCount** returns the number of [Slice](https://docs.microsoft.com/javascript/api/office/office.slice) objects (discussed later in this article) in the file.</span></span>

<span data-ttu-id="0a7ec-p110">使用以下代码时，将通过  **Document.getFileAsync** 方法以 **File** 对象的形式获取 PowerPoint 或 Word 文档，然后调用本地定义的 `getSlice` 函数。请注意，在调用匿名对象中的 `getSlice` 时，将传入 **File** 对象（一个计数器变量）以及文件中切片的总数。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p110">Use the following code to get the PowerPoint or Word document as a  **File** object using the **Document.getFileAsync** method and then makes a call to the locally defined `getSlice` function. Note that the **File** object, a counter variable, and the total number of slices in the file are passed along in the call to `getSlice` in an anonymous object.</span></span>

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

<span data-ttu-id="0a7ec-p111">本地函数  `getSlice` 可对 **File.getSliceAsync** 方法进行调用，以从 **File** 对象中检索切片。 **getSliceAsync** 方法返回切片集合中的 **Slice** 对象。它具有两个必需参数： _sliceIndex_ 和 _callback_。 _sliceIndex_ 参数将整数作为切块集合中的索引器。与 JavaScript API for Office 中的其他函数一样， **getSliceAsync** 方法还将回调函数作为参数，以处理方法调用的结果。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p111">The local function  `getSlice` makes a call to the **File.getSliceAsync** method to retrieve a slice from the **File** object. The **getSliceAsync** method returns a **Slice** object from the collection of slices. It has two required parameters, _sliceIndex_ and _callback_. The  _sliceIndex_ parameter takes an integer as an indexer into the collection of slices. Like other functions in the JavaScript API for Office, the **getSliceAsync** method also takes a callback function as a parameter to handle the results from the method call.</span></span>

<span data-ttu-id="0a7ec-152">**切片**对象提供对文件中包含的数据的访问权限。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-152">The **Slice** object gives you access to the data contained in the file.</span></span> <span data-ttu-id="0a7ec-153">**切片**对象的大小为 4 MB，除非 **getFileAsync** 方法的_选项_参数中另有指定。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-153">Unless otherwise specified in the _options_ parameter of the **getFileAsync** method, the **Slice** object is 4 MB in size.</span></span> <span data-ttu-id="0a7ec-154">**切片**对象公开三个属性：[大小](https://docs.microsoft.com/javascript/api/office/office.slice#size)、[数据](https://docs.microsoft.com/javascript/api/office/office.slice#data)和[索引](https://docs.microsoft.com/javascript/api/office/office.slice#index)。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-154">The **Slice** object exposes three properties: [size](https://docs.microsoft.com/javascript/api/office/office.slice#size), [data](https://docs.microsoft.com/javascript/api/office/office.slice#data), and [index](https://docs.microsoft.com/javascript/api/office/office.slice#index).</span></span> <span data-ttu-id="0a7ec-155">**大小**属性获取以字节为单位的切片大小。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-155">The **size** property gets the size, in bytes, of the slice.</span></span> <span data-ttu-id="0a7ec-156">**索引**属性获取一个整数，表示切片在切片集合中的位置。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-156">The **index** property gets an integer that represents the slice's position in the collection of slices.</span></span>

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

<span data-ttu-id="0a7ec-p113">
            *\*Slice.data\*\* 属性以字节数组形式返回文件的原始数据。如果数据采用文本格式（即 XML 或纯文本），则切片包含原始文本。如果您为 \*\*Document.getFileAsync** 的 _fileType_ 参数传入 \*\*Office.FileType.Compressed\*\*，则切片将以字节数组形式包含文件的二进制数据。对于 PowerPoint 或 Word 文件，切片包含字节数组。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p113">The  **Slice.data** property returns the raw data of the file as a byte array. If the data is in text format (that is, XML or plain text), the slice contains the raw text. If you pass in **Office.FileType.Compressed** for the _fileType_ parameter of **Document.getFileAsync**, the slice contains the binary data of the file as a byte array. In the case of a PowerPoint or Word file, the slices contain byte arrays.</span></span>

<span data-ttu-id="0a7ec-p114">您必须实施自己的函数（或使用可用库），将字节数组数据转换为 Base64 编码的字符串。有关使用 JavaScript 进行 Base64 编码的信息，请参阅 [Base64 编码和解码](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding)。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p114">You must implement your own function (or use an available library) to convert byte array data to a Base64-encoded string. For information about Base64 encoding with JavaScript, see [Base64 encoding and decoding](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).</span></span>

<span data-ttu-id="0a7ec-163">将数据转换为 Base64 后，即可通过多种方法（包括作为 HTTP POST 请求的正文）将其传输到 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-163">Once you have converted the data to Base64, you can then transmit it to a web server in several ways -- including as the body of an HTTP POST request.</span></span>

<span data-ttu-id="0a7ec-164">添加以下代码，将切片发送到 Web 服务。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-164">Add the following code to send a slice to a web service.</span></span>

> [!NOTE]
> <span data-ttu-id="0a7ec-p115">此代码通过多个切片将 PowerPoint 或 Word 文件发送到 Web 服务器。Web 服务器或服务必须将每个单独切片编译为一个 .pptx 文件，之后您才能对其执行任何操作。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p115">This code sends a PowerPoint or Word file to the web server in multiple slices. The web server or service must compile each individual slice into a single .pptx file before you can perform any manipulations on it.</span></span>

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

<span data-ttu-id="0a7ec-p116">顾名思义， **File.closeAsync** 方法会关闭与文档的连接并释放资源。虽然 Office 外接程序沙盒垃圾可回收对文件的范围外引用，但在使用这些文件完成您的代码后，最好显式关闭它们。 **closeAsync** 方法有一个参数 _callback_，可指定调用完成时要调用的函数。</span><span class="sxs-lookup"><span data-stu-id="0a7ec-p116">As the name implies, the  **File.closeAsync** method closes the connection to the document and frees up resources. Although the Office Add-ins sandbox garbage collects out-of-scope references to files, it is still a best practice to explicitly close files once your code is done with them. The **closeAsync** method has a single parameter, _callback_, that specifies the function to call on the completion of the call.</span></span>

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