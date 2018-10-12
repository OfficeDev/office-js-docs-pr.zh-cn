# <a name="word-javascript-api-overview"></a><span data-ttu-id="4333c-101">Word JavaScript API 概述</span><span class="sxs-lookup"><span data-stu-id="4333c-101">Word JavaScript API usage overview</span></span>

<span data-ttu-id="4333c-p101">Word 提供了一组丰富的 API，你可以使用它们创建与文档内容和元数据进行交互的外接程序。使用这些 API 可以为用户带来与 Word 融为一体并扩展 Word 的精彩体验。你可以导入和导出内容、组合来自不同数据源的新文档，并能与文档工作流进行集成，从而创建自定义文档解决方案。</span><span class="sxs-lookup"><span data-stu-id="4333c-p101">Word provides a rich set of APIs that you can use to create add-ins that interact with document content and metadata. Use these APIs to create compelling experiences that integrate with and extend Word. You can import and export content, assemble new documents from different data sources, and integrate with document workflows to create custom document solutions.</span></span>

<span data-ttu-id="4333c-105">你可以使用以下两个 JavaScript API 与 Word 文档中的对象和元数据进行交互：</span><span class="sxs-lookup"><span data-stu-id="4333c-105">You can use two JavaScript APIs to interact with the objects and metadata in a Word document:</span></span>

- <span data-ttu-id="4333c-106">Word JavaScript API - 在 Office 2016 中引入。</span><span class="sxs-lookup"><span data-stu-id="4333c-106">Word JavaScript API - Introduced in Office 2016.</span></span>
- <span data-ttu-id="4333c-107">[适用于 Office 的 JavaScript API](../javascript-api-for-office.md) (Office.js) - 在 Office 2013 中引入。</span><span class="sxs-lookup"><span data-stu-id="4333c-107">[JavaScript API for Office](../javascript-api-for-office.md) (Office.js) - Introduced in Office 2013.</span></span>

## <a name="word-javascript-api"></a><span data-ttu-id="4333c-108">Word JavaScript API</span><span class="sxs-lookup"><span data-stu-id="4333c-108">Word JavaScript API</span></span>

<span data-ttu-id="4333c-p102">Word JavaScript API 通过 Office.js 进行加载，它改变了你与文档和段落等对象的交互方式。Word JavaScript API 不提供各个用于检索和更新每个对象的异步 API，而是提供与 Word 中运行的真实对象对应的“代理”JavaScript 对象。你可以通过同步读取和写入这些代理对象的属性，并调用对其执行操作的同步方法，从而与这些代理对象进行交互。与代理对象的这些交互不会立即在运行的脚本中实现。**context.sync** 方法通过执行已排入队列的指令并检索可供在脚本中使用的已加载 Word 对象的属性，在运行的 JavaScript 和 Office 真实对象之间同步状态。</span><span class="sxs-lookup"><span data-stu-id="4333c-p102">The Word JavaScript API is loaded by Office.js. The Word JavaScript API changes the way that you can interact with objects like documents and paragraphs. Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the Word JavaScript API provides “proxy” JavaScript objects that correspond to the real objects running in Word. You can interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them. These interactions with proxy objects aren’t immediately realized in the running script. The **context.sync** method synchronizes the state between your running JavaScript and the real objects in Office by executing queued instructions and retrieving properties of loaded Word objects for use in your script.</span></span>

## <a name="javascript-api-for-office"></a><span data-ttu-id="4333c-115">适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="4333c-115">JavaScript API for Office</span></span>

<span data-ttu-id="4333c-116">你可以从以下位置引用 Office.js：</span><span class="sxs-lookup"><span data-stu-id="4333c-116">You can reference Office.js from the following locations:</span></span>

* <span data-ttu-id="4333c-117">https://appsforoffice.microsoft.com/lib/1/hosted/office.js - 在生产外接程序中使用此资源用。</span><span class="sxs-lookup"><span data-stu-id="4333c-117">https://appsforoffice.microsoft.com/lib/1/hosted/office.js - use this resource for production add-ins.</span></span>
* <span data-ttu-id="4333c-118">https://appsforoffice.microsoft.com/lib/beta/hosted/office.js - 尝试预览功能时使用此资源。</span><span class="sxs-lookup"><span data-stu-id="4333c-118">https://appsforoffice.microsoft.com/lib/beta/hosted/office.js - use this resource when you're trying out preview features.</span></span>

<span data-ttu-id="4333c-p103">如果你使用的是 [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs)，则可以下载 [Office 开发人员工具](https://www.visualstudio.com/features/office-tools-vs.aspx)，从而获取包含 Office.js 的项目模板。你还可以使用 [nuget 获取 Office.js](https://www.nuget.org/packages/Microsoft.Office.js/)。</span><span class="sxs-lookup"><span data-stu-id="4333c-p103">If you're using [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs), you can download the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) to get project templates that include Office.js.  You can also use [nuget to get Office.js](https://www.nuget.org/packages/Microsoft.Office.js/).</span></span>

<span data-ttu-id="4333c-121">如果你使用的是 TypeScript 并且拥有 npm，则可以在命令行接口中键入以下命令，从而获取 TypeScript 定义：`typings install office-js --ambient`。</span><span class="sxs-lookup"><span data-stu-id="4333c-121">If you use TypeScript and have npm, you can get the the TypeScript definitions by typing this in your command line interface: `typings install office-js --ambient`.</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="4333c-122">运行 Word 外接程序</span><span class="sxs-lookup"><span data-stu-id="4333c-122">Running Word add-ins</span></span>

<span data-ttu-id="4333c-p104">若要运行外接程序，请使用 Office.initialize 事件处理程序。若要详细了解如何初始化外接程序，请参阅[了解 API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)。</span><span class="sxs-lookup"><span data-stu-id="4333c-p104">To run your add-in, use an Office.initialize event handler. For more information about add-in initialization, see [Understanding the API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office) .</span></span>

<span data-ttu-id="4333c-125">以 Word 2016 或更高版本为目标的外接程序通过将一个函数传递到 **Word.run()** 方法来执行。</span><span class="sxs-lookup"><span data-stu-id="4333c-125">Add-ins that target Word 2016 or later execute by passing a function into the **Word.run()** method.</span></span> <span data-ttu-id="4333c-126">向 **run** 方法传递的函数必须具有上下文参数。</span><span class="sxs-lookup"><span data-stu-id="4333c-126">The function passed into the **run** method must have a context argument.</span></span> <span data-ttu-id="4333c-127">此[ context 对象](/javascript/api/word/word.requestcontext) 与从 Office 对象获取的 context 对象不同，但它也用于与 Word 运行时环境交互。</span><span class="sxs-lookup"><span data-stu-id="4333c-127">This [context object](/javascript/api/word/word.requestcontext) is different than the context object you get from the Office object, but it is also used to interact with the Word runtime environment.</span></span> <span data-ttu-id="4333c-128">Context 对象提供对 Word JavaScript API 对象模型的访问。</span><span class="sxs-lookup"><span data-stu-id="4333c-128">The context object provides access to the Word JavaScript API object model.</span></span> <span data-ttu-id="4333c-129">下面的示例显示如何使用 **Word.run()** 方法初始化并执行 Word 加载项。</span><span class="sxs-lookup"><span data-stu-id="4333c-129">The following example shows how to initialize and execute a Word add-in by using the **Word.run()** method.</span></span>

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a><span data-ttu-id="4333c-130">将 Word 文档与 Word JavaScript API 代理对象进行同步</span><span class="sxs-lookup"><span data-stu-id="4333c-130">Synchronizing Word documents with Word JavaScript API proxy objects</span></span>

<span data-ttu-id="4333c-p106">Word JavaScript API 对象模型与 Word 中的对象松散耦合。Word JavaScript API 对象是 Word 文档中对象的代理。在文档状态完成同步前，对代理对象执行的操作不会在 Word 中实现。反过来说，在文档状态完成同步前，Word 文档的状态也不会在代理对象中实现。若要同步文档状态，请运行 **context.sync()** 方法。下面的示例创建了代理正文对象以及用于在代理正文对象上加载文本属性的已排入队列命令，并使用 **context.sync()** 方法将 Word 文档正文与正文代理对象同步。</span><span class="sxs-lookup"><span data-stu-id="4333c-p106">The Word JavaScript API object model is loosely coupled with the objects in Word. Word JavaScript API objects are proxies for objects in a Word document. Actions taken on proxy objects are not realized in Word until the document state has been synchronized. Conversely, the state of the Word document is not realized in the proxy objects until the document state has been synchronized. To synchronize the document state, you run the **context.sync()** method. The following example creates a proxy body object and a queued command to load the text property on the proxy body object, and uses the **context.sync()** method to synchronize the body of the Word document with the body proxy object.</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a><span data-ttu-id="4333c-137">执行一批命令</span><span class="sxs-lookup"><span data-stu-id="4333c-137">Executing a batch of commands</span></span>

<span data-ttu-id="4333c-p107">Word 代理对象具有用于访问和更新对象模型的方法。这些方法按其在批处理中的排队顺序依次执行。调用 context.sync() 后，批处理中已排入队列的所有命令都会得到执行。</span><span class="sxs-lookup"><span data-stu-id="4333c-p107">The Word proxy objects have methods for accessing and updating the object model. These methods are executed sequentially in the order in which they were queued in the batch. All of the commands that are queued in the batch are executed when context.sync() is called.</span></span>

<span data-ttu-id="4333c-p108">下面的示例展示了命令队列的工作原理。调用 **context.sync()** 时，用于加载正文文本的命令会在 Word 中执行。然后，用于在正文中插入文本的命令会在 Word 中执行。接下来，结果会返回到正文代理对象。Word JavaScript API 中 **body.text** 属性的值为在将文本插入 Word 文档<u>之前</u> Word 文档正文的值。</span><span class="sxs-lookup"><span data-stu-id="4333c-p108">The following example shows how the command queue works. When **context.sync()** is called, the command to load the body text is executed in Word. Then, the command to insert text into the body in Word occurs. The results are then returned to the body proxy object. The value of the **body.text** property in the Word JavaScript API is the value of the Word document body <u>before</u> the text was inserted into Word document.</span></span>


```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    context.load(body, 'text');

    // Queue a command to insert text into the end of the Word document body.
    body.insertText('This is text inserted after loading the body.text property',
                    Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

## <a name="word-javascript-api-open-specifications"></a><span data-ttu-id="4333c-146">Word 的 JavaScript API 开放性规范</span><span class="sxs-lookup"><span data-stu-id="4333c-146">Word JavaScript API open specifications</span></span>

<span data-ttu-id="4333c-p109">在我们设计和开发新的 API 以用于创建 Word 外接程序时，我们会公开它们，以便你可以在我们的[开放性 API 规范](../openspec.md)页面上提供反馈。了解即将推出的面向 Word JavaScript API 的新功能，并提供你对我们的设计规范的宝贵意见。</span><span class="sxs-lookup"><span data-stu-id="4333c-p109">As we design and develop new APIs for Word add-ins, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline for the Word JavaScript APIs, and provide your input on our design specifications.</span></span>

## <a name="word-javascript-api-reference"></a><span data-ttu-id="4333c-149">Word JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="4333c-149">Word JavaScript API reference</span></span>

<span data-ttu-id="4333c-150">有关 Word JavaScript API 的详细信息，请参阅 [Word JavaScript API 参考文档](/javascript/api/word)。</span><span class="sxs-lookup"><span data-stu-id="4333c-150">For detailed information about the Word JavaScript API, see the [Word JavaScript API reference documentation](/javascript/api/word).</span></span>

## <a name="see-also"></a><span data-ttu-id="4333c-151">另请参阅</span><span class="sxs-lookup"><span data-stu-id="4333c-151">See also</span></span>

* [<span data-ttu-id="4333c-152">Word 外接程序概述</span><span class="sxs-lookup"><span data-stu-id="4333c-152">Word add-ins overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/word/word-add-ins-programming-overview)
* [<span data-ttu-id="4333c-153">Office 外接程序平台概述</span><span class="sxs-lookup"><span data-stu-id="4333c-153">Office Add-ins platform overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* [<span data-ttu-id="4333c-154">GitHub 上的 Word 外接程序示例</span><span class="sxs-lookup"><span data-stu-id="4333c-154">Word add-in samples on GitHub</span></span>](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Word)
