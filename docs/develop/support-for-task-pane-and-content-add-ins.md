---
title: 对 Office 2013 内容和任务窗格加载项的 Office JavaScript API 支持
description: ''
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: a9eb67ca78f89888860ff3ed11ae1632ff62b690
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42323818"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a><span data-ttu-id="e56f8-102">对 Office 2013 中内容和任务窗格外接程序的 Office JavaScript API 支持</span><span class="sxs-lookup"><span data-stu-id="e56f8-102">Office JavaScript API support for content and task pane add-ins in Office 2013</span></span>

[!include[information about the common API](../includes/alert-common-api-info.md)]

<span data-ttu-id="e56f8-p101">您可以使用 [Office JavaScript API](../reference/javascript-api-for-office.md) 创建 Office 2013 主机应用程序的任务窗格或内容外接程序。已对内容和任务窗格外接程序支持的对象和方法进行如下分类：</span><span class="sxs-lookup"><span data-stu-id="e56f8-p101">You can use the [Office JavaScript API](../reference/javascript-api-for-office.md) to create task pane or content add-ins for Office 2013 host applications. The objects and methods that content and task pane add-ins support are categorized as follows:</span></span>

1. <span data-ttu-id="e56f8-105">**与其他 Office 外接程序共享的常见对象。** 这些对象包括[Office](/javascript/api/office)、 [Context](/javascript/api/office/office.context)和[AsyncResult](/javascript/api/office/office.asyncresult)。</span><span class="sxs-lookup"><span data-stu-id="e56f8-105">**Common objects shared with other Office Add-ins.** These objects include [Office](/javascript/api/office), [Context](/javascript/api/office/office.context), and [AsyncResult](/javascript/api/office/office.asyncresult).</span></span> <span data-ttu-id="e56f8-106">`Office`对象是 OFFICE JavaScript API 的根对象。</span><span class="sxs-lookup"><span data-stu-id="e56f8-106">The `Office` object is the root object of the Office JavaScript API.</span></span> <span data-ttu-id="e56f8-107">该`Context`对象表示加载项的运行时环境。</span><span class="sxs-lookup"><span data-stu-id="e56f8-107">The `Context` object represents the add-in's runtime environment.</span></span> <span data-ttu-id="e56f8-108">`Office`和`Context`都是适用于任何 Office 外接程序的基本对象。</span><span class="sxs-lookup"><span data-stu-id="e56f8-108">Both `Office` and `Context` are the fundamental objects for any Office Add-in.</span></span> <span data-ttu-id="e56f8-109">`AsyncResult`对象表示异步操作的结果，如返回到`getSelectedDataAsync`方法的数据，该方法读取用户在文档中选定的内容。</span><span class="sxs-lookup"><span data-stu-id="e56f8-109">The `AsyncResult` object represents the results of an asynchronous operation, such as the data returned to the `getSelectedDataAsync` method, which reads what a user has selected in a document.</span></span>

2. <span data-ttu-id="e56f8-110">**Document 对象。**</span><span class="sxs-lookup"><span data-stu-id="e56f8-110">**The Document object.**</span></span> <span data-ttu-id="e56f8-111">可通过 [Document](/javascript/api/office/office.document) 对象的方法、属性和事件公开大多数可用于内容和任务窗格加载项的 API。</span><span class="sxs-lookup"><span data-stu-id="e56f8-111">The majority of the API available to content and task pane add-ins is exposed through the methods, properties, and events of the [Document](/javascript/api/office/office.document) object.</span></span> <span data-ttu-id="e56f8-112">内容或任务窗格加载项可以使用[CustomXmlParts 属性访问](/javascript/api/office/office.context#document) **document**对象，通过它，可以访问用于处理文档中的数据（如[绑定](/javascript/api/office/office.bindings)和[](/javascript/api/office/office.customxmlparts)对象）的 API 的关键成员，以及[getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-)、 [document.setselecteddataasync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-)和[document.getfileasync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)方法。</span><span class="sxs-lookup"><span data-stu-id="e56f8-112">A content or task pane add-in can use the [Office.context.document](/javascript/api/office/office.context#document) property to access the **Document** object, and through it, can access the key members of the API for working with data in documents, such as the [Bindings](/javascript/api/office/office.bindings) and [CustomXmlParts](/javascript/api/office/office.customxmlparts) objects, and the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-), [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-), and [getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-) methods.</span></span> <span data-ttu-id="e56f8-113">该`Document`对象还提供了[mode](/javascript/api/office/office.document#mode)属性，用于确定文档是否为只读或处于编辑模式、用于获取当前文档的 url 的[Url](/javascript/api/office/office.document#url)属性，以及对[Settings](/javascript/api/office/office.settings)对象的访问权限。</span><span class="sxs-lookup"><span data-stu-id="e56f8-113">The `Document` object also provides the [mode](/javascript/api/office/office.document#mode) property for determining whether a document is read-only or in edit mode, the [url](/javascript/api/office/office.document#url) property to get the URL of the current document, and access to the [Settings](/javascript/api/office/office.settings) object.</span></span> <span data-ttu-id="e56f8-114">该`Document`对象还支持添加[SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs)事件的事件处理程序，以便您可以检测用户何时更改文档中的选定内容。</span><span class="sxs-lookup"><span data-stu-id="e56f8-114">The `Document` object also supports adding event handlers for the [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event, so you can detect when a user changes their selection in the document.</span></span>

   <span data-ttu-id="e56f8-115">仅在加载了 DOM 和运行时环境之后， `Document`内容或任务窗格外接程序才能访问该对象，这通常是在运行的 " [Office initialize](/javascript/api/office) " 事件的事件处理程序中。</span><span class="sxs-lookup"><span data-stu-id="e56f8-115">A content or task pane add-in can access the `Document` object only after the DOM and runtime environment has been loaded, typically in the event handler for the [Office.initialize](/javascript/api/office) event.</span></span> <span data-ttu-id="e56f8-116">有关应用程序初始化时的事件流以及如何检查 DOM 和运行时是否成功加载的信息，请参阅[加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)。</span><span class="sxs-lookup"><span data-stu-id="e56f8-116">For information about the flow of events when an add-in is initialized, and how to check that the DOM and runtime and loaded successfully, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

3. <span data-ttu-id="e56f8-p105">**使用特定的功能的对象。** 若要使用 API 的特定功能，请使用下面的对象和方法：</span><span class="sxs-lookup"><span data-stu-id="e56f8-p105">**Objects for working with specific features.** To work with specific features of the API, use the following objects and methods:</span></span>

    - <span data-ttu-id="e56f8-119">创建或获取绑定的 [Bindings](/javascript/api/office/office.bindings) 对象的方法，以及使用数据的 [Binding](/javascript/api/office/office.binding) 对象的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="e56f8-119">The methods of the [Bindings](/javascript/api/office/office.bindings) object to create or get bindings, and the methods and properties of the [Binding](/javascript/api/office/office.binding) object to work with data.</span></span>

    - <span data-ttu-id="e56f8-120">创建和操控 Word 文档中自定义的 XML 部件的 [CustomXmlParts](/javascript/api/office/office.customxmlparts)、[CustomXmlPart](/javascript/api/office/office.customxmlpart) 和关联的对象。</span><span class="sxs-lookup"><span data-stu-id="e56f8-120">The [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart) and associated objects to create and manipulate custom XML parts in Word documents.</span></span>

    - <span data-ttu-id="e56f8-121">创建整个文档的副本，将它分解成多个块或“切片”，然后读取或传输这些切片中数据的 [File](/javascript/api/office/office.file) 和 [Slice](/javascript/api/office/office.slice) 对象。</span><span class="sxs-lookup"><span data-stu-id="e56f8-121">The [File](/javascript/api/office/office.file) and [Slice](/javascript/api/office/office.slice) objects to create a copy of the entire document, break it into chunks or "slices", and then read or transmit the data in those slices.</span></span>

    - <span data-ttu-id="e56f8-122">[Settings](/javascript/api/office/office.settings) 对象，用于保存自定义数据（如用户偏好设置）和加载项状态。</span><span class="sxs-lookup"><span data-stu-id="e56f8-122">The [Settings](/javascript/api/office/office.settings) object to save custom data, such as user preferences, and add-in state.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="e56f8-123">并不是所有能够承载内容和任务窗格加载项的 Office 应用程序都支持一些 API 成员。要确定支持哪些成员，请参阅以下任一资源：</span><span class="sxs-lookup"><span data-stu-id="e56f8-123">Some of the API members aren't supported across all Office applications that can host content and task pane add-ins. To determine which members are supported, see any of the following:</span></span>

<span data-ttu-id="e56f8-124">有关跨 Office 主机应用程序的 Office JavaScript API 支持的摘要，请参阅[了解 Office JAVASCRIPT api](understanding-the-javascript-api-for-office.md)。</span><span class="sxs-lookup"><span data-stu-id="e56f8-124">For a summary of Office JavaScript API support across Office host applications, see [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md).</span></span>


## <a name="reading-and-writing-to-an-active-selection"></a><span data-ttu-id="e56f8-125">在活动的选择内容中读取和写入</span><span class="sxs-lookup"><span data-stu-id="e56f8-125">Reading and writing to an active selection</span></span>

<span data-ttu-id="e56f8-126">您可以在文档、电子表格或演示文稿的用户当前选定内容中读取和写入。</span><span class="sxs-lookup"><span data-stu-id="e56f8-126">You can read or write to the user's current selection in a document, spreadsheet, or presentation.</span></span> <span data-ttu-id="e56f8-127">根据加载项的主机应用程序，您可以在 [Document](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) 对象的 [getSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 和 [setSelectedDataAsync](/javascript/api/office/office.document) 方法中指定要作为参数来读取或写入的数据结构类型。</span><span class="sxs-lookup"><span data-stu-id="e56f8-127">Depending on the host application for your add-in, you can specify the type of data structure to read or write as a parameter in the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) and [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) methods of the [Document](/javascript/api/office/office.document) object.</span></span> <span data-ttu-id="e56f8-128">例如，您可以指定任何用于 Word 的数据类型（文本、HTML、表格数据或 Office Open XML）、用于 Excel 的文本和表格数据，以及用于 PowerPoint 和 Project 的文本。</span><span class="sxs-lookup"><span data-stu-id="e56f8-128">For example, you can specify any type of data (text, HTML, tabular data, or Office Open XML) for Word, text and tabular data for Excel, and text for PowerPoint and Project.</span></span> <span data-ttu-id="e56f8-129">您还可以创建事件处理程序来检测对用户选择内容的更改。</span><span class="sxs-lookup"><span data-stu-id="e56f8-129">You can also create event handlers to detect changes to the user's selection.</span></span> <span data-ttu-id="e56f8-130">下面的示例使用`getSelectedDataAsync`方法以文本形式获取选定内容中的数据。</span><span class="sxs-lookup"><span data-stu-id="e56f8-130">The following example gets data from the selection as text using the `getSelectedDataAsync` method.</span></span>


```js
Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        }
        else {
            write('Selected data: ' + asyncResult.value);
        }
    });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}

```

<span data-ttu-id="e56f8-131">有关详细信息和示例，请参阅[将数据读取和写入到文档或电子表格中的活动选择区](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。</span><span class="sxs-lookup"><span data-stu-id="e56f8-131">For more details and examples, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).</span></span>


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a><span data-ttu-id="e56f8-132">绑定到文档或电子表格中的区域</span><span class="sxs-lookup"><span data-stu-id="e56f8-132">Binding to a region in a document or spreadsheet</span></span>

<span data-ttu-id="e56f8-133">您可以使用`getSelectedDataAsync`和`setSelectedDataAsync`方法在文档、电子表格或演示文稿中读取或写入用户的*当前*选择。</span><span class="sxs-lookup"><span data-stu-id="e56f8-133">You can use the `getSelectedDataAsync` and `setSelectedDataAsync` methods to read or write to the user's *current* selection in a document, spreadsheet, or presentation.</span></span> <span data-ttu-id="e56f8-134">但是，如果您想在不要求用户选定内容的情况下，在运行您外接程序的各个会话中访问文档中的同一区域，您应首先绑定到该区域。</span><span class="sxs-lookup"><span data-stu-id="e56f8-134">However, if you would like to access the same region in a document across sessions of running your add-in without requiring the user to make a selection, you should first bind to that region.</span></span> <span data-ttu-id="e56f8-135">您还可以订阅该绑定区域的数据和选定内容更改事件。</span><span class="sxs-lookup"><span data-stu-id="e56f8-135">You can also subscribe to data and selection change events for that bound region.</span></span>

<span data-ttu-id="e56f8-p108">可以使用 [Bindings](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-) 对象的 [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-)、[addFromPromptAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) 或 [addFromSelectionAsync](/javascript/api/office/office.bindings) 方法添加绑定。这些方法可以返回一个标识符，您可以用它访问绑定中的数据或者订阅数据更改或选择更改事件。</span><span class="sxs-lookup"><span data-stu-id="e56f8-p108">You can add a binding by using [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-), [addFromPromptAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-), or [addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) methods of the [Bindings](/javascript/api/office/office.bindings) object. These methods return an identifier that you can use to access data in the binding, or to subscribe to its data change or selection change events.</span></span>

<span data-ttu-id="e56f8-138">下面的示例演示如何使用`Bindings.addFromSelectionAsync`方法将绑定添加到文档中的当前选定文本。</span><span class="sxs-lookup"><span data-stu-id="e56f8-138">The following is an example that adds a binding to the currently selected text in a document, by using the `Bindings.addFromSelectionAsync` method.</span></span>



```js
Office.context.document.bindings.addFromSelectionAsync(
    Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' +
            asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="e56f8-139">有关详细信息和示例，请参阅[绑定到文档或电子表格中的区域](bind-to-regions-in-a-document-or-spreadsheet.md)。</span><span class="sxs-lookup"><span data-stu-id="e56f8-139">For more details and examples, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>


## <a name="getting-entire-documents"></a><span data-ttu-id="e56f8-140">获取整个文档</span><span class="sxs-lookup"><span data-stu-id="e56f8-140">Getting entire documents</span></span>

<span data-ttu-id="e56f8-141">如果任务窗格外接程序在 PowerPoint 或 Word 中运行，您可以使用 [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)、[File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-) 和 [File.closeAsync](/javascript/api/office/office.file#closeasync-callback-) 方法获取整个演示文稿或文档。</span><span class="sxs-lookup"><span data-stu-id="e56f8-141">If your task pane add-in runs in PowerPoint or Word, you can use the [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-), [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-), and [File.closeAsync](/javascript/api/office/office.file#closeasync-callback-) methods to get an entire presentation or document.</span></span>

<span data-ttu-id="e56f8-142">当您调用`Document.getFileAsync`时，将在[File](/javascript/api/office/office.file)对象中获取文档的副本。</span><span class="sxs-lookup"><span data-stu-id="e56f8-142">When you call `Document.getFileAsync` you get a copy of the document in a [File](/javascript/api/office/office.file) object.</span></span> <span data-ttu-id="e56f8-143">`File`对象提供对以[切片](/javascript/api/office/office.slice)对象表示的 "块" 中的文档的访问。</span><span class="sxs-lookup"><span data-stu-id="e56f8-143">The `File` object provides access to the document in "chunks" represented as [Slice](/javascript/api/office/office.slice) objects.</span></span> <span data-ttu-id="e56f8-144">调用`getFileAsync`时，可以指定文件类型（文本或压缩的 OPEN Office XML 格式）以及切片的大小（最高为4mb）。</span><span class="sxs-lookup"><span data-stu-id="e56f8-144">When you call `getFileAsync`, you can specify the file type (text or compressed Open Office XML format), and size of the slices (up to 4MB).</span></span> <span data-ttu-id="e56f8-145">若要访问`File`对象的内容，您可以在切片`File.getSliceAsync`的[data](/javascript/api/office/office.slice#data)属性中调用返回原始数据。</span><span class="sxs-lookup"><span data-stu-id="e56f8-145">To access the contents of the `File` object, you then call `File.getSliceAsync` which returns the raw data in the [Slice.data](/javascript/api/office/office.slice#data) property.</span></span> <span data-ttu-id="e56f8-146">如果您指定了压缩格式，则获取作为字节数组的文件数据。</span><span class="sxs-lookup"><span data-stu-id="e56f8-146">If you specified compressed format, you will get the file data as a byte array.</span></span> <span data-ttu-id="e56f8-147">如果您在将文件传输给 Web 服务，则可以在提交前将压缩的原始数据转换为 base64 编码的字符串。</span><span class="sxs-lookup"><span data-stu-id="e56f8-147">If you are transmitting the file to a web service, you can transform the compressed raw data to a base64-encoded string before submission.</span></span> <span data-ttu-id="e56f8-148">最后，当您完成文件的切片的获取后，请使用`File.closeAsync`方法关闭该文档。</span><span class="sxs-lookup"><span data-stu-id="e56f8-148">Finally, when you are finished getting slices of the file, use the `File.closeAsync` method to close the document.</span></span>

<span data-ttu-id="e56f8-149">有关详细信息，请参阅如何[从 PowerPoint 或 Word 外接程序中获取整个文档](../word/get-the-whole-document-from-an-add-in-for-word.md)。</span><span class="sxs-lookup"><span data-stu-id="e56f8-149">For more details, see how to [get the whole document from an add-in for PowerPoint or Word](../word/get-the-whole-document-from-an-add-in-for-word.md).</span></span>


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a><span data-ttu-id="e56f8-150">读取和写入 Word 文档的自定义 XML 部件</span><span class="sxs-lookup"><span data-stu-id="e56f8-150">Reading and writing custom XML parts of a Word document</span></span>

<span data-ttu-id="e56f8-p110">通过使用 Open Office XML 文件格式和内容控件，您可以将自定义 XML 部件添加到 Word 文档，并将 XML 部件中的元素绑定到文档的内容控件。打开文档时，Word 读取并自动使用自定义 XML 部件中的数据填充绑定的内容控件。用户还可以将数据写入内容控件，且在用户保存文档时，控件中的数据也将保存到绑定的 XML 部件。适用于 Word 的任务窗格外接程序可以使用 [Document.customXmlParts](/javascript/api/office/office.document#customxmlparts) 属性、[CustomXmlParts](/javascript/api/office/office.customxmlparts)、[CustomXmlPart](/javascript/api/office/office.customxmlpart) 和 [CustomXmlNode](/javascript/api/office/office.customxmlnode) 对象来动态读取文档中的数据和将数据写入文档中。</span><span class="sxs-lookup"><span data-stu-id="e56f8-p110">Using the Open Office XML file format and content controls, you can add custom XML parts to a Word document and bind elements in the XML parts to content controls in that document. When you open the document, Word reads and automatically populates bound content controls with data from the custom XML parts. Users can also write data into the content controls, and when the user saves the document, the data in the controls will be saved to the bound XML parts. Task pane add-ins for Word, can use the [Document.customXmlParts](/javascript/api/office/office.document#customxmlparts) property,[CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart), and [CustomXmlNode](/javascript/api/office/office.customxmlnode) objects to read and write data dynamically to the document.</span></span>

<span data-ttu-id="e56f8-p111">自定义 XML 部件可能与命名空间相关联。若要从命名空间的自定义 XML 部件获取数据，请使用 [CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#getbynamespaceasync-ns--options--callback-) 方法。</span><span class="sxs-lookup"><span data-stu-id="e56f8-p111">Custom XML parts may be associated with namespaces. To get data from custom XML parts in a namespace, use the [CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#getbynamespaceasync-ns--options--callback-) method.</span></span>

<span data-ttu-id="e56f8-p112">您还可以使用 [CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) 方法通过其 GUID 访问自定义 XML 部件。在获取自定义 XML 部件后，使用 [CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#getxmlasync-options--callback-) 方法获取 XML 数据。</span><span class="sxs-lookup"><span data-stu-id="e56f8-p112">You can also use the [CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) method to access custom XML parts by their GUIDs. After getting a custom XML part, use the [CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#getxmlasync-options--callback-) method to get the XML data.</span></span>

<span data-ttu-id="e56f8-159">若要向文档中添加新的自定义 XML 部件， `Document.customXmlParts`请使用属性获取文档中的自定义 xml 部件，然后调用[CustomXmlParts](/javascript/api/office/office.customxmlparts#addasync-xml--options--callback-)方法。</span><span class="sxs-lookup"><span data-stu-id="e56f8-159">To add a new custom XML part to a document, use the `Document.customXmlParts` property to get the custom XML parts that are in the document, and call the [CustomXmlParts.addAsync](/javascript/api/office/office.customxmlparts#addasync-xml--options--callback-) method.</span></span>

<span data-ttu-id="e56f8-160">有关如何使用含有任务窗格外接程序的自定义 XML 部件的详细信息，请参阅[使用 Office Open XML 创建更好的 Word 外接程序](../word/create-better-add-ins-for-word-with-office-open-xml.md)。</span><span class="sxs-lookup"><span data-stu-id="e56f8-160">For detailed information about how to work with custom XML parts with a task pane add-in, see [Creating Better Add-ins for Word with Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).</span></span>


## <a name="persisting-add-in-settings"></a><span data-ttu-id="e56f8-161">保留加载项设置</span><span class="sxs-lookup"><span data-stu-id="e56f8-161">Persisting add-in settings</span></span>


<span data-ttu-id="e56f8-162">通常需要保存外接程序的自定义数据，例如用户的首选项或外接程序的状态，并在下一次打开外接程序时访问该数据。</span><span class="sxs-lookup"><span data-stu-id="e56f8-162">Often you need to save custom data for your add-in, such as a user's preferences or the add-in's state, and access that data the next time the add-in is opened.</span></span> <span data-ttu-id="e56f8-163">可以使用通用的 Web 编程技术保存该数据，例如浏览器 cookie 或 HTML 5 Web 存储。</span><span class="sxs-lookup"><span data-stu-id="e56f8-163">You can use common web programming techniques to save that data, such as browser cookies or HTML 5 web storage.</span></span> <span data-ttu-id="e56f8-164">或者，如果你的外接程序在 Excel、PowerPoint 或 Word 中运行，则可以使用 [设置](/javascript/api/office/office.settings) 对象的方法。</span><span class="sxs-lookup"><span data-stu-id="e56f8-164">Alternatively, if your add-in runs in Excel, PowerPoint, or Word, you can use the methods of the [Settings](/javascript/api/office/office.settings) object.</span></span> <span data-ttu-id="e56f8-165">使用该`Settings`对象创建的数据存储在电子表格、演示文稿或文档中，该外接程序插入并与之一起保存。</span><span class="sxs-lookup"><span data-stu-id="e56f8-165">Data created with the `Settings` object is stored in the spreadsheet, presentation, or document that the add-in was inserted into and saved with.</span></span> <span data-ttu-id="e56f8-166">此数据仅用于创建它的外接程序。</span><span class="sxs-lookup"><span data-stu-id="e56f8-166">This data is available to only the add-in that created it.</span></span>

<span data-ttu-id="e56f8-167">若要避免向存储文档的服务器往返，在运行时使用该`Settings`对象创建的数据在内存中进行管理。</span><span class="sxs-lookup"><span data-stu-id="e56f8-167">To avoid roundtrips to the server where the document is stored, data created with the `Settings` object is managed in memory at run time.</span></span> <span data-ttu-id="e56f8-168">之前保存的设置数据在初始化外接程序时加载到内存中，并在调用 [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) 方法时，仅将对数据的更改保存回文档。</span><span class="sxs-lookup"><span data-stu-id="e56f8-168">Previously saved settings data is loaded into memory when the add-in is initialized, and changes to that data are only saved back to the document when you call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method.</span></span> <span data-ttu-id="e56f8-169">在内部，将该数据作为名称/值对存储在序列化的 JSON 对象中。</span><span class="sxs-lookup"><span data-stu-id="e56f8-169">Internally, the data is stored in a serialized JSON object as name/value pairs.</span></span> <span data-ttu-id="e56f8-170">可以使用 [Settings](/javascript/api/office/office.settings#get-name-) 对象的 [get](/javascript/api/office/office.settings#set-name--value-)、[set](/javascript/api/office/office.settings#remove-name-) 和 **remove** 方法从数据的内存副本中读取、写入和删除项目。</span><span class="sxs-lookup"><span data-stu-id="e56f8-170">You use the [get](/javascript/api/office/office.settings#get-name-), [set](/javascript/api/office/office.settings#set-name--value-), and [remove](/javascript/api/office/office.settings#remove-name-) methods of the **Settings** object, to read, write, and delete items from the in-memory copy of the data.</span></span> <span data-ttu-id="e56f8-171">以下代码行显示如何创建名为 `themeColor` 的设置，并将它的值设置为“green”。</span><span class="sxs-lookup"><span data-stu-id="e56f8-171">The following line of code shows how to create a setting named `themeColor` and set its value to 'green'.</span></span>




```js
Office.context.document.settings.set('themeColor', 'green');
```

<span data-ttu-id="e56f8-172">由于使用`set`和`remove`方法创建或删除的设置数据针对的是内存中的数据副本，因此您必须调用`saveAsync`以将对设置数据所做的更改保存到加载项所使用的文档中。</span><span class="sxs-lookup"><span data-stu-id="e56f8-172">Because settings data created or deleted with the `set` and `remove` methods is acting on an in-memory copy of the data, you must call `saveAsync` to persist changes to settings data into the document your add-in is working with.</span></span>

<span data-ttu-id="e56f8-173">有关使用`Settings`对象的方法处理自定义数据的详细信息，请参阅[保留外接程序状态和设置](persisting-add-in-state-and-settings.md)。</span><span class="sxs-lookup"><span data-stu-id="e56f8-173">For more details about working with custom data using the methods of the `Settings` object, see [Persisting add-in state and settings](persisting-add-in-state-and-settings.md).</span></span>


## <a name="reading-properties-of-a-project-document"></a><span data-ttu-id="e56f8-174">读取项目文档的属性</span><span class="sxs-lookup"><span data-stu-id="e56f8-174">Reading properties of a project document</span></span>

<span data-ttu-id="e56f8-175">如果您的任务窗格外接程序在 Project 中运行，则它可以从活动项目的某些项目字段、资源和任务字段中读取数据。</span><span class="sxs-lookup"><span data-stu-id="e56f8-175">If your task pane add-in runs in Project, your add-in can read data from some of the project fields, resource, and task fields in the active project.</span></span> <span data-ttu-id="e56f8-176">若要执行此操作，请使用[ProjectDocument](/javascript/api/office/office.document)对象的方法和事件，该对象将`Document`扩展对象以提供其他特定于项目的功能。</span><span class="sxs-lookup"><span data-stu-id="e56f8-176">To do that, you use the methods and events of the [ProjectDocument](/javascript/api/office/office.document) object, which extends the `Document` object to provide additional Project-specific functionality.</span></span>

<span data-ttu-id="e56f8-177">有关读取 Project 数据的示例，请参阅[使用文本编辑器创建您第一个用于 Project 2013 的任务窗格外接程序](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)。</span><span class="sxs-lookup"><span data-stu-id="e56f8-177">For examples of reading Project data, see [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span></span>


## <a name="permissions-model-and-governance"></a><span data-ttu-id="e56f8-178">权限模型和管治</span><span class="sxs-lookup"><span data-stu-id="e56f8-178">Permissions model and governance</span></span>

<span data-ttu-id="e56f8-179">您的外接程序使用`Permissions`其清单中的元素，以请求从 OFFICE JavaScript API 访问所需功能级别的权限。</span><span class="sxs-lookup"><span data-stu-id="e56f8-179">Your add-in uses the `Permissions` element in its manifest to request permission to access the level of functionality it requires from the Office JavaScript API.</span></span> <span data-ttu-id="e56f8-180">例如，如果您的外接程序需要对文档的读/写访问权限，则其清单`ReadWriteDocument`必须在其`Permissions`元素中指定为文本值。</span><span class="sxs-lookup"><span data-stu-id="e56f8-180">For example, if your add-in requires read/write access to the document, its manifest must specify `ReadWriteDocument` as the text value in its `Permissions` element.</span></span> <span data-ttu-id="e56f8-181">因为权限的存在是为了保护用户的隐私和安全，因此最佳做法应当是，请求功能所需的最低级别的权限。</span><span class="sxs-lookup"><span data-stu-id="e56f8-181">Because permissions exist to protect a user's privacy and security, as a best practice you should request the minimum level of permissions it needs for its features.</span></span> <span data-ttu-id="e56f8-182">以下示例显示如何在任务窗格清单中请求 **ReadDocument** 权限。</span><span class="sxs-lookup"><span data-stu-id="e56f8-182">The following example shows how to request the **ReadDocument** permission in a task pane's manifest.</span></span>


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???
</OfficeApp>

```

<span data-ttu-id="e56f8-183">有关详细信息，请参阅在[外接程序中请求 API 使用的权限](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="e56f8-183">For more information, see [Requesting permissions for API use in add-ins](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).</span></span>


## <a name="see-also"></a><span data-ttu-id="e56f8-184">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e56f8-184">See also</span></span>

- [<span data-ttu-id="e56f8-185">Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="e56f8-185">Office JavaScript API</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="e56f8-186">Office 外接程序清单的架构参考</span><span class="sxs-lookup"><span data-stu-id="e56f8-186">Schema reference for Office Add-ins manifests</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="e56f8-187">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="e56f8-187">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
