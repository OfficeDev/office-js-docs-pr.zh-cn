---
title: 对 Office 2013 内容和任务窗格加载项的 Office JavaScript API 支持
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 2aab577e3536ed11c8f2e9810f6f200bdf5d1768
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437589"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a><span data-ttu-id="972d2-102">对 Office 2013 内容和任务窗格加载项的 Office JavaScript API 支持</span><span class="sxs-lookup"><span data-stu-id="972d2-102">Office JavaScript API support for content and task pane add-ins in Office 2013</span></span>


<span data-ttu-id="972d2-p101">您可以使用 [Office JavaScript API](https://dev.office.com/reference/add-ins/javascript-api-for-office) 创建 Office 2013 主机应用程序的任务窗格或内容外接程序。已对内容和任务窗格外接程序支持的对象和方法进行如下分类：</span><span class="sxs-lookup"><span data-stu-id="972d2-p101">You can use the [Office JavaScript API](https://dev.office.com/reference/add-ins/javascript-api-for-office) to create task pane or content add-ins for Office 2013 host applications. The objects and methods that content and task pane add-ins support are categorized as follows:</span></span>


1. <span data-ttu-id="972d2-p102">**与其他 Office 外接程序共享的常见对象。** 这些对象包括 [Office](https://dev.office.com/reference/add-ins/shared/office)、[Context](https://dev.office.com/reference/add-ins/shared/office.context) 和 [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult)。**Office** 对象是 Office JavaScript API 的根对象。**Context** 对象表示外接程序的运行时环境。**Office** 和 **Context** 都是适用于任何 Office 外接程序的基础对象。**AsyncResult** 对象表示异步操作的结果，比如返回到 **getSelectedDataAsync** 方法的数据，其中该方法可以读取用户在文档中选择的内容。</span><span class="sxs-lookup"><span data-stu-id="972d2-p102">**Common objects shared with other Office Add-ins.** These objects include [Office](https://dev.office.com/reference/add-ins/shared/office), [Context](https://dev.office.com/reference/add-ins/shared/office.context), and [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult). The  **Office** object is the root object of the Office JavaScript API. The **Context** object represents the add-in's runtime environment. Both **Office** and **Context** are the fundamental objects for any Office Add-in. The **AsyncResult** object represents the results of an asynchronous operation, such as the data returned to the **getSelectedDataAsync** method, which reads what a user has selected in a document.</span></span>
    
2.  <span data-ttu-id="972d2-110">**Document 对象。**</span><span class="sxs-lookup"><span data-stu-id="972d2-110">**The Document object**</span></span> <span data-ttu-id="972d2-111">可用于内容和任务窗格外接程序的大部分 API 通过 [Document](https://dev.office.com/reference/add-ins/shared/document) 对象的方法、属性和事件公开。</span><span class="sxs-lookup"><span data-stu-id="972d2-111">The majority of the API available to content and task pane add-ins is exposed through the methods, properties, and events of the [Document](https://dev.office.com/reference/add-ins/shared/document) object. Using this subset of the API, your content or task pane add-in can perform the tasks described later in this topic.</span></span> <span data-ttu-id="972d2-112">内容或任务窗格外接程序可以使用 [Office.context.document](https://dev.office.com/reference/add-ins/shared/office.context.document) 属性访问 **Document** 对象，并通过它访问 API 的关键成员以处理文档中的数据，例如 [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings) 和 [CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts) 对象，以及 [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync)、[setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) 和 [getFileAsync](https://dev.office.com/reference/add-ins/shared/document.getfileasync) 方法。</span><span class="sxs-lookup"><span data-stu-id="972d2-112">A content or task pane add-in can use the [Office.context.document](https://dev.office.com/reference/add-ins/shared/office.context.document) property to access the **Document** object, and through it, can access the key members of the API for working with data in documents, such as the [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings) and [CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts) objects, and the [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync), [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync), and [getFileAsync](https://dev.office.com/reference/add-ins/shared/document.getfileasync) methods.</span></span> <span data-ttu-id="972d2-113">**Document** 对象也提供了用于确定文档是只读还是处于编辑模式的 [mode](https://dev.office.com/reference/add-ins/shared/document.mode) 属性、获取当前文档 URL 的 [url](https://dev.office.com/reference/add-ins/shared/document.url) 属性，以及对 [Settings](https://dev.office.com/reference/add-ins/shared/settings) 对象的访问。</span><span class="sxs-lookup"><span data-stu-id="972d2-113">The **Document** object also provides the [mode](https://dev.office.com/reference/add-ins/shared/document.mode) property for determining whether a document is read-only or in edit mode, the [url](https://dev.office.com/reference/add-ins/shared/document.url) property to get the URL of the current document, and access to the [Settings](https://dev.office.com/reference/add-ins/shared/settings) object.</span></span> <span data-ttu-id="972d2-114">**Document<** 对象还支持添加 [SelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) 事件的事件处理程序，因此你可以检测用户何时更改文档中的选择内容。</span><span class="sxs-lookup"><span data-stu-id="972d2-114">The **Document** object also supports adding event handlers for the [SelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) event, so you can detect when a user changes their selection in the document.</span></span>
    
   <span data-ttu-id="972d2-p104">内容或任务窗格外接程序只能在加载 DOM 和运行时环境后访问 **Document** 对象，通常是在 [Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) 事件的事件处理程序中加载。有关应用程序初始化时的事件流以及如何检查 DOM 和运行时是否成功加载的信息，请参阅[加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)。</span><span class="sxs-lookup"><span data-stu-id="972d2-p104">A content or task pane add-in can access the  **Document** object only after the DOM and runtime environment has been loaded, typically in the event handler for the [Office.initialize](https://dev.office.com/reference/add-ins/shared/office.initialize) event. For information about the flow of events when an add-in is initialized, and how to check that the DOM and runtime and loaded successfully, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>
    
3.  <span data-ttu-id="972d2-p105">**使用特定的功能的对象。** 若要使用 API 的特定功能，请使用下面的对象和方法：</span><span class="sxs-lookup"><span data-stu-id="972d2-p105">**Objects for working with specific features.** To work with specific features of the API, use the following objects and methods:</span></span>
    
    - <span data-ttu-id="972d2-119">创建或获取绑定的 [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings) 对象的方法，以及使用数据的 [Binding](https://dev.office.com/reference/add-ins/shared/binding) 对象的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="972d2-119">The methods of the [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings) object to create or get bindings, and the methods and properties of the [Binding](https://dev.office.com/reference/add-ins/shared/binding) object to work with data.</span></span>
    
    - <span data-ttu-id="972d2-120">[CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts)、[CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) 和关联对象，用于创建和控制 Word 文档中的自定义 XML 部分。</span><span class="sxs-lookup"><span data-stu-id="972d2-120">The [CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts), [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) and associated objects to create and manipulate custom XML parts in Word documents.</span></span>
    
    - <span data-ttu-id="972d2-121">创建整个文档的副本，将它分解成多个块或“切片”，然后读取或传输这些切片中数据的 [File](https://dev.office.com/reference/add-ins/shared/file) 和 [Slice](https://dev.office.com/reference/add-ins/shared/slice) 对象。</span><span class="sxs-lookup"><span data-stu-id="972d2-121">The [File](https://dev.office.com/reference/add-ins/shared/file) and [Slice](https://dev.office.com/reference/add-ins/shared/slice) objects to create a copy of the entire document, break it into chunks or "slices", and then read or transmit the data in those slices.</span></span>
    
    - <span data-ttu-id="972d2-122">[Settings](https://dev.office.com/reference/add-ins/shared/settings) 对象，用于保存自定义数据（如用户偏好设置）和加载项状态。</span><span class="sxs-lookup"><span data-stu-id="972d2-122">The [Settings](https://dev.office.com/reference/add-ins/shared/settings) object to save custom data, such as user preferences, and add-in state.</span></span>
    

> [!IMPORTANT]
> <span data-ttu-id="972d2-123">并不是所有能够承载内容和任务窗格加载项的 Office 应用程序都支持一些 API 成员。要确定支持哪些成员，请参阅以下任一资源：</span><span class="sxs-lookup"><span data-stu-id="972d2-123">Some of the API members aren't supported across all Office applications that can host content and task pane add-ins. To determine which members are supported, see any of the following:</span></span>

<span data-ttu-id="972d2-124">若要概览各 Office 主机应用提供的 Office JavaScript API 支持，请参阅[了解适用于 Office 的 JavaScript API](understanding-the-javascript-api-for-office.md)。</span><span class="sxs-lookup"><span data-stu-id="972d2-124">For a summary of Office JavaScript API support across Office host applications, see [Understanding the JavaScript API for Office](understanding-the-javascript-api-for-office.md).</span></span>


## <a name="reading-and-writing-to-an-active-selection"></a><span data-ttu-id="972d2-125">在活动的选择内容中读取和写入</span><span class="sxs-lookup"><span data-stu-id="972d2-125">Reading and writing to an active selection</span></span>

<span data-ttu-id="972d2-p106">您可以在文档、电子表格或演示文稿的用户当前选定内容中读取和写入。根据加载项的主机应用程序，您可以在 [Document](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) 对象的 [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) 和 [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document) 方法中指定要作为参数来读取或写入的数据结构类型。例如，您可以指定任何用于 Word 的数据类型（文本、HTML、表格数据或 Office Open XML）、用于 Excel 的文本和表格数据，以及用于 PowerPoint 和 Project 的文本。您还可以创建事件处理程序来检测对用户选择内容的更改。以下示例使用 **getSelectedDataAsync** 方法从作为文本的选择内容中获取数据。</span><span class="sxs-lookup"><span data-stu-id="972d2-p106">You can read or write to the user's current selection in a document, spreadsheet, or presentation. Depending on the host application for your add-in, you can specify the type of data structure to read or write as a parameter in the [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) and [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) methods of the [Document](https://dev.office.com/reference/add-ins/shared/document) object. For example, you can specify any type of data (text, HTML, tabular data, or Office Open XML) for Word, text and tabular data for Excel, and text for PowerPoint and Project. You can also create event handlers to detect changes to the user's selection. The following example gets data from the selection as text using the **getSelectedDataAsync** method.</span></span>


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

<span data-ttu-id="972d2-131">有关详细信息和示例，请参阅[对文档或电子表格中的活动选择执行数据读取和写入操作](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。</span><span class="sxs-lookup"><span data-stu-id="972d2-131">For more details and examples, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).</span></span>


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a><span data-ttu-id="972d2-132">绑定到文档或电子表格中的区域</span><span class="sxs-lookup"><span data-stu-id="972d2-132">Binding to a region in a document or spreadsheet</span></span>

<span data-ttu-id="972d2-p107">您可以使用 **getSelectedDataAsync** 和 **setSelectedDataAsync** 方法在文档、电子表格或演示文稿中的用户*当前*选定内容中读取和写入。但是，如果您想在不要求用户选定内容的情况下，在运行您外接程序的各个会话中访问文档中的同一区域，您应首先绑定到该区域。您还可以订阅该绑定区域的数据和选定内容更改事件。</span><span class="sxs-lookup"><span data-stu-id="972d2-p107">You can use the  **getSelectedDataAsync** and **setSelectedDataAsync** methods to read or write to the user's *current* selection in a document, spreadsheet, or presentation. However, if you would like to access the same region in a document across sessions of running your add-in without requiring the user to make a selection, you should first bind to that region. You can also subscribe to data and selection change events for that bound region.</span></span>

<span data-ttu-id="972d2-p108">可以使用 [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync) 对象的 [addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync)、[addFromPromptAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync) 或 [addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.bindings) 方法添加绑定。这些方法可以返回一个标识符，您可以用它访问绑定中的数据或者订阅数据更改或选择更改事件。</span><span class="sxs-lookup"><span data-stu-id="972d2-p108">You can add a binding by using [addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync), [addFromPromptAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync), or [addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync) methods of the [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings) object. These methods return an identifier that you can use to access data in the binding, or to subscribe to its data change or selection change events.</span></span>

<span data-ttu-id="972d2-138">以下是使用 **Bindings.addFromSelectionAsync** 方法添加绑定到文档中当前选定文本的示例。</span><span class="sxs-lookup"><span data-stu-id="972d2-138">The following is an example that adds a binding to the currently selected text in a document, by using the  **Bindings.addFromSelectionAsync** method.</span></span>



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

<span data-ttu-id="972d2-139">有关详细信息和示例，请参阅[绑定到文档或电子表格中的区域](bind-to-regions-in-a-document-or-spreadsheet.md)。</span><span class="sxs-lookup"><span data-stu-id="972d2-139">For more details and examples, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>


## <a name="getting-entire-documents"></a><span data-ttu-id="972d2-140">获取整个文档</span><span class="sxs-lookup"><span data-stu-id="972d2-140">Getting entire documents</span></span>

<span data-ttu-id="972d2-141">如果任务窗格外接程序在 PowerPoint 或 Word 中运行，您可以使用 [Document.getFileAsync](https://dev.office.com/reference/add-ins/shared/document.getfileasync)、[File.getSliceAsync](https://dev.office.com/reference/add-ins/shared/file.getsliceasync) 和 [File.closeAsync](https://dev.office.com/reference/add-ins/shared/file.closeasync) 方法获取整个演示文稿或文档。</span><span class="sxs-lookup"><span data-stu-id="972d2-141">If your task pane add-in runs in PowerPoint or Word, you can use the [Document.getFileAsync](https://dev.office.com/reference/add-ins/shared/document.getfileasync), [File.getSliceAsync](https://dev.office.com/reference/add-ins/shared/file.getsliceasync), and [File.closeAsync](https://dev.office.com/reference/add-ins/shared/file.closeasync) methods to get an entire presentation or document.</span></span>

<span data-ttu-id="972d2-p109">您在调用 **Document.getFileAsync** 时，获取了 [File](https://dev.office.com/reference/add-ins/shared/file) 对象中的文档副本。**File** 对象提供对表示为 [Slice](https://dev.office.com/reference/add-ins/shared/document) 对象的“块”中文档的访问。当调用 **getFileAsync** 时，您可以指定文件类型（文本或压缩的 Open Office XML 格式）和切片的大小（高达 4MB）。若要访问 **File** 对象的内容，您可以调用在 **Slice.data** 属性中返回原始数据的 [File.getSliceAsync](https://dev.office.com/reference/add-ins/shared/slice.data)。如果您指定了压缩格式，则获取作为字节数组的文件数据。如果您在将文件传输给 Web 服务，则可以在提交前将压缩的原始数据转换为 base64 编码的字符串。最后，在完成获取文件切片后，使用 **File.closeAsync** 方法关闭文档。</span><span class="sxs-lookup"><span data-stu-id="972d2-p109">When you call  **Document.getFileAsync**, you get a copy of the document in a [File](https://dev.office.com/reference/add-ins/shared/file) object. The **File** object provides access to the document in "chunks" represented as [Slice](https://dev.office.com/reference/add-ins/shared/document) objects. When you call **getFileAsync**, you can specify the file type (text or compressed Open Office XML format), and size of the slices (up to 4MB). To access the contents of the  **File** object, you then call **File.getSliceAsync** which returns the raw data in the [Slice.data](https://dev.office.com/reference/add-ins/shared/slice.data) property. If you specified compressed format, you will get the file data as a byte array. If you are transmitting the file to a web service, you can transform the compressed raw data to a base64-encoded string before submission. Finally, when you are finished getting slices of the file, use the **File.closeAsync** method to close the document.</span></span>

<span data-ttu-id="972d2-149">有关详细信息，请参阅如何[通过 PowerPoint 或 Word 加载项获取整个文档](../word/get-the-whole-document-from-an-add-in-for-word.md)。</span><span class="sxs-lookup"><span data-stu-id="972d2-149">For more details, see how to [get the whole document from an add-in for PowerPoint or Word](../word/get-the-whole-document-from-an-add-in-for-word.md).</span></span> 


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a><span data-ttu-id="972d2-150">读取和写入 Word 文档的自定义 XML 部件</span><span class="sxs-lookup"><span data-stu-id="972d2-150">Reading and writing custom XML parts of a Word document</span></span>

<span data-ttu-id="972d2-p110">通过使用 Open Office XML 文件格式和内容控件，您可以将自定义 XML 部件添加到 Word 文档，并将 XML 部件中的元素绑定到文档的内容控件。打开文档时，Word 读取并自动使用自定义 XML 部件中的数据填充绑定的内容控件。用户还可以将数据写入内容控件，且在用户保存文档时，控件中的数据也将保存到绑定的 XML 部件。适用于 Word 的任务窗格外接程序可以使用 [Document.customXmlParts](https://dev.office.com/reference/add-ins/shared/document.customxmlparts) 属性、[CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts)、[CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) 和 [CustomXmlNode](https://dev.office.com/reference/add-ins/shared/customxmlnode.customxmlnode) 对象来动态读取文档中的数据和将数据写入文档中。</span><span class="sxs-lookup"><span data-stu-id="972d2-p110">Using the Open Office XML file format and content controls, you can add custom XML parts to a Word document and bind elements in the XML parts to content controls in that document. When you open the document, Word reads and automatically populates bound content controls with data from the custom XML parts. Users can also write data into the content controls, and when the user saves the document, the data in the controls will be saved to the bound XML parts. Task pane add-ins for Word, can use the [Document.customXmlParts](https://dev.office.com/reference/add-ins/shared/document.customxmlparts) property,[CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts), [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart), and [CustomXmlNode](https://dev.office.com/reference/add-ins/shared/customxmlnode.customxmlnode) objects to read and write data dynamically to the document.</span></span>

<span data-ttu-id="972d2-p111">自定义 XML 部分可能与命名空间相关联。若要从命名空间中的自定义 XML 部分获取数据，请使用 [CustomXmlParts.getByNamespaceAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbynamespaceasync) 方法。</span><span class="sxs-lookup"><span data-stu-id="972d2-p111">Custom XML parts may be associated with namespaces. To get data from custom XML parts in a namespace, use the [CustomXmlParts.getByNamespaceAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbynamespaceasync) method.</span></span>

<span data-ttu-id="972d2-p112">您还可以使用 [CustomXmlParts.getByIdAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbyidasync) 方法通过其 GUID 访问自定义 XML 部件。在获取自定义 XML 部件后，使用 [CustomXmlPart.getXmlAsync](https://dev.office.com/reference/add-ins/shared/customxmlpart.getxmlasync) 方法获取 XML 数据。</span><span class="sxs-lookup"><span data-stu-id="972d2-p112">You can also use the [CustomXmlParts.getByIdAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.getbyidasync) method to access custom XML parts by their GUIDs. After getting a custom XML part, use the [CustomXmlPart.getXmlAsync](https://dev.office.com/reference/add-ins/shared/customxmlpart.getxmlasync) method to get the XML data.</span></span>

<span data-ttu-id="972d2-159">若要将新的自定义 XML 部件添加到文档，请使用 **Document.customXmlParts** 属性获取文档中的自定义 XML 部件，并调用 [CustomXmlParts.addAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.addasync) 方法。</span><span class="sxs-lookup"><span data-stu-id="972d2-159">To add a new custom XML part to a document, use the  **Document.customXmlParts** property to get the custom XML parts that are in the document, and call the [CustomXmlParts.addAsync](https://dev.office.com/reference/add-ins/shared/customxmlparts.addasync) method.</span></span>

<span data-ttu-id="972d2-160">有关如何使用含有任务窗格外接程序的自定义 XML 部件的详细信息，请参阅[使用 Office Open XML 创建更好的 Word 外接程序](../word/create-better-add-ins-for-word-with-office-open-xml.md)。</span><span class="sxs-lookup"><span data-stu-id="972d2-160">For detailed information about how to work with custom XML parts with a task pane add-in, see [Creating Better Add-ins for Word with Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).</span></span>


## <a name="persisting-add-in-settings"></a><span data-ttu-id="972d2-161">保留加载项设置</span><span class="sxs-lookup"><span data-stu-id="972d2-161">Persisting add-in settings</span></span>


<span data-ttu-id="972d2-p113">通常需要保存外接程序的自定义数据，例如用户的首选项或外接程序的状态，并在下一次打开外接程序时访问该数据。可以使用通用的 Web 编程技术保存该数据，例如浏览器 cookie 或 HTML 5 Web 存储。或者，如果你的外接程序在 Excel、PowerPoint 或 Word 中运行，则可以使用 [设置](https://dev.office.com/reference/add-ins/shared/settings) 对象的方法。使用**设置**对象创建的数据存储在电子表格、演示文档或植入和保存外接程序的文档中。此数据仅用于创建它的外接程序。</span><span class="sxs-lookup"><span data-stu-id="972d2-p113">Often you need to save custom data for your add-in, such as a user's preferences or the add-in's state, and access that data the next time the add-in is opened. You can use common web programming techniques to save that data, such as browser cookies or HTML 5 web storage. Alternatively, if your add-in runs in Excel, PowerPoint, or Word, you can use the methods of the [Settings](https://dev.office.com/reference/add-ins/shared/settings) object. Data created with the **Settings** object is stored in the spreadsheet, presentation, or document that the add-in was inserted into and saved with. This data is available to only the add-in that created it.</span></span>

<span data-ttu-id="972d2-p114">为了避免往返于存储文档的服务器，使用 **Settings** 对象创建的数据运行时在内存中进行管理。之前保存的设置数据在初始化外接程序时加载到内存中，并在调用 [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) 方法时，仅将对数据的更改保存回文档。在内部，将该数据作为名称/值对存储在序列化的 JSON 对象中。可以使用 [Settings](https://dev.office.com/reference/add-ins/shared/settings.get) 对象的 [get](https://dev.office.com/reference/add-ins/shared/settings.set)、[set](https://dev.office.com/reference/add-ins/shared/settings.removehandlerasync) 和 **remove** 方法从数据的内存副本中读取、写入和删除项目。以下代码行显示如何创建名为 `themeColor` 的设置，并将它的值设置为“green”。</span><span class="sxs-lookup"><span data-stu-id="972d2-p114">To avoid roundtrips to the server where the document is stored, data created with the  **Settings** object is managed in memory at run time. Previously saved settings data is loaded into memory when the add-in is initialized, and changes to that data are only saved back to the document when you call the [Settings.saveAsync](https://dev.office.com/reference/add-ins/shared/settings.saveasync) method. Internally, the data is stored in a serialized JSON object as name/value pairs. You use the [get](https://dev.office.com/reference/add-ins/shared/settings.get), [set](https://dev.office.com/reference/add-ins/shared/settings.set), and [remove](https://dev.office.com/reference/add-ins/shared/settings.removehandlerasync) methods of the **Settings** object, to read, write, and delete items from the in-memory copy of the data. The following line of code shows how to create a setting named `themeColor` and set its value to 'green'.</span></span>




```js
Office.context.document.settings.set('themeColor', 'green');
```

<span data-ttu-id="972d2-172">因为使用 **set** 和 **remove** 方法创建或删除的设置数据对数据的内存副本有影响，您必须调用 **saveAsync** 将对设置数据的更改保存到外接程序的工作文档。</span><span class="sxs-lookup"><span data-stu-id="972d2-172">Because settings data created or deleted with the  **set** and **remove** methods is acting on an in-memory copy of the data, you must call **saveAsync** to persist changes to settings data into the document your add-in is working with.</span></span>

<span data-ttu-id="972d2-173">若要详细了解如何使用 **Settings** 对象的方法处理自定义数据，请参阅[暂留加载项状态和设置](persisting-add-in-state-and-settings.md)。</span><span class="sxs-lookup"><span data-stu-id="972d2-173">For more details about working with custom data using the methods of the  **Settings** object, see [Persisting add-in state and settings](persisting-add-in-state-and-settings.md).</span></span>


## <a name="reading-properties-of-a-project-document"></a><span data-ttu-id="972d2-174">读取项目文档的属性</span><span class="sxs-lookup"><span data-stu-id="972d2-174">Reading properties of a project document</span></span>

<span data-ttu-id="972d2-p115">如果您的任务窗格外接程序在 Project 中运行，则它可以从活动项目的某些项目字段、资源和任务字段中读取数据。为此，可以使用将 [Document](https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument) 对象扩展为提供其他特定于 Project 功能的 **ProjectDocument** 对象的方法和事件。</span><span class="sxs-lookup"><span data-stu-id="972d2-p115">If your task pane add-in runs in Project, your add-in can read data from some of the project fields, resource, and task fields in the active project. To do that, you use the methods and events of the [ProjectDocument](https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument) object, which extends the **Document** object to provide additional Project-specific functionality.</span></span>

<span data-ttu-id="972d2-177">有关读取 Project 数据的示例，请参阅[使用文本编辑器创建首个 Project 2013 任务窗格加载项](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)。</span><span class="sxs-lookup"><span data-stu-id="972d2-177">For examples of reading Project data, see [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span></span>


## <a name="permissions-model-and-governance"></a><span data-ttu-id="972d2-178">权限模型和管治</span><span class="sxs-lookup"><span data-stu-id="972d2-178">Permissions model and governance</span></span>

<span data-ttu-id="972d2-p116">您的外接程序使用其清单中的 **Permissions** 元素请求对 Office JavaScript API 中功能级别的访问权限。例如，如果您的外接程序请求对文档的读取/写入访问权限，它的清单必须将 `ReadWriteDocument` 指定为其 **Permissions** 元素中的文本值。因为权限的存在是为了保护用户的隐私和安全，因此最佳做法应当是，请求功能所需的最低级别的权限。以下示例显示如何在任务窗格清单中请求 **ReadDocument** 权限。</span><span class="sxs-lookup"><span data-stu-id="972d2-p116">Your add-in uses the  **Permissions** element in its manifest to request permission to access the level of functionality it requires from the Office JavaScript API. For example, if your add-in requires read/write access to the document, its manifest must specify `ReadWriteDocument` as the text value in its **Permissions** element. Because permissions exist to protect a user's privacy and security, as a best practice you should request the minimum level of permissions it needs for its features. The following example shows how to request the **ReadDocument** permission in a task pane's manifest.</span></span>


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

<span data-ttu-id="972d2-183">有关详细信息，请参阅[在内容和任务窗格加载项中请求获取 API 使用权限](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="972d2-183">For more information, see [Requesting permissions for API use in content and task pane add-ins](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).</span></span>


## <a name="see-also"></a><span data-ttu-id="972d2-184">另请参阅</span><span class="sxs-lookup"><span data-stu-id="972d2-184">See also</span></span>

- [<span data-ttu-id="972d2-185">Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="972d2-185">Office JavaScript API</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)
- [<span data-ttu-id="972d2-186">Office 加载项清单的架构参考</span><span class="sxs-lookup"><span data-stu-id="972d2-186">Schema reference for Office Add-ins manifests</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="972d2-187">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="972d2-187">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
    
