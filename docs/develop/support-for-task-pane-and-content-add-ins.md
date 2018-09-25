---
title: 对 Office 2013 内容和任务窗格加载项的 Office JavaScript API 支持
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: cb4bb003966639fd5518fefcd3983ee9ca2fb101
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/25/2018
ms.locfileid: "25005009"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>对 Office 2013 内容和任务窗格加载项的 Office JavaScript API 支持


您可以使用 [Office JavaScript API](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js) 创建 Office 2013 主机应用程序的任务窗格或内容外接程序。已对内容和任务窗格外接程序支持的对象和方法进行如下分类：


1. **与其他 Office 外接程序共享的常见对象。** 这些对象包括 [Office](https://docs.microsoft.com/javascript/api/office?view=office-js)、[Context](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js) 和 [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js)。**Office** 对象是 Office JavaScript API 的根对象。**Context** 对象表示外接程序的运行时环境。**Office** 和 **Context** 都是适用于任何 Office 外接程序的基础对象。**AsyncResult** 对象表示异步操作的结果，比如返回到 **getSelectedDataAsync** 方法的数据，其中该方法可以读取用户在文档中选择的内容。
    
2.  **Document 对象。** 可用于内容和任务窗格外接程序的大部分 API 通过 [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) 对象的方法、属性和事件公开。 内容或任务窗格外接程序可以使用 [Office.context.document](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#document) 属性访问 **Document** 对象，并通过它访问 API 的关键成员以处理文档中的数据，例如 [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js) 和 [CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js) 对象，以及 [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-)、[setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) 和 [getFileAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfileasync-filetype--options--callback-) 方法。 **Document** 对象也提供了用于确定文档是只读还是处于编辑模式的 [mode](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#mode) 属性、获取当前文档 URL 的 [url](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#url) 属性，以及对 [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) 对象的访问。 **Document<** 对象还支持添加 [SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) 事件的事件处理程序，因此你可以检测用户何时更改文档中的选择内容。
    
   内容或任务窗格外接程序只能在加载 DOM 和运行时环境后访问 **Document** 对象，通常是在 [Office.initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) 事件的事件处理程序中加载。有关应用程序初始化时的事件流以及如何检查 DOM 和运行时是否成功加载的信息，请参阅[加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)。
    
3.  **使用特定的功能的对象。** 若要使用 API 的特定功能，请使用下面的对象和方法：
    
    - 创建或获取绑定的 [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js) 对象的方法，以及使用数据的 [Binding](https://docs.microsoft.com/javascript/api/office/office.binding?view=office-js) 对象的方法和属性。
    
    - [CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js)、[CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js) 和关联对象，用于创建和控制 Word 文档中的自定义 XML 部分。
    
    - 创建整个文档的副本，将它分解成多个块或“切片”，然后读取或传输这些切片中数据的 [File](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js) 和 [Slice](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js) 对象。
    
    - [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) 对象，用于保存自定义数据（如用户偏好设置）和加载项状态。
    

> [!IMPORTANT]
> 并不是所有能够承载内容和任务窗格加载项的 Office 应用程序都支持一些 API 成员。要确定支持哪些成员，请参阅以下任一资源：

若要概览各 Office 主机应用提供的 Office JavaScript API 支持，请参阅[了解适用于 Office 的 JavaScript API](understanding-the-javascript-api-for-office.md)。


## <a name="reading-and-writing-to-an-active-selection"></a>在活动的选择内容中读取和写入

您可以在文档、电子表格或演示文稿的用户当前选定内容中读取和写入。根据加载项的主机应用程序，您可以在 [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) 对象的 [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) 和 [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) 方法中指定要作为参数来读取或写入的数据结构类型。例如，您可以指定任何用于 Word 的数据类型（文本、HTML、表格数据或 Office Open XML）、用于 Excel 的文本和表格数据，以及用于 PowerPoint 和 Project 的文本。您还可以创建事件处理程序来检测对用户选择内容的更改。以下示例使用 **getSelectedDataAsync** 方法从作为文本的选择内容中获取数据。


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

有关详细信息和示例，请参阅[对文档或电子表格中的活动选择执行数据读取和写入操作](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a>绑定到文档或电子表格中的区域

您可以使用 **getSelectedDataAsync** 和 **setSelectedDataAsync** 方法在文档、电子表格或演示文稿中的用户*当前*选定内容中读取和写入。但是，如果您想在不要求用户选定内容的情况下，在运行您外接程序的各个会话中访问文档中的同一区域，您应首先绑定到该区域。您还可以订阅该绑定区域的数据和选定内容更改事件。

可以使用 [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromnameditemasync-itemname--bindingtype--options--callback-) 对象的 [addFromNamedItemAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfrompromptasync-bindingtype--options--callback-)、[addFromPromptAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-) 或 [addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js) 方法添加绑定。这些方法可以返回一个标识符，您可以用它访问绑定中的数据或者订阅数据更改或选择更改事件。

以下是使用 **Bindings.addFromSelectionAsync** 方法添加绑定到文档中当前选定文本的示例。



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

有关详细信息和示例，请参阅[绑定到文档或电子表格中的区域](bind-to-regions-in-a-document-or-spreadsheet.md)。


## <a name="getting-entire-documents"></a>获取整个文档

如果任务窗格外接程序在 PowerPoint 或 Word 中运行，您可以使用 [Document.getFileAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfileasync-filetype--options--callback-)、[File.getSliceAsync](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js#getsliceasync-sliceindex--callback-) 和 [File.closeAsync](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js#closeasync-callback-) 方法获取整个演示文稿或文档。

您在调用 **Document.getFileAsync** 时，获取了 [File](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js) 对象中的文档副本。**File** 对象提供对表示为 [Slice](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js) 对象的“块”中文档的访问。当调用 **getFileAsync** 时，您可以指定文件类型（文本或压缩的 Open Office XML 格式）和切片的大小（高达 4MB）。若要访问 **File** 对象的内容，您可以调用在 **Slice.data** 属性中返回原始数据的 [File.getSliceAsync](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js#data)。如果您指定了压缩格式，则获取作为字节数组的文件数据。如果您在将文件传输给 Web 服务，则可以在提交前将压缩的原始数据转换为 base64 编码的字符串。最后，在完成获取文件切片后，使用 **File.closeAsync** 方法关闭文档。

有关详细信息，请参阅如何[通过 PowerPoint 或 Word 加载项获取整个文档](../word/get-the-whole-document-from-an-add-in-for-word.md)。 


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a>读取和写入 Word 文档的自定义 XML 部件

通过使用 Open Office XML 文件格式和内容控件，您可以将自定义 XML 部件添加到 Word 文档，并将 XML 部件中的元素绑定到文档的内容控件。打开文档时，Word 读取并自动使用自定义 XML 部件中的数据填充绑定的内容控件。用户还可以将数据写入内容控件，且在用户保存文档时，控件中的数据也将保存到绑定的 XML 部件。适用于 Word 的任务窗格外接程序可以使用 [Document.customXmlParts](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js.customxmlparts) 属性、[CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js)、[CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js) 和 [CustomXmlNode](https://docs.microsoft.com/javascript/api/office/office.customxmlnode?view=office-js) 对象来动态读取文档中的数据和将数据写入文档中。

自定义 XML 部分可能与命名空间相关联。若要从命名空间中的自定义 XML 部分获取数据，请使用 [CustomXmlParts.getByNamespaceAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js#getbynamespaceasync-ns--options--callback-) 方法。

您还可以使用 [CustomXmlParts.getByIdAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js#getbyidasync-id--options--callback-) 方法通过其 GUID 访问自定义 XML 部件。在获取自定义 XML 部件后，使用 [CustomXmlPart.getXmlAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js#getxmlasync-options--callback-) 方法获取 XML 数据。

若要将新的自定义 XML 部件添加到文档，请使用 **Document.customXmlParts** 属性获取文档中的自定义 XML 部件，并调用 [CustomXmlParts.addAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js#addasync-xml--options--callback-) 方法。

有关如何使用含有任务窗格外接程序的自定义 XML 部件的详细信息，请参阅[使用 Office Open XML 创建更好的 Word 外接程序](../word/create-better-add-ins-for-word-with-office-open-xml.md)。


## <a name="persisting-add-in-settings"></a>保留加载项设置


通常需要保存外接程序的自定义数据，例如用户的首选项或外接程序的状态，并在下一次打开外接程序时访问该数据。可以使用通用的 Web 编程技术保存该数据，例如浏览器 cookie 或 HTML 5 Web 存储。或者，如果你的外接程序在 Excel、PowerPoint 或 Word 中运行，则可以使用 [设置](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) 对象的方法。使用**设置**对象创建的数据存储在电子表格、演示文档或植入和保存外接程序的文档中。此数据仅用于创建它的外接程序。

若要避免往返存储文档的服务器，运行时在内存中管理使用 **Settings** 对象创建的数据。 以前保存的设置数据在外接程序初始化时加载到内存中，对数据的更改仅在调用 [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#saveasync-options--callback-) 方法时保存回文档。 内部的数据以名称/值对的形式存储在序列化的 JSON 对象。 使用 **Settings** 对象的 [get](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#get-name-)、[set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#set-name--value-) 和 [remove](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#remove-name-) 方法读取、写入和删除数据在内存副本中的项目。 下面的代码行演示如何创建名为 `themeColor` 的设置并设置其值为 'green'。




```js
Office.context.document.settings.set('themeColor', 'green');
```

因为使用 **set** 和 **remove** 方法创建或删除的设置数据对数据的内存副本有影响，您必须调用 **saveAsync** 将对设置数据的更改保存到外接程序的工作文档。

若要详细了解如何使用 **Settings** 对象的方法处理自定义数据，请参阅[暂留加载项状态和设置](persisting-add-in-state-and-settings.md)。


## <a name="reading-properties-of-a-project-document"></a>读取项目文档的属性

如果您的任务窗格外接程序在 Project 中运行，则它可以从活动项目的某些项目字段、资源和任务字段中读取数据。为此，可以使用将 [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) 对象扩展为提供其他特定于 Project 功能的 **ProjectDocument** 对象的方法和事件。

有关读取 Project 数据的示例，请参阅[使用文本编辑器创建首个 Project 2013 任务窗格加载项](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)。


## <a name="permissions-model-and-governance"></a>权限模型和管治

您的外接程序使用其清单中的 **Permissions** 元素请求对 Office JavaScript API 中功能级别的访问权限。例如，如果您的外接程序请求对文档的读取/写入访问权限，它的清单必须将 `ReadWriteDocument` 指定为其 **Permissions** 元素中的文本值。因为权限的存在是为了保护用户的隐私和安全，因此最佳做法应当是，请求功能所需的最低级别的权限。以下示例显示如何在任务窗格清单中请求 **ReadDocument** 权限。


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???
</OfficeApp>

```

有关详细信息，请参阅[在内容和任务窗格加载项中请求获取 API 使用权限](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)。


## <a name="see-also"></a>另请参阅

- [Office JavaScript API](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)
- [Office 加载项清单的架构参考](../develop/add-in-manifests.md)
- [排查 Office 加载项中的用户错误](../testing/testing-and-troubleshooting.md)
    
