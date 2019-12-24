---
title: 对 Office 2013 内容和任务窗格加载项的 Office JavaScript API 支持
description: ''
ms.date: 09/26/2019
localization_priority: Normal
ms.openlocfilehash: b9a6d1f81cb86cb84c0f825394d14e8559017d11
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851381"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>对 Office 2013 中内容和任务窗格外接程序的 Office JavaScript API 支持


您可以使用 [Office JavaScript API](/office/dev/add-ins/reference/javascript-api-for-office) 创建 Office 2013 主机应用程序的任务窗格或内容外接程序。已对内容和任务窗格外接程序支持的对象和方法进行如下分类：


1. **与其他 Office 外接程序共享的常见对象。** 这些对象包括 [Office](/javascript/api/office)、[Context](/javascript/api/office/office.context) 和 [AsyncResult](/javascript/api/office/office.asyncresult)。**Office** 对象是 Office JavaScript API 的根对象。**Context** 对象表示外接程序的运行时环境。**Office** 和 **Context** 都是适用于任何 Office 外接程序的基础对象。**AsyncResult** 对象表示异步操作的结果，比如返回到 **getSelectedDataAsync** 方法的数据，其中该方法可以读取用户在文档中选择的内容。

2. **Document 对象。** 可通过 [Document](/javascript/api/office/office.document) 对象的方法、属性和事件公开大多数可用于内容和任务窗格加载项的 API。 内容或任务窗格加载项可以使用[CustomXmlParts 属性访问](/javascript/api/office/office.context#document) **document**对象，通过它，可以访问用于处理文档中的数据（如[绑定](/javascript/api/office/office.bindings)和[](/javascript/api/office/office.customxmlparts)对象）的 API 的关键成员，以及[getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-)、 [document.setselecteddataasync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-)和[document.getfileasync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)方法。 **Document** 对象还提供用于确定文档是只读模式还是编辑模式的 [mode](/javascript/api/office/office.document#mode) 属性，[url](/javascript/api/office/office.document#url) 属性可以获取当前文档的 URL，并访问 [Settings](/javascript/api/office/office.settings) 对象。 **Document**对象还支持添加[SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs)事件的事件处理程序，以便您可以在用户更改文档中的选定内容时进行检测。

   内容或任务窗格外接程序只能在加载 DOM 和运行时环境后访问 **Document** 对象，通常是在 [Office.initialize](/javascript/api/office) 事件的事件处理程序中加载。有关应用程序初始化时的事件流以及如何检查 DOM 和运行时是否成功加载的信息，请参阅[加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)。

3. **使用特定的功能的对象。** 若要使用 API 的特定功能，请使用下面的对象和方法：

    - 创建或获取绑定的 [Bindings](/javascript/api/office/office.bindings) 对象的方法，以及使用数据的 [Binding](/javascript/api/office/office.binding) 对象的方法和属性。

    - 创建和操控 Word 文档中自定义的 XML 部件的 [CustomXmlParts](/javascript/api/office/office.customxmlparts)、[CustomXmlPart](/javascript/api/office/office.customxmlpart) 和关联的对象。

    - 创建整个文档的副本，将它分解成多个块或“切片”，然后读取或传输这些切片中数据的 [File](/javascript/api/office/office.file) 和 [Slice](/javascript/api/office/office.slice) 对象。

    - [Settings](/javascript/api/office/office.settings) 对象，用于保存自定义数据（如用户偏好设置）和加载项状态。


> [!IMPORTANT]
> 并不是所有能够承载内容和任务窗格加载项的 Office 应用程序都支持一些 API 成员。要确定支持哪些成员，请参阅以下任一资源：

有关对各 Office 主机应用程序的 Office JavaScript API 支持的摘要信息，请参阅[了解适用于 Office 的 JavaScript API](understanding-the-javascript-api-for-office.md)。


## <a name="reading-and-writing-to-an-active-selection"></a>在活动的选择内容中读取和写入

您可以在文档、电子表格或演示文稿的用户当前选定内容中读取和写入。根据加载项的主机应用程序，您可以在 [Document](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) 对象的 [getSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 和 [setSelectedDataAsync](/javascript/api/office/office.document) 方法中指定要作为参数来读取或写入的数据结构类型。例如，您可以指定任何用于 Word 的数据类型（文本、HTML、表格数据或 Office Open XML）、用于 Excel 的文本和表格数据，以及用于 PowerPoint 和 Project 的文本。您还可以创建事件处理程序来检测对用户选择内容的更改。以下示例使用 **getSelectedDataAsync** 方法从作为文本的选择内容中获取数据。


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

有关详细信息和示例，请参阅[将数据读取和写入到文档或电子表格中的活动选择区](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a>绑定到文档或电子表格中的区域

您可以使用 **getSelectedDataAsync** 和 **setSelectedDataAsync** 方法在文档、电子表格或演示文稿中的用户*当前*选定内容中读取和写入。但是，如果您想在不要求用户选定内容的情况下，在运行您外接程序的各个会话中访问文档中的同一区域，您应首先绑定到该区域。您还可以订阅该绑定区域的数据和选定内容更改事件。

可以使用 [Bindings](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-) 对象的 [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-)、[addFromPromptAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) 或 [addFromSelectionAsync](/javascript/api/office/office.bindings) 方法添加绑定。这些方法可以返回一个标识符，您可以用它访问绑定中的数据或者订阅数据更改或选择更改事件。

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

如果任务窗格外接程序在 PowerPoint 或 Word 中运行，您可以使用 [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)、[File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-) 和 [File.closeAsync](/javascript/api/office/office.file#closeasync-callback-) 方法获取整个演示文稿或文档。

您在调用 **Document.getFileAsync** 时，获取了 [File](/javascript/api/office/office.file) 对象中的文档副本。**File** 对象提供对表示为 [Slice](/javascript/api/office/office.slice) 对象的“块”中文档的访问。当调用 **getFileAsync** 时，您可以指定文件类型（文本或压缩的 Open Office XML 格式）和切片的大小（高达 4MB）。若要访问 **File** 对象的内容，您可以调用在 **Slice.data** 属性中返回原始数据的 [File.getSliceAsync](/javascript/api/office/office.slice#data)。如果您指定了压缩格式，则获取作为字节数组的文件数据。如果您在将文件传输给 Web 服务，则可以在提交前将压缩的原始数据转换为 base64 编码的字符串。最后，在完成获取文件切片后，使用 **File.closeAsync** 方法关闭文档。

有关详细信息，请参阅如何[从 PowerPoint 或 Word 外接程序中获取整个文档](../word/get-the-whole-document-from-an-add-in-for-word.md)。


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a>读取和写入 Word 文档的自定义 XML 部件

通过使用 Open Office XML 文件格式和内容控件，您可以将自定义 XML 部件添加到 Word 文档，并将 XML 部件中的元素绑定到文档的内容控件。打开文档时，Word 读取并自动使用自定义 XML 部件中的数据填充绑定的内容控件。用户还可以将数据写入内容控件，且在用户保存文档时，控件中的数据也将保存到绑定的 XML 部件。适用于 Word 的任务窗格外接程序可以使用 [Document.customXmlParts](/javascript/api/office/office.document#customxmlparts) 属性、[CustomXmlParts](/javascript/api/office/office.customxmlparts)、[CustomXmlPart](/javascript/api/office/office.customxmlpart) 和 [CustomXmlNode](/javascript/api/office/office.customxmlnode) 对象来动态读取文档中的数据和将数据写入文档中。

自定义 XML 部件可能与命名空间相关联。若要从命名空间的自定义 XML 部件获取数据，请使用 [CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#getbynamespaceasync-ns--options--callback-) 方法。

您还可以使用 [CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) 方法通过其 GUID 访问自定义 XML 部件。在获取自定义 XML 部件后，使用 [CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#getxmlasync-options--callback-) 方法获取 XML 数据。

若要将新的自定义 XML 部件添加到文档，请使用 **Document.customXmlParts** 属性获取文档中的自定义 XML 部件，并调用 [CustomXmlParts.addAsync](/javascript/api/office/office.customxmlparts#addasync-xml--options--callback-) 方法。

有关如何使用含有任务窗格外接程序的自定义 XML 部件的详细信息，请参阅[使用 Office Open XML 创建更好的 Word 外接程序](../word/create-better-add-ins-for-word-with-office-open-xml.md)。


## <a name="persisting-add-in-settings"></a>保留外接程序设置


通常需要保存外接程序的自定义数据，例如用户的首选项或外接程序的状态，并在下一次打开外接程序时访问该数据。 可以使用通用的 Web 编程技术保存该数据，例如浏览器 cookie 或 HTML 5 Web 存储。 或者，如果你的外接程序在 Excel、PowerPoint 或 Word 中运行，则可以使用 [设置](/javascript/api/office/office.settings) 对象的方法。 使用**设置**对象创建的数据存储在电子表格、演示文档或植入和保存外接程序的文档中。 此数据仅用于创建它的外接程序。

若要避免对存储文档的服务器的往返，使用**Settings**对象创建的数据在运行时在内存中进行管理。 之前保存的设置数据在初始化外接程序时加载到内存中，并在调用 [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) 方法时，仅将对数据的更改保存回文档。 在内部，将该数据作为名称/值对存储在序列化的 JSON 对象中。 可以使用 [Settings](/javascript/api/office/office.settings#get-name-) 对象的 [get](/javascript/api/office/office.settings#set-name--value-)、[set](/javascript/api/office/office.settings#remove-name-) 和 **remove** 方法从数据的内存副本中读取、写入和删除项目。 以下代码行显示如何创建名为 `themeColor` 的设置，并将它的值设置为“green”。




```js
Office.context.document.settings.set('themeColor', 'green');
```

因为使用 **set** 和 **remove** 方法创建或删除的设置数据对数据的内存副本有影响，您必须调用 **saveAsync** 将对设置数据的更改保存到外接程序的工作文档。

有关通过 **Settings** 对象的方法使用自定义数据的详细信息，请参阅[保留外接程序的状态和设置](persisting-add-in-state-and-settings.md)。


## <a name="reading-properties-of-a-project-document"></a>读取项目文档的属性

如果您的任务窗格外接程序在 Project 中运行，则它可以从活动项目的某些项目字段、资源和任务字段中读取数据。为此，可以使用将 [Document](/javascript/api/office/office.document) 对象扩展为提供其他特定于 Project 功能的 **ProjectDocument** 对象的方法和事件。

有关读取 Project 数据的示例，请参阅[使用文本编辑器创建您第一个用于 Project 2013 的任务窗格外接程序](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)。


## <a name="permissions-model-and-governance"></a>权限模型和管治

您的外接程序使用其清单中的 **Permissions** 元素请求对 Office JavaScript API 中功能级别的访问权限。例如，如果您的外接程序请求对文档的读取/写入访问权限，它的清单必须将 `ReadWriteDocument` 指定为其 **Permissions** 元素中的文本值。因为权限的存在是为了保护用户的隐私和安全，因此最佳做法应当是，请求功能所需的最低级别的权限。以下示例显示如何在任务窗格清单中请求 **ReadDocument** 权限。


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

有关详细信息，请参阅在[外接程序中请求 API 使用的权限](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)。


## <a name="see-also"></a>另请参阅

- [Office JavaScript API](/office/dev/add-ins/reference/javascript-api-for-office)
- [Office 外接程序清单的架构参考](../develop/add-in-manifests.md)
- [排查 Office 加载项中的用户错误](../testing/testing-and-troubleshooting.md)
