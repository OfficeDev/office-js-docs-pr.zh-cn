---
title: 对 Office 2013 内容和任务窗格加载项的 Office JavaScript API 支持
description: 使用 Office JavaScript API 在 Office 2013 中创建任务窗格。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: a6072538fe7328a71767394adf67398ebe4f0911
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422864"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>对 Office 2013 内容和任务窗格加载项的 Office JavaScript API 支持

[!include[information about the common API](../includes/alert-common-api-info.md)]

可以使用 [Office JavaScript API](../reference/javascript-api-for-office.md) 为 Office 2013 客户端应用程序创建任务窗格或内容加载项。 已对内容和任务窗格外接程序支持的对象和方法进行如下分类：

1. **与其他 Office 加载项共享的常见对象。** 这些对象包括 [Office](/javascript/api/office)、 [Context](/javascript/api/office/office.context) 和 [AsyncResult](/javascript/api/office/office.asyncresult)。 该 `Office` 对象是 Office JavaScript API 的根对象。 该 `Context` 对象表示加载项的运行时环境。 `Context`这两者都是`Office`任何 Office 加载项的基本对象。 该 `AsyncResult` 对象表示异步操作的结果，例如返回到 `getSelectedDataAsync` 方法的数据，该方法读取用户在文档中选择的内容。

2. **Document 对象。** 可通过 [Document](/javascript/api/office/office.document) 对象的方法、属性和事件公开大多数可用于内容和任务窗格加载项的 API。 内容或任务窗格外接程序可以使用 [Office.context.document](/javascript/api/office/office.context#office-office-context-document-member) 属性访问 **Document** 对象，并通过它访问 API 的关键成员，以便处理文档中的数据，例如 [绑定](/javascript/api/office/office.bindings) 和 [CustomXmlParts](/javascript/api/office/office.customxmlparts) 对象，以及 [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1))、 [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) 和 [getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) 方法。 该 `Document` 对象还提供 [模式](/javascript/api/office/office.document#office-office-document-mode-member) 属性，用于确定文档是只读还是处于编辑模式， [URL](/javascript/api/office/office.document#office-office-document-url-member) 属性用于获取当前文档的 URL，以及对 [Settings](/javascript/api/office/office.settings) 对象的访问权限。 该 `Document` 对象还支持为 [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) 事件添加事件处理程序，以便可以检测用户何时更改文档中的选择。

   内容或任务窗格外接程序只能在加载 DOM 和运行时环境后访问 `Document` 对象，通常在 [Office.initialize](/javascript/api/office) 事件的事件处理程序中。 有关应用程序初始化时的事件流以及如何检查 DOM 和运行时是否成功加载的信息，请参阅[加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)。

3. **使用特定的功能的对象。** 若要使用 API 的特定功能，请使用以下对象和方法。

    - 创建或获取绑定的 [Bindings](/javascript/api/office/office.bindings) 对象的方法，以及使用数据的 [Binding](/javascript/api/office/office.binding) 对象的方法和属性。

    - 创建和操控 Word 文档中自定义的 XML 部件的 [CustomXmlParts](/javascript/api/office/office.customxmlparts)、[CustomXmlPart](/javascript/api/office/office.customxmlpart) 和关联的对象。

    - 创建整个文档的副本，将它分解成多个块或“切片”，然后读取或传输这些切片中数据的 [File](/javascript/api/office/office.file) 和 [Slice](/javascript/api/office/office.slice) 对象。

    - [Settings](/javascript/api/office/office.settings) 对象，用于保存自定义数据（如用户偏好设置）和加载项状态。

> [!IMPORTANT]
> 并不是所有能够承载内容和任务窗格加载项的 Office 应用程序都支持一些 API 成员。要确定支持哪些成员，请参阅以下任一资源：

有关 Office 客户端应用程序中 Office JavaScript API 支持的摘要，请参 [阅了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)。

## <a name="read-and-write-to-an-active-selection-in-a-document-spreadsheet-or-presentation"></a>读取和写入文档、电子表格或演示文稿中的活动选择

您可以在文档、电子表格或演示文稿的用户当前选定内容中读取和写入。 根据外接程序的 Office 应用程序，可以在 [Document](/javascript/api/office/office.document) 对象的 [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) 和 [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) 方法中指定要读取或写入的数据结构类型作为参数。 例如，您可以指定任何用于 Word 的数据类型（文本、HTML、表格数据或 Office Open XML）、用于 Excel 的文本和表格数据，以及用于 PowerPoint 和 Project 的文本。 您还可以创建事件处理程序来检测对用户选择内容的更改。 下面的示例使用该方法以文本形式从所选内容中 `getSelectedDataAsync` 获取数据。


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

## <a name="bind-to-a-region-in-a-document-or-spreadsheet"></a>绑定到文档或电子表格中的区域

可以使用 `getSelectedDataAsync` 和 `setSelectedDataAsync` 方法在文档、电子表格或演示文稿中读取或写入用户 *的当前* 选择。 但是，如果您想在不要求用户选定内容的情况下，在运行您外接程序的各个会话中访问文档中的同一区域，您应首先绑定到该区域。 您还可以订阅该绑定区域的数据和选定内容更改事件。

可以使用 [Bindings](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1)) 对象的 [addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1))、[addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)) 或 [addFromSelectionAsync](/javascript/api/office/office.bindings) 方法添加绑定。这些方法可以返回一个标识符，您可以用它访问绑定中的数据或者订阅数据更改或选择更改事件。

下面是一个示例，它使用 `Bindings.addFromSelectionAsync` 该方法将绑定添加到文档中当前选定的文本。

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

## <a name="get-entire-documents"></a>获取整个文档

如果任务窗格外接程序在 PowerPoint 或 Word 中运行，您可以使用 [Document.getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1))、[File.getSliceAsync](/javascript/api/office/office.file#office-office-file-getsliceasync-member(1)) 和 [File.closeAsync](/javascript/api/office/office.file#office-office-file-closeasync-member(1)) 方法获取整个演示文稿或文档。

调用 `Document.getFileAsync` 时，会在 File 对象中获取文 [档](/javascript/api/office/office.file) 的副本。 该 `File` 对象提供对以 [切片](/javascript/api/office/office.slice) 对象表示的“区块”中的文档的访问权限。 调用 `getFileAsync`时，可以指定文件类型 (文本或压缩的 Open Office XML 格式) ，切片的大小 (高达 4MB) 。 若要访问对象的内容 `File` ，请调用 `File.getSliceAsync` 该对象返回 [Slice.data](/javascript/api/office/office.slice#office-office-slice-data-member) 属性中的原始数据。 如果您指定了压缩格式，则获取作为字节数组的文件数据。 如果您在将文件传输给 Web 服务，则可以在提交前将压缩的原始数据转换为 base64 编码的字符串。 最后，完成文件切片后，使用该 `File.closeAsync` 方法关闭文档。

有关详细信息，请参阅如何[从 PowerPoint 或 Word 外接程序中获取整个文档](../word/get-the-whole-document-from-an-add-in-for-word.md)。

## <a name="read-and-write-custom-xml-parts-of-a-word-document"></a>读取和写入 Word 文档的自定义 XML 部分

通过使用 Open Office XML 文件格式和内容控件，您可以将自定义 XML 部件添加到 Word 文档，并将 XML 部件中的元素绑定到文档的内容控件。打开文档时，Word 读取并自动使用自定义 XML 部件中的数据填充绑定的内容控件。用户还可以将数据写入内容控件，且在用户保存文档时，控件中的数据也将保存到绑定的 XML 部件。适用于 Word 的任务窗格外接程序可以使用 [Document.customXmlParts](/javascript/api/office/office.document#office-office-document-customxmlparts-member) 属性、[CustomXmlParts](/javascript/api/office/office.customxmlparts)、[CustomXmlPart](/javascript/api/office/office.customxmlpart) 和 [CustomXmlNode](/javascript/api/office/office.customxmlnode) 对象来动态读取文档中的数据和将数据写入文档中。

自定义 XML 部件可能与命名空间相关联。若要从命名空间的自定义 XML 部件获取数据，请使用 [CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbynamespaceasync-member(1)) 方法。

您还可以使用 [CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbyidasync-member(1)) 方法通过其 GUID 访问自定义 XML 部件。在获取自定义 XML 部件后，使用 [CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#office-office-customxmlpart-getxmlasync-member(1)) 方法获取 XML 数据。

若要向文档添加新的自定义 XML 部件，请使用该 `Document.customXmlParts` 属性获取文档中的自定义 XML 部件，并调用 [CustomXmlParts.addAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-addasync-member(1)) 方法。

有关如何使用任务窗格加载项管理自定义 XML 部件的详细信息，请参 [阅了解何时以及如何在 Word 加载项中使用 Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md)。

## <a name="persisting-add-in-settings"></a>保留加载项设置

通常需要保存外接程序的自定义数据，例如用户的首选项或外接程序的状态，并在下一次打开外接程序时访问该数据。 可以使用通用的 Web 编程技术保存该数据，例如浏览器 cookie 或 HTML 5 Web 存储。 或者，如果你的外接程序在 Excel、PowerPoint 或 Word 中运行，则可以使用 [设置](/javascript/api/office/office.settings) 对象的方法。 使用 `Settings` 对象创建的数据存储在外接程序插入并保存的电子表格、演示文稿或文档中。 此数据仅用于创建它的外接程序。

为了避免往返到存储文档的服务器，运行时会在内存中管理使用 `Settings` 该对象创建的数据。 之前保存的设置数据在初始化外接程序时加载到内存中，并在调用 [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) 方法时，仅将对数据的更改保存回文档。 在内部，将该数据作为名称/值对存储在序列化的 JSON 对象中。 可以使用 [Settings](/javascript/api/office/office.settings#office-office-settings-get-member(1)) 对象的 [get](/javascript/api/office/office.settings#office-office-settings-set-member(1))、[set](/javascript/api/office/office.settings#office-office-settings-remove-member(1)) 和 **remove** 方法从数据的内存副本中读取、写入和删除项目。 以下代码行显示如何创建名为 `themeColor` 的设置，并将它的值设置为“green”。

```js
Office.context.document.settings.set('themeColor', 'green');
```

由于使用 `set` 这些设置创建或删除的设置数据和 `remove` 方法正在处理数据的内存中副本，因此必须调用 `saveAsync` 将设置数据的更改保存到加载项正在使用的文档中。

有关使用对象的方法使用自定义数据的 `Settings` 更多详细信息，请参阅 [持久保存加载项状态和设置](persisting-add-in-state-and-settings.md)。

## <a name="read-properties-of-a-project-document"></a>读取项目文档的属性

如果您的任务窗格外接程序在 Project 中运行，则它可以从活动项目的某些项目字段、资源和任务字段中读取数据。 为此，请使用 [ProjectDocument](/javascript/api/office/office.document) 对象的方法和事件，该对象扩展 `Document` 该对象以提供其他特定于 Project 的功能。

有关读取 Project 数据的示例，请参阅[使用文本编辑器创建您第一个用于 Project 2013 的任务窗格外接程序](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)。

## <a name="permissions-model-and-governance"></a>权限模型和管治

外接程序使用 `Permissions` 其清单中的元素请求权限，以从 Office JavaScript API 访问所需的功能级别。 例如，如果加载项需要对文档进行读/写访问，则其清单必须指定 `ReadWriteDocument` 为其 `Permissions` 元素中的文本值。 因为权限的存在是为了保护用户的隐私和安全，因此最佳做法应当是，请求功能所需的最低级别的权限。 以下示例显示如何在任务窗格清单中请求 **ReadDocument** 权限。

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

有关详细信息，请参阅 [在加载项中使用 API 的请求](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)权限。

## <a name="see-also"></a>另请参阅

- [Office JavaScript API](../reference/javascript-api-for-office.md)
- [Office 外接程序清单的架构参考](../develop/add-in-manifests.md)
- [解决 Office 外接程序中的用户错误](../testing/testing-and-troubleshooting.md)
- [Office 加载项中的运行时](../testing/runtimes.md)
