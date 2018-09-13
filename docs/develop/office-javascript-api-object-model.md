---
title: Office JavaScript API 对象模型
description: ''
ms.date: 07/27/2018
ms.openlocfilehash: 0842d9deafd8df411f3074dcddca04ebe0f9ed02
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945672"
---
# <a name="office-javascript-api-object-model"></a>Office JavaScript API 对象模型
Office JavaScript 加载项可以访问主机的基础功能。 大多数访问都通过一些重要的对象。 [Context](#context-object) 对象可以访问初始化后的运行时环境。 [Document](#document-object) 对象使用户可以控制 Excel、PowerPoint 或Word 文档。 [Mailbox](#mailbox-object) 对象提供对消息和用户配置文件的 Outlook 加载项访问权限。 理解这些高级对象之间的关系是 JavaScript 加载项的基础。

## <a name="context-object"></a>Context 对象

**适用于：** 所有加载项类型

加载项[初始化](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in)时，它有许多可以在运行时环境中交互的不同对象。 加载项运行时上下文通过 [Context](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js) 对象反映在 API 中。 **Context** 是用于访问最重要 API 对象（例如 [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) 和 [Mailbox](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) 对象）的主对象，这些对象继而提供对文档和邮箱内容的访问。

例如，在任务窗格或内容外接程序中，可以使用 [Context](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#document) 对象的 **document** 属性访问 **Document** 对象的属性和方法，以便与 Word 文档、Excel 工作表或 Project 计划的内容交互。类似地，在 Outlook 外接程序中，可以使用 [Context](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) 对象的 **mailbox** 属性访问 **Mailbox** 对象的属性和方法，以便与邮件、会议请求或约会内容交互。

**Context** 对象还提供对 [contentLanguage](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#contentlanguage) 和 [displayLanguage](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#displaylanguage) 属性的访问，以便于确定文档或项中使用的或由主机应用使用的区域设置（语言）。 [roamingSettings](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#roamingsettings) 属性允许您访问 [ RoamingSettings](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#roamingsettings) 对象的成员，用于存储特定于各个用户邮箱的加载项的设置。 最后，**Context** 对象提供一个允许你的外接程序启动弹出对话框的 [ui](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) 属性。


## <a name="document-object"></a>Document 对象

**适用于：** 内容和任务窗格加载项类型

为了与 Excel、PowerPoint 和 Word 中的文档数据交互，API 提供 [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) 对象。您可以使用 **Document** 对象成员通过以下方法访问数据：

- 读取和写入文本形式、连续单元格（矩阵）或表格中的活动选区。
    
- 表格数据（矩阵或表格）。
    
- 绑定（通过 **Bindings** 对象的“add”方法创建）。
    
- 自定义 XML 部件（仅适用于 Word）。
    
- 文档上按加载项保留的设置或加载项状态。
    
也可以使用 **Document** 对象与 Project 文档中的数据交互。特定于 Project 的 API 功能记录在成员 [ProjectDocument](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) 抽象类中。有关为 Project 创建任务窗格加载项的详细信息，请参阅[适用于 Project 的任务窗格加载项](../project/project-add-ins.md)。

所有这些形式的数据访问都起始于抽象 **Document** 对象的实例。

可以在使用 **Context** 对象的 [document](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#document) 属性初始化任务窗格或内容加载项时访问 **Document** 对象的实例。**Document** 对象定义跨 Word 和 Excel 文档共享的通用数据访问函数，还提供对 Word 文档的 **CustomXmlParts** 对象的访问权限。

**Document** 对象支持四种方式以供开发人员访问文档内容：


- 基于选区的访问
    
- 基于绑定的访问
    
- 基于自定义 XML 部件的访问（仅适用于 Word）
    
- 基于整个文档的访问（仅适用于 PowerPoint 和 Word）
    
为了帮助您理解基于选区和绑定的数据访问方法的工作方式，我们将首先解释数据访问 API 如何跨不同的 Office 应用程序提供一致的数据访问。


### <a name="consistent-data-access-across-office-applications"></a>跨 Office 应用程序的一致数据访问

 **适用于：** 内容和任务窗格加载项类型

为了创建跨不同 Office 文档无缝工作的扩展，适用于 Office 的 JavaScript API 通过通用数据类型和强制将不同文档内容划分为三种通用数据类型的功能抽象出每个 Office 应用程序的细节。


#### <a name="common-data-types"></a>通用数据类型

在基于选区和基于绑定的数据访问中，文档内容通过跨所有受支持的 Office 应用程序通用的数据类型来公开。在 Office 2013 中，支持三种主要的数据类型：



|**数据类型**|**说明**|**主机应用程序支持**|
|:-----|:-----|:-----|
|文本|提供选定范围或绑定中数据的字符串表示形式。|在 Excel 2013、Project 2013 和 PowerPoint 2013 中，仅支持纯文本。在 Word 2013 中，支持三种文本格式：纯文本、HTML 和 Office Open XML (OOXML)。如果选中的是 Excel 单元格中的文本，基于选定范围的方法会对单元格的全部内容执行读取和写入操作，即使仅选中了单元格中的部分文本，也不例外。如果选中的是 Word 和 PowerPoint 中的文本，基于选定范围的方法只会对选中的一系列字符执行读取和写入操作。Project 2013 和 PowerPoint 2013 仅支持基于选定范围的数据访问。|
|矩阵|将选定范围或绑定中的数据作为二维 **Array** 提供，这在 JavaScript 中实现为一组数组。例如，两行两列 **string** 值为 ` [['a', 'b'], ['c', 'd']]`，而三行一列则为 `[['a'], ['b'], ['c']]`。|仅 Excel 2013 和 Word 2013 支持矩阵数据访问。|
|表格|将选区或绑定中的数据作为 [TableData](https://docs.microsoft.com/javascript/api/office/office.tabledata?view=office-js) 对象提供。**TableData** 对象通过 **headers** 和 **rows** 属性公开数据。|表格数据访问仅在 Excel 2013 和 Word 2013 中受支持。|

#### <a name="data-type-coercion"></a>数据类型强制转换

适用于 **Document** 和 [Binding](https://docs.microsoft.com/javascript/api/office/office.binding?view=office-js) 对象的数据访问方法支持使用这些方法的 _coercionType_ 参数以及相应的 [CoercionType](https://docs.microsoft.com/javascript/api/office/office.coerciontype?view=office-js) 枚举值，指定所需的数据类型。不管绑定的实际形状如何，不同的 Office 应用都支持常见数据类型，具体是通过尝试将数据强制转换为请求的数据类型。例如，如果选中了某个 Word 表格或段落，开发人员可以指定以纯文本、HTML、Office Open XML 或表格的形式读取它，而 API 实现则负责处理必要的转换和数据转换。


> [!TIP]
> **何时应使用矩阵与表格 coercionType 数据访问？** 如果需要表格数据在添加行和列时动态增长，且必须处理表格标题，应使用表格数据类型（具体操作是将 **Document** 或 **Binding** 对象数据访问方法的 _coercionType_ 参数指定为 `"table"` 或 **Office.CoercionType.Table**）。虽然表格数据和矩阵数据都支持在数据结构内添加行和列，但只有表格数据支持追加行和列。如果不打算添加行和列，且数据不需要使用标题功能，应使用矩阵数据类型（具体操作是将数据访问方法的 _coercionType_ 参数指定为 `"matrix"` 或 **Office.CoercionType.Matrix**），它提供了更简单的数据交互模型。

如果无法将数据强制转换为指定的类型，那么回调中的 [AsyncResult.status](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js.error) 属性返回 `"failed"`，并且你可以使用 [AsyncResult.error](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js.context) 属性访问 [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js) 对象，其中包括方法调用失败原因的信息。


## <a name="working-with-selections-using-the-document-object"></a>使用 Document 对象处理选择内容


**Document** 对象显示方法，使用户可以采用“获取和设置”方式读取和写入用户的当前选择内容。要执行此操作，**Document** 对象提供 **getSelectedDataAsync** 和 **setSelectedDataAsync** 方法。

有关演示如何使用选区执行任务的代码示例，请参阅[在文档或电子表格的活动选区中读取和写入数据](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>使用 Bindings 和 Binding 对象处理绑定


基于绑定的数据访问使内容和任务窗格加载项能够通过与绑定相关联的标识符一致地访问文档或电子表格的特定区域。加载项首先需要通过调用将文档的某一部分与唯一标识符相关联的以下某个方法来建立绑定：[addFromPromptAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfrompromptasync-bindingtype--options--callback-)、[addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-) 或 [addFromNamedItemAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromnameditemasync-itemname--bindingtype--options--callback-)。建立绑定后，加载项可以使用提供的标识符访问文档或电子表格的关联区域中包含的数据。创建绑定可为加载项提供以下值：


- 允许访问跨支持的 Office 应用程序的通用数据结构，例如：表、区域或文本（一系列连续字符）。
    
- 允许读/写操作，而不需要用户做出选择。
    
- 在加载项和文档中的数据之间建立关系。绑定会保留在文档中，以后可以进行访问。
    
建立绑定还允许您订阅仅限文档或电子表格的特定区域的数据和选择更改事件。这意味着，加载项只会收到绑定区域内发生的更改的通知，而不是收到整个文档或电子表格内的常规更改的通知。

[Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js) 对象公开 [getAllAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#getallasync-options--callback-) 方法，通过该方法可以访问在文档或电子表格中建立的所有绑定的集合。可使用 [Bindings.getBindingByIdAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#getbyidasync-id--options--callback-) 或 [Office.select](https://docs.microsoft.com/javascript/api/office?view=office-js) 方法按 ID 访问单个绑定。可使用 **Bindings** 对象的以下方法之一建立新绑定和删除现有绑定：[addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-)、[addFromPromptAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfrompromptasync-bindingtype--options--callback-)、[addFromNamedItemAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromnameditemasync-itemname--bindingtype--options--callback-) 或 [releaseByIdAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#releasebyidasync-id--options--callback-)。

在使用  _addFromSelectionAsync_ 、 **addFromPromptAsync** 或 **addFromNamedItemAsync** 方法创建绑定时，可通过 **bindingType** 参数指定三种不同的绑定类型：



|**绑定类型**|**说明**|**主机应用程序支持**|
|:-----|:-----|:-----|
|文本绑定|绑定到可以文本形式表示的文档区域。|在 Word 中，大多数连续选区都是有效的，而在 Excel 中，只有单个单元格选区才能作为文本绑定的目标。在 Excel 中，只支持纯文本。在 Word 中，支持以下三种格式：纯文本、HTML 和 Open XML for Office。|
|矩阵绑定|绑定到包含表格数据（不带标题）的文档的固定区域。矩阵绑定中的数据以二维 **Array**（在 JavaScript 中实现为一组数组）的形式进行写入或读取。例如，两行两列 **string** 值可以写入或读取为 ` [['a', 'b'], ['c', 'd']]`，而三行单列则可以写入或读取为 `[['a'], ['b'], ['c']]`。|在 Excel 中，任何连续选择的单元格都可用于建立矩阵绑定。在 Word 中，只有表格支持矩阵绑定。|
|表格绑定|绑定到包含表格（带标题）的文档的区域。表格绑定中的数据以 [TableData](https://docs.microsoft.com/javascript/api/office/office.tabledata?view=office-js) 对象的形式进行写入或读取。**TableData** 对象通过 **headers** 和 **rows** 属性公开数据。|任何 Excel 或 Word 表格均可作为表格绑定的基础。建立表格绑定后，用户添加到表格中的每个新行或新列都自动包含在绑定中。 |

<br/>

使用 **Bindings** 对象的三个“add”方法之一创建绑定后，可以通过相应对象的方法处理绑定的数据和属性：[MatrixBinding](https://docs.microsoft.com/javascript/api/office/office.matrixbinding?view=office-js)、[TableBinding](https://docs.microsoft.com/javascript/api/office/office.tablebinding?view=office-js) 或 [TextBinding](https://docs.microsoft.com/javascript/api/office/office.textbinding?view=office-js)。这三个对象全部继承 [Binding](https://docs.microsoft.com/javascript/api/office/office.binding?view=office-js#getdataasync-options--callback-) 对象的 [getDataAsync](https://docs.microsoft.com/javascript/api/office/office.binding?view=office-js#setdataasync-data--options--callback-) 和 **setDataAsync** 方法，使你能够与绑定的数据交互。

有关演示如何使用绑定执行任务的代码示例，请参阅[绑定到文档或电子表格中的区域](bind-to-regions-in-a-document-or-spreadsheet.md)。


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>使用 CustomXmlParts 和 CustomXmlPart 对象处理自定义 XML 部件


 **适用于：** Word 的任务窗格加载项

API 的 [CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js) 和 [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js) 对象提供对 Word 文档中自定义 XML 部件的访问，从而基于 XML 对文档内容的执行操作。有关使用 **CustomXmlParts** 和 **CustomXmlPart** 对象的演示，请参阅 [Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts) 代码示例。


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>使用 getFileAsync 方法处理整个文档


 **适用于：** Word 和 PowerPoint 任务窗格加载项

[Document.getFileAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfileasync-filetype--options--callback-) 方法以及 [File](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js) 和 [Slice](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js) 对象的成员可用于一次性获取整个 Word 和 PowerPoint 文档文件，所有切片（区块）的总大小上限为 4MB。有关详细信息，请参阅[通过 PowerPoint 或 Word 加载项获取整个文档](../word/get-the-whole-document-from-an-add-in-for-word.md)。


## <a name="mailbox-object"></a>Mailbox 对象

**适用于：** Outlook 外接程序

Outlook 外接程序主要使用通过 [Mailbox](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) 对象公开的 API 的子集。要访问专用于 Outlook 外接程序的对象和成员（例如 [Item](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) 对象），可以使用 [Context](https://docs.microsoft.com/javascript/api/outlook/Office.mailbox?view=office-js) 对象的 **mailbox** 属性访问 **Mailbox** 对象，如下面的代码行所示。




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

另外，Outlook 外接程序可以使用以下对象：


-  **Office** 对象：用于初始化。
    
-  **Context** 对象：用于访问内容和显示语言属性。
    
-  **RoamingSettings** 对象：用于将 Outlook 加载项专用自定义设置保存到安装了加载项的用户邮箱。
    
若要了解如何在 Outlook 加载项中使用 JavaScript，请参阅 [Outlook 加载项](https://docs.microsoft.com/outlook/add-ins/)。