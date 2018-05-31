---
title: 了解适用于 Office 的 JavaScript API
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 1ff65e8cf081330c0ce5fe8d048f703b259a5ef3
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437680"
---
# <a name="understanding-the-javascript-api-for-office"></a>了解适用于 Office 的 JavaScript API

本文提供了有关适用于 Office 的 JavaScript API 的信息以及使用方法。有关参考信息，请参阅 [适用于 Office 的 JavaScript API](https://dev.office.com/reference/add-ins/javascript-api-for-office)。有关将 Visual Studio 项目文件更新到适用于 Office 的 JavaScript API 的最新当前版本的信息，请参阅 [更新适用于 Office 的 JavaScript API 版本和清单架构文件](update-your-javascript-api-for-office-and-manifest-schema-version.md)。

> [!NOTE]
> 如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](https://docs.microsoft.com/en-us/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性](../overview/office-add-in-availability.md)页面）。 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>在加载项中引用适用于 Office 的 JavaScript API 库

[适用于 Office 的 JavaScript](https://dev.office.com/reference/add-ins/javascript-api-for-office) 库包含 Office.js 文件和关联的特定于主机应用程序的 .js 文件，例如 Excel-15.js 和 Outlook-15.js。引用该 API 最简单的方法是通过添加以下 `<script>` 到你的页面的 `<head>` 标记来使用我们的 CDN：  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

这将在加载项首次加载时下载并缓存适用于 Office 的 JavaScript API 文件，以确保对特定版本使用 Office.js 及其关联文件的最新实现。

有关 Office.js CDN 的更多详细信息（包括如何处理版本控制和向后兼容性），请参阅[从适用于 Office 的 JavaScript API 库的内容交付网络 (CDN) 对其引用](referencing-the-javascript-api-for-office-library-from-its-cdn.md)。

## <a name="initializing-your-add-in"></a>初始化加载项

**适用于：** 所有加载项类型

Office.js 提供初始化事件，API 完全加载并准备与用户开始交互时会触发该事件。你可以使用 **initialize** 事件处理程序实现常见的外接程序初始化方案，例如，可以提示用户选择 Excel 中的一些单元格，然后插入使用选定值初始化的图表。还可以使用 initialize 事件处理程序初始化外接程序的其他自定义逻辑，例如建立绑定、提示默认外接程序设置值等。

至少，initialize 事件应类似下面的示例：     

```js
Office.initialize = function () { };
```
如果你使用其他 JavaScript 框架，其中包括它们自己的初始化处理程序或测试，那么它们应放置在 Office.initialize 事件内。例如，会对 [JQuery](https://jquery.com) `$(document).ready()` 函数进行以下引用：

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {        
        // The document is ready
    });
  };
```

Office 外接程序中的所有页面需要向 initialize 事件 (**Office.initialize**) 分配一个事件处理程序。如果未能分配一个事件处理程序，则外接程序可能会在启动时出现错误。而且，如果某个用户尝试通过 Office Online Web 客户端（例如 Excel Online、PowerPoint Online 或 Outlook Web App）使用你的外接程序，则外接程序会无法运行。如果无需任何初始化代码，则向 **Office.initialize** 分配的函数的正文可以如同上述第一个示例中一样为空。

若要详细了解加载项初始化时的事件发生顺序，请参阅[加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)。

#### <a name="initialization-reason"></a>初始化原因
Office.initialize 为任务窗格和内容外接程序提供其他“_reason_”参数。此参数可用于确定如何将外接程序添加到当前文档。你可以使用此参数针对首次插入加载项时和加载项已存在于文档中时实施不同的逻辑。 

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```
有关详细信息，请参阅 [Office.initialize 事件](https://dev.office.com/reference/add-ins/shared/office.initialize)和 [InitializationReason 枚举](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration)。 

## <a name="context-object"></a>Context 对象

**适用于：** 所有加载项类型

加载项初始化时，它有许多可以在运行时环境中交互的不同对象。加载项运行时上下文通过 [Context](https://dev.office.com/reference/add-ins/shared/office.context) 对象反映在 API 中。**Context** 是提供对于最重要 API 对象（例如 [Document](https://dev.office.com/reference/add-ins/shared/document) 和 [Mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) 对象）的访问权限的主对象，这些对象继而提供对文档和邮箱内容的访问权限。

例如，在任务窗格或内容外接程序中，可以使用 [Context](https://dev.office.com/reference/add-ins/shared/office.context.document) 对象的 **document** 属性访问 **Document** 对象的属性和方法，以便与 Word 文档、Excel 工作表或 Project 计划的内容交互。类似地，在 Outlook 外接程序中，可以使用 [Context](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) 对象的 **mailbox** 属性访问 **Mailbox** 对象的属性和方法，以便与邮件、会议请求或约会内容交互。

**Context** 对象还提供对 [contentLanguage](https://dev.office.com/reference/add-ins/shared/office.context.contentlanguage) 和 [displayLanguage](https://dev.office.com/reference/add-ins/shared/office.context.displaylanguage) 属性的访问权限，以便于确定文档或项中使用的或由主机应用使用的区域设置（语言）。另外，使用 [roamingSettings](https://dev.office.com/reference/add-ins/outlook/Office.context) 属性，还可以访问 [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) 对象的成员。最后，**Context** 对象提供 [ui](https://dev.office.com/reference/add-ins/shared/officeui) 属性，以便于加载项启动弹出对话框。


## <a name="document-object"></a>Document 对象

**适用于：** 内容和任务窗格加载项类型

为了与 Excel、PowerPoint 和 Word 中的文档数据交互，API 提供 [Document](https://dev.office.com/reference/add-ins/shared/document) 对象。您可以使用 **Document** 对象成员通过以下方法访问数据：

- 读取和写入文本形式、连续单元格（矩阵）或表格中的活动选区。
    
- 表格数据（矩阵或表格）。
    
- 绑定（通过 **Bindings** 对象的“add”方法创建）。
    
- 自定义 XML 部件（仅适用于 Word）。
    
- 文档上按加载项保留的设置或加载项状态。
    
也可以使用 **Document** 对象与 Project 文档中的数据交互。特定于 Project 的 API 功能记录在成员 [ProjectDocument](https://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument) 抽象类中。有关为 Project 创建任务窗格加载项的详细信息，请参阅[适用于 Project 的任务窗格加载项](../project/project-add-ins.md)。

所有这些形式的数据访问都起始于抽象 **Document** 对象的实例。

可以在使用 **Context** 对象的 [document](https://dev.office.com/reference/add-ins/shared/office.context.document) 属性初始化任务窗格或内容加载项时访问 **Document** 对象的实例。**Document** 对象定义跨 Word 和 Excel 文档共享的通用数据访问函数，还提供对 Word 文档的 **CustomXmlParts** 对象的访问权限。

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
|表格|将选区或绑定中的数据作为 [TableData](https://dev.office.com/reference/add-ins/shared/tabledata) 对象提供。**TableData** 对象通过 **headers** 和 **rows** 属性公开数据。|表格数据访问仅在 Excel 2013 和 Word 2013 中受支持。|

#### <a name="data-type-coercion"></a>数据类型强制转换

适用于 **Document** 和 [Binding](https://dev.office.com/reference/add-ins/shared/binding) 对象的数据访问方法支持使用这些方法的 _coercionType_ 参数以及相应的 [CoercionType](https://dev.office.com/reference/add-ins/shared/coerciontype-enumeration) 枚举值，指定所需的数据类型。不管绑定的实际形状如何，不同的 Office 应用都支持常见数据类型，具体是通过尝试将数据强制转换为请求的数据类型。例如，如果选中了某个 Word 表格或段落，开发人员可以指定以纯文本、HTML、Office Open XML 或表格的形式读取它，而 API 实现则负责处理必要的转换和数据转换。


> [!TIP]
> **何时应使用矩阵与表格 coercionType 数据访问？** 如果需要表格数据在添加行和列时动态增长，且必须处理表格标题，应使用表格数据类型（具体操作是将 **Document** 或 **Binding** 对象数据访问方法的 _coercionType_ 参数指定为 `"table"` 或 **Office.CoercionType.Table**）。虽然表格数据和矩阵数据都支持在数据结构内添加行和列，但只有表格数据支持追加行和列。如果不打算添加行和列，且数据不需要使用标题功能，应使用矩阵数据类型（具体操作是将数据访问方法的 _coercionType_ 参数指定为 `"matrix"` 或 **Office.CoercionType.Matrix**），它提供了更简单的数据交互模型。

如果无法将数据强制转换为指定的类型，那么回调中的 [AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.error) 属性返回 `"failed"`，并且你可以使用 [AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.context) 属性访问 [Error](https://dev.office.com/reference/add-ins/shared/error) 对象，其中包括方法调用失败原因的信息。


## <a name="working-with-selections-using-the-document-object"></a>使用 Document 对象处理选择内容


**Document** 对象显示方法，使用户可以采用“获取和设置”方式读取和写入用户的当前选择内容。要执行此操作，**Document** 对象提供 **getSelectedDataAsync** 和 **setSelectedDataAsync** 方法。

有关演示如何使用选区执行任务的代码示例，请参阅[在文档或电子表格的活动选区中读取和写入数据](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>使用 Bindings 和 Binding 对象处理绑定


基于绑定的数据访问使内容和任务窗格加载项能够通过与绑定相关联的标识符一致地访问文档或电子表格的特定区域。加载项首先需要通过调用将文档的某一部分与唯一标识符相关联的以下某个方法来建立绑定：[addFromPromptAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync)、[addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync) 或 [addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync)。建立绑定后，加载项可以使用提供的标识符访问文档或电子表格的关联区域中包含的数据。创建绑定可为加载项提供以下值：


- 允许访问跨支持的 Office 应用程序的通用数据结构，例如：表、区域或文本（一系列连续字符）。
    
- 允许读/写操作，而不需要用户做出选择。
    
- 在加载项和文档中的数据之间建立关系。绑定会保留在文档中，以后可以进行访问。
    
建立绑定还允许您订阅仅限文档或电子表格的特定区域的数据和选择更改事件。这意味着，加载项只会收到绑定区域内发生的更改的通知，而不是收到整个文档或电子表格内的常规更改的通知。

[Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings) 对象公开 [getAllAsync](https://dev.office.com/reference/add-ins/shared/bindings.getallasync) 方法，通过该方法可以访问在文档或电子表格中建立的所有绑定的集合。可使用 [Bindings.getBindingByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) 或 [Office.select](https://dev.office.com/reference/add-ins/shared/office.select) 方法按 ID 访问单个绑定。可使用 **Bindings** 对象的以下方法之一建立新绑定和删除现有绑定：[addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync)、[addFromPromptAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync)、[addFromNamedItemAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync) 或 [releaseByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync)。

在使用  _addFromSelectionAsync_ 、 **addFromPromptAsync** 或 **addFromNamedItemAsync** 方法创建绑定时，可通过 **bindingType** 参数指定三种不同的绑定类型：



|**绑定类型**|**说明**|**主机应用程序支持**|
|:-----|:-----|:-----|
|文本绑定|绑定到可以文本形式表示的文档区域。|在 Word 中，大多数连续选区都是有效的，而在 Excel 中，只有单个单元格选区才能作为文本绑定的目标。在 Excel 中，只支持纯文本。在 Word 中，支持以下三种格式：纯文本、HTML 和 Open XML for Office。|
|矩阵绑定|绑定到包含表格数据（不带标题）的文档的固定区域。矩阵绑定中的数据以二维 **Array**（在 JavaScript 中实现为一组数组）的形式进行写入或读取。例如，两行两列 **string** 值可以写入或读取为 ` [['a', 'b'], ['c', 'd']]`，而三行单列则可以写入或读取为 `[['a'], ['b'], ['c']]`。|在 Excel 中，任何连续选择的单元格都可用于建立矩阵绑定。在 Word 中，只有表格支持矩阵绑定。|
|表格绑定|绑定到包含表格（带标题）的文档的区域。表格绑定中的数据以 [TableData](https://dev.office.com/reference/add-ins/shared/tabledata) 对象的形式进行写入或读取。**TableData** 对象通过 **headers** 和 **rows** 属性公开数据。|任何 Excel 或 Word 表格均可作为表格绑定的基础。建立表格绑定后，用户添加到表格中的每个新行或新列都自动包含在绑定中。 |

<br/>

使用 **Bindings** 对象的三个“add”方法之一创建绑定后，可以通过相应对象的方法处理绑定的数据和属性：[MatrixBinding](https://dev.office.com/reference/add-ins/shared/binding.matrixbinding)、[TableBinding](https://dev.office.com/reference/add-ins/shared/binding.tablebinding) 或 [TextBinding](https://dev.office.com/reference/add-ins/shared/binding.textbinding)。这三个对象全部继承 [Binding](https://dev.office.com/reference/add-ins/shared/binding.getdataasync) 对象的 [getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.setdataasync) 和 **setDataAsync** 方法，使你能够与绑定的数据交互。

有关演示如何使用绑定执行任务的代码示例，请参阅[绑定到文档或电子表格中的区域](bind-to-regions-in-a-document-or-spreadsheet.md)。


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>使用 CustomXmlParts 和 CustomXmlPart 对象处理自定义 XML 部件


 **适用于：** Word 的任务窗格加载项

API 的 [CustomXmlParts](https://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts) 和 [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) 对象提供访问 Word 文档中自定义 XML 部件的权限，从而启用文档内容的 XML 驱动操作。有关使用 **CustomXmlParts** 和 **CustomXmlPart** 对象的演示，请参阅 [Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts) 代码示例。


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>使用 getFileAsync 方法处理整个文档


 **适用于：** Word 和 PowerPoint 任务窗格加载项

[Document.getFileAsync](https://dev.office.com/reference/add-ins/shared/document.getfileasync) 方法以及 [File](https://dev.office.com/reference/add-ins/shared/file) 和 [Slice](https://dev.office.com/reference/add-ins/shared/slice) 对象的成员可用于一次性获取整个 Word 和 PowerPoint 文档文件，所有切片（区块）的总大小上限为 4MB。有关详细信息，请参阅[通过 PowerPoint 或 Word 加载项获取整个文档](../word/get-the-whole-document-from-an-add-in-for-word.md)。


## <a name="mailbox-object"></a>Mailbox 对象


 **适用于：** Outlook 外接程序

Outlook 外接程序主要使用通过 [Mailbox](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) 对象公开的 API 的子集。要访问专用于 Outlook 外接程序的对象和成员（例如 [Item](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) 对象），可以使用 [Context](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) 对象的 **mailbox** 属性访问 **Mailbox** 对象，如下面的代码行所示。




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

另外，Outlook 外接程序可以使用以下对象：


-  **Office** 对象：用于初始化。
    
-  **Context** 对象：用于访问内容和显示语言属性。
    
-  **RoamingSettings** 对象：用于将 Outlook 加载项专用自定义设置保存到安装了加载项的用户邮箱。
    
若要了解如何在 Outlook 加载项中使用 JavaScript，请参阅 [Outlook 加载项](https://docs.microsoft.com/en-us/outlook/add-ins/)。


## <a name="api-support-matrix"></a>API 支持矩阵


下表总结了各种类型的加载项（内容、任务窗格和 Outlook）支持的 API 和功能，以及使用[适用于 Office 的 JavaScript API v1.1 支持的 1.1 加载项清单架构和功能](update-your-javascript-api-for-office-and-manifest-schema-version.md)指定加载项支持的 Office 主机应用时，可以托管它们的 Office 应用。


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**主机名**|数据库|工作簿|邮箱|演示文稿|文档|项目|
||**支持的****主机应用程序**|Access Web App|Excel、<br/>Excel 在线|Outlook、<br/>Outlook Web App、<br/>适用于设备的 OWA|PowerPoint、<br/>PowerPoint 联机|Word|项目|
|**支持的外接程序类型**|内容|是|是||是|||
||任务窗格||是||是|是|是|
||Outlook|||是||||
|**支持的 API 功能**|读/写文本||是||是|是|是<br/>（只读）|
||读/写矩阵||是|||是||
||读/写表||是|||是||
||读/写 HTML|||||是||
||读/写<br/>Office Open XML|||||是||
||读取任务、资源、视图和字段属性||||||是|
||选择已更改事件||是|||是||
||获取整个文档||||是|是||
||绑定和绑定事件|是<br/>（仅限完全和部分表格绑定）|是|||是||
||读/写自定义 XML 部分|||||是||
||暂留加载项状态数据（设置）|是<br/>（每主机加载项）|是<br/>（每文档）|是<br/>（每邮箱）|是<br/>（每文档）|是<br/>（每文档）||
||设置更改事件|是|是||是|是||
||获取活动视图模式<br/>和视图更改事件||||是|||
||转到文档中<br/>的相应位置||是||是|是||
||使用规则和 RegEx<br/>执行上下文式激活|||是||||
||读取项目属性|||是||||
||读取用户配置文件|||是||||
||获取附件|||是||||
||获取用户标识令牌|||是||||
||调用 Exchange Web 服务|||是||||
