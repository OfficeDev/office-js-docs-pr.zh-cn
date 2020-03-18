---
title: 常见 JavaScript API 对象模型
description: 了解 Office JavaScript 通用 API 对象模型
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 0944ed36f2d8e4a4ed557dbd25e9f21be137cdaf
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719446"
---
# <a name="common-javascript-api-object-model"></a>常见 JavaScript API 对象模型

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office JavaScript 加载项可提供对主机的基础功能的访问权限。 大多数此类访问权限可以访问一些重要的对象。 [Context](#context-object) 对象提供在初始化之后对运行时环境的访问权限。 [Document](#document-object) 对象使用户能够控制 Excel、PowerPoint 或 Word 文档。 [Mailbox](#mailbox-object) 对象使 Outlook 加载项能够访问邮件和用户配置文件。 理解这些高级对象之间的关系是 JavaScript 加载项的基础。

## <a name="context-object"></a>Context 对象

**适用于：** 所有加载项类型

如果加载项[已初始化](initialize-add-in.md)，则它具有许多可在运行时环境中交互的不同对象。 加载项的运行时上下文通过 [Context](/javascript/api/office/office.context) 对象反应在 API 中。 **Context** 是主要对象，它提供对大部分 API 最重要对象的访问权限，例如 [Document](/javascript/api/office/office.document) 和 [Mailbox](/javascript/api/outlook/Office.mailbox) 对象，二者反过来又提供对文档和邮箱内容的访问权限。

例如，在任务窗格或内容外接程序中，可以使用 **Context** 对象的 [document](/javascript/api/office/office.context#document) 属性访问 **Document** 对象的属性和方法，以便与 Word 文档、Excel 工作表或 Project 计划的内容交互。类似地，在 Outlook 外接程序中，可以使用 **Context** 对象的 [mailbox](/javascript/api/outlook/Office.mailbox) 属性访问 **Mailbox** 对象的属性和方法，以便与邮件、会议请求或约会内容交互。

**Context** 对象还提供对 [contentLanguage](/javascript/api/office/office.context#contentlanguage) 和 [displayLanguage](/javascript/api/office/office.context#displaylanguage) 属性的访问权限，这些属性允许你确定文档或项目中或由宿主应用程序使用的区域设置（语言）。 [roamingSettings](/javascript/api/office/office.context#roamingsettings) 属性使你能够访问 [RoamingSettings](/javascript/api/office/office.context#roamingsettings) 对象的成员，该对象用于存储各用户邮箱的加载项特定的设置。 最后，**Context** 对象提供一个允许你的加载项启动弹出对话框的 [ui](/javascript/api/office/office.ui) 属性。


## <a name="document-object"></a>Document 对象

**适用于：** 内容和任务窗格加载项类型

为了与 Excel、PowerPoint 和 Word 中的文档数据交互，API 提供 [Document](/javascript/api/office/office.document) 对象。 您可以使用`Document`对象成员从以下几个方面访问数据：

- 读取和写入文本形式、连续单元格（矩阵）或表格中的活动选区。

- 表格数据（矩阵或表格）。

- 绑定（使用`Bindings`对象的 "add" 方法创建）。

- 自定义 XML 部件（仅适用于 Word）。

- 文档上按加载项保留的设置或加载项状态。

您还可以使用`Document`对象与项目文档中的数据进行交互。 特定于 Project 的 API 功能记录在成员 [ProjectDocument](/javascript/api/office/office.document) 抽象类中。 有关为 Project 创建任务窗格加载项的详细信息，请参阅[适用于 Project 的任务窗格加载项](../project/project-add-ins.md)。

所有这些形式的数据访问都是从抽象`Document`对象的实例开始的。

使用`Context`对象的[document](/javascript/api/office/office.context#document)属性初始化任务`Document`窗格或内容加载项时，可以访问对象的实例。 该`Document`对象定义在 Word 和 Excel 文档中共享的常见数据访问函数，还提供对 word `CustomXmlParts`文档的对象的访问权限。

`Document`对象支持四种方式以供开发人员访问文档内容：


- 基于选区的访问

- 基于绑定的访问

- 基于自定义 XML 部件的访问（仅适用于 Word）

- 基于整个文档的访问（仅适用于 PowerPoint 和 Word）

为了帮助您理解基于选区和绑定的数据访问方法的工作方式，我们将首先解释数据访问 API 如何跨不同的 Office 应用程序提供一致的数据访问。


### <a name="consistent-data-access-across-office-applications"></a>跨 Office 应用程序的一致数据访问

 **适用于：** 内容和任务窗格加载项类型

为了创建可跨不同的 Office 文档无缝工作的扩展功能，Office JavaScript API 通过常见数据类型抽象化了每个 Office 应用程序的 particularities，并能够将不同的文档内容强制转换为三个常见的数据类型。


#### <a name="common-data-types"></a>通用数据类型

在基于选区和基于绑定的数据访问中，文档内容通过跨所有受支持的 Office 应用程序通用的数据类型来公开。在 Office 2013 中，支持三种主要的数据类型：



|**数据类型**|**说明**|**主机应用程序支持**|
|:-----|:-----|:-----|
|文本|提供选定范围或绑定中数据的字符串表示形式。|在 Excel 2013、Project 2013 和 PowerPoint 2013 中，仅支持纯文本。在 Word 2013 中，支持三种文本格式：纯文本、HTML 和 Office Open XML (OOXML)。如果选中的是 Excel 单元格中的文本，基于选定范围的方法会对单元格的全部内容执行读取和写入操作，即使仅选中了单元格中的部分文本，也不例外。如果选中的是 Word 和 PowerPoint 中的文本，基于选定范围的方法只会对选中的一系列字符执行读取和写入操作。Project 2013 和 PowerPoint 2013 仅支持基于选定范围的数据访问。|
|矩阵|将选定范围或绑定中的数据作为二维 **Array** 提供，这在 JavaScript 中实现为一组数组。例如，两行两列 **string** 值为 ` [['a', 'b'], ['c', 'd']]`，而三行一列则为 `[['a'], ['b'], ['c']]`。|仅 Excel 2013 和 Word 2013 支持矩阵数据访问。|
|Table|将选区或绑定中的数据作为 [TableData](/javascript/api/office/office.tabledata) 对象提供。 `TableData`对象通过`headers`和`rows`属性公开数据。|表格数据访问仅在 Excel 2013 和 Word 2013 中受支持。|

#### <a name="data-type-coercion"></a>数据类型强制转换

和 Binding 对象上的数据访问方法支持使用这些方法的_coercionType_参数以及相应的[coercionType](/javascript/api/office/office.coerciontype)枚举值指定所需的数据类型。 [Binding](/javascript/api/office/office.binding) `Document` 不管绑定的实际形状如何，不同的 Office 应用程序都通过尝试将数据强制转换为请求的数据类型来支持通用的数据类型。 例如，如果选中了某个 Word 表格或段落，开发人员可以指定以纯文本、HTML、Office Open XML 或表格的形式读取它，而 API 实现则负责处理必要的转换和数据转换。


> [!TIP]
> **何时应使用矩阵与表格 coercionType 数据访问？** 如果需要在添加行和列时动态增大表格数据，并且必须使用表格标题，则应使用 table 数据类型（通过将`Document` `Binding` _coercionType_参数指定为`"table"`或的对象数据访问方法`Office.CoercionType.Table`）。 表格数据和矩阵数据中都支持在数据结构内添加行和列，但仅支持对表格数据追加行和列。 如果您不打算添加行和列，并且您的数据不需要标头功能，则应使用矩阵数据类型（通过将 data access 方法的_coercionType_参数指定为`"matrix"`或`Office.CoercionType.Matrix`），这提供了与数据交互的更简单的模型。

如果无法将数据强制转换为指定的类型，那么回调中的 [AsyncResult.status](/javascript/api/office/office.asyncresult#status) 属性返回 `"failed"`，并且你可以使用 [AsyncResult.error](/javascript/api/office/office.asyncresult#error) 属性访问 [Error](/javascript/api/office/office.error) 对象，其中包括方法调用失败原因的信息。


## <a name="working-with-selections-using-the-document-object"></a>使用 Document 对象处理选择内容


该`Document`对象公开的方法允许您以 "获取和设置" 方式读取和写入用户当前所选内容。 若要执行此操作`Document` ，对象将`getSelectedDataAsync`提供`setSelectedDataAsync`和方法。

有关演示如何使用选区执行任务的代码示例，请参阅[在文档或电子表格的活动选区中读取和写入数据](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>使用 Bindings 和 Binding 对象处理绑定


基于绑定的数据访问使内容和任务窗格加载项能够通过与绑定相关联的标识符一致地访问文档或电子表格的特定区域。加载项首先需要通过调用将文档的某一部分与唯一标识符相关联的以下某个方法来建立绑定：[addFromPromptAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-)、[addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) 或 [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-)。建立绑定后，加载项可以使用提供的标识符访问文档或电子表格的关联区域中包含的数据。创建绑定可为加载项提供以下值：


- 允许访问跨支持的 Office 应用程序的通用数据结构，例如：表、区域或文本（一系列连续字符）。

- 允许读/写操作，而不需要用户做出选择。

- 在加载项和文档中的数据之间建立关系。绑定会保留在文档中，以后可以进行访问。

建立绑定还允许您订阅仅限文档或电子表格的特定区域的数据和选择更改事件。这意味着，加载项只会收到绑定区域内发生的更改的通知，而不是收到整个文档或电子表格内的常规更改的通知。

[Bindings](/javascript/api/office/office.bindings) 对象公开 [getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) 方法，通过该方法可以访问在文档或电子表格中建立的所有绑定的集合。 可使用 [Bindings.getBindingByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) 或 [Office.select](/javascript/api/office) 方法按 ID 访问单个绑定。 您可以`Bindings`使用以下对象的方法之一来建立新绑定，并删除现有绑定： [addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-)、 [addFromPromptAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-)、 [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-)或[releaseByIdAsync](/javascript/api/office/office.bindings#releasebyidasync-id--options--callback-)。

在使用或`addFromNamedItemAsync`方法创建绑定时`addFromSelectionAsync`， `addFromPromptAsync`可以使用_bindingType_参数指定三种不同类型的绑定：



|**绑定类型**|**说明**|**主机应用程序支持**|
|:-----|:-----|:-----|
|文本绑定|绑定到可以文本形式表示的文档区域。|在 Word 中，大多数连续选区都是有效的，而在 Excel 中，只有单个单元格选区才能作为文本绑定的目标。在 Excel 中，只支持纯文本。在 Word 中，支持以下三种格式：纯文本、HTML 和 Open XML for Office。|
|矩阵绑定|绑定到包含表格数据（不带标题）的文档的固定区域。矩阵绑定中的数据以二维 **Array**（在 JavaScript 中实现为一组数组）的形式进行写入或读取。例如，两行两列 **string** 值可以写入或读取为 ` [['a', 'b'], ['c', 'd']]`，而三行单列则可以写入或读取为 `[['a'], ['b'], ['c']]`。|在 Excel 中，任何连续选择的单元格都可用于建立矩阵绑定。在 Word 中，只有表格支持矩阵绑定。|
|表格绑定|绑定到包含带标题的表格的文档区域。 表格绑定中的数据作为 [TableData](/javascript/api/office/office.tabledata) 对象写入或读取。 `TableData`对象通过**标头**和**行**属性公开数据。|任何 Excel 或 Word 表格均可作为表格绑定的基础。建立表格绑定后，用户添加到表格中的每个新行或新列都自动包含在绑定中。 |

<br/>

使用`Bindings`该对象的三个 "添加" 方法之一创建绑定之后，可以使用相应对象的方法来处理绑定的数据和属性： [MatrixBinding](/javascript/api/office/office.matrixbinding)、 [TableBinding](/javascript/api/office/office.tablebinding)或[textbinding 添加](/javascript/api/office/office.textbinding)。 所有这三个对象都继承`Binding`对象的[binding.getdataasync](/javascript/api/office/office.binding#getdataasync-options--callback-)和[binding.setdataasync](/javascript/api/office/office.binding#setdataasync-data--options--callback-)方法，您可以使用这些方法与绑定的数据进行交互。

有关演示如何使用绑定执行任务的代码示例，请参阅[绑定到文档或电子表格中的区域](bind-to-regions-in-a-document-or-spreadsheet.md)。


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>使用 CustomXmlParts 和 CustomXmlPart 对象处理自定义 XML 部件


 **适用于：** Word 的任务窗格加载项

API 的 [CustomXmlParts](/javascript/api/office/office.customxmlparts) 和 [CustomXmlPart](/javascript/api/office/office.customxmlpart) 对象提供访问 Word 文档中自定义 XML 部件的权限，从而启用文档内容的 XML 驱动操作。 有关使用`CustomXmlParts`和`CustomXmlPart`对象的演示，请参阅[Word 加载项-与工作的自定义 XML 部分](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts)的代码示例。


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>使用 getFileAsync 方法处理整个文档


 **适用于：** Word 和 PowerPoint 任务窗格加载项

[Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-) 方法以及 [File](/javascript/api/office/office.file) 和 [Slice](/javascript/api/office/office.slice) 对象的成员可用于一次性获取整个 Word 和 PowerPoint 文档文件，所有切片（区块）的总大小上限为 4MB。有关详细信息，请参阅[通过 PowerPoint 或 Word 加载项获取整个文档](../word/get-the-whole-document-from-an-add-in-for-word.md)。


## <a name="mailbox-object"></a>Mailbox 对象

**适用于：** Outlook 外接程序

Outlook 外接程序主要使用通过 [Mailbox](/javascript/api/outlook/Office.mailbox) 对象公开的 API 的子集。要访问专用于 Outlook 外接程序的对象和成员（例如 [Item](/javascript/api/outlook/Office.mailbox) 对象），可以使用 [Context](/javascript/api/outlook/Office.mailbox) 对象的 **mailbox** 属性访问 **Mailbox** 对象，如下面的代码行所示。

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

另外，Outlook 外接程序可以使用以下对象：

- `Office`对象：用于初始化。

- `Context`对象：用于访问内容和显示语言属性。

- `RoamingSettings`对象：用于将特定于 Outlook 外接程序的自定义设置保存到安装外接程序的用户邮箱中。

有关在 Outlook 加载项中使用 JavaScript 的信息，请参阅 [Outlook 加载项](../outlook/outlook-add-ins-overview.md)。

