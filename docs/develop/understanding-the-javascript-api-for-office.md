
# <a name="understanding-the-javascript-api-for-office"></a>了解适用于 Office 的 JavaScript API



本文提供了有关适用于 Office 的 JavaScript API 的信息以及使用方法。有关参考信息，请参阅 [适用于 Office 的 JavaScript API](http://dev.office.com/reference/add-ins/javascript-api-for-office)。有关将 Visual Studio 项目文件更新到适用于 Office 的 JavaScript API 的最新当前版本的信息，请参阅 [更新适用于 Office 的 JavaScript API 版本和清单架构文件](../develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)。

> [!NOTE]
> 生成外接程序时，如果计划将外接程序[发布](../publish/publish.md)到 Office 应用商店，请务必遵循 [Office 应用商店验证策略](https://msdn.microsoft.com/zh-cn/library/jj220035.aspx)。例如，外接程序必须适用于支持你定义的方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://msdn.microsoft.com/zh-cn/library/jj220035.aspx#Anchor_3)以及 [Office 外接程序主机和可用性](https://dev.office.com/add-in-availability)页）。

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>在外接程序中引用适用于 Office 的 JavaScript API 库

[适用于 Office 的 JavaScript](http://dev.office.com/reference/add-ins/javascript-api-for-office) 库包含 Office.js 文件和关联的特定于主机应用程序的 .js 文件，例如 Excel-15.js 和 Outlook-15.js。引用该 API 最简单的方法是通过添加以下 `<script>` 到你的页面的 `<head>` 标记来使用我们的 CDN：  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

这将在加载项首次加载时下载并缓存适用于 Office 的 JavaScript API 文件，以确保对特定版本使用 Office.js 及其关联文件的最新实现。

有关 Office.js CDN 的详细信息，包括如何处理版本控制和向后兼容性，请参阅[从适用于 Office 的 JavaScript API 的内容传送网络 (CDN) 引用适用于 Office 的 JavaScript API 库](referencing-the-javascript-api-for-office-library-from-its-cdn.md)。

## <a name="initializing-your-add-in"></a>初始化加载项


 **适用于：**所有加载项类型


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

有关初始化外接程序时的事件顺序的详细信息，请参阅 [加载 DOM 和运行时环境](../develop/loading-the-dom-and-runtime-environment.md)。

#### <a name="initialization-reason"></a>Initialization Reason
Office.initialize 为任务窗格和内容外接程序提供其他“_reason_”参数。此参数可用于确定如何将外接程序添加到当前文档。你可以使用此参数针对首次插入加载项时和加载项已存在于文档中时实施不同的逻辑。 

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
      switch (reason) {
        case 'inserted': console.log('The add-in was just inserted.');
        case 'documentOpened': console.log('The add-in is already part of the document.');
    }
}
```
有关详细信息，请参阅 [Office.initialize 事件](http://dev.office.com/reference/add-ins/shared/office.initialize)和 [InitializationReason 枚举](http://dev.office.com/reference/add-ins/shared/initializationreason-enumeration) 

## <a name="context-object"></a>Context 对象

 **适用于：**所有加载项类型

加载项初始化时，它有许多可以在运行时环境中交互的不同对象。加载项运行时上下文通过 [Context](http://dev.office.com/reference/add-ins/shared/office.context) 对象反映在 API 中。**Context** 是提供对于最重要 API 对象（例如 [Document](http://dev.office.com/reference/add-ins/shared/document) 和 [Mailbox](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) 对象）的访问权限的主对象，这些对象继而提供对文档和邮箱内容的访问权限。

例如，在任务窗格或内容外接程序中，可以使用 [Context](http://dev.office.com/reference/add-ins/shared/office.context.document) 对象的 **document** 属性访问 **Document** 对象的属性和方法，以便与 Word 文档、Excel 工作表或 Project 计划的内容交互。类似地，在 Outlook 外接程序中，可以使用 [Context](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) 对象的 **mailbox** 属性访问 **Mailbox** 对象的属性和方法，以便与邮件、会议请求或约会内容交互。

**Context** 对象还提供对 [contentLanguage](http://dev.office.com/reference/add-ins/shared/office.context.contentlanguage) 和 [displayLanguage](http://dev.office.com/reference/add-ins/shared/office.context.displaylanguage) 属性的访问权限，这些属性允许你确定文档或项目中或由宿主应用程序使用的区域设置（语言）。另外，[roamingSettings](http://dev.office.com/reference/add-ins/outlook/Office.context) 属性允许你访问 [RoamingSettings](http://dev.office.com/reference/add-ins/outlook/RoamingSettings) 对象的成员。最后，**Context** 对象提供一个允许你的外接程序启动弹出对话框的 [ui](http://dev.office.com/reference/add-ins/shared/officeui) 属性。


## <a name="document-object"></a>Document 对象


 **适用于：**内容和任务窗格加载项类型

为了与 Excel、PowerPoint 和 Word 中的文档数据交互，API 提供 [Document](http://dev.office.com/reference/add-ins/shared/document) 对象。您可以使用 **Document** 对象成员通过以下方法访问数据：


- 读取和写入文本形式、连续单元格（矩阵）或表格中的活动选区。
    
- 表格数据（矩阵或表格）。
    
- 绑定（通过 **Bindings** 对象的“add”方法创建）。
    
- 自定义 XML 部件（仅适用于 Word）。
    
- 文档上按加载项保留的设置或加载项状态。
    
也可以使用 **Document** 对象与 Project 文档中的数据交互。特定于 Project 的 API 功能记录在成员 [ProjectDocument](http://dev.office.com/reference/add-ins/shared/projectdocument.projectdocument) 抽象类中。有关为 Project 创建任务窗格加载项的详细信息，请参阅[适用于 Project 的任务窗格加载项](../project/project-add-ins.md)。

所有这些形式的数据访问都起始于抽象 **Document** 对象的实例。

可以在使用 **Context** 对象的 [document](http://dev.office.com/reference/add-ins/shared/office.context.document) 属性初始化任务窗格或内容加载项时访问 **Document** 对象的实例。**Document** 对象定义跨 Word 和 Excel 文档共享的通用数据访问函数，还提供对 Word 文档的 **CustomXmlParts** 对象的访问权限。

**Document** 对象支持四种方式以供开发人员访问文档内容：


- 基于选区的访问
    
- 基于绑定的访问
    
- 基于自定义 XML 部件的访问（仅适用于 Word）
    
- 基于整个文档的访问（仅适用于 PowerPoint 和 Word）
    
为了帮助您理解基于选区和绑定的数据访问方法的工作方式，我们将首先解释数据访问 API 如何跨不同的 Office 应用程序提供一致的数据访问。


### <a name="consistent-data-access-across-office-applications"></a>跨 Office 应用程序的一致数据访问

 **适用于：**内容和任务窗格加载项类型

为了创建跨不同 Office 文档无缝工作的扩展，适用于 Office 的 JavaScript API 通过通用数据类型和强制将不同文档内容划分为三种通用数据类型的功能抽象出每个 Office 应用程序的细节。


#### <a name="common-data-types"></a>通用数据类型

在基于选区和基于绑定的数据访问中，文档内容通过跨所有受支持的 Office 应用程序通用的数据类型来公开。在 Office 2013 中，支持三种主要的数据类型：



|**数据类型**|**说明**|**主机应用程序支持**|
|:-----|:-----|:-----|
|文本|提供选区或绑定中数据的字符串表示形式。|在 Excel 2013、Project 2013 和 PowerPoint 2013 中，仅支持纯文本。在 Word 2013 中，支持三种文本格式：纯文本、HTML 和 Office Open XML (OOXML)。当选中 Excel 单元格中的文本时，基于选区的方法会读取和写入单元格的整个内容，即使仅选择了单元格中的部分文本也是如此。当选中 Word 和 PowerPoint 中的文本时，基于选区的方法会仅读取和写入选中的一系列字符。Project 2013 和 PowerPoint 2013 仅支持基于选区的数据访问。|
|Matrix|将选区或绑定中的数据作为双维的 **Array** 提供，这在 JavaScript 中作为数组的数组来实现。例如，两行 **string** 值以两列表示则为 ` [['a', 'b'], ['c', 'd']]`，而单列三行则为 `[['a'], ['b'], ['c']]`。|矩阵数据访问仅在 Excel 2013 和 Word 2013 中受支持。|
|Table|将选区或绑定中的数据作为 [TableData](http://dev.office.com/reference/add-ins/shared/tabledata) 对象提供。**TableData** 对象通过 **headers** 和 **rows** 属性公开数据。|表格数据访问仅在 Excel 2013 和 Word 2013 中受支持。|

#### <a name="data-type-coercion"></a>数据类型强制转换

**Document** 和 [Binding](http://dev.office.com/reference/add-ins/shared/binding) 对象上的数据访问方法支持使用这些方法的 _coercionType_ 参数以及相应的 [CoercionType](http://dev.office.com/reference/add-ins/shared/coerciontype-enumeration) 枚举值指定所需的数据类型。不管绑定的实际形状如何，不同的 Office 应用程序都通过尝试将数据强制转换为请求的数据类型来支持通用的数据类型。例如，如果选中了某个 Word 表格或段落，开发人员可以指定将其读取为纯文本、HTML、Office Open XML 或表格，而 API 实现处理必要的转换和数据转换。


 >**提示**   **你应该在何时使用矩阵和表格 coercionType 进行数据访问？**若你需要在添加行和列时使表格数据动态增大，且必须使用表格标题，则应该使用表格数据类型（通过将 **Document** 或 **Binding** 对象数据访问方法的 _coercionType_ 参数指定为 `"table"` 或 **Office.CoercionType.Table**）。表格数据和矩阵数据中都支持在数据结构内添加行和列，但仅支持对表格数据追加行和列。若你不计划添加行和列，且数据不需要标题功能，则应使用矩阵数据类型（通过将数据访问方法的 _coercionType_ 参数指定为 `"matrix"` 或 **Office.CoercionType.Matrix**），它提供了与数据交互更简单的模型。

如果无法将数据强制转换为指定的类型，那么回调中的 [AsyncResult.status](http://dev.office.com/reference/add-ins/shared/asyncresult.error) 属性返回 `"failed"`，并且你可以使用 [AsyncResult.error](http://dev.office.com/reference/add-ins/shared/asyncresult.context) 属性访问 [Error](http://dev.office.com/reference/add-ins/shared/error) 对象，其中包括方法调用失败原因的信息。


## <a name="working-with-selections-using-the-document-object"></a>使用 Document 对象处理选择内容


**Document** 对象显示方法，使用户可以采用“获取和设置”方式读取和写入用户的当前选择内容。要执行此操作，**Document** 对象提供 **getSelectedDataAsync** 和 **setSelectedDataAsync** 方法。

有关演示如何使用选区执行任务的代码示例，请参阅[在文档或电子表格的活动选区中读取和写入数据](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>使用 Bindings 和 Binding 对象处理绑定


基于绑定的数据访问使内容和任务窗格加载项能够通过与绑定相关联的标识符一致地访问文档或电子表格的特定区域。加载项首先需要通过调用将文档的某一部分与唯一标识符相关联的以下某个方法来建立绑定：[addFromPromptAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync)、[addFromSelectionAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync) 或 [addFromNamedItemAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync)。建立绑定后，加载项可以使用提供的标识符访问文档或电子表格的关联区域中包含的数据。创建绑定可为加载项提供以下值：


- 允许访问跨支持的 Office 应用程序的通用数据结构，例如：表、区域或文本（一系列连续字符）。
    
- 允许读/写操作，而不需要用户做出选择。
    
- 在加载项和文档中的数据之间建立关系。绑定会保留在文档中，以后可以进行访问。
    
建立绑定还允许您订阅仅限文档或电子表格的特定区域的数据和选择更改事件。这意味着，加载项只会收到绑定区域内发生的更改的通知，而不是收到整个文档或电子表格内的常规更改的通知。

[Bindings](http://dev.office.com/reference/add-ins/shared/bindings.bindings) 对象公开 [getAllAsync](http://dev.office.com/reference/add-ins/shared/bindings.getallasync) 方法，通过该方法可以访问在文档或电子表格中建立的所有绑定的集合。可使用 [Bindings.getBindingByIdAsync](http://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) 或 [Office.select](http://dev.office.com/reference/add-ins/shared/office.select) 方法按 ID 访问单个绑定。可使用 **Bindings** 对象的以下方法之一建立新绑定和删除现有绑定：[addFromSelectionAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync)、[addFromPromptAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync)、[addFromNamedItemAsync](http://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync) 或 [releaseByIdAsync](http://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync)。

在使用  _addFromSelectionAsync_ 、 **addFromPromptAsync** 或 **addFromNamedItemAsync** 方法创建绑定时，可通过 **bindingType** 参数指定三种不同的绑定类型：



|**绑定类型**|**说明**|**主机应用程序支持**|
|:-----|:-----|:-----|
|文本绑定|绑定到可以文本形式表示的文档区域。|在 Word 中，大多数连续选区都是有效的，而在 Excel 中，只有单个单元格选区才能作为文本绑定的目标。在 Excel 中，只支持纯文本。在 Word 中，支持以下三种格式：纯文本、HTML 和 Open XML for Office。|
|矩阵绑定|绑定到包含没有标题的表格数据的文档的某个固定区域。矩阵绑定中的数据作为二维  **Array** 写入和读取（在 JavaScript 中作为数组的数组实现）。例如，两列中的两行 **string** 值可以作为 ` [['a', 'b'], ['c', 'd']]` 写入或读取，而包含三行的单列可以作为 `[['a'], ['b'], ['c']]` 写入或读取。|在 Excel 中，任何连续的单元格选区都可用于建立矩阵绑定。在 Word 中，只有表格支持矩阵绑定。|
|表格绑定|绑定到包含带标题的表格的文档区域。表格绑定中的数据作为 [TableData](http://dev.office.com/reference/add-ins/shared/tabledata) 对象写入或读取。**TableData** 对象通过 **headers** 和 **rows** 属性公开数据。|任何 Excel 或 Word 表格均可作为表格绑定的基础。建立表格绑定后，用户添加到表格中的每个新行或新列都自动包含在绑定中。 |
使用 **Bindings** 对象的三个“add”方法之一创建绑定后，可以通过相应对象的方法处理绑定的数据和属性：[MatrixBinding](http://dev.office.com/reference/add-ins/shared/binding.matrixbinding)、[TableBinding](http://dev.office.com/reference/add-ins/shared/binding.tablebinding) 或 [TextBinding](http://dev.office.com/reference/add-ins/shared/binding.textbinding)。这三个对象全部继承 [Binding](http://dev.office.com/reference/add-ins/shared/binding.getdataasync) 对象的 [getDataAsync](http://dev.office.com/reference/add-ins/shared/binding.setdataasync) 和 **setDataAsync** 方法，使你能够与绑定的数据交互。

有关演示如何使用绑定执行任务的代码示例，请参阅[绑定到文档或电子表格中的区域](../develop/bind-to-regions-in-a-document-or-spreadsheet.md)。


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>使用 CustomXmlParts 和 CustomXmlPart 对象处理自定义 XML 部件


 **适用于：**Word 的任务窗格加载项

API 的 [CustomXmlParts](http://dev.office.com/reference/add-ins/shared/customxmlparts.customxmlparts) 和 [CustomXmlPart](http://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart) 对象提供访问 Word 文档中自定义 XML 部件的权限，从而启用文档内容的 XML 驱动操作。有关使用 **CustomXmlParts** 和 **CustomXmlPart** 对象的演示，请参阅 [Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts) 代码示例。


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>使用 getFileAsync 方法处理整个文档


 **适用于：**Word 和 PowerPoint 的任务窗格加载项

[Document.getFileAsync](http://dev.office.com/reference/add-ins/shared/document.getfileasync) 方法以及 [File](http://dev.office.com/reference/add-ins/shared/file) 和 [Slice](http://dev.office.com/reference/add-ins/shared/slice) 对象的成员可提供以达 4 MB 的切块（块）形式一次性获取整个 Word 和 PowerPoint 文档文件的功能。有关详细信息，请参阅[操作方法：获取某加载项中文档的全部文件内容](../develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)。


## <a name="mailbox-object"></a>Mailbox 对象


 **适用于：**Outlook 外接程序

Outlook 外接程序主要使用通过 [Mailbox](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) 对象公开的 API 的子集。要访问专用于 Outlook 外接程序的对象和成员（例如 [Item](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) 对象），可以使用 [Context](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) 对象的 **mailbox** 属性访问 **Mailbox** 对象，如下面的代码行所示。




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

另外，Outlook 外接程序可以使用以下对象：


-  **Office** 对象：用于初始化。
    
-  **Context** 对象：用于访问内容和显示语言属性。
    
-  **RoamingSettings** 对象：用于将特定于 Outlook 外接程序的自定义设置保存到安装外接程序的用户邮箱。
    
有关在 Outlook 外接程序中使用 JavaScript 的信息，请参阅 [Outlook 外接程序](../outlook/outlook-add-ins.md)和 [Outlook 外接程序体系结构和功能概述](../outlook/overview.md)。


## <a name="api-support-matrix"></a>API 支持矩阵


此表总结了当使用[受适用于 Office 的 v1.1 JavaScript API 支持的 1.1 外接程序清单架构和功能](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx)指定[您的外接程序支持的 Office 主机应用程序](../develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)时，在各种外接程序类型（内容、任务窗格和 Outlook）之间均受支持的 API 和功能以及可以托管它们的 Office 应用程序。


|||||||||
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
||**主机名**|数据库|工作簿|邮箱|演示文稿|文档|Project|
||**支持的****主机应用程序**|Access Web 应用程序|ExcelExcel Online|OutlookOutlook Web AppOWA for Devices|PowerPointPowerPoint Online|Word|项目|
|**支持的外接程序类型**|内容|Y|Y||Y|||
||任务窗格||Y||Y|Y|Y|
||Outlook|||Y||||
|**支持的 API 功能**|读/写文本||Y||Y|Y|Y（只读）|
||读/写矩阵||Y|||Y||
||读/写表||Y|||Y||
||读/写 HTML|||||Y||
||读/写Office Open XML|||||Y||
||读取任务、资源、视图和字段属性||||||Y|
||选择已更改事件||Y|||Y||
||获取整个文档||||Y|Y||
||绑定和绑定事件|Y（仅限完全和部分绑定）|Y|||Y||
||读/写自定义 Xml 部件|||||Y||
||保存加载项状态数据（设置）|Y（每个主机加载项）|Y（每篇文档）|Y（每个邮箱）|Y（每篇文档）|Y（每篇文档）||
||设置已更改事件|Y|Y||Y|Y||
||获取活动视图模式并查看已更改的事件||||Y|||
||导航到文档中的相应位置||Y||Y|Y||
||使用规则和 RegEx 根据上下文激活|||Y||||
||读取项目属性|||Y||||
||读取用户配置文件|||Y||||
||获取附件|||Y||||
||获取用户标识令牌|||Y||||
||调用 Exchange Web 服务|||Y||||
