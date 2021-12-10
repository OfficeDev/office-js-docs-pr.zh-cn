---
title: Excel JavaScript 预览 API
description: 有关即将推出的 JavaScript Excel的详细信息。
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 32a2f5d355086c51cbf165dd7ed335e96c96647a
ms.sourcegitcommit: ddb1d85186fd6e77d732159430d20eb7395b9a33
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/10/2021
ms.locfileid: "61406639"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript 预览 API

新的 Excel JavaScript API 首先在“预览版”中引入，在进行充分测试并获得用户反馈后，它将成为编号的特定要求集的一部分。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

下表提供了 API 的简要摘要，而后续 [的 API](#api-list) 列表表提供了一个详细列表。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [Data types](../../excel/excel-data-types-overview.md) | 现有数字数据类型Excel扩展，包括对格式化数字和 Web 图像的支持。 | [](/javascript/api/excel/excel.arraycellvalue)ArrayCellValue、BooleanCellValue、CellValueAttributionAttributes、CellValueProviderAttributes、DoubleCellValue、EmptyCellValue、EntityCellValue、FormattedNumberCellValue、StringCellValue、ValueTypeNotAvailableCellValue、WebImageCellValue [](/javascript/api/excel/excel.booleancellvalue) [](/javascript/api/excel/excel.cellvalueattributionattributes) [](/javascript/api/excel/excel.cellvalueproviderattributes) [](/javascript/api/excel/excel.doublecellvalue) [](/javascript/api/excel/excel.emptycellvalue) [](/javascript/api/excel/excel.entitycellvalue) [](/javascript/api/excel/excel.formattednumbercellvalue) [](/javascript/api/excel/excel.stringcellvalue) [](/javascript/api/excel/excel.valuetypenotavailablecellvalue) [](/javascript/api/excel/excel.webimagecellvalue) |
| [数据类型错误](../../excel/excel-data-types-concepts.md#improved-error-support) | 支持扩展数据类型的错误对象。 | BlockedErrorCellValue、BusyErrorCellValue、CalcErrorCellValue、ConnectErrorCellValue、Div0ErrorCellValue、FieldErrorCellValue、GettingDataErrorCellValue、NotAvailableErrorCellValue、NameErrorCellValue、NullErrorCellValue、NumErrorCellValue、RefErrorCellValue、SpillErrorCellValue、ValueErrorCellValue [](/javascript/api/excel/excel.blockederrorcellvalue) [](/javascript/api/excel/excel.busyerrorcellvalue) [](/javascript/api/excel/excel.calcerrorcellvalue) [](/javascript/api/excel/excel.connecterrorcellvalue) [](/javascript/api/excel/excel.div0errorcellvalue) [](/javascript/api/excel/excel.fielderrorcellvalue) [](/javascript/api/excel/excel.gettingdataerrorcellvalue) [](/javascript/api/excel/excel.notavailableerrorcellvalue) [](/javascript/api/excel/excel.nameerrorcellvalue) [](/javascript/api/excel/excel.nullerrorcellvalue) [](/javascript/api/excel/excel.numerrorcellvalue) [](/javascript/api/excel/excel.referrorcellvalue) [](/javascript/api/excel/excel.spillerrorcellvalue) [](/javascript/api/excel/excel.valueerrorcellvalue)|
| 记录任务 | 将注释转换为分配给用户的任务。 | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| 身份 | 管理用户标识，包括显示名称和电子邮件地址。 | [](/javascript/api/excel/excel.identity) [Identity、IdentityCollection、IdentityEntity](/javascript/api/excel/excel.identitycollection) [](/javascript/api/excel/excel.identityentity) |
| 链接的数据类型 | 添加对从外部源连接到Excel类型的支持。 | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype) [、LinkedDataTypeAddedEventArgs、LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypeaddedeventargs) [](/javascript/api/excel/excel.linkeddatatypecollection) |
| 表样式 | 提供对字体、边框、填充颜色以及表格样式的其他方面的控制。 | [表](/javascript/api/excel/excel.table)、[数据透视表](/javascript/api/excel/excel.pivottable)[、切片器](/javascript/api/excel/excel.slicer) |
| 工作表保护 | 防止未经授权的用户对工作表中的指定区域进行更改。 | [](/javascript/api/excel/excel.worksheetprotection)WorksheetProtection、AllowEditRange、AllowEditRangeCollection、AllowEditRangeOptions [](/javascript/api/excel/excel.alloweditrange) [](/javascript/api/excel/excel.alloweditrangecollection) [](/javascript/api/excel/excel.alloweditrangeoptions) |

## <a name="api-list"></a>API 列表

下表列出了当前预览Excel JavaScript API 的列表。 有关所有 JavaScript EXCEL的完整列表 (包括预览 API 和以前发布的 API) ，请参阅所有 Excel [JavaScript API。](/javascript/api/excel?view=excel-js-preview&preserve-view=true)

| 类 | 域 | 说明 |
|:---|:---|:---|
|[AllowEditRange](/javascript/api/excel/excel.alloweditrange)|[address](/javascript/api/excel/excel.alloweditrange#address)|指定与对象关联的区域。|
||[delete()](/javascript/api/excel/excel.alloweditrange#delete__)|从 中删除此对象 `AllowEditRangeCollection` 。|
||[isPasswordProtected](/javascript/api/excel/excel.alloweditrange#isPasswordProtected)|指定 是否 `AllowEditRange` 受密码保护。|
||[pauseProtection (password？： string) ](/javascript/api/excel/excel.alloweditrange#pauseProtection_password_)|暂停给定会话中用户 `AllowEditRange` 给定对象的工作表保护。|
||[setPassword (password？： string) ](/javascript/api/excel/excel.alloweditrange#setPassword_password_)|更改与 关联的密码 `AllowEditRange` 。|
||[title](/javascript/api/excel/excel.alloweditrange#title)|指定对象的标题。|
|[AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)|[add (title： string， rangeAddress： string， options？： Excel.AllowEditRangeOptions) ](/javascript/api/excel/excel.alloweditrangecollection#add_title__rangeAddress__options_)|向 `AllowEditRange` 集合添加对象。|
||[getCount()](/javascript/api/excel/excel.alloweditrangecollection#getCount__)|返回集合 `AllowEditRange` 中对象的数量。|
||[getItem(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItem_key_)|按 `AllowEditRange` 对象的标题获取对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.alloweditrangecollection#getItemAt_index_)|按 `AllowEditRange` 对象在集合中的索引返回对象。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItemOrNullObject_key_)|按 `AllowEditRange` 对象的标题获取对象。|
||[items](/javascript/api/excel/excel.alloweditrangecollection#items)|获取此集合中已加载的子项。|
||[pauseProtection (password： string) ](/javascript/api/excel/excel.alloweditrangecollection#pauseProtection_password_)|暂停对集合中具有给定会话中用户给定密码 `AllowEditRange` 的所有对象的工作表保护。|
|[AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions)|[password](/javascript/api/excel/excel.alloweditrangeoptions#password)|与 关联的密码 `AllowEditRange` 。|
|[ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)|[basicType](/javascript/api/excel/excel.arraycellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.arraycellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[elements](/javascript/api/excel/excel.arraycellvalue#elements)|表示数组的元素。|
||[type](/javascript/api/excel/excel.arraycellvalue#type)|表示此单元格值的类型。|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[basicType](/javascript/api/excel/excel.blockederrorcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.blockederrorcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#errorSubType)|表示 的类型 `BlockedErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.blockederrorcellvalue#errorType)|表示 的类型 `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.blockederrorcellvalue#type)|表示此单元格值的类型。|
|[BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)|[basicType](/javascript/api/excel/excel.booleancellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.booleancellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[type](/javascript/api/excel/excel.booleancellvalue#type)|表示此单元格值的类型。|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[basicType](/javascript/api/excel/excel.busyerrorcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.busyerrorcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#errorSubType)|表示 的类型 `BusyErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.busyerrorcellvalue#errorType)|表示 的类型 `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.busyerrorcellvalue#type)|表示此单元格值的类型。|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[basicType](/javascript/api/excel/excel.calcerrorcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.calcerrorcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#errorSubType)|表示 的类型 `CalcErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.calcerrorcellvalue#errorType)|表示 的类型 `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.calcerrorcellvalue#type)|表示此单元格值的类型。|
|[CardLayoutListSection](/javascript/api/excel/excel.cardlayoutlistsection)|[layout](/javascript/api/excel/excel.cardlayoutlistsection#layout)|表示此节的布局类型。|
|[CardLayoutPropertyReference](/javascript/api/excel/excel.cardlayoutpropertyreference)|[property](/javascript/api/excel/excel.cardlayoutpropertyreference#property)|卡片布局所引用的属性的名称。|
|[CardLayoutSectionStandardProperties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties)|[collapsed](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#collapsed)|表示卡片的此部分最初是否折叠。|
||[可折叠](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#collapsible)|表示卡片的此部分是否可折叠。|
||[properties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#properties)|表示此部分中属性的名称。|
||[title](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#title)|表示卡片的此部分的标题。|
|[CardLayoutStandardProperties](/javascript/api/excel/excel.cardlayoutstandardproperties)|[mainImage](/javascript/api/excel/excel.cardlayoutstandardproperties#mainImage)|指定将用作卡片主图像的属性。|
||[sections](/javascript/api/excel/excel.cardlayoutstandardproperties#sections)|表示卡片的各个部分。|
||[subTitle](/javascript/api/excel/excel.cardlayoutstandardproperties#subTitle)|表示包含卡片副标题的属性的规范。|
||[title](/javascript/api/excel/excel.cardlayoutstandardproperties#title)|表示卡片的标题或包含卡片标题的属性的规范。|
|[CardLayoutTableSection](/javascript/api/excel/excel.cardlayouttablesection)|[layout](/javascript/api/excel/excel.cardlayouttablesection#layout)|表示此节的布局类型。|
|[CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)|[licenseAddress](/javascript/api/excel/excel.cellvalueattributionattributes#licenseAddress)|表示指向描述如何使用此属性的许可证或源的 URL。|
||[licenseText](/javascript/api/excel/excel.cellvalueattributionattributes#licenseText)|表示管理此属性的许可证的名称。|
||[sourceAddress](/javascript/api/excel/excel.cellvalueattributionattributes#sourceAddress)|表示指向 的源的 `CellValue` URL。|
||[sourceText](/javascript/api/excel/excel.cellvalueattributionattributes#sourceText)|表示 的源的名称 `CellValue` 。|
|[CellValuePropertyMetadata](/javascript/api/excel/excel.cellvaluepropertymetadata)|[attribution](/javascript/api/excel/excel.cellvaluepropertymetadata#attribution)|表示用于描述使用此属性的来源和许可证要求的属性信息。|
||[excludeFrom](/javascript/api/excel/excel.cellvaluepropertymetadata#excludeFrom)|表示从中排除此属性的功能。|
||[sublabel](/javascript/api/excel/excel.cellvaluepropertymetadata#sublabel)|表示卡片视图中显示的此属性的子标签。|
|[CellValuePropertyMetadataExclusions](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions)|[autoComplete](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#autoComplete)|True 表示属性从自动完成显示的属性中排除。|
||[calcCompare](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#calcCompare)|如果为 True，则从用于重新计算期间比较单元格值的属性中排除该属性。|
||[cardView](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#cardView)|True 表示属性从卡片视图显示的属性中排除。|
||[dotNotation](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#dotNotation)|True 表示属性被从可通过 FIELDVALUE 函数访问的属性中排除。|
|[CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)|[description](/javascript/api/excel/excel.cellvalueproviderattributes#description)|表示在未指定徽标时在卡片视图中使用的提供程序说明属性。|
||[logoSourceAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoSourceAddress)|表示用于下载将在卡片视图中用作徽标的图像的 URL。|
||[logoTargetAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoTargetAddress)|表示一个 URL，如果用户单击卡片视图中的徽标元素，该 URL 即为导航目标。|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask (：Identity) ](/javascript/api/excel/excel.comment#assignTask_assignee_)|将附加到注释的任务作为委派者分配给给定用户。|
||[getTask () ](/javascript/api/excel/excel.comment#getTask__)|获取与此注释关联的任务。|
||[getTaskOrNullObject () ](/javascript/api/excel/excel.comment#getTaskOrNullObject__)|获取与此注释关联的任务。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask (：Identity) ](/javascript/api/excel/excel.commentreply#assignTask_assignee_)|将附加到注释的任务分配给指定用户作为唯一的代理人。|
||[getTask () ](/javascript/api/excel/excel.commentreply#getTask__)|获取与此批注回复线程相关联的任务。|
||[getTaskOrNullObject () ](/javascript/api/excel/excel.commentreply#getTaskOrNullObject__)|获取与此批注回复线程相关联的任务。|
|[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)|[basicType](/javascript/api/excel/excel.connecterrorcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.connecterrorcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.connecterrorcellvalue#errorSubType)|表示 的类型 `ConnectErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.connecterrorcellvalue#errorType)|表示 的类型 `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.connecterrorcellvalue#type)|表示此单元格值的类型。|
|[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)|[basicType](/javascript/api/excel/excel.div0errorcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.div0errorcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[errorType](/javascript/api/excel/excel.div0errorcellvalue#errorType)|表示 的类型 `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.div0errorcellvalue#type)|表示此单元格值的类型。|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[assignees](/javascript/api/excel/excel.documenttask#assignees)|返回任务的被分配者的集合。|
||[更改](/javascript/api/excel/excel.documenttask#changes)|获取任务的更改记录。|
||[comment](/javascript/api/excel/excel.documenttask#comment)|获取与任务关联的注释。|
||[completedBy](/javascript/api/excel/excel.documenttask#completedBy)|获取完成任务的最新用户。|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completedDateTime)|获取任务的完成日期和时间。|
||[createdBy](/javascript/api/excel/excel.documenttask#createdBy)|获取创建任务的用户。|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createdDateTime)|获取任务的创建日期和时间。|
||[id](/javascript/api/excel/excel.documenttask#id)|获取任务的 ID。|
||[percentComplete](/javascript/api/excel/excel.documenttask#percentComplete)|指定任务的完成百分比。|
||[priority](/javascript/api/excel/excel.documenttask#priority)|指定任务的优先级。|
||[setStartAndDueDateTime (startDateTime： Date， dueDateTime： Date) ](/javascript/api/excel/excel.documenttask#setStartAndDueDateTime_startDateTime__dueDateTime_)|更改任务的开始日期和截止日期。|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startAndDueDateTime)|获取或设置任务应开始和到期的日期和时间。|
||[title](/javascript/api/excel/excel.documenttask#title)|指定任务的标题。|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[被分派人](/javascript/api/excel/excel.documenttaskchange#assignee)|表示分配给更改记录类型的任务的用户，或者从更改记录类型的任务 `assign` 中取消 `unassign` 分配的用户。|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedBy)|表示创建或更改任务的用户。|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentId)|表示 任务更改锁定的 或 `Comment` `CommentReply` 的 ID。|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createdDateTime)|表示任务更改记录的创建日期和时间。|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#dueDateTime)|表示任务的截止日期和时间，以 UTC 时区表示。|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|任务更改记录的 ID。|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#percentComplete)|表示任务的完成百分比。|
||[priority](/javascript/api/excel/excel.documenttaskchange#priority)|表示任务的优先级。|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startDateTime)|表示任务的开始日期和时间，以 UTC 时区表示。|
||[title](/javascript/api/excel/excel.documenttaskchange#title)|表示任务的标题。|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|表示任务更改记录的操作类型。|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undoHistoryId)|表示 `DocumentTaskChange.id` 对更改记录类型撤消 `undo` 的属性。|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getCount__)|获取任务集合中的更改记录数。|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getItemAt_index_)|使用任务更改记录在集合中的索引获取该记录。|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|获取此集合中已加载的子项。|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getCount__)|获取集合中的任务数。|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getItem_key_)|使用其 ID 获取任务。|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getItemAt_index_)|按任务在集合中的索引获取任务。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getItemOrNullObject_key_)|使用其 ID 获取任务。|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|获取此集合中已加载的子项。|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#dueDateTime)|获取任务到期的日期和时间。|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startDateTime)|获取任务应开始的日期和时间。|
|[DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)|[basicType](/javascript/api/excel/excel.doublecellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.doublecellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[type](/javascript/api/excel/excel.doublecellvalue#type)|表示此单元格值的类型。|
|[EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)|[basicType](/javascript/api/excel/excel.emptycellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.emptycellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[type](/javascript/api/excel/excel.emptycellvalue#type)|表示此单元格值的类型。|
|[EntityCardLayout](/javascript/api/excel/excel.entitycardlayout)|[layout](/javascript/api/excel/excel.entitycardlayout#layout)|表示此布局的类型。|
|[EntityCellValue](/javascript/api/excel/excel.entitycellvalue)|[basicType](/javascript/api/excel/excel.entitycellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.entitycellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[cardLayout](/javascript/api/excel/excel.entitycellvalue#cardLayout)|表示卡片视图中此实体的布局。|
||[properties： { [key： string]](/javascript/api/excel/excel.entitycellvalue#properties)|表示此实体的属性及其元数据。|
||[text](/javascript/api/excel/excel.entitycellvalue#text)|表示呈现具有此值的单元格时显示的文本。|
||[type](/javascript/api/excel/excel.entitycellvalue#type)|表示此单元格值的类型。|
|[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)|[basicType](/javascript/api/excel/excel.fielderrorcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.fielderrorcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.fielderrorcellvalue#errorSubType)|表示 的类型 `FieldErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.fielderrorcellvalue#errorType)|表示 的类型 `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.fielderrorcellvalue#type)|表示此单元格值的类型。|
|[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)|[basicType](/javascript/api/excel/excel.formattednumbercellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.formattednumbercellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[numberFormat](/javascript/api/excel/excel.formattednumbercellvalue#numberFormat)|返回用于显示此值的数值格式字符串。|
||[type](/javascript/api/excel/excel.formattednumbercellvalue#type)|表示此单元格值的类型。|
|[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)|[basicType](/javascript/api/excel/excel.gettingdataerrorcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.gettingdataerrorcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[errorType](/javascript/api/excel/excel.gettingdataerrorcellvalue#errorType)|表示 的类型 `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.gettingdataerrorcellvalue#type)|表示此单元格值的类型。|
|[标识](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayName)|表示用户的显示名称。|
||[email](/javascript/api/excel/excel.identity#email)|表示用户的电子邮件地址。|
||[id](/javascript/api/excel/excel.identity#id)|表示用户的唯一 ID。|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[添加 (：标识) ](/javascript/api/excel/excel.identitycollection#add_assignee_)|向集合添加用户标识。|
||[clear()](/javascript/api/excel/excel.identitycollection#clear__)|从集合中删除所有的用户标识。|
||[getCount()](/javascript/api/excel/excel.identitycollection#getCount__)|获取集合中项的数目。|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getItemAt_index_)|使用文档在集合中的索引获取文档用户标识。|
||[items](/javascript/api/excel/excel.identitycollection#items)|获取此集合中已加载的子项。|
||[remove (assignee： Identity) ](/javascript/api/excel/excel.identitycollection#remove_assignee_)|从集合中删除用户标识。|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#displayName)|表示用户的显示名称。|
||[email](/javascript/api/excel/excel.identityentity#email)|表示用户的电子邮件地址。|
||[id](/javascript/api/excel/excel.identityentity#id)|表示用户的唯一 ID。|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataProvider)|链接数据提供程序的数据提供程序数据类型。|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastRefreshed)|自上次刷新链接工作簿时打开工作簿以来的本地数据类型日期和时间。|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|链接对象数据类型。|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicRefreshInterval)|链接对象刷新的频率（以秒数据类型设置为 `refreshMode` "Periodic"时刷新。|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshMode)|用于检索链接数据数据类型的机制。|
||[requestRefresh () ](/javascript/api/excel/excel.linkeddatatype#requestRefresh__)|请求刷新链接数据类型。|
||[requestSetRefreshMode (refreshMode： Excel。LinkedDataTypeRefreshMode) ](/javascript/api/excel/excel.linkeddatatype#requestSetRefreshMode_refreshMode_)|请求更改此链接的刷新数据类型。|
||[服务 Id](/javascript/api/excel/excel.linkeddatatype#serviceId)|链接对象的唯一数据类型。|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedRefreshModes)|返回一个数组，该数组包含链接对象支持的所有数据类型。|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[服务 Id](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceId)|新链接对象的唯一 ID 数据类型。|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|获取事件的类型。|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getCount__)|获取集合中链接的数据类型的数量。|
||[getItem (键：number) ](/javascript/api/excel/excel.linkeddatatypecollection#getItem_key_)|按服务 ID 数据类型链接的标识符。|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemAt_index_)|按集合数据类型索引获取链接对象。|
||[getItemOrNullObject (键：number) ](/javascript/api/excel/excel.linkeddatatypecollection#getItemOrNullObject_key_)|按 ID 数据类型链接的标识符。|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|获取此集合中已加载的子项。|
||[requestRefreshAll () ](/javascript/api/excel/excel.linkeddatatypecollection#requestRefreshAll__)|请求刷新集合中所有链接的数据类型。|
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[basicType](/javascript/api/excel/excel.nameerrorcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.nameerrorcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[errorType](/javascript/api/excel/excel.nameerrorcellvalue#errorType)|表示 的类型 `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.nameerrorcellvalue#type)|表示此单元格值的类型。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[valueAsJson](/javascript/api/excel/excel.nameditem#valueAsJson)|此已命名项中值的 JSON 表示形式。|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[valuesAsJson](/javascript/api/excel/excel.nameditemarrayvalues#valuesAsJson)|此区域单元格中的值的 JSON 表示形式。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItemOrNullObject_key_)|使用工作表视图的名称获取工作表视图。|
|[NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)|[basicType](/javascript/api/excel/excel.notavailableerrorcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.notavailableerrorcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[errorType](/javascript/api/excel/excel.notavailableerrorcellvalue#errorType)|表示 的类型 `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.notavailableerrorcellvalue#type)|表示此单元格值的类型。|
|[NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)|[basicType](/javascript/api/excel/excel.nullerrorcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.nullerrorcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[errorType](/javascript/api/excel/excel.nullerrorcellvalue#errorType)|表示 的类型 `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.nullerrorcellvalue#type)|表示此单元格值的类型。|
|[NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)|[basicType](/javascript/api/excel/excel.numerrorcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.numerrorcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[errorType](/javascript/api/excel/excel.numerrorcellvalue#errorType)|表示 的类型 `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.numerrorcellvalue#type)|表示此单元格值的类型。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getCell_dataHierarchy__rowItems__columnItems_)|根据数据层次结构以及各自层次结构的行和列项，获取数据透视表中的唯一单元格。|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotStyle)|应用于数据透视表的样式。|
||[setStyle (样式：string \| PivotTableStyle \| BuiltInPivotTableStyle) ](/javascript/api/excel/excel.pivotlayout#setStyle_style_)|设置应用于数据透视表的样式。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[getDataSourceString () ](/javascript/api/excel/excel.pivottable#getDataSourceString__)|返回数据透视表数据源的字符串表示形式。|
||[getDataSourceType () ](/javascript/api/excel/excel.pivottable#getDataSourceType__)|获取数据透视表的数据源的类型。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject () ](/javascript/api/excel/excel.pivottablescopedcollection#getFirstOrNullObject__)|获取集合中的第一个数据透视表。|
|[区域](/javascript/api/excel/excel.range)|[getDependents () ](/javascript/api/excel/excel.range#getDependents__)|返回一个对象，该对象表示包含同一工作表或多个工作表中单元格的所有从属 `WorkbookRangeAreas` 单元格的范围。|
||[valuesAsJson](/javascript/api/excel/excel.range#valuesAsJson)|此区域单元格中的值的 JSON 表示形式。|
|[RangeView](/javascript/api/excel/excel.rangeview)|[valuesAsJson](/javascript/api/excel/excel.rangeview#valuesAsJson)|此区域单元格中的值的 JSON 表示形式。|
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[basicType](/javascript/api/excel/excel.referrorcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.referrorcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.referrorcellvalue#errorSubType)|表示 的类型 `RefErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.referrorcellvalue#errorType)|表示 的类型 `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.referrorcellvalue#type)|表示此单元格值的类型。|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshMode)|链接的数据类型刷新模式。|
||[服务 Id](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceId)|刷新模式已更改的对象的唯一 ID。|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|获取事件的类型。|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[已刷新](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|指示刷新请求是否成功。|
||[服务 Id](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceId)|已完成刷新请求的对象的唯一 ID。|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|获取事件源。|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|获取事件的类型。|
||[warnings](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|包含从刷新请求生成的任何警告的数组。|
|[Shape](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#displayName)|获取显示名称的大小。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addSvg_xml_)|从 XML 字符串创建可缩放的矢量图形 (SVG) 并将其添加到工作表。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameInFormula)|表示公式中使用切片器名称。|
||[setStyle (样式：string \| SlicerStyle \| BuiltInSlicerStyle) ](/javascript/api/excel/excel.slicer#setStyle_style_)|设置应用于切片器的样式。|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerStyle)|应用于切片器的样式。|
|[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)|[basicType](/javascript/api/excel/excel.spillerrorcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.spillerrorcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.spillerrorcellvalue#errorSubType)|表示 的类型 `SpillErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.spillerrorcellvalue#errorType)|表示 的类型 `ErrorCellValue` 。|
||[spilledColumns](/javascript/api/excel/excel.spillerrorcellvalue#spilledColumns)|表示如果没有数据，将溢出的#SPILL！ error。|
||[spilledRows](/javascript/api/excel/excel.spillerrorcellvalue#spilledRows)|表示如果没有数据，则溢出的#SPILL！ error。|
||[type](/javascript/api/excel/excel.spillerrorcellvalue#type)|表示此单元格值的类型。|
|[StringCellValue](/javascript/api/excel/excel.stringcellvalue)|[basicType](/javascript/api/excel/excel.stringcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.stringcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[type](/javascript/api/excel/excel.stringcellvalue#type)|表示此单元格值的类型。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearStyle__)|将表格更改为使用默认表格样式。|
||[onFiltered](/javascript/api/excel/excel.table#onFiltered)|在将筛选器应用于特定表时发生。|
||[setStyle (样式：string \| TableStyle \| BuiltInTableStyle) ](/javascript/api/excel/excel.table#setStyle_style_)|设置应用于表格的样式。|
||[tableStyle](/javascript/api/excel/excel.table#tableStyle)|应用于表格的样式。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onFiltered)|在工作簿或工作表的任何表上应用筛选器时发生。|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[valuesAsJson](/javascript/api/excel/excel.tablecolumn#valuesAsJson)|此表列中单元格中的值的 JSON 表示形式。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableId)|获取应用筛选器的表的 ID。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetId)|获取包含表格的工作表的 ID。|
|[TableRow](/javascript/api/excel/excel.tablerow)|[valuesAsJson](/javascript/api/excel/excel.tablerow#valuesAsJson)|此表行中单元格中的值的 JSON 表示形式。|
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[basicType](/javascript/api/excel/excel.valueerrorcellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.valueerrorcellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#errorSubType)|表示 的类型 `ValueErrorCellValue` 。|
||[errorType](/javascript/api/excel/excel.valueerrorcellvalue#errorType)|表示 的类型 `ErrorCellValue` 。|
||[type](/javascript/api/excel/excel.valueerrorcellvalue#type)|表示此单元格值的类型。|
|[ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)|[basicType](/javascript/api/excel/excel.valuetypenotavailablecellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[type](/javascript/api/excel/excel.valuetypenotavailablecellvalue#type)|表示此单元格值的类型。|
|[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)|[address](/javascript/api/excel/excel.webimagecellvalue#address)|表示将下载映像的 URL。|
||[altText](/javascript/api/excel/excel.webimagecellvalue#altText)|表示可在辅助功能方案中用来描述图像表示的内容的备用文本。|
||[attribution](/javascript/api/excel/excel.webimagecellvalue#attribution)|表示用于描述使用此映像的源和许可证要求的属性信息。|
||[basicType](/javascript/api/excel/excel.webimagecellvalue#basicType)|表示由具有此值 `Range.valueTypes` 的单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.webimagecellvalue#basicValue)|表示由具有此值 `Range.values` 的单元格返回的值。|
||[提供程序](/javascript/api/excel/excel.webimagecellvalue#provider)|表示描述提供图像的实体或个人的信息。|
||[relatedImagesAddress](/javascript/api/excel/excel.webimagecellvalue#relatedImagesAddress)|表示包含图像被视为与此 相关的网页的 `WebImageCellValue` URL。|
||[type](/javascript/api/excel/excel.webimagecellvalue#type)|表示此单元格值的类型。|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkedDataTypes)|返回属于工作簿的链接数据类型的集合。|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showPivotFieldList)|指定是否在工作簿级别显示数据透视表的字段列表窗格。|
||[任务](/javascript/api/excel/excel.workbook#tasks)|返回工作簿中的任务集合。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904DateSystem)|如果工作簿使用 1904 日期系统，则为 True。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onFiltered)|在将筛选器应用于特定工作表时发生。|
||[任务](/javascript/api/excel/excel.worksheet#tasks)|返回工作表中的任务集合。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addFromBase64_base64File__sheetNamesToInsert__positionType__relativeTo_)|将工作簿的指定工作表插入当前工作簿。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onFiltered)|在工作簿中应用任何工作表的筛选器时发生。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetId)|获取应用筛选器的工作表的 ID。|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[allowEditRanges](/javascript/api/excel/excel.worksheetprotection#allowEditRanges)|指定 `AllowEditRangeCollection` 在此工作表中找到的 。|
||[canPauseProtection](/javascript/api/excel/excel.worksheetprotection#canPauseProtection)|指定是否可暂停此工作表的保护。|
||[checkPassword (password？： string) ](/javascript/api/excel/excel.worksheetprotection#checkPassword_password_)|指定密码是否可用于解锁工作表保护。|
||[isPasswordProtected](/javascript/api/excel/excel.worksheetprotection#isPasswordProtected)|指定工作表是否受密码保护。|
||[isPaused](/javascript/api/excel/excel.worksheetprotection#isPaused)|指定是否暂停工作表保护。|
||[pauseProtection (password？： string) ](/javascript/api/excel/excel.worksheetprotection#pauseProtection_password_)|暂停给定会话中用户的给定工作表对象的工作表保护。|
||[resumeProtection () ](/javascript/api/excel/excel.worksheetprotection#resumeProtection__)|为给定会话中的用户恢复给定工作表对象的工作表保护。|
||[setPassword (password？： string) ](/javascript/api/excel/excel.worksheetprotection#setPassword_password_)|更改与对象关联的 `WorksheetProtection` 密码。|
||[updateOptions (选项：Excel。WorksheetProtectionOptions) ](/javascript/api/excel/excel.worksheetprotection#updateOptions_options_)|更改与对象关联的工作表保护 `WorksheetProtection` 选项。|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#allowEditRangesChanged)|指定是否更改了 `AllowEditRange` 任何对象。|
||[protectionOptionsChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#protectionOptionsChanged)|指定 是否 `WorksheetProtectionOptions` 已更改。|
||[sheetPasswordChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#sheetPasswordChanged)|指定工作表密码是否已更改。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
