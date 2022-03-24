---
title: Excel JavaScript 预览 API
description: 有关即将推出的 JavaScript Excel的详细信息。
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: f15a72631f83a5102fb4e042cc1357d179d1fa3d
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63747180"
---
# <a name="excel-javascript-preview-apis"></a>Excel JavaScript 预览 API

新的 Excel JavaScript API 首先在“预览版”中引入，在进行充分测试并获得用户反馈后，它将成为编号的特定要求集的一部分。

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

下表提供了 API 的简要摘要，而后续 [的 API](#api-list) 列表表提供了一个详细列表。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| [Data types](../../excel/excel-data-types-overview.md) | 现有数字数据类型Excel扩展，包括对格式化数字和 Web 图像的支持。 | [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)、 [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)、 [CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)、 [CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)、 [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)、 [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)、 [EntityCellValue](/javascript/api/excel/excel.entitycellvalue)、 [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)、 [StringCellValue](/javascript/api/excel/excel.stringcellvalue)、 [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)、 [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) |
| [数据类型错误](../../excel/excel-data-types-concepts.md#improved-error-support) | 支持扩展数据类型的错误对象。 | [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)、 [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)、 [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)、 [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)、 [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)、 [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)、 [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)、 [NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)、 [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)、 [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)、 [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)、 [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)、 [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)、 [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|
| 记录任务 | 将注释转换为分配给用户的任务。 | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| 身份 | 管理用户标识，包括显示名称电子邮件地址。 | [Identity](/javascript/api/excel/excel.identity)、 [IdentityCollection](/javascript/api/excel/excel.identitycollection)、 [IdentityEntity](/javascript/api/excel/excel.identityentity) |
| 链接的数据类型 | 添加对从外部源连接到Excel类型的支持。 | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)、 [LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)、 [LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection) |
| 表样式 | 提供对字体、边框、填充颜色以及表格样式的其他方面的控制。 | [表](/javascript/api/excel/excel.table)、[数据透视表](/javascript/api/excel/excel.pivottable)[、切片器](/javascript/api/excel/excel.slicer) |
| 工作表保护 | 防止未经授权的用户对工作表中的指定区域进行更改。 | [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)、 [AllowEditRange](/javascript/api/excel/excel.alloweditrange)、 [AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)、 [AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions) |

## <a name="api-list"></a>API 列表

下表列出了当前预览Excel JavaScript API 的列表。 有关所有 JavaScript API Excel列表 (包括预览 API 和以前发布的 API) ，请参阅所有 Excel [JavaScript API](/javascript/api/excel?view=excel-js-preview&preserve-view=true)。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[AllowEditRange](/javascript/api/excel/excel.alloweditrange)|[address](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-address-member)|指定与对象关联的区域。|
||[delete()](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-delete-member(1))|从 中删除此对象 `AllowEditRangeCollection`。|
||[isPasswordProtected](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-ispasswordprotected-member)|指定 是否 `AllowEditRange` 受密码保护。|
||[pauseProtection (password？： string) ](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-pauseprotection-member(1))|暂停给定会话中用户给定 `AllowEditRange` 对象的工作表保护。|
||[setPassword (password？： string) ](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-setpassword-member(1))|更改与 关联的密码 `AllowEditRange`。|
||[title](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-title-member)|指定对象的标题。|
|[AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)|[add (title： string， rangeAddress： string， options？： Excel.AllowEditRangeOptions) ](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-add-member(1))|`AllowEditRange`向集合添加对象。|
||[getCount()](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getcount-member(1))|返回集合中 `AllowEditRange` 对象的数量。|
||[getItem(key: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitem-member(1))|按对象 `AllowEditRange` 的标题获取对象。|
||[getItemAt(index: number)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitemat-member(1))|按对象 `AllowEditRange` 在集合中的索引返回对象。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitemornullobject-member(1))|按对象 `AllowEditRange` 的标题获取对象。|
||[items](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-items-member)|获取此集合中已加载的子项。|
||[pauseProtection (password： string) ](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-pauseprotection-member(1))|暂停对集合中具有 `AllowEditRange` 给定会话中用户给定密码的所有对象的工作表保护。|
|[AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions)|[password](/javascript/api/excel/excel.alloweditrangeoptions#excel-excel-alloweditrangeoptions-password-member)|与 关联的密码 `AllowEditRange`。|
|[ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)|[basicType](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[elements](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-elements-member)|表示数组的元素。|
||[type](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-type-member)|表示此单元格值的类型。|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[basicType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-errorsubtype-member)|表示 的类型 `BlockedErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-type-member)|表示此单元格值的类型。|
|[BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)|[basicType](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[type](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-type-member)|表示此单元格值的类型。|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[basicType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-errorsubtype-member)|表示 的类型 `BusyErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-type-member)|表示此单元格值的类型。|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[basicType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-errorsubtype-member)|表示 的类型 `CalcErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-type-member)|表示此单元格值的类型。|
|[CardLayoutListSection](/javascript/api/excel/excel.cardlayoutlistsection)|[layout](/javascript/api/excel/excel.cardlayoutlistsection#excel-excel-cardlayoutlistsection-layout-member)|表示此节的布局类型。|
|[CardLayoutPropertyReference](/javascript/api/excel/excel.cardlayoutpropertyreference)|[property](/javascript/api/excel/excel.cardlayoutpropertyreference#excel-excel-cardlayoutpropertyreference-property-member)|卡片布局所引用的属性的名称。|
|[CardLayoutSectionStandardProperties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties)|[collapsed](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-collapsed-member)|表示卡片的此部分最初是否折叠。|
||[可折叠](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-collapsible-member)|表示卡片的此部分是否可折叠。|
||[properties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-properties-member)|表示此部分中属性的名称。|
||[title](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-title-member)|表示卡片的此部分的标题。|
|[CardLayoutStandardProperties](/javascript/api/excel/excel.cardlayoutstandardproperties)|[mainImage](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-mainimage-member)|指定将用作卡片主图像的属性。|
||[sections](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-sections-member)|表示卡片的各个部分。|
||[subTitle](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-subtitle-member)|表示包含卡片副标题的属性的规范。|
||[title](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-title-member)|表示卡片的标题或包含卡片标题的属性的规范。|
|[CardLayoutTableSection](/javascript/api/excel/excel.cardlayouttablesection)|[layout](/javascript/api/excel/excel.cardlayouttablesection#excel-excel-cardlayouttablesection-layout-member)|表示此节的布局类型。|
|[CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)|[licenseAddress](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-licenseaddress-member)|表示指向描述如何使用此属性的许可证或源的 URL。|
||[licenseText](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-licensetext-member)|表示管理此属性的许可证的名称。|
||[sourceAddress](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-sourceaddress-member)|表示指向 的源的 `CellValue`URL。|
||[sourceText](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-sourcetext-member)|表示 的源的名称 `CellValue`。|
|[CellValuePropertyMetadata](/javascript/api/excel/excel.cellvaluepropertymetadata)|[attribution](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-attribution-member)|表示用于描述使用此属性的来源和许可证要求的属性信息。|
||[excludeFrom](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-excludefrom-member)|表示从中排除此属性的功能。|
||[sublabel](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-sublabel-member)|表示卡片视图中显示的此属性的子标签。|
|[CellValuePropertyMetadataExclusions](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions)|[autoComplete](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-autocomplete-member)|True 表示属性从自动完成显示的属性中排除。|
||[calcCompare](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-calccompare-member)|如果为 True，则从用于重新计算期间比较单元格值的属性中排除该属性。|
||[cardView](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-cardview-member)|True 表示属性从卡片视图显示的属性中排除。|
||[dotNotation](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-dotnotation-member)|True 表示属性被从可通过 FIELDVALUE 函数访问的属性中排除。|
|[CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)|[说明](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-description-member)|表示在未指定徽标时在卡片视图中使用的提供程序说明属性。|
||[logoSourceAddress](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-logosourceaddress-member)|表示用于下载将在卡片视图中用作徽标的图像的 URL。|
||[logoTargetAddress](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-logotargetaddress-member)|表示一个 URL，如果用户单击卡片视图中的徽标元素，该 URL 即为导航目标。|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask (：Identity) ](/javascript/api/excel/excel.comment#excel-excel-comment-assigntask-member(1))|将附加到注释的任务作为委派者分配给给定用户。|
||[getTask () ](/javascript/api/excel/excel.comment#excel-excel-comment-gettask-member(1))|获取与此注释关联的任务。|
||[getTaskOrNullObject () ](/javascript/api/excel/excel.comment#excel-excel-comment-gettaskornullobject-member(1))|获取与此注释关联的任务。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask (：Identity) ](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-assigntask-member(1))|将附加到注释的任务分配给指定用户作为唯一的代理人。|
||[getTask () ](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettask-member(1))|获取与此批注回复线程相关联的任务。|
||[getTaskOrNullObject () ](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettaskornullobject-member(1))|获取与此批注回复线程相关联的任务。|
|[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)|[basicType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-errorsubtype-member)|表示 的类型 `ConnectErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-type-member)|表示此单元格值的类型。|
|[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)|[basicType](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorType](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-type-member)|表示此单元格值的类型。|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[assignees](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-assignees-member)|返回任务的被分配者的集合。|
||[更改](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-changes-member)|获取任务的更改记录。|
||[comment](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-comment-member)|获取与任务关联的注释。|
||[completedBy](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-completedby-member)|获取完成任务的最新用户。|
||[completedDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-completeddatetime-member)|获取任务的完成日期和时间。|
||[createdBy](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-createdby-member)|获取创建任务的用户。|
||[createdDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-createddatetime-member)|获取任务的创建日期和时间。|
||[id](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-id-member)|获取任务的 ID。|
||[percentComplete](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-percentcomplete-member)|指定任务的完成百分比。|
||[priority](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-priority-member)|指定任务的优先级。|
||[setStartAndDueDateTime (startDateTime： Date， dueDateTime： Date) ](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-setstartandduedatetime-member(1))|更改任务的开始日期和截止日期。|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-startandduedatetime-member)|获取或设置任务应开始和到期的日期和时间。|
||[title](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-title-member)|指定任务的标题。|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[被分派人](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-assignee-member)|表示分配给更改记录 `assign` 类型的任务的用户，或者从更改记录类型的任务中 `unassign` 取消分配的用户。|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-changedby-member)|表示创建或更改任务的用户。|
||[commentId](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-commentid-member)|表示 任务 `Comment` 更改锁定 `CommentReply` 的 或 的 ID。|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-createddatetime-member)|表示任务更改记录的创建日期和时间。|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-duedatetime-member)|表示任务的截止日期和时间，以 UTC 时区表示。|
||[id](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-id-member)|任务更改记录的 ID。|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-percentcomplete-member)|表示任务的完成百分比。|
||[priority](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-priority-member)|表示任务的优先级。|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-startdatetime-member)|表示任务的开始日期和时间，以 UTC 时区表示。|
||[title](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-title-member)|表示任务的标题。|
||[type](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-type-member)|表示任务更改记录的操作类型。|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-undohistoryid-member)|表示 `DocumentTaskChange.id` 对更改记录类型撤消 `undo` 的属性。|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-getcount-member(1))|获取任务集合中的更改记录数。|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-getitemat-member(1))|使用任务更改记录在集合中的索引获取该记录。|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-items-member)|获取此集合中已加载的子项。|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getcount-member(1))|获取集合中的任务数。|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitem-member(1))|使用其 ID 获取任务。|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitemat-member(1))|按任务在集合中的索引获取任务。|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitemornullobject-member(1))|使用其 ID 获取任务。|
||[items](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-items-member)|获取此集合中已加载的子项。|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#excel-excel-documenttaskschedule-duedatetime-member)|获取任务到期的日期和时间。|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#excel-excel-documenttaskschedule-startdatetime-member)|获取任务应开始的日期和时间。|
|[DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)|[basicType](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[type](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-type-member)|表示此单元格值的类型。|
|[EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)|[basicType](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[type](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-type-member)|表示此单元格值的类型。|
|[EntityCardLayout](/javascript/api/excel/excel.entitycardlayout)|[layout](/javascript/api/excel/excel.entitycardlayout#excel-excel-entitycardlayout-layout-member)|表示此布局的类型。|
|[EntityCellValue](/javascript/api/excel/excel.entitycellvalue)|[basicType](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[cardLayout](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-cardlayout-member)|表示卡片视图中此实体的布局。|
||[properties： { [key： string]](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-properties-member)|表示此实体的属性及其元数据。|
||[text](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-text-member)|表示呈现具有此值的单元格时显示的文本。|
||[type](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-type-member)|表示此单元格值的类型。|
|[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)|[basicType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-errorsubtype-member)|表示 的类型 `FieldErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-type-member)|表示此单元格值的类型。|
|[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)|[basicType](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[numberFormat](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-numberformat-member)|返回用于显示此值的数值格式字符串。|
||[type](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-type-member)|表示此单元格值的类型。|
|[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)|[basicType](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorType](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-type-member)|表示此单元格值的类型。|
|[标识](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#excel-excel-identity-displayname-member)|表示用户的显示名称。|
||[email](/javascript/api/excel/excel.identity#excel-excel-identity-email-member)|表示用户的电子邮件地址。|
||[id](/javascript/api/excel/excel.identity#excel-excel-identity-id-member)|表示用户的唯一 ID。|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[添加 (：标识) ](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-add-member(1))|向集合添加用户标识。|
||[clear()](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-clear-member(1))|从集合中删除所有的用户标识。|
||[getCount()](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-getcount-member(1))|获取集合中项的数目。|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-getitemat-member(1))|使用文档在集合中的索引获取文档用户标识。|
||[items](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-items-member)|获取此集合中已加载的子项。|
||[remove (assignee： Identity) ](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-remove-member(1))|从集合中删除用户标识。|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-displayname-member)|表示用户的显示名称。|
||[email](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-email-member)|表示用户的电子邮件地址。|
||[id](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-id-member)|表示用户的唯一 ID。|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-dataprovider-member)|链接数据提供程序的数据提供程序数据类型。|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-lastrefreshed-member)|自上次刷新链接工作簿时打开工作簿以来的本地数据类型日期和时间。|
||[name](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-name-member)|链接对象数据类型。|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-periodicrefreshinterval-member)|链接对象刷新的频率（以秒 `refreshMode` 数据类型设置为"Periodic"时刷新。|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-refreshmode-member)|用于检索链接数据数据类型的机制。|
||[requestRefresh () ](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-requestrefresh-member(1))|请求刷新链接数据类型。|
||[requestSetRefreshMode (refreshMode： Excel。LinkedDataTypeRefreshMode) ](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-requestsetrefreshmode-member(1))|请求更改此链接的刷新数据类型。|
||[服务 Id](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-serviceid-member)|链接对象的唯一数据类型。|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-supportedrefreshmodes-member)|返回一个数组，该数组包含链接对象支持的所有刷新数据类型。|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[服务 Id](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-serviceid-member)|新链接对象的唯一 ID 数据类型。|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-source-member)|获取事件源。|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-type-member)|获取事件的类型。|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getcount-member(1))|获取集合中链接的数据类型的数量。|
||[getItem (键：number) ](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitem-member(1))|按服务 ID 数据类型链接的标识。|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitemat-member(1))|按集合数据类型索引获取链接对象。|
||[getItemOrNullObject (键：number) ](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitemornullobject-member(1))|按 ID 获取数据类型链接对象。|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-items-member)|获取此集合中已加载的子项。|
||[requestRefreshAll () ](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-requestrefreshall-member(1))|请求刷新集合中所有链接的数据类型。|
|[LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)|[basicType](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[id](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-id-member)|表示此值中提供的信息的服务源。|
||[properties： { [key： string]： CellValue & { propertyMetadata](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-properties-member)|表示此实体的属性及其元数据。|
||[propertyMetadata](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-propertymetadata-member)||
||[提供程序](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-provider-member)|表示描述提供映像的服务的信息。|
||[text](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-text-member)|表示呈现具有此值的单元格时显示的文本。|
||[type](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-type-member)|表示此单元格值的类型。|
|[LinkedEntityId](/javascript/api/excel/excel.linkedentityid)|[culture](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-culture-member)|表示用于创建此 的语言区域性 `CellValue`。|
||[domainId](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-domainid-member)|表示特定于用于创建 的服务的域 `CellValue`。|
||[entityId](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-entityid-member)|表示特定于用于创建 的服务的标识符 `CellValue`。|
||[服务 Id](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-serviceid-member)|表示用于创建 的服务 `CellValue`。|
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[basicType](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorType](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-type-member)|表示此单元格值的类型。|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[valueAsJson](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-valueasjson-member)|此已命名项中值的 JSON 表示形式。|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[valuesAsJson](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-valuesasjson-member)|此区域单元格中的值的 JSON 表示形式。|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitemornullobject-member(1))|使用工作表视图的名称获取工作表视图。|
|[NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)|[basicType](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorType](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-type-member)|表示此单元格值的类型。|
|[NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)|[basicType](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorType](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-type-member)|表示此单元格值的类型。|
|[NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)|[basicType](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorType](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-type-member)|表示此单元格值的类型。|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getcell-member(1))|根据数据层次结构以及各自层次结构的行和列项，获取数据透视表中的唯一单元格。|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-pivotstyle-member)|应用于数据透视表的样式。|
||[setStyle (样式：string \| PivotTableStyle \| BuiltInPivotTableStyle) ](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setstyle-member(1))|设置应用于数据透视表的样式。|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[getDataSourceString () ](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-getdatasourcestring-member(1))|返回数据透视表数据源的字符串表示形式。|
||[getDataSourceType () ](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-getdatasourcetype-member(1))|获取数据透视表的数据源的类型。|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject () ](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getfirstornullobject-member(1))|获取集合中的第一个数据透视表。|
|[PlaceholderErrorCellValue](/javascript/api/excel/excel.placeholdererrorcellvalue)|[basicType](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorType](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[target](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-target-member)|`PlaceholderErrorCellValue` 在处理期间，在下载数据期间使用。|
||[type](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-type-member)|表示此单元格值的类型。|
|[范围](/javascript/api/excel/excel.range)|[getDependents () ](/javascript/api/excel/excel.range#excel-excel-range-getdependents-member(1))|返回一 `WorkbookRangeAreas` 个对象，该对象表示包含同一工作表或多个工作表中单元格的所有从属单元格的范围。|
||[valuesAsJson](/javascript/api/excel/excel.range#excel-excel-range-valuesasjson-member)|此区域单元格中的值的 JSON 表示形式。|
|[RangeView](/javascript/api/excel/excel.rangeview)|[valuesAsJson](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-valuesasjson-member)|此区域单元格中的值的 JSON 表示形式。|
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[basicType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-errorsubtype-member)|表示 的类型 `RefErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-type-member)|表示此单元格值的类型。|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-refreshmode-member)|链接的数据类型刷新模式。|
||[服务 Id](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-serviceid-member)|刷新模式已更改的对象的唯一 ID。|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-source-member)|获取事件源。|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-type-member)|获取事件的类型。|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[已刷新](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-refreshed-member)|指示刷新请求是否成功。|
||[服务 Id](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-serviceid-member)|已完成刷新请求的对象的唯一 ID。|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-source-member)|获取事件源。|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-type-member)|获取事件的类型。|
||[warnings](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-warnings-member)|包含从刷新请求生成的任何警告的数组。|
|[Shape](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#excel-excel-shape-displayname-member)|获取显示名称的大小。|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1))|从 XML 字符串创建可缩放的矢量图形 (SVG) 并将其添加到工作表。|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#excel-excel-slicer-nameinformula-member)|表示公式中使用切片器名称。|
||[setStyle (样式：字符串 \| SlicerStyle \| BuiltInSlicerStyle) ](/javascript/api/excel/excel.slicer#excel-excel-slicer-setstyle-member(1))|设置应用于切片器的样式。|
||[slicerStyle](/javascript/api/excel/excel.slicer#excel-excel-slicer-slicerstyle-member)|应用于切片器的样式。|
|[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)|[basicType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-errorsubtype-member)|表示 的类型 `SpillErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[spilledColumns](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-spilledcolumns-member)|表示如果没有数据，将溢出的#SPILL！ error。|
||[spilledRows](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-spilledrows-member)|表示如果没有溢出的行数#SPILL！ error。|
||[type](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-type-member)|表示此单元格值的类型。|
|[StringCellValue](/javascript/api/excel/excel.stringcellvalue)|[basicType](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[type](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-type-member)|表示此单元格值的类型。|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#excel-excel-table-clearstyle-member(1))|将表格更改为使用默认表格样式。|
||[onFiltered](/javascript/api/excel/excel.table#excel-excel-table-onfiltered-member)|在将筛选器应用于特定表时发生。|
||[setStyle (样式：string \| TableStyle \| BuiltInTableStyle) ](/javascript/api/excel/excel.table#excel-excel-table-setstyle-member(1))|设置应用于表格的样式。|
||[tableStyle](/javascript/api/excel/excel.table#excel-excel-table-tablestyle-member)|应用于表格的样式。|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onfiltered-member)|在工作簿或工作表的任何表上应用筛选器时发生。|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[valuesAsJson](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-valuesasjson-member)|此表列中单元格中的值的 JSON 表示形式。|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-tableid-member)|获取应用筛选器的表的 ID。|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-worksheetid-member)|获取包含表格的工作表的 ID。|
|[TableRow](/javascript/api/excel/excel.tablerow)|[valuesAsJson](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-valuesasjson-member)|此表行中单元格中的值的 JSON 表示形式。|
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[basicType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-errorsubtype-member)|表示 的类型 `ValueErrorCellValue`。|
||[errorType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-errortype-member)|表示 的类型 `ErrorCellValue`。|
||[type](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-type-member)|表示此单元格值的类型。|
|[ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)|[basicType](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[type](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-type-member)|表示此单元格值的类型。|
|[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)|[address](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-address-member)|表示将下载映像的 URL。|
||[altText](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-alttext-member)|表示可在辅助功能方案中用来描述图像表示的内容的备用文本。|
||[attribution](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-attribution-member)|表示用于描述使用此映像的源和许可证要求的属性信息。|
||[basicType](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-basictype-member)|表示由具有此值的 `Range.valueTypes` 单元格返回的值。|
||[basicValue](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-basicvalue-member)|表示由具有此值的 `Range.values` 单元格返回的值。|
||[提供程序](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-provider-member)|表示描述提供图像的实体或个人的信息。|
||[relatedImagesAddress](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-relatedimagesaddress-member)|表示包含图像被视为与此 相关的网页的 URL `WebImageCellValue`。|
||[type](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-type-member)|表示此单元格值的类型。|
|[Workbook](/javascript/api/excel/excel.workbook)|[getLinkedEntityCellValue (linkedEntityCellValueId：LinkedEntityId) ](/javascript/api/excel/excel.workbook#excel-excel-workbook-getlinkedentitycellvalue-member(1))|基于提供的 `LinkedEntityCellValue` 返回 `LinkedEntityId`。|
||[linkedDataTypes](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkeddatatypes-member)|返回属于工作簿的链接数据类型的集合。|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#excel-excel-workbook-showpivotfieldlist-member)|指定是否在工作簿级别显示数据透视表的字段列表窗格。|
||[任务](/javascript/api/excel/excel.workbook#excel-excel-workbook-tasks-member)|返回工作簿中的任务集合。|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#excel-excel-workbook-use1904datesystem-member)|如果工作簿使用 1904 日期系统，则为 True。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onfiltered-member)|在将筛选器应用于特定工作表时发生。|
||[任务](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tasks-member)|返回工作表中的任务集合。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-addfrombase64-member(1))|将工作簿的指定工作表插入当前工作簿。|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onfiltered-member)|在工作簿中应用任何工作表的筛选器时发生。|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-worksheetid-member)|获取应用筛选器的工作表的 ID。|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[allowEditRanges](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-alloweditranges-member)|指定在此 `AllowEditRangeCollection` 工作表中找到的对象。|
||[canPauseProtection](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-canpauseprotection-member)|指定是否可暂停此工作表的保护。|
||[checkPassword (password？： string) ](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-checkpassword-member(1))|指定密码是否可用于解锁工作表保护。|
||[isPasswordProtected](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-ispasswordprotected-member)|指定工作表是否受密码保护。|
||[isPaused](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-ispaused-member)|指定是否暂停工作表保护。|
||[pauseProtection (password？： string) ](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-pauseprotection-member(1))|暂停给定会话中用户的给定工作表对象的工作表保护。|
||[resumeProtection () ](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-resumeprotection-member(1))|为给定会话中的用户恢复给定工作表对象的工作表保护。|
||[setPassword (password？： string) ](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-setpassword-member(1))|更改与对象关联的 `WorksheetProtection` 密码。|
||[updateOptions (选项：Excel。WorksheetProtectionOptions) ](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-updateoptions-member(1))|更改与对象关联的工作表保护 `WorksheetProtection` 选项。|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-alloweditrangeschanged-member)|指定是否更改了 `AllowEditRange` 任何对象。|
||[protectionOptionsChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-protectionoptionschanged-member)|指定 是否 `WorksheetProtectionOptions` 已更改。|
||[sheetPasswordChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-sheetpasswordchanged-member)|指定工作表密码是否已更改。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
