---
title: Excel JavaScript API 要求集 1.11
description: 有关 ExcelApi 1.11 要求集的详细信息。
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
---

# <a name="whats-new-in-excel-javascript-api-111"></a>JavaScript API 1.11 Excel新增功能

ExcelApi 1.11 改进了对注释和工作簿级控件的支持 (例如保存和关闭工作簿) 。 它还添加了对区域性设置的访问权限，以帮助说明本地化。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 评论 [提及](../../excel/excel-add-ins-comments.md#mentions) |通过注释标记并通知其他工作簿用户。 | [Comment](/javascript/api/excel/excel.comment)、 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| 注释 [解析](../../excel/excel-add-ins-comments.md#resolve-comment-threads) | 解析注释线程并获取解析状态。 | [Comment](/javascript/api/excel/excel.comment) |
| [区域性设置](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | 获取工作簿的区域性系统设置，如数字格式。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo)、 [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [应用程序](/javascript/api/excel/excel.application) |
| [剪切并粘贴 (moveTo) ](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | 复制 Range 的 Excel中的剪切和粘贴功能。 | [区域](/javascript/api/excel/excel.range) |
| 工作簿[保存](../../excel/excel-add-ins-workbooks.md#save-the-workbook)和[关闭](../../excel/excel-add-ins-workbooks.md#close-the-workbook) | 保存和关闭工作簿。 | [Workbook](/javascript/api/excel/excel.workbook) |
| 工作表事件 | 工作表计算和隐藏行的其他事件和事件信息。 | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)、 [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>API 列表

下表列出了 JavaScript API 要求集 1.11 Excel中的 API。 若要查看受 Excel JavaScript API 要求集 1.11 或更早版本支持的所有 API 的 API 参考文档，请参阅[要求集 1.11](/javascript/api/excel?view=excel-js-1.11&preserve-view=true) 或更早中的 Excel API。

| 类 | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#excel-excel-application-cultureinfo-member)|基于当前系统区域性设置提供相关信息。|
||[decimalSeparator](/javascript/api/excel/excel.application#excel-excel-application-decimalseparator-member)|获取用作数值的小数分隔符的字符串。|
||[thousandsSeparator](/javascript/api/excel/excel.application#excel-excel-application-thousandsseparator-member)|获取用于分隔数字值小数左侧的一组数字的字符串。|
||[useSystemSeparators](/javascript/api/excel/excel.application#excel-excel-application-usesystemseparators-member)|指定是否启用Excel分隔符。|
|[Comment](/javascript/api/excel/excel.comment)|[提及](/javascript/api/excel/excel.comment#excel-excel-comment-mentions-member)|获取 (实体，例如) 中提到的人员。|
||[已解决](/javascript/api/excel/excel.comment#excel-excel-comment-resolved-member)|注释线程状态。|
||[richContent](/javascript/api/excel/excel.comment#excel-excel-comment-richcontent-member)|获取丰富的注释内容 (例如，注释和批注中的) 。|
||[updateMentions (contentWithMentions： Excel。CommentRichContent) ](/javascript/api/excel/excel.comment#excel-excel-comment-updatementions-member(1))|使用特殊格式的字符串和提及列表更新注释内容。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add (cellAddress： Range \| string， content： CommentRichContent \| string， contentType？： Excel.ContentType) ](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-add-member(1))|使用给定单元格上的给定内容创建新批注。|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-email-member)|注释中提到的实体的电子邮件地址。|
||[id](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-id-member)|实体的 ID。|
||[name](/javascript/api/excel/excel.commentmention#excel-excel-commentmention-name-member)|注释中提到的实体的名称。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[提及](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-mentions-member)|实体 (，例如) 中提到的人员。|
||[已解决](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-resolved-member)|批注回复状态。|
||[richContent](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-richcontent-member)|丰富的评论内容 (例如，注释和批注) 。|
||[updateMentions (contentWithMentions： Excel。CommentRichContent) ](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-updatementions-member(1))|使用特殊格式的字符串和提及列表更新注释内容。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add (content： CommentRichContent \| string， contentType？： Excel.ContentType) ](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-add-member(1))|为批注创建批注回复。|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[提及](/javascript/api/excel/excel.commentrichcontent#excel-excel-commentrichcontent-mentions-member)|一个包含注释 (实体的数组，例如) 人。|
||[richContent](/javascript/api/excel/excel.commentrichcontent#excel-excel-commentrichcontent-richcontent-member)|指定批注内容的丰富内容 (例如，提及评论内容，第一个提及实体的 ID 属性为 0，第二个提及实体的 ID 属性为 1) 。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-name-member)|获取语言代码 2-国家/地区代码2 格式的区域性名称 (例如，"zh-cn"或"en-us") 。|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-numberformat-member)|定义在文化上适合显示数字的格式。|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#excel-excel-numberformatinfo-numberdecimalseparator-member)|获取用作数值的小数分隔符的字符串。|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#excel-excel-numberformatinfo-numbergroupseparator-member)|获取用于分隔数字值小数左侧的一组数字的字符串。|
|[区域](/javascript/api/excel/excel.range)|[moveTo (destinationRange：Range \| string) ](/javascript/api/excel/excel.range#excel-excel-range-moveto-member(1))|将单元格值、格式设置和公式从当前区域移动到目标区域，替换这些单元格中的旧信息。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent (amount： number) ](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-adjustindent-member(1))|调整区域格式的缩进。|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#excel-excel-workbook-close-member(1))|关闭当前工作簿。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#excel-excel-workbook-save-member(1))|保存当前工作簿。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowhiddenchanged-member)|当特定工作表上一行或多行的隐藏状态发生更改时发生。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-address-member)|完成计算的范围的地址。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowhiddenchanged-member)|当特定工作表上一行或多行的隐藏状态发生更改时发生。|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-address-member)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-changetype-member)|获取表示如何触发事件的更改类型。|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-source-member)|获取事件源。|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-type-member)|获取事件的类型。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#excel-excel-worksheetrowhiddenchangedeventargs-worksheetid-member)|获取其中的数据发生更改的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
