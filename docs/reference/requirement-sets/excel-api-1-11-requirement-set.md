---
title: Excel JavaScript API 要求集1.11
description: 有关 ExcelApi 1.11 要求集的详细信息
ms.date: 05/06/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a7bbb3dc48902e914be8ea3bcbec369e1a64bf42
ms.sourcegitcommit: 735bf94ac3c838f580a992e7ef074dbc8be2b0ea
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/08/2020
ms.locfileid: "44170843"
---
# <a name="whats-new-in-excel-javascript-api-111"></a>Excel JavaScript API 1.11 中的新增功能

ExcelApi 1.11 改进了对注释和工作簿级别的控件（如保存和关闭工作簿）的支持。 此外，它还添加了对区域性设置的访问权限，以帮助帐户进行本地化。

| 功能区域 | 说明 | 相关对象 |
|:--- |:--- |:--- |
| 注释[提到](../../excel/excel-add-ins-comments.md#mentions) |标记，并通过注释通知其他工作簿用户。 | [Comment](/javascript/api/excel/excel.comment)、 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| 批注[分辨率](../../excel/excel-add-ins-comments.md#resolve-comment-threads) | 解析注释线程并获取解决状态。 | [注释](/javascript/api/excel/excel.comment) |
| [区域性设置](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | 获取工作簿的区域性系统设置，如数字格式。 | [CultureInfo](/javascript/api/excel/excel.cultureinfo)、 [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [应用程序](/javascript/api/excel/excel.application) |
| [剪切和粘贴（moveTo）](../../excel/excel-add-ins-ranges-advanced.md#cut-copy-and-paste) | 在 Excel 中复制区域的剪切和粘贴功能。 | [区域](/javascript/api/excel/excel.range) |
| 工作簿[保存](../../excel/excel-add-ins-workbooks.md#save-the-workbook)和[关闭](../../excel/excel-add-ins-workbooks.md#close-the-workbook) | 保存和关闭工作簿。 | [Workbook](/javascript/api/excel/excel.workbook) |
| 工作表事件 | 工作表计算和隐藏行的其他事件和事件信息。 | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)、 [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>API 列表

下表列出了 Excel JavaScript API 要求集1.11 中的 Api。 若要查看 Excel JavaScript API 要求集1.11 或更早版本支持的所有 Api 的 API 参考文档，请参阅[要求集1.10 或更早版本中的 Excel api](/javascript/api/excel?view=excel-js-1.11)。

| Class | 域 | 说明 |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|基于当前系统区域性设置提供信息。 这包括区域性名称、数字格式和其他区域性相关设置。|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|获取用作数值的小数分隔符的字符串。 这是基于 Excel 的本地设置。|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|获取一个字符串，用于将数字值的小数位数与小数的左边隔开。 这是基于 Excel 的本地设置。|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|指定是否启用 Excel 的系统分隔符。|
|[注释](/javascript/api/excel/excel.comment)|[提及](/javascript/api/excel/excel.comment#mentions)|获取注释中提到的实体（例如，人员）。|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|获取丰富的注释内容（例如，注释中的提到）。 此字符串不应显示给最终用户。 您的外接程序应仅使用此信息分析丰富的注释内容。|
||[经过](/javascript/api/excel/excel.comment#resolved)|注释线程状态。 值为 "true" 表示解析注释线程。|
||[updateMentions （contentWithMentions： CommentRichContent）](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|使用特殊格式的字符串和提及列表更新注释内容。|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add （cellAddress： Range \| string，Content： CommentRichContent \| string，contenttype？： Excel. contenttype）](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|使用给定单元格上的给定内容创建新批注。 如果`InvalidArgument`提供的范围大于一个单元格，则会引发错误。|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Comment 中提到的实体的电子邮件地址。|
||[id](/javascript/api/excel/excel.commentmention#id)|实体的 id。 Id 与中`CommentRichContent.richContent`的 id 之一相匹配。|
||[name](/javascript/api/excel/excel.commentmention#name)|Comment 中提到的实体的名称。|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[提及](/javascript/api/excel/excel.commentreply#mentions)|注释中提到的实体（例如，人员）。|
||[经过](/javascript/api/excel/excel.commentreply#resolved)|批注答复状态。 值为 "true" 表示答复处于 "已解决" 状态。|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|丰富的注释内容（例如，注释中提到的内容）。 此字符串不应显示给最终用户。 您的外接程序应仅使用此信息分析丰富的注释内容。|
||[updateMentions （contentWithMentions： CommentRichContent）](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|使用特殊格式的字符串和提及列表更新注释内容。|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add （content： CommentRichContent \| String，contenttype？： Excel. contenttype）](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|为批注创建批注回复。|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[提及](/javascript/api/excel/excel.commentrichcontent#mentions)|包含注释中提到的所有实体（例如，人员）的数组。|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|指定注释的丰富内容（例如，注释内容与提及，第一个提到的实体的 id 属性为0，第二个提到的实体的 id 属性为1）。|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|以 languagecode2/regioncode2 格式获取区域性名称（例如，"zh-tw-cn" 或 "en-us"）。 这取决于当前的系统设置。|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|定义适当的区域性格式，以显示数字。 这取决于当前的系统区域性设置。|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|获取用作数值的小数分隔符的字符串。 这取决于当前的系统设置。|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|获取一个字符串，用于将数字值的小数位数与小数的左边隔开。 这取决于当前的系统设置。|
|[区域](/javascript/api/excel/excel.range)|[moveTo （destinationRange： Range \|字符串）](/javascript/api/excel/excel.range#moveto-destinationrange-)|将单元格的值、格式和公式从当前区域移动到目标区域，替换这些单元格中的旧信息。|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent （金额：数字）](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|调整范围格式的缩进量。 缩进值的范围为0到250，以字符为单位。|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|关闭当前工作簿。|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|保存当前工作簿。|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|在特定工作表上的一个或多个行的隐藏状态更改时发生。|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|完成计算的区域的地址。|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|在特定工作表上的一个或多个行的隐藏状态更改时发生。|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|获取区域地址，该地址表示特定工作表上的更改区域。|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|获取表示事件触发方式的更改类型。 有关详细信息，请参阅 `Excel.RowHiddenChangeType`。|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|获取事件源。 有关详细信息，请参阅 Excel.EventSource。|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|获取事件的类型。 有关详细信息，请参阅 Excel.EventType。|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|获取其中的数据发生更改的工作表的 ID。|

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-1.11)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)