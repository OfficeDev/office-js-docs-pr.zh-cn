---
title: Excel 外接程序疑难解答
description: 了解如何解决 Excel 外接程序中的开发错误。
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 1bdd96772d3a221ca3a02e3d5dfcfa16561dd5f1
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409380"
---
# <a name="troubleshooting-excel-add-ins"></a>Excel 外接程序疑难解答

本文讨论了 Excel 特有的故障排除问题。 请使用页面底部的反馈工具建议可添加到文章中的其他问题。

## <a name="api-limitations-when-the-active-workbook-switches"></a>活动工作簿切换时的 API 限制

Excel 相关外接程序用于一次运行单个工作簿。 当运行加载项的工作簿获得焦点时，可能会出现错误。 仅当焦点更改时要调用的特定方法时，才会发生这种情况。

此工作簿开关会影响以下 Api：

|Excel JavaScript API | 引发的错误 |
|--|--|
| `Chart.activate` | GeneralException |
| `Range.select` | GeneralException |
| `Table.clearFilters` | GeneralException |
| `Workbook.getActiveCell`  | InvalidSelection|
| `Workbook.getSelectedRange` | InvalidSelection|
| `Workbook.getSelectedRanges`  | InvalidSelection|
| `Worksheet.activate` | GeneralException |
| `Worksheet.delete`  | InvalidSelection|
| `Worksheet.gridlines` | GeneralException |
| `Worksheet.showHeadings` | GeneralException |
| `WorksheetCollection.add` | GeneralException |
| `WorksheetFreezePanes.freezeAt` | GeneralException |
| `WorksheetFreezePanes.freezeColumns` | GeneralException |
| `WorksheetFreezePanes.freezeRows` | GeneralException |
| `WorksheetFreezePanes.getLocationOrNullObject`| GeneralException |
| `WorksheetFreezePanes.unfreeze` | GeneralException |

> [!NOTE]
> 这仅适用于在 Windows 或 Mac 上打开的多个 Excel 工作簿。

## <a name="coauthoring"></a>共同创作

请参阅 [Excel 外接程序中](co-authoring-in-excel-add-ins.md) 用于共同创作环境中事件的模式的合著。 本文还讨论了使用某些 Api （例如）时的潜在合并冲突 [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) 。

## <a name="see-also"></a>另请参阅

- [解决 Office 外接程序的开发错误](../testing/troubleshoot-development-errors.md)
- [排查 Office 加载项中的用户错误](../testing/testing-and-troubleshooting.md)
