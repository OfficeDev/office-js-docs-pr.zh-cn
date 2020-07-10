---
title: 使用 Excel 加载项共同创作
description: 了解如何 coauthor 存储在 OneDrive、OneDrive for Business 或 SharePoint Online 中的 Excel 工作簿。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 4414bf64f05c29328c63d0857a6e498495712ff1
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093474"
---
# <a name="coauthoring-in-excel-add-ins"></a>使用 Excel 加载项共同创作  

With [coauthoring](https://support.office.com/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104), multiple people can work together and edit the same Excel workbook simultaneously. All coauthors of a workbook can see another coauthor's changes as soon as that coauthor saves the workbook. To coauthor an Excel workbook, the workbook must be stored in OneDrive, OneDrive for Business, or SharePoint Online.

> [!IMPORTANT]
> 在 Excel for Microsoft 365 中，你会发现左上角有 "自动保存"。 启用“自动保存”后，将实时向合著者显示你的更改。 请考虑此行为对 Excel 外接程序设计的影响。 用户可以通过 Excel 窗口左上方的开关禁用“自动保存”。

## <a name="coauthoring-overview"></a>共同创作功能概述

如果工作簿的内容有变化，Excel 会自动跨所有共同创作者同步这些更改。 共同创作者可以更改工作簿的内容，而 Excel 加载项中运行的代码也可以这样做。 例如，在 Office 加载项中运行以下 JavaScript 代码时，范围值会设置为 Contoso：

```js
range.values = [['Contoso']];
```
跨所有共同创作者同步“Contoso”后，同一个工作簿的所有用户或其中运行的所有加载项都能看到新范围值。

Coauthoring only synchronizes the content within the shared workbook. Values copied from the workbook to JavaScript variables in an Excel add-in are not synchronized. For example, if your add-in stores the value of a cell (such as 'Contoso') in a JavaScript variable, and then a coauthor changes the value of the cell to 'Example', after synchronization all coauthors see 'Example' in the cell. However, the value of the JavaScript variable is still set to 'Contoso'. Furthermore, when multiple coauthors use the same add-in, each coauthor has their own copy of the variable, which is not synchronized. When you use variables that use workbook content, be sure you check for updated values in the workbook before you use the variable.

## <a name="use-events-to-manage-the-in-memory-state-of-your-add-in"></a>使用事件管理外接程序的内存中状态

Excel add-ins can read workbook content (from hidden worksheets and a setting object), and then store it in data structures such as variables. After the original values are copied into any of these data structures, coauthors can update the original workbook content. This means that the copied values in the data structures are now out of sync with the workbook content. When you build your add-ins, be sure to account for this separation of workbook content and values stored in data structures.

For example, you might build a content add-in that displays custom visualizations. The state of your custom visualizations might be saved in a hidden worksheet. When coauthors use the same workbook, the following scenario can occur:

- User A opens the document and the custom visualizations are shown in the workbook. The custom visualizations read data from a hidden worksheet (for example, the color of the visualizations is set to blue).
- User B opens the same document, and starts modifying the custom visualizations. User B sets the color of the custom visualizations to orange. Orange is saved to the hidden worksheet.
- 用户 A 的隐藏工作表更新为新值橙色。
- 用户 A 的自定义可视化效果仍为蓝色。

If you want User A's custom visualizations to respond to changes made by coauthors on the hidden worksheet, use the [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event. This ensures that changes to workbook content made by coauthors is reflected in the state of your add-in.

## <a name="caveats-to-using-events-with-coauthoring"></a>使用事件进行共同创作的注意事项

As described earlier, in some scenarios, triggering events for all coauthors provides an improved user experience. However, be aware that in some scenarios this behavior can produce poor user experiences. 

例如，在数据验证应用场景下，通常通过显示 UI 来响应事件。 本地用户或合著者（远程）通过绑定更改工作簿内容时，会运行前面部分中所述的 [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) 事件。 如果事件的事件处理程序 `BindingDataChanged` 显示 ui，则用户将看到与工作簿中正在工作的更改无关的 ui，从而导致较差的用户体验。 在外接程序中使用事件时，请避免显示 UI。

## <a name="see-also"></a>另请参阅

- [Excel 中的共同创作功能的相关信息 (VBA)](/office/vba/excel/concepts/about-coauthoring-in-excel)
- [自动保存如何影响外接程序和宏 (VBA)](/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)
