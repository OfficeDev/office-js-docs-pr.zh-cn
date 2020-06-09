---
title: 使用 Excel 加载项共同创作
description: 了解如何 coauthor 存储在 OneDrive、OneDrive for Business 或 SharePoint Online 中的 Excel 工作簿。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: b70db9c6a0f1f9582288f1078561277b395d3815
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609350"
---
# <a name="coauthoring-in-excel-add-ins"></a>使用 Excel 加载项共同创作  

借助[共同创作功能](https://support.office.com/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104)，多个人可以共同协作，同时编辑同一个 Excel 工作簿。 在另一个共同创作者更改并保存工作簿后，此工作簿的所有共同创作者都可以立即看这些更改。 若要共同创作 Excel 工作簿，必须将工作簿存储在 OneDrive、OneDrive for Business 或 SharePoint Online 中。

> [!IMPORTANT]
> 在 Excel for Office 365 中，便会发现左上角显示有“自动保存”。 启用“自动保存”后，将实时向合著者显示你的更改。 请考虑此行为对 Excel 外接程序设计的影响。 用户可以通过 Excel 窗口左上方的开关禁用“自动保存”。

## <a name="coauthoring-overview"></a>共同创作功能概述

如果工作簿的内容有变化，Excel 会自动跨所有共同创作者同步这些更改。 共同创作者可以更改工作簿的内容，而 Excel 加载项中运行的代码也可以这样做。 例如，在 Office 加载项中运行以下 JavaScript 代码时，范围值会设置为 Contoso：

```js
range.values = [['Contoso']];
```
跨所有共同创作者同步“Contoso”后，同一个工作簿的所有用户或其中运行的所有加载项都能看到新范围值。

共同创作功能仅同步共享工作簿中的内容。 不会同步从工作簿复制到 Excel 加载项中 JavaScript 变量的值。 例如，如果加载项将单元格值（如“Contoso”）存储在 JavaScript 变量中，然后共同创作者将此单元格值更改为“Example”，那么在同步后所有共同创作者都能在单元格中看到“Example”。 不过，JavaScript 变量值仍设置为“Contoso”。 此外，如果多个共同创作者使用同一个加载项，每个共同创作者都会拥有自己的变量副本，此副本是不会同步的。 如果你使用的变量使用工作簿内容，那么，在使用此变量前，请务必查看工作簿中的更新值。

## <a name="use-events-to-manage-the-in-memory-state-of-your-add-in"></a>使用事件管理外接程序的内存中状态

Excel 外接程序可以读取工作簿内容（通过隐藏工作表和设置对象），然后将内容存储在变量等数据结构中。 将原始值复制到其中的任意一个数据结构后，合著者可以更新原始的工作簿内容。 这表示现在数据结构中的复制值与工作簿内容不同步。 生成外接程序时，请务必考虑工作簿内容与数据结构中存储的值之间的这种独立性。

例如，你可能要生成一个显示自定义可视化效果的内容外接程序。 自定义可视化效果的状态可能保存在隐藏工作表中。 当合著者使用同一个工作簿时，可能会发生以下情况：

- 用户 A 打开文档，自定义可视化效果在工作簿中显示。 自定义可视化效果从隐藏工作表中读取数据（例如，将可视化效果的颜色设置为蓝色）。
- 用户 B 打开同一个文档，并开始修改自定义可视化效果。 用户 B 将自定义可视化效果的颜色设置为橙色。 橙色保存到隐藏工作表中。
- 用户 A 的隐藏工作表更新为新值橙色。
- 用户 A 的自定义可视化效果仍为蓝色。

如果希望用户 A 的自定义可视化效果响应共同创作者对隐藏工作表所做的更改，请使用 [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) 事件。 这可确保共同创作者对工作簿内容所做的更改反映到加载项状态中。

## <a name="caveats-to-using-events-with-coauthoring"></a>使用事件进行共同创作的注意事项

如上所述，在某些情况下，对所有共同创作者触发事件可提升用户体验。 但是，请注意在一些应用场景下，此行为可能会导致不良的用户体验。 

例如，在数据验证应用场景下，通常通过显示 UI 来响应事件。 本地用户或合著者（远程）通过绑定更改工作簿内容时，会运行前面部分中所述的 [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) 事件。 如果事件的事件处理程序 `BindingDataChanged` 显示 ui，则用户将看到与工作簿中正在工作的更改无关的 ui，从而导致较差的用户体验。 在外接程序中使用事件时，请避免显示 UI。

## <a name="see-also"></a>另请参阅

- [Excel 中的共同创作功能的相关信息 (VBA)](/office/vba/excel/concepts/about-coauthoring-in-excel)
- [自动保存如何影响外接程序和宏 (VBA)](/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)
