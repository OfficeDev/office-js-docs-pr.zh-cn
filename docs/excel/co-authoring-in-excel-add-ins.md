---
title: 使用 Excel 加载项共同创作
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 774eee0c81f9fb99424070be0ee42860e44e46f5
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449286"
---
# <a name="coauthoring-in-excel-add-ins"></a><span data-ttu-id="8abae-102">使用 Excel 加载项共同创作</span><span class="sxs-lookup"><span data-stu-id="8abae-102">Coauthoring in Excel add-ins</span></span>  

<span data-ttu-id="8abae-p101">借助[共同创作功能](https://support.office.com/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104)，多个人可以共同协作，同时编辑同一个 Excel 工作簿。 在另一个共同创作者更改并保存工作簿后，此工作簿的所有共同创作者都可以立即看这些更改。 若要共同创作 Excel 工作簿，必须将工作簿存储在 OneDrive、OneDrive for Business 或 SharePoint Online 中。</span><span class="sxs-lookup"><span data-stu-id="8abae-p101">With [coauthoring](https://support.office.com/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104), multiple people can work together and edit the same Excel workbook simultaneously. All coauthors of a workbook can see another coauthor's changes as soon as that coauthor saves the workbook. To coauthor an Excel workbook, the workbook must be stored in OneDrive, OneDrive for Business, or SharePoint Online.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8abae-106">在 Excel for Office 365 中，便会发现左上角显示有“自动保存”。</span><span class="sxs-lookup"><span data-stu-id="8abae-106">In Excel for Office 365, you will notice AutoSave in the upper-left corner.</span></span> <span data-ttu-id="8abae-107">启用“自动保存”后，将实时向合著者显示你的更改。</span><span class="sxs-lookup"><span data-stu-id="8abae-107">When AutoSave is turned on, coauthors see your changes in real time.</span></span> <span data-ttu-id="8abae-108">请考虑此行为对 Excel 外接程序设计的影响。</span><span class="sxs-lookup"><span data-stu-id="8abae-108">Consider the impact of this behavior on the design of your Excel add-in.</span></span> <span data-ttu-id="8abae-109">用户可以通过 Excel 窗口左上方的开关禁用“自动保存”。</span><span class="sxs-lookup"><span data-stu-id="8abae-109">Users can turn off AutoSave via the switch in the upper left of the Excel window.</span></span>

<span data-ttu-id="8abae-110">共同创作功能在以下平台上可用：</span><span class="sxs-lookup"><span data-stu-id="8abae-110">Coauthoring is available on the following platforms:</span></span>

- <span data-ttu-id="8abae-111">Excel Online</span><span class="sxs-lookup"><span data-stu-id="8abae-111">Excel Online</span></span>
- <span data-ttu-id="8abae-112">Excel for Android</span><span class="sxs-lookup"><span data-stu-id="8abae-112">Excel for Android</span></span>
- <span data-ttu-id="8abae-113">Excel for iOS</span><span class="sxs-lookup"><span data-stu-id="8abae-113">Excel for iOS</span></span>
- <span data-ttu-id="8abae-114">Excel Mobile for Windows 10</span><span class="sxs-lookup"><span data-stu-id="8abae-114">Excel Mobile for Windows 10</span></span>
- <span data-ttu-id="8abae-115">适用于 Office 365 客户的 Excel for Windows Desktop（Windows 桌面内部版本 16.0.8326.2076 或更高版本，当前的渠道客户自 2017 年 8 月起可获取这些版本）</span><span class="sxs-lookup"><span data-stu-id="8abae-115">Excel for Windows Desktop for Office 365 customers (Windows desktop build 16.0.8326.2076 or later, which is available to current channel customers effective August 2017)</span></span>

## <a name="coauthoring-overview"></a><span data-ttu-id="8abae-116">共同创作功能概述</span><span class="sxs-lookup"><span data-stu-id="8abae-116">Coauthoring overview</span></span>

<span data-ttu-id="8abae-117">如果工作簿的内容有变化，Excel 会自动跨所有共同创作者同步这些更改。</span><span class="sxs-lookup"><span data-stu-id="8abae-117">When you change a workbook's content, Excel automatically synchronizes those changes across all coauthors.</span></span> <span data-ttu-id="8abae-118">共同创作者可以更改工作簿的内容，而 Excel 加载项中运行的代码也可以这样做。</span><span class="sxs-lookup"><span data-stu-id="8abae-118">Coauthors can change the content of a workbook, but so can code running within an Excel add-in.</span></span> <span data-ttu-id="8abae-119">例如，在 Office 加载项中运行以下 JavaScript 代码时，范围值会设置为 Contoso：</span><span class="sxs-lookup"><span data-stu-id="8abae-119">For example, when the following JavaScript code runs in an Office Add-in, the value of a range is set to Contoso:</span></span>

```js
range.values = [['Contoso']];
```
<span data-ttu-id="8abae-120">跨所有共同创作者同步“Contoso”后，同一个工作簿的所有用户或其中运行的所有加载项都能看到新范围值。</span><span class="sxs-lookup"><span data-stu-id="8abae-120">After 'Contoso' synchronizes across all coauthors, any user or add-in running in the same workbook will see the new value of the range.</span></span> 

<span data-ttu-id="8abae-p104">共同创作功能仅同步共享工作簿中的内容。 不会同步从工作簿复制到 Excel 加载项中 JavaScript 变量的值。 例如，如果加载项将单元格值（如“Contoso”）存储在 JavaScript 变量中，然后共同创作者将此单元格值更改为“Example”，那么在同步后所有共同创作者都能在单元格中看到“Example”。 不过，JavaScript 变量值仍设置为“Contoso”。 此外，如果多个共同创作者使用同一个加载项，每个共同创作者都会拥有自己的变量副本，此副本是不会同步的。 如果你使用的变量使用工作簿内容，那么，在使用此变量前，请务必查看工作簿中的更新值。</span><span class="sxs-lookup"><span data-stu-id="8abae-p104">Coauthoring only synchronizes the content within the shared workbook. Values copied from the workbook to JavaScript variables in an Excel add-in are not synchronized. For example, if your add-in stores the value of a cell (such as 'Contoso') in a JavaScript variable, and then a coauthor changes the value of the cell to 'Example', after synchronization all coauthors see 'Example' in the cell. However, the value of the JavaScript variable is still set to 'Contoso'. Furthermore, when multiple coauthors use the same add-in, each coauthor has their own copy of the variable, which is not synchronized. When you use variables that use workbook content, be sure you check for updated values in the workbook before you use the variable.</span></span>

## <a name="use-events-to-manage-the-in-memory-state-of-your-add-in"></a><span data-ttu-id="8abae-127">使用事件管理外接程序的内存中状态</span><span class="sxs-lookup"><span data-stu-id="8abae-127">Use events to manage the in-memory state of your add-in</span></span>

<span data-ttu-id="8abae-p105">Excel 外接程序可以读取工作簿内容（通过隐藏工作表和设置对象），然后将内容存储在变量等数据结构中。 将原始值复制到其中的任意一个数据结构后，合著者可以更新原始的工作簿内容。 这表示现在数据结构中的复制值与工作簿内容不同步。 生成外接程序时，请务必考虑工作簿内容与数据结构中存储的值之间的这种独立性。</span><span class="sxs-lookup"><span data-stu-id="8abae-p105">Excel add-ins can read workbook content (from hidden worksheets and a setting object), and then store it in data structures such as variables. After the original values are copied into any of these data structures, coauthors can update the original workbook content. This means that the copied values in the data structures are now out of sync with the workbook content. When you build your add-ins, be sure to account for this separation of workbook content and values stored in data structures.</span></span>

<span data-ttu-id="8abae-p106">例如，你可能要生成一个显示自定义可视化效果的内容外接程序。 自定义可视化效果的状态可能保存在隐藏工作表中。 当合著者使用同一个工作簿时，可能会发生以下情况：</span><span class="sxs-lookup"><span data-stu-id="8abae-p106">For example, you might build a content add-in that displays custom visualizations. The state of your custom visualizations might be saved in a hidden worksheet. When coauthors use the same workbook, the following scenario can occur:</span></span>

- <span data-ttu-id="8abae-p107">用户 A 打开文档，自定义可视化效果在工作簿中显示。 自定义可视化效果从隐藏工作表中读取数据（例如，将可视化效果的颜色设置为蓝色）。</span><span class="sxs-lookup"><span data-stu-id="8abae-p107">User A opens the document and the custom visualizations are shown in the workbook. The custom visualizations read data from a hidden worksheet (for example, the color of the visualizations is set to blue).</span></span>
- <span data-ttu-id="8abae-p108">用户 B 打开同一个文档，并开始修改自定义可视化效果。 用户 B 将自定义可视化效果的颜色设置为橙色。 橙色保存到隐藏工作表中。</span><span class="sxs-lookup"><span data-stu-id="8abae-p108">User B opens the same document, and starts modifying the custom visualizations. User B sets the color of the custom visualizations to orange. Orange is saved to the hidden worksheet.</span></span>
- <span data-ttu-id="8abae-140">用户 A 的隐藏工作表更新为新值橙色。</span><span class="sxs-lookup"><span data-stu-id="8abae-140">User A's hidden worksheet is updated with the new value of orange.</span></span>
- <span data-ttu-id="8abae-141">用户 A 的自定义可视化效果仍为蓝色。</span><span class="sxs-lookup"><span data-stu-id="8abae-141">User A's custom visualizations are still blue.</span></span>

<span data-ttu-id="8abae-p109">如果希望用户 A 的自定义可视化效果响应共同创作者对隐藏工作表所做的更改，请使用 [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) 事件。 这可确保共同创作者对工作簿内容所做的更改反映到加载项状态中。</span><span class="sxs-lookup"><span data-stu-id="8abae-p109">If you want User A's custom visualizations to respond to changes made by coauthors on the hidden worksheet, use the [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event. This ensures that changes to workbook content made by coauthors is reflected in the state of your add-in.</span></span>

## <a name="caveats-to-using-events-with-coauthoring"></a><span data-ttu-id="8abae-144">使用事件进行共同创作的注意事项</span><span class="sxs-lookup"><span data-stu-id="8abae-144">Caveats to using events with coauthoring</span></span>

<span data-ttu-id="8abae-p110">如上所述，在某些情况下，对所有共同创作者触发事件可提升用户体验。 但是，请注意在一些应用场景下，此行为可能会导致不良的用户体验。</span><span class="sxs-lookup"><span data-stu-id="8abae-p110">As described earlier, in some scenarios, triggering events for all coauthors provides an improved user experience. However, be aware that in some scenarios this behavior can produce poor user experiences.</span></span> 

<span data-ttu-id="8abae-p111">例如，在数据验证应用场景下，通常通过显示 UI 来响应事件。 本地用户或合著者（远程）通过绑定更改工作簿内容时，会运行前面部分中所述的 [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) 事件。 如果 **BindingDataChanged** 事件的事件处理程序显示 UI，用户就会看到与他们在工作簿中进行的更改无关的 UI，从而导致不良的用户体验。 在外接程序中使用事件时，请避免显示 UI。</span><span class="sxs-lookup"><span data-stu-id="8abae-p111">For example, in data validation scenarios, it is common to display UI in response to events. The [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event described in the previous section runs when either a local user or coauthor (remote) changes the workbook content within the binding. If the event handler of the **BindingDataChanged** event displays UI, users will see UI that is unrelated to changes they were working on in the workbook, leading to a poor user experience. Avoid displaying UI when using events in your add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="8abae-151">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8abae-151">See also</span></span>

- [<span data-ttu-id="8abae-152">Excel 中的共同创作功能的相关信息 (VBA)</span><span class="sxs-lookup"><span data-stu-id="8abae-152">About coauthoring in Excel (VBA)</span></span>](/office/vba/excel/concepts/about-coauthoring-in-excel)
- [<span data-ttu-id="8abae-153">自动保存如何影响外接程序和宏 (VBA)</span><span class="sxs-lookup"><span data-stu-id="8abae-153">How AutoSave impacts add-ins and macros (VBA)</span></span>](/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)
