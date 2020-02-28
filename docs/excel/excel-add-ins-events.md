---
title: 使用 Excel JavaScript API 处理事件
description: Excel JavaScript 对象的事件列表。 其中包括有关使用事件处理程序和关联模式的信息。
ms.date: 02/11/2020
localization_priority: Normal
ms.openlocfilehash: f1a1faf9acc370e7183a078aeeba34019e54900f
ms.sourcegitcommit: d85efbf41a3382ca7d3ab08f2c3f0664d4b26c53
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/28/2020
ms.locfileid: "42327773"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="ccc4d-104">使用 Excel JavaScript API 处理事件</span><span class="sxs-lookup"><span data-stu-id="ccc4d-104">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="ccc4d-105">本文介绍了与处理 Excel 中事件相关的重要概念，并提供了代码示例，以展示如何使用 Excel JavaScript API 注册事件处理程序、处理事件和删除事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-105">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="ccc4d-106">Excel 中的事件</span><span class="sxs-lookup"><span data-stu-id="ccc4d-106">Events in Excel</span></span>

<span data-ttu-id="ccc4d-p102">每当 Excel 工作簿中出现某种类型的更改时，就会触发事件通知。 使用 Excel JavaScript API，可以注册事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。 下列事件暂不受支持。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-p102">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="ccc4d-110">事件</span><span class="sxs-lookup"><span data-stu-id="ccc4d-110">Event</span></span> | <span data-ttu-id="ccc4d-111">说明</span><span class="sxs-lookup"><span data-stu-id="ccc4d-111">Description</span></span> | <span data-ttu-id="ccc4d-112">支持的对象</span><span class="sxs-lookup"><span data-stu-id="ccc4d-112">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="ccc4d-113">激活对象时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-113">Occurs when an object is activated.</span></span> | <span data-ttu-id="ccc4d-114">[**Chart**](/javascript/api/excel/excel.chart)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**Shape**](/javascript/api/excel/excel.shape)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="ccc4d-114">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Shape**](/javascript/api/excel/excel.shape), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onAdded` | <span data-ttu-id="ccc4d-115">当向集合中添加对象时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-115">Occurs when an object is added to the collection.</span></span> | <span data-ttu-id="ccc4d-116">[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="ccc4d-116">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onAutoSaveSettingChanged` | <span data-ttu-id="ccc4d-117">在工作簿上更改 `autoSave` 设置时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-117">Occurs when the `autoSave` setting is changed on the workbook.</span></span> | [<span data-ttu-id="ccc4d-118">**Workbook**</span><span class="sxs-lookup"><span data-stu-id="ccc4d-118">**Workbook**</span></span>](/javascript/api/excel/excel.workbook) |
| `onCalculated` | <span data-ttu-id="ccc4d-119">工作表完成计算（或集合的所有工作表都已完成）时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-119">Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="ccc4d-120">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="ccc4d-120">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="ccc4d-121">更改单元格内的数据时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-121">Occurs when data within cells is changed.</span></span> | <span data-ttu-id="ccc4d-122">[**Table**](/javascript/api/excel/excel.table)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="ccc4d-122">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onColumnSorted` | <span data-ttu-id="ccc4d-123">在已对一个或多个列进行排序时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-123">Occurs when one or more columns have been sorted.</span></span> <span data-ttu-id="ccc4d-124">这是从左到右排序操作的结果。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-124">This happens as the result of a left-to-right sort operation.</span></span> | <span data-ttu-id="ccc4d-125">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="ccc4d-125">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="ccc4d-126">当绑定内的数据或格式变化时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-126">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="ccc4d-127">**Binding**</span><span class="sxs-lookup"><span data-stu-id="ccc4d-127">**Binding**</span></span>](/javascript/api/excel/excel.binding) |
| `onDeactivated` | <span data-ttu-id="ccc4d-128">停用对象时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-128">Occurs when an object is deactivated.</span></span> | <span data-ttu-id="ccc4d-129">[**Chart**](/javascript/api/excel/excel.chart)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**Shape**](/javascript/api/excel/excel.shape)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="ccc4d-129">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Shape**](/javascript/api/excel/excel.shape), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="ccc4d-130">当从集合中删除对象时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-130">Occurs when an object is deleted from the collection.</span></span> | <span data-ttu-id="ccc4d-131">[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="ccc4d-131">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onFormatChanged` | <span data-ttu-id="ccc4d-132">在工作表上的格式变化时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-132">Occurs when the format is changed on a worksheet.</span></span> | <span data-ttu-id="ccc4d-133">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="ccc4d-133">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onRowSorted` | <span data-ttu-id="ccc4d-134">在已对一个或多个行进行排序时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-134">Occurs when one or more rows have been sorted.</span></span> <span data-ttu-id="ccc4d-135">这是从上到下排序操作的结果。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-135">This happens as the result of a top-to-bottom sort operation.</span></span> | <span data-ttu-id="ccc4d-136">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="ccc4d-136">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="ccc4d-137">当活动单元格或选定范围更改时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-137">Occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="ccc4d-138">[**Binding**](/javascript/api/excel/excel.binding)、[**Table**](/javascript/api/excel/excel.table)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="ccc4d-138">[**Binding**](/javascript/api/excel/excel.binding), [**Table**](/javascript/api/excel/excel.table),  [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="ccc4d-139">当文档中的设置变化时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-139">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="ccc4d-140">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="ccc4d-140">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection) |
| `onSingleClicked` | <span data-ttu-id="ccc4d-141">在工作表中进行左键单击/点击操作时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-141">Occurs when left-clicked/tapped action occurs in the worksheet.</span></span> | <span data-ttu-id="ccc4d-142">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="ccc4d-142">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |

> [!WARNING]
> <span data-ttu-id="ccc4d-143">`onSelectionChanged`目前不稳定。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-143">`onSelectionChanged` is currently unstable.</span></span> <span data-ttu-id="ccc4d-144">可通过某种方法可靠地使用 `onSelectionChanged`。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-144">There is a workaround to reliably use `onSelectionChanged`.</span></span> <span data-ttu-id="ccc4d-145">将下面的代码添加到 HTML 主页的 `<head>` 部分：</span><span class="sxs-lookup"><span data-stu-id="ccc4d-145">Add the following code to the `<head>` section of your HTML home page:</span></span>
>
> ```HTML
> <script> MutationObserver=null; </script>
> ```
>
> <span data-ttu-id="ccc4d-146">有关此问题的完整讨论，可在 [office-js GitHub repo](https://github.com/OfficeDev/office-js/issues/533) 上找到。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-146">A full discussion of the issue can be found on the [office-js GitHub repo](https://github.com/OfficeDev/office-js/issues/533).</span></span>

### <a name="events-in-preview"></a><span data-ttu-id="ccc4d-147">预览版中的事件</span><span class="sxs-lookup"><span data-stu-id="ccc4d-147">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="ccc4d-148">以下事件当前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-148">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="ccc4d-149">事件</span><span class="sxs-lookup"><span data-stu-id="ccc4d-149">Event</span></span> | <span data-ttu-id="ccc4d-150">说明</span><span class="sxs-lookup"><span data-stu-id="ccc4d-150">Description</span></span> | <span data-ttu-id="ccc4d-151">支持的对象</span><span class="sxs-lookup"><span data-stu-id="ccc4d-151">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onFiltered` | <span data-ttu-id="ccc4d-152">当将筛选器应用于对象时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-152">Occurs when a filter is applied to an object.</span></span> | <span data-ttu-id="ccc4d-153">[**Table**](/javascript/api/excel/excel.table)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="ccc4d-153">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onRowHiddenChanged` | <span data-ttu-id="ccc4d-154">在特定工作表上的行隐藏状态更改时发生。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-154">Occurs when the row-hidden state changes on a specific worksheet.</span></span> | <span data-ttu-id="ccc4d-155">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="ccc4d-155">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="ccc4d-156">事件触发器</span><span class="sxs-lookup"><span data-stu-id="ccc4d-156">Event triggers</span></span>

<span data-ttu-id="ccc4d-157">Excel 工作簿内的事件可以通过下列方式触发：</span><span class="sxs-lookup"><span data-stu-id="ccc4d-157">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="ccc4d-158">更改工作簿的 Excel 用户界面 (UI) 用户交互</span><span class="sxs-lookup"><span data-stu-id="ccc4d-158">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="ccc4d-159">更改工作簿的 Office 加载项 (JavaScript) 代码</span><span class="sxs-lookup"><span data-stu-id="ccc4d-159">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="ccc4d-160">更改工作簿的 VBA 加载项（宏）代码</span><span class="sxs-lookup"><span data-stu-id="ccc4d-160">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="ccc4d-161">任何符合 Excel 默认行为的更改都会在工作簿中触发一个或多个相应事件。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-161">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="ccc4d-162">事件处理程序的生命周期</span><span class="sxs-lookup"><span data-stu-id="ccc4d-162">Lifecycle of an event handler</span></span>

<span data-ttu-id="ccc4d-163">当加载项注册事件处理程序时，将创建事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-163">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="ccc4d-164">当加载项取消注册事件处理程序或者刷新、重新加载或关闭加载项时，将销毁事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-164">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="ccc4d-165">事件处理程序不会暂留为 Excel 文件的一部分，也不会跨与 Excel 网页版的会话保留。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-165">Event handlers do not persist as part of the Excel file, or across sessions with Excel on the web.</span></span>

> [!CAUTION]
> <span data-ttu-id="ccc4d-166">删除了注册事件的对象（例如，注册 `onChanged` 事件的表）时，事件处理程序不再触发但会保留在内存中，直到加载项或 Excel 会话刷新或关闭为止。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-166">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="ccc4d-167">事件和共同创作</span><span class="sxs-lookup"><span data-stu-id="ccc4d-167">Events and coauthoring</span></span>

<span data-ttu-id="ccc4d-p108">借助[共同创作功能](co-authoring-in-excel-add-ins.md)，多个人可以共同协作，同时编辑同一个 Excel 工作簿。对于可由共同创作者触发的事件（如 `onChanged`），相应的 **Event** 对象会包含 **source** 属性，以指示事件是由当前用户在本地触发 (`event.source = Local`)，还是由远程共同创作者触发 (`event.source = Remote`)。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-p108">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="ccc4d-170">注册事件处理程序</span><span class="sxs-lookup"><span data-stu-id="ccc4d-170">Register an event handler</span></span>

<span data-ttu-id="ccc4d-p109">下面的代码示例为 `onChanged` 工作表中的 \*\*\*\* 事件注册事件处理程序。 此代码指定 `handleDataChange` 函数应在工作表中的数据有变化时运行。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-p109">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

```js
Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);
```

## <a name="handle-an-event"></a><span data-ttu-id="ccc4d-173">处理事件</span><span class="sxs-lookup"><span data-stu-id="ccc4d-173">Handle an event</span></span>

<span data-ttu-id="ccc4d-p110">如上一示例所示，注册事件处理程序时，指定函数应在指定事件发生时运行。 可以将函数设计为执行方案所需的任何操作。 下面的代码示例展示了事件处理程序函数如何直接将事件信息写入控制台。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-p110">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

```js
function handleChange(event)
{
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Change type of event: " + event.changeType);
                console.log("Address of event: " + event.address);
                console.log("Source of event: " + event.source);
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="remove-an-event-handler"></a><span data-ttu-id="ccc4d-177">删除事件处理程序</span><span class="sxs-lookup"><span data-stu-id="ccc4d-177">Remove an event handler</span></span>

<span data-ttu-id="ccc4d-178">下面的代码示例为 **Sample** 工作表中的 `onSelectionChanged` 事件注册事件处理程序，并将 `handleSelectionChange` 函数定义为在事件发生时运行。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-178">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="ccc4d-179">它还定义了随后可以调用的 `remove()` 函数，以删除相应事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-179">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span> <span data-ttu-id="ccc4d-180">请注意， `RequestContext`需要使用来创建事件处理程序才能将其删除。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-180">Note that the `RequestContext` used to create the event handler is needed to remove it.</span></span> 

```js
var eventResult;

Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);

function handleSelectionChange(event)
{
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Address of current selection: " + event.address);
            });
    }).catch(errorHandlerFunction);
}

function remove() {
    return Excel.run(eventResult.context, function (context) {
        eventResult.remove();

        return context.sync()
            .then(function() {
                eventResult = null;
                console.log("Event handler successfully removed.");
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="enable-and-disable-events"></a><span data-ttu-id="ccc4d-181">启用和禁用事件</span><span class="sxs-lookup"><span data-stu-id="ccc4d-181">Enable and disable events</span></span>

<span data-ttu-id="ccc4d-182">可以通过禁用事件来改进加载项性能。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-182">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="ccc4d-183">例如，你的应用可能永远不需要接收事件，也可能在执行多个实体的批量编辑时忽略事件。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-183">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="ccc4d-184">启用和禁用事件是在[运行时](/javascript/api/excel/excel.runtime)级别进行的。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-184">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="ccc4d-185">`enableEvents` 属性确定是否触发事件并激活其处理程序。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-185">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="ccc4d-186">以下代码示例展示了如何打开和关闭事件。</span><span class="sxs-lookup"><span data-stu-id="ccc4d-186">The following code sample shows how to toggle events on and off.</span></span>

```js
Excel.run(function (context) {
    context.runtime.load("enableEvents");
    return context.sync()
        .then(function () {
            var eventBoolean = !context.runtime.enableEvents;
            context.runtime.enableEvents = eventBoolean;
            if (eventBoolean) {
                console.log("Events are currently on.");
            } else {
                console.log("Events are currently off.");
            }
        }).then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="ccc4d-187">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ccc4d-187">See also</span></span>

- [<span data-ttu-id="ccc4d-188">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="ccc4d-188">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
