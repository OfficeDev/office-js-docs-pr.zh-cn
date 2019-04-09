---
title: 使用 Excel JavaScript API 处理事件
description: ''
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: 7f05263f5220c2d60d0cebcfc686e1fed3f07900
ms.sourcegitcommit: 63219bcc1bb5e3bed7eb6c6b0adb73a4829c7e8f
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/05/2019
ms.locfileid: "31479709"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="c394c-102">使用 Excel JavaScript API 处理事件</span><span class="sxs-lookup"><span data-stu-id="c394c-102">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="c394c-103">本文介绍了与处理 Excel 中事件相关的重要概念，并提供了代码示例，以展示如何使用 Excel JavaScript API 注册事件处理程序、处理事件和删除事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="c394c-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="c394c-104">Excel 中的事件</span><span class="sxs-lookup"><span data-stu-id="c394c-104">Events in Excel</span></span>

<span data-ttu-id="c394c-p101">每当 Excel 工作簿中出现某种类型的更改时，就会触发事件通知。 使用 Excel JavaScript API，可以注册事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。 下列事件暂不受支持。</span><span class="sxs-lookup"><span data-stu-id="c394c-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="c394c-108">事件</span><span class="sxs-lookup"><span data-stu-id="c394c-108">Event</span></span> | <span data-ttu-id="c394c-109">说明</span><span class="sxs-lookup"><span data-stu-id="c394c-109">Description</span></span> | <span data-ttu-id="c394c-110">支持的对象</span><span class="sxs-lookup"><span data-stu-id="c394c-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="c394c-111">激活对象时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-111">Event that occurs when an object is added.</span></span> | <span data-ttu-id="c394c-112">[**Chart**](/javascript/api/excel/excel.chart)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="c394c-112">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Worksheet**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onAdded` | <span data-ttu-id="c394c-113">添加对象时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-113">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="c394c-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="c394c-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onCalculated` | <span data-ttu-id="c394c-115">工作表完成计算（或集合的所有工作表都已完成）时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-115">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="c394c-116">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="c394c-116">[**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Worksheet**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="c394c-117">更改单元格内的数据时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-117">Occurs when data within the binding is changed.</span></span> | <span data-ttu-id="c394c-118">[**Table**](/javascript/api/excel/excel.table)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="c394c-118">[**Worksheet**](/javascript/api/excel/excel.table), [**Table**](/javascript/api/excel/excel.tablecollection), [**TableCollection**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onDataChanged` | <span data-ttu-id="c394c-119">当绑定内的数据或格式变化时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-119">Occurs when data or formatting within the binding is changed.</span></span> | [**<span data-ttu-id="c394c-120">Binding</span><span class="sxs-lookup"><span data-stu-id="c394c-120">Binding</span></span>**](/javascript/api/excel/excel.binding) |
| `onDeactivated` | <span data-ttu-id="c394c-121">停用对象时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-121">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="c394c-122">[**Chart**](/javascript/api/excel/excel.chart)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="c394c-122">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Worksheet**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="c394c-123">删除对象时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-123">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="c394c-124">[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="c394c-124">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="c394c-125">更改活动单元格或选定范围时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-125">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="c394c-126">[**Binding**](/javascript/api/excel/excel.binding)、[**Table**](/javascript/api/excel/excel.table)、[**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="c394c-126">[**Binding**](/javascript/api/excel/excel.binding), [**Table**](/javascript/api/excel/excel.table),  [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="c394c-127">当文档中的设置变化时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-127">Occurs when the Settings in the document are changed.</span></span> | [**<span data-ttu-id="c394c-128">SettingCollection</span><span class="sxs-lookup"><span data-stu-id="c394c-128">SettingCollection</span></span>**](/javascript/api/excel/excel.settingcollection) |

### <a name="events-in-preview"></a><span data-ttu-id="c394c-129">预览版中的事件</span><span class="sxs-lookup"><span data-stu-id="c394c-129">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="c394c-130">以下事件当前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="c394c-130">The  and  methods are currently available only in public preview (beta).</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="c394c-131">事件</span><span class="sxs-lookup"><span data-stu-id="c394c-131">Event</span></span> | <span data-ttu-id="c394c-132">说明</span><span class="sxs-lookup"><span data-stu-id="c394c-132">Description</span></span> | <span data-ttu-id="c394c-133">支持的对象</span><span class="sxs-lookup"><span data-stu-id="c394c-133">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="c394c-134">当激活形状时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-134">Occurs when the shape is activated.</span></span> | [**<span data-ttu-id="c394c-135">Shape</span><span class="sxs-lookup"><span data-stu-id="c394c-135">Shape</span></span>**](/javascript/api/excel/excel.shape)|
| `onAdded` | <span data-ttu-id="c394c-136">在工作簿中添加新表格时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-136">Occurs when new table is added in a workbook.</span></span> | [**<span data-ttu-id="c394c-137">TableCollection</span><span class="sxs-lookup"><span data-stu-id="c394c-137">tableCollection</span></span>**](/javascript/api/excel/excel.tablecollection)|
| `onAutoSaveSettingChanged` | <span data-ttu-id="c394c-138">在工作簿上更改 `autoSave` 设置时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-138">Occurs when the autoSave setting is changed on the workbook.</span></span> | [**<span data-ttu-id="c394c-139">Workbook</span><span class="sxs-lookup"><span data-stu-id="c394c-139">Workbook</span></span>**](/javascript/api/excel/excel.workbook) |
| `onChanged` | <span data-ttu-id="c394c-140">在更改工作簿中的任何工作表时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-140">Occurs when any worksheet in the workbook is changed.</span></span> | [**<span data-ttu-id="c394c-141">WorksheetCollection</span><span class="sxs-lookup"><span data-stu-id="c394c-141">worksheetCollection</span></span>**](/javascript/api/excel/excel.worksheetcollection)|
| `onDeactivated` | <span data-ttu-id="c394c-142">当停用形状时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-142">Occurs when the shape is deactivated.</span></span> | [**<span data-ttu-id="c394c-143">Shape</span><span class="sxs-lookup"><span data-stu-id="c394c-143">Shape</span></span>**](/javascript/api/excel/excel.shape)|
| `onDeleted` | <span data-ttu-id="c394c-144">在工作簿中删除指定的表格时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-144">Occurs when the specified table is deleted in a workbook.</span></span> | [**<span data-ttu-id="c394c-145">TableCollection</span><span class="sxs-lookup"><span data-stu-id="c394c-145">tableCollection</span></span>**](/javascript/api/excel/excel.tablecollection)|
| `onFiltered` | <span data-ttu-id="c394c-146">在对象上应用筛选器时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-146">Occurs when filter is applied on an object.</span></span> | <span data-ttu-id="c394c-147">[**Table**](/javascript/api/excel/excel.table)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="c394c-147">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onFormatChanged` | <span data-ttu-id="c394c-148">在工作表上的格式变化时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-148">Occurs when format changed on a specific worksheet.</span></span> | <span data-ttu-id="c394c-149">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="c394c-149">[**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Worksheet**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="c394c-150">在任何工作表上更改选择时发生。</span><span class="sxs-lookup"><span data-stu-id="c394c-150">Occurs when the selection changes on any worksheet.</span></span> | [**<span data-ttu-id="c394c-151">WorksheetCollection</span><span class="sxs-lookup"><span data-stu-id="c394c-151">worksheetCollection</span></span>**](/javascript/api/excel/excel.worksheetcollection) |

### <a name="event-triggers"></a><span data-ttu-id="c394c-152">事件触发器</span><span class="sxs-lookup"><span data-stu-id="c394c-152">Event triggers</span></span>

<span data-ttu-id="c394c-153">Excel 工作簿内的事件可以通过下列方式触发：</span><span class="sxs-lookup"><span data-stu-id="c394c-153">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="c394c-154">更改工作簿的 Excel 用户界面 (UI) 用户交互</span><span class="sxs-lookup"><span data-stu-id="c394c-154">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="c394c-155">更改工作簿的 Office 加载项 (JavaScript) 代码</span><span class="sxs-lookup"><span data-stu-id="c394c-155">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="c394c-156">更改工作簿的 VBA 加载项（宏）代码</span><span class="sxs-lookup"><span data-stu-id="c394c-156">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="c394c-157">任何符合 Excel 默认行为的更改都会在工作簿中触发一个或多个相应事件。</span><span class="sxs-lookup"><span data-stu-id="c394c-157">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="c394c-158">事件处理程序的生命周期</span><span class="sxs-lookup"><span data-stu-id="c394c-158">Lifecycle of an event handler</span></span>

<span data-ttu-id="c394c-159">当加载项注册事件处理程序时，将创建事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="c394c-159">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="c394c-160">当加载项取消注册事件处理程序或者刷新、重新加载或关闭加载项时，将销毁事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="c394c-160">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="c394c-161">事件处理程序不会作为 Excel 文件的一部分保留，也不会在与 Excel Online 的会话之间保留。</span><span class="sxs-lookup"><span data-stu-id="c394c-161">Event handlers do not persist as part of the Excel file, or across sessions with Excel Online.</span></span>

> [!CAUTION]
> <span data-ttu-id="c394c-162">删除了注册事件的对象（例如，注册 `onChanged` 事件的表）时，事件处理程序不再触发但会保留在内存中，直到加载项或 Excel 会话刷新或关闭为止。</span><span class="sxs-lookup"><span data-stu-id="c394c-162">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="c394c-163">事件和共同创作</span><span class="sxs-lookup"><span data-stu-id="c394c-163">Events and coauthoring</span></span>

<span data-ttu-id="c394c-p104">借助[共同创作功能](co-authoring-in-excel-add-ins.md)，多个人可以共同协作，同时编辑同一个 Excel 工作簿。对于可由共同创作者触发的事件（如 `onChanged`），相应的 **Event** 对象会包含 **source** 属性，以指示事件是由当前用户在本地触发 (`event.source = Local`)，还是由远程共同创作者触发 (`event.source = Remote`)。</span><span class="sxs-lookup"><span data-stu-id="c394c-p104">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="c394c-166">注册事件处理程序</span><span class="sxs-lookup"><span data-stu-id="c394c-166">Register an event handler</span></span>

<span data-ttu-id="c394c-p105">下面的代码示例为 `onChanged` 工作表中的 \*\*\*\* 事件注册事件处理程序。 此代码指定 `handleDataChange` 函数应在工作表中的数据有变化时运行。</span><span class="sxs-lookup"><span data-stu-id="c394c-p105">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="c394c-169">处理事件</span><span class="sxs-lookup"><span data-stu-id="c394c-169">Handle an event</span></span>

<span data-ttu-id="c394c-p106">如上一示例所示，注册事件处理程序时，指定函数应在指定事件发生时运行。 可以将函数设计为执行方案所需的任何操作。 下面的代码示例展示了事件处理程序函数如何直接将事件信息写入控制台。</span><span class="sxs-lookup"><span data-stu-id="c394c-p106">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="c394c-173">删除事件处理程序</span><span class="sxs-lookup"><span data-stu-id="c394c-173">Remove an event handler</span></span>

<span data-ttu-id="c394c-p107">下面的代码示例为 `onSelectionChanged` 工作表中的 \*\*\*\* 事件注册事件处理程序，并将 `handleSelectionChange` 函数定义为在事件发生时运行。 它还定义了随后可以调用的 `remove()` 函数，以删除相应事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="c394c-p107">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="c394c-176">启用和禁用事件</span><span class="sxs-lookup"><span data-stu-id="c394c-176">Enable and disable events</span></span>

<span data-ttu-id="c394c-177">可以通过禁用事件来改进加载项性能。</span><span class="sxs-lookup"><span data-stu-id="c394c-177">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="c394c-178">例如，你的应用可能永远不需要接收事件，也可能在执行多个实体的批量编辑时忽略事件。</span><span class="sxs-lookup"><span data-stu-id="c394c-178">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="c394c-179">启用和禁用事件是在[运行时](/javascript/api/excel/excel.runtime)级别进行的。</span><span class="sxs-lookup"><span data-stu-id="c394c-179">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="c394c-180">`enableEvents` 属性确定是否触发事件并激活其处理程序。</span><span class="sxs-lookup"><span data-stu-id="c394c-180">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="c394c-181">以下代码示例展示了如何打开和关闭事件。</span><span class="sxs-lookup"><span data-stu-id="c394c-181">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="c394c-182">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c394c-182">See also</span></span>

- [<span data-ttu-id="c394c-183">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="c394c-183">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
