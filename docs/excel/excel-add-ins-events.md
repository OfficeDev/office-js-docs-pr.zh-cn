---
title: 使用 Excel JavaScript API 处理事件
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 08653a84c051709d16371d89672d3f7ebe2030b7
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872016"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="fc520-102">使用 Excel JavaScript API 处理事件</span><span class="sxs-lookup"><span data-stu-id="fc520-102">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="fc520-103">本文介绍了与处理 Excel 中事件相关的重要概念，并提供了代码示例，以展示如何使用 Excel JavaScript API 注册事件处理程序、处理事件和删除事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="fc520-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="fc520-104">Excel 中的事件</span><span class="sxs-lookup"><span data-stu-id="fc520-104">Events in Excel</span></span>

<span data-ttu-id="fc520-p101">每当 Excel 工作簿中出现某种类型的更改时，就会触发事件通知。 使用 Excel JavaScript API，可以注册事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。 下列事件暂不受支持。</span><span class="sxs-lookup"><span data-stu-id="fc520-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="fc520-108">事件</span><span class="sxs-lookup"><span data-stu-id="fc520-108">Event</span></span> | <span data-ttu-id="fc520-109">说明</span><span class="sxs-lookup"><span data-stu-id="fc520-109">Description</span></span> | <span data-ttu-id="fc520-110">支持的对象</span><span class="sxs-lookup"><span data-stu-id="fc520-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="fc520-111">添加对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="fc520-111">Event that occurs when an object is added.</span></span> | <span data-ttu-id="fc520-112">[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="fc520-112">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="fc520-113">删除对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="fc520-113">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="fc520-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="fc520-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onActivated` | <span data-ttu-id="fc520-115">启用对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="fc520-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="fc520-116">[**Chart**](/javascript/api/excel/excel.chart)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="fc520-116">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onDeactivated` | <span data-ttu-id="fc520-117">停用对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="fc520-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="fc520-118">[**Chart**](/javascript/api/excel/excel.chart)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="fc520-118">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onCalculated` | <span data-ttu-id="fc520-119">工作表完成计算（或集合的所有工作表都已完成）时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="fc520-119">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="fc520-120">[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="fc520-120">[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onChanged` | <span data-ttu-id="fc520-121">更改单元格内数据时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="fc520-121">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="fc520-122">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**Table**](/javascript/api/excel/excel.table)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)</span><span class="sxs-lookup"><span data-stu-id="fc520-122">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="fc520-123">绑定内的数据或格式变化时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="fc520-123">Event that occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="fc520-124">**Binding**</span><span class="sxs-lookup"><span data-stu-id="fc520-124">**Binding**</span></span>](/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | <span data-ttu-id="fc520-125">更改活动单元格或选定范围时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="fc520-125">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="fc520-126">[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**Table**](/javascript/api/excel/excel.table)、[**Binding**](/javascript/api/excel/excel.binding)</span><span class="sxs-lookup"><span data-stu-id="fc520-126">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**Table**](/javascript/api/excel/excel.table), [**Binding**](/javascript/api/excel/excel.binding)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="fc520-127">更改文档中的设置时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="fc520-127">Event that occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="fc520-128">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="fc520-128">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a><span data-ttu-id="fc520-129">事件触发器</span><span class="sxs-lookup"><span data-stu-id="fc520-129">Event triggers</span></span>

<span data-ttu-id="fc520-130">Excel 工作簿内的事件可以通过下列方式触发：</span><span class="sxs-lookup"><span data-stu-id="fc520-130">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="fc520-131">更改工作簿的 Excel 用户界面 (UI) 用户交互</span><span class="sxs-lookup"><span data-stu-id="fc520-131">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="fc520-132">更改工作簿的 Office 加载项 (JavaScript) 代码</span><span class="sxs-lookup"><span data-stu-id="fc520-132">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="fc520-133">更改工作簿的 VBA 加载项（宏）代码</span><span class="sxs-lookup"><span data-stu-id="fc520-133">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="fc520-134">任何符合 Excel 默认行为的更改都会在工作簿中触发一个或多个相应事件。</span><span class="sxs-lookup"><span data-stu-id="fc520-134">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="fc520-135">事件处理程序的生命周期</span><span class="sxs-lookup"><span data-stu-id="fc520-135">Lifecycle of an event handler</span></span>

<span data-ttu-id="fc520-136">当加载项注册事件处理程序时，将创建事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="fc520-136">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="fc520-137">当加载项取消注册事件处理程序或者刷新、重新加载或关闭加载项时，将销毁事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="fc520-137">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="fc520-138">事件处理程序不会作为 Excel 文件的一部分保留，也不会在与 Excel Online 的会话之间保留。</span><span class="sxs-lookup"><span data-stu-id="fc520-138">Event handlers do not persist as part of the Excel file, or across sessions with Excel Online.</span></span>

> [!CAUTION]
> <span data-ttu-id="fc520-139">删除了注册事件的对象（例如，注册 `onChanged` 事件的表）时，事件处理程序不再触发但会保留在内存中，直到加载项或 Excel 会话刷新或关闭为止。</span><span class="sxs-lookup"><span data-stu-id="fc520-139">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="fc520-140">事件和共同创作</span><span class="sxs-lookup"><span data-stu-id="fc520-140">Events and coauthoring</span></span>

<span data-ttu-id="fc520-p103">借助[共同创作功能](co-authoring-in-excel-add-ins.md)，多个人可以共同协作，同时编辑同一个 Excel 工作簿。对于可由共同创作者触发的事件（如 `onChanged`），相应的 **Event** 对象会包含 **source** 属性，以指示事件是由当前用户在本地触发 (`event.source = Local`)，还是由远程共同创作者触发 (`event.source = Remote`)。</span><span class="sxs-lookup"><span data-stu-id="fc520-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="fc520-143">注册事件处理程序</span><span class="sxs-lookup"><span data-stu-id="fc520-143">Register an event handler</span></span>

<span data-ttu-id="fc520-p104">下面的代码示例为 `onChanged` 工作表中的 \*\*\*\* 事件注册事件处理程序。 此代码指定 `handleDataChange` 函数应在工作表中的数据有变化时运行。</span><span class="sxs-lookup"><span data-stu-id="fc520-p104">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="fc520-146">处理事件</span><span class="sxs-lookup"><span data-stu-id="fc520-146">Handle an event</span></span>

<span data-ttu-id="fc520-p105">如上一示例所示，注册事件处理程序时，指定函数应在指定事件发生时运行。 可以将函数设计为执行方案所需的任何操作。 下面的代码示例展示了事件处理程序函数如何直接将事件信息写入控制台。</span><span class="sxs-lookup"><span data-stu-id="fc520-p105">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="fc520-150">删除事件处理程序</span><span class="sxs-lookup"><span data-stu-id="fc520-150">Remove an event handler</span></span>

<span data-ttu-id="fc520-p106">下面的代码示例为 `onSelectionChanged` 工作表中的 \*\*\*\* 事件注册事件处理程序，并将 `handleSelectionChange` 函数定义为在事件发生时运行。 它还定义了随后可以调用的 `remove()` 函数，以删除相应事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="fc520-p106">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="fc520-153">启用和禁用事件</span><span class="sxs-lookup"><span data-stu-id="fc520-153">Enable and disable events</span></span>

<span data-ttu-id="fc520-154">可以通过禁用事件来改进加载项性能。</span><span class="sxs-lookup"><span data-stu-id="fc520-154">The performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="fc520-155">例如，你的应用可能永远不需要接收事件，也可能在执行多个实体的批量编辑时忽略事件。</span><span class="sxs-lookup"><span data-stu-id="fc520-155">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="fc520-156">启用和禁用事件是在[运行时](/javascript/api/excel/excel.runtime)级别进行的。</span><span class="sxs-lookup"><span data-stu-id="fc520-156">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="fc520-157">`enableEvents` 属性确定是否触发事件并激活其处理程序。</span><span class="sxs-lookup"><span data-stu-id="fc520-157">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="fc520-158">以下代码示例展示了如何打开和关闭事件。</span><span class="sxs-lookup"><span data-stu-id="fc520-158">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="fc520-159">另请参阅</span><span class="sxs-lookup"><span data-stu-id="fc520-159">See also</span></span>

- [<span data-ttu-id="fc520-160">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="fc520-160">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
