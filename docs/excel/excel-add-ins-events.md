---
title: 使用 Excel JavaScript API 处理事件
description: ''
ms.date: 05/25/2018
ms.openlocfilehash: fbeb0e6efabe37afb0f73ab8e7448d8cf01ebace
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2018
ms.locfileid: "23943976"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="d7a2b-102">使用 Excel JavaScript API 处理事件</span><span class="sxs-lookup"><span data-stu-id="d7a2b-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="d7a2b-103">本文介绍了与处理 Excel 中事件相关的重要概念，并提供了代码示例，以展示如何使用 Excel JavaScript API 注册事件处理程序、处理事件和删除事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="d7a2b-104">Excel 中的事件</span><span class="sxs-lookup"><span data-stu-id="d7a2b-104">Events in Excel</span></span>

<span data-ttu-id="d7a2b-105">每当 Excel 工作簿中出现某种类型的更改时，就会触发事件通知。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-105">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="d7a2b-106">使用 Excel JavaScript API，可以注册事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-106">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="d7a2b-107">下列事件暂不受支持。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-107">The following events are currently supported.</span></span>

| <span data-ttu-id="d7a2b-108">事件</span><span class="sxs-lookup"><span data-stu-id="d7a2b-108">Event</span></span> | <span data-ttu-id="d7a2b-109">说明</span><span class="sxs-lookup"><span data-stu-id="d7a2b-109">Description</span></span> | <span data-ttu-id="d7a2b-110">支持的对象</span><span class="sxs-lookup"><span data-stu-id="d7a2b-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="d7a2b-111">添加对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-111">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="d7a2b-112">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="d7a2b-112">**WorksheetCollection**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onDeleted` | <span data-ttu-id="d7a2b-113">删除对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-113">Event that occurs when an object is deleted.</span></span> | [<span data-ttu-id="d7a2b-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="d7a2b-114">**WorksheetCollection**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onActivated` | <span data-ttu-id="d7a2b-115">启用对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="d7a2b-116">[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**工作表**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="d7a2b-116">**WorksheetCollection**, [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeactivated` | <span data-ttu-id="d7a2b-117">停用对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="d7a2b-118">[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**工作表**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="d7a2b-118">**WorksheetCollection**, [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="d7a2b-119">更改单元格内数据时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="d7a2b-120">[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)、[**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table)[**、TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection)</span><span class="sxs-lookup"><span data-stu-id="d7a2b-120">[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="d7a2b-121">数据绑定中的数据或格式更改时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-121">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="d7a2b-122">**数据绑定**</span><span class="sxs-lookup"><span data-stu-id="d7a2b-122">**Binding**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | <span data-ttu-id="d7a2b-123">更改活动单元格或选定范围时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-123">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="d7a2b-124">[**工作表**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)、[**表格**](https://docs.microsoft.com/javascript/api/excel/excel.table)、[**数据绑定**](https://docs.microsoft.com/javascript/api/excel/excel.binding)</span><span class="sxs-lookup"><span data-stu-id="d7a2b-124">**Worksheet**, [Table](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), **Binding**</span></span> |
| `onSettingsChanged` | <span data-ttu-id="d7a2b-125">当文档设置变更时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-125">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="d7a2b-126">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="d7a2b-126">**SettingCollection**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection) |

## <a name="preview-beta-events-in-excel"></a><span data-ttu-id="d7a2b-127">在 Excel 中预览（Beta）事件</span><span class="sxs-lookup"><span data-stu-id="d7a2b-127">Preview (Beta) Events in Excel</span></span>

> [!NOTE]
> <span data-ttu-id="d7a2b-128">这些事件目前仅适用于公开预览版（测试版）。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-128">These samples use APIs currently available only in public preview (beta).</span></span> <span data-ttu-id="d7a2b-129">要使用这些功能，您必须使用 Office.js CDN 的 beta 库： https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-129">To use these features, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

| <span data-ttu-id="d7a2b-130">事件</span><span class="sxs-lookup"><span data-stu-id="d7a2b-130">Event</span></span> | <span data-ttu-id="d7a2b-131">说明</span><span class="sxs-lookup"><span data-stu-id="d7a2b-131">Description</span></span> | <span data-ttu-id="d7a2b-132">支持的对象</span><span class="sxs-lookup"><span data-stu-id="d7a2b-132">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="d7a2b-133">添加图表时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-133">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="d7a2b-134">**ChartCollection**</span><span class="sxs-lookup"><span data-stu-id="d7a2b-134">**chartCollection**</span></span>](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md) |
| `onDeleted` | <span data-ttu-id="d7a2b-135">删除图表时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-135">Event that occurs when an object is deleted.</span></span> | [<span data-ttu-id="d7a2b-136">**ChartCollection**</span><span class="sxs-lookup"><span data-stu-id="d7a2b-136">**chartCollection**</span></span>](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md) |
| `onActivated` | <span data-ttu-id="d7a2b-137">激活图表时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-137">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="d7a2b-138">[**Chart**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)， [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span><span class="sxs-lookup"><span data-stu-id="d7a2b-138">[**Chart**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md), [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span></span> |
| `onDeactivated` | <span data-ttu-id="d7a2b-139">停用图表时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-139">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="d7a2b-140">[**Chart**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)， [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span><span class="sxs-lookup"><span data-stu-id="d7a2b-140">[**Chart**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md), [**ChartCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span></span> |
| `onCalculated` | <span data-ttu-id="d7a2b-141">工作表完成计算（或集合的所有工作表都已完成）时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-141">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="d7a2b-142">[**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)、[**工作表**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span><span class="sxs-lookup"><span data-stu-id="d7a2b-142">**WorksheetCollection**, [Worksheet](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/new-events.md)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="d7a2b-143">事件触发器</span><span class="sxs-lookup"><span data-stu-id="d7a2b-143">Event triggers</span></span>

<span data-ttu-id="d7a2b-144">Excel 工作簿内的事件可以通过下列方式触发：</span><span class="sxs-lookup"><span data-stu-id="d7a2b-144">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="d7a2b-145">更改工作簿的 Excel 用户界面 (UI) 用户交互</span><span class="sxs-lookup"><span data-stu-id="d7a2b-145">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="d7a2b-146">更改工作簿的 Office 加载项 (JavaScript) 代码</span><span class="sxs-lookup"><span data-stu-id="d7a2b-146">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="d7a2b-147">更改工作簿的 VBA 加载项（宏）代码</span><span class="sxs-lookup"><span data-stu-id="d7a2b-147">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="d7a2b-148">任何符合 Excel 默认行为的更改都会在工作簿中触发一个或多个相应事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-148">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="d7a2b-149">事件处理程序的生命周期</span><span class="sxs-lookup"><span data-stu-id="d7a2b-149">Lifecycle of an event handler</span></span>

<span data-ttu-id="d7a2b-p103">事件处理程序在加载项注册事件处理程序时创建完成，并在加载项取消注册事件处理程序或加载项关闭时销毁。事件处理程序不会暂留为 Excel 文件的一部分。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-p103">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="d7a2b-152">事件和共同创作</span><span class="sxs-lookup"><span data-stu-id="d7a2b-152">Events and coauthoring</span></span>

<span data-ttu-id="d7a2b-p104">借助[共同创作功能](co-authoring-in-excel-add-ins.md)，多个人可以共同协作，同时编辑同一个 Excel 工作簿。对于可由共同创作者触发的事件（如 `onChanged`），相应的 **Event** 对象会包含 **source** 属性，以指示事件是由当前用户在本地触发 (`event.source = Local`)，还是由远程共同创作者触发 (`event.source = Remote`)。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-p104">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="d7a2b-155">注册事件处理程序</span><span class="sxs-lookup"><span data-stu-id="d7a2b-155">Register an event handler</span></span>

<span data-ttu-id="d7a2b-156">下面的代码示例为 **Sample** 工作表中的 `onChanged` 事件注册事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-156">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="d7a2b-157">此代码指定 `handleDataChange` 函数应在工作表中的数据有变化时运行。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-157">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="d7a2b-158">处理事件</span><span class="sxs-lookup"><span data-stu-id="d7a2b-158">Handle an event</span></span>

<span data-ttu-id="d7a2b-159">如上一示例所示，注册事件处理程序时，指定函数应在指定事件发生时运行。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-159">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="d7a2b-160">可以将函数设计为执行方案所需的任何操作。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-160">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="d7a2b-161">下面的代码示例展示了事件处理程序函数如何直接将事件信息写入控制台。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-161">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="d7a2b-162">删除事件处理程序</span><span class="sxs-lookup"><span data-stu-id="d7a2b-162">Remove an event handler</span></span>

<span data-ttu-id="d7a2b-163">下面的代码示例为 **Sample** 工作表中的 `onSelectionChanged` 事件注册事件处理程序，并将 `handleSelectionChange` 函数定义为在事件发生时运行。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-163">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="d7a2b-164">它还定义了随后可以调用的 `remove()` 函数，以删除相应事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-164">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="d7a2b-165">启用和禁用事件</span><span class="sxs-lookup"><span data-stu-id="d7a2b-165">Enable and disable events</span></span>

> [!NOTE]
> <span data-ttu-id="d7a2b-166">此功能目前仅在公共预览版（测试版）中可用。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-166">This sample uses APIs that are currently available only in public preview (beta).</span></span> <span data-ttu-id="d7a2b-167">要使用该功能，您必须引用 Office.js CDN 的 beta 库：https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-167">To use these features, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

<span data-ttu-id="d7a2b-168">可以通过禁用事件来提高加载项的性能。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-168">The performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="d7a2b-169">例如，您的应用可能永远不需要接收事件，也可能在执行多个实体的批量编辑时忽略事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-169">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span> 

<span data-ttu-id="d7a2b-170">在[运行时](https://docs.microsoft.com/javascript/api/excel/excel.runtime)级别启用或禁用事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-170">Events are turned on and off at the [runtime](https://docs.microsoft.com/javascript/api/excel/excel.runtime) level.</span></span> <span data-ttu-id="d7a2b-171">属性判断是否会触发事件，并激活其处理程序。`enableEvents`</span><span class="sxs-lookup"><span data-stu-id="d7a2b-171">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span> 

<span data-ttu-id="d7a2b-172">下面的代码示例演示如何打开和关闭事件。</span><span class="sxs-lookup"><span data-stu-id="d7a2b-172">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="d7a2b-173">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d7a2b-173">See also</span></span>

- [<span data-ttu-id="d7a2b-174">Excel JavaScript API 核心概念</span><span class="sxs-lookup"><span data-stu-id="d7a2b-174">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="d7a2b-175">Excel JavaScript API 开放性规范</span><span class="sxs-lookup"><span data-stu-id="d7a2b-175">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)