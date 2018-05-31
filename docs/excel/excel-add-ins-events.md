---
title: 使用 Excel JavaScript API 处理事件
description: ''
ms.date: 01/29/2018
ms.openlocfilehash: 4e04b31e7a130f21d6a9c94d041dc2a122a5890e
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437470"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="4e7ae-102">使用 Excel JavaScript API 处理事件</span><span class="sxs-lookup"><span data-stu-id="4e7ae-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="4e7ae-103">本文介绍了与处理 Excel 中事件相关的重要概念，并提供了代码示例，以展示如何使用 Excel JavaScript API 注册事件处理程序、处理事件和删除事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="4e7ae-104">本文中介绍的 API 暂仅为公共预览版 (beta)，不适用于生产环境。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-104">The APIs described in this article are currently available only in public preview (beta) and are not intended for use in production environments.</span></span> <span data-ttu-id="4e7ae-105">若要运行本文中的代码示例，必须使用最新版 Office，并参考 Office.js CDN 的 beta 库：https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-105">To run the code samples that this article contains, you must use a sufficiently recent build of Office and reference the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="4e7ae-106">Excel 中的事件</span><span class="sxs-lookup"><span data-stu-id="4e7ae-106">Events in Excel</span></span>

<span data-ttu-id="4e7ae-107">每当 Excel 工作簿中出现某种类型的更改时，就会触发事件通知。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-107">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="4e7ae-108">使用 Excel JavaScript API，可以注册事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-108">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="4e7ae-109">下列事件暂不受支持。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-109">The following events are currently supported.</span></span>

| <span data-ttu-id="4e7ae-110">事件</span><span class="sxs-lookup"><span data-stu-id="4e7ae-110">Event</span></span> | <span data-ttu-id="4e7ae-111">说明</span><span class="sxs-lookup"><span data-stu-id="4e7ae-111">Description</span></span> | <span data-ttu-id="4e7ae-112">支持的对象</span><span class="sxs-lookup"><span data-stu-id="4e7ae-112">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="4e7ae-113">添加对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-113">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="4e7ae-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="4e7ae-114">**WorksheetCollection**</span></span>](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetaddedeventargs.md) |
| `onActivated` | <span data-ttu-id="4e7ae-115">启用对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="4e7ae-116">[**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md)、[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="4e7ae-116">**WorksheetCollection**, [Worksheet](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md)</span></span> |
| `onDeactivated` | <span data-ttu-id="4e7ae-117">停用对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="4e7ae-118">[**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md)、[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="4e7ae-118">**WorksheetCollection**, [Worksheet](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md)</span></span> |
| `onChanged` | <span data-ttu-id="4e7ae-119">更改单元格内数据时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="4e7ae-120">[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetchangedeventargs.md)、[**Table**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md)、[**TableCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md)、[**Binding**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingdatachangedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="4e7ae-120">**Worksheet**, [Table](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetchangedeventargs.md), **TableCollection**, [Binding](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="4e7ae-121">更改活动单元格或选定范围时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-121">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="4e7ae-122">[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetselectionchangedeventargs.md)、[**Table**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tableselectionchangedeventargs.md)、[**Binding**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingselectionchangedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="4e7ae-122">**Worksheet**, [Table](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetselectionchangedeventargs.md), **Binding**</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="4e7ae-123">事件触发器</span><span class="sxs-lookup"><span data-stu-id="4e7ae-123">Event triggers</span></span>

<span data-ttu-id="4e7ae-124">Excel 工作簿内的事件可以通过下列方式触发：</span><span class="sxs-lookup"><span data-stu-id="4e7ae-124">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="4e7ae-125">更改工作簿的 Excel 用户界面 (UI) 用户交互</span><span class="sxs-lookup"><span data-stu-id="4e7ae-125">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="4e7ae-126">更改工作簿的 Office 加载项 (JavaScript) 代码</span><span class="sxs-lookup"><span data-stu-id="4e7ae-126">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="4e7ae-127">更改工作簿的 VBA 加载项（宏）代码</span><span class="sxs-lookup"><span data-stu-id="4e7ae-127">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="4e7ae-128">任何符合 Excel 默认行为的更改都会在工作簿中触发一个或多个相应事件。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-128">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="4e7ae-129">事件处理程序的生命周期</span><span class="sxs-lookup"><span data-stu-id="4e7ae-129">Lifecycle of an event handler</span></span>

<span data-ttu-id="4e7ae-p103">事件处理程序在加载项注册事件处理程序时创建完成，并在加载项取消注册事件处理程序或加载项关闭时销毁。事件处理程序不会暂留为 Excel 文件的一部分。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-p103">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="4e7ae-132">事件和共同创作</span><span class="sxs-lookup"><span data-stu-id="4e7ae-132">Events and coauthoring</span></span>

<span data-ttu-id="4e7ae-p104">借助[共同创作功能](co-authoring-in-excel-add-ins.md)，多个人可以共同协作，同时编辑同一个 Excel 工作簿。对于可由共同创作者触发的事件（如 `onChanged`），相应的 **Event** 对象会包含 **source** 属性，以指示事件是由当前用户在本地触发 (`event.source = Local`)，还是由远程共同创作者触发 (`event.source = Remote`)。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-p104">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="4e7ae-135">注册事件处理程序</span><span class="sxs-lookup"><span data-stu-id="4e7ae-135">Register an event handler</span></span>

<span data-ttu-id="4e7ae-136">下面的代码示例为 **Sample** 工作表中的 `onChanged` 事件注册事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-136">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="4e7ae-137">此代码指定 `handleDataChange` 函数应在工作表中的数据有变化时运行。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-137">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="4e7ae-138">处理事件</span><span class="sxs-lookup"><span data-stu-id="4e7ae-138">Handle an event</span></span>

<span data-ttu-id="4e7ae-139">如上一示例所示，注册事件处理程序时，指定函数应在指定事件发生时运行。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-139">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="4e7ae-140">可以将函数设计为执行方案所需的任何操作。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-140">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="4e7ae-141">下面的代码示例展示了事件处理程序函数如何直接将事件信息写入控制台。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-141">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="4e7ae-142">删除事件处理程序</span><span class="sxs-lookup"><span data-stu-id="4e7ae-142">Remove an event handler</span></span>

<span data-ttu-id="4e7ae-143">下面的代码示例为 **Sample** 工作表中的 `onSelectionChanged` 事件注册事件处理程序，并将 `handleSelectionChange` 函数定义为在事件发生时运行。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-143">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="4e7ae-144">它还定义了随后可以调用的 `remove()` 函数，以删除相应事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="4e7ae-144">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="4e7ae-145">另请参阅</span><span class="sxs-lookup"><span data-stu-id="4e7ae-145">See also</span></span>

- [<span data-ttu-id="4e7ae-146">Excel JavaScript API 核心概念</span><span class="sxs-lookup"><span data-stu-id="4e7ae-146">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="4e7ae-147">Excel JavaScript API 开放性规范</span><span class="sxs-lookup"><span data-stu-id="4e7ae-147">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="4e7ae-148">Excel 事件功能简介（预览）</span><span class="sxs-lookup"><span data-stu-id="4e7ae-148">Introduction to Excel Event Features (preview)</span></span>](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/Event_README.md)
