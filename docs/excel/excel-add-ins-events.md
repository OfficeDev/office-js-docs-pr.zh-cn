---
title: 使用 Excel JavaScript API 处理事件
description: ''
ms.date: 05/25/2018
ms.openlocfilehash: b928910cc673cfe8ff99906259b51fa2c3afdca4
ms.sourcegitcommit: 17f60431644b448a4816913039aaebfa328f9b0a
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/25/2018
ms.locfileid: "19476478"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="d1198-102">使用 Excel JavaScript API 处理事件</span><span class="sxs-lookup"><span data-stu-id="d1198-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="d1198-103">本文介绍了与处理 Excel 中事件相关的重要概念，并提供了代码示例，以展示如何使用 Excel JavaScript API 注册事件处理程序、处理事件和删除事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="d1198-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="d1198-104">Excel 中的事件</span><span class="sxs-lookup"><span data-stu-id="d1198-104">Events in Excel</span></span>

<span data-ttu-id="d1198-105">每当 Excel 工作簿中出现某种类型的更改时，就会触发事件通知。</span><span class="sxs-lookup"><span data-stu-id="d1198-105">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="d1198-106">使用 Excel JavaScript API，可以注册事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。</span><span class="sxs-lookup"><span data-stu-id="d1198-106">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="d1198-107">下列事件暂不受支持。</span><span class="sxs-lookup"><span data-stu-id="d1198-107">The following events are currently supported.</span></span>

| <span data-ttu-id="d1198-108">事件</span><span class="sxs-lookup"><span data-stu-id="d1198-108">Event</span></span> | <span data-ttu-id="d1198-109">说明</span><span class="sxs-lookup"><span data-stu-id="d1198-109">Description</span></span> | <span data-ttu-id="d1198-110">支持的对象</span><span class="sxs-lookup"><span data-stu-id="d1198-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="d1198-111">添加对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d1198-111">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="d1198-112">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="d1198-112">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onDeleted` | <span data-ttu-id="d1198-113">删除对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d1198-113">Event that occurs when an object is deleted.</span></span> | [<span data-ttu-id="d1198-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="d1198-114">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onActivated` | <span data-ttu-id="d1198-115">启用对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d1198-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="d1198-116">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection)、[**工作表**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="d1198-116">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onDeactivated` | <span data-ttu-id="d1198-117">停用对象时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d1198-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="d1198-118">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection)、[**工作表**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="d1198-118">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="d1198-119">更改单元格内数据时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d1198-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="d1198-120">[**工作表**](https://dev.office.com/reference/add-ins/excel/worksheet)、[**表**](https://dev.office.com/reference/add-ins/excel/table)[**、TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span><span class="sxs-lookup"><span data-stu-id="d1198-120">[**Worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Table**](https://dev.office.com/reference/add-ins/excel/table), [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="d1198-121">更改捆绑中的数据或格式时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d1198-121">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="d1198-122">**捆绑**</span><span class="sxs-lookup"><span data-stu-id="d1198-122">**Binding**</span></span>](https://dev.office.com/reference/add-ins/excel/binding) |
| `onSelectionChanged` | <span data-ttu-id="d1198-123">更改活动单元格或选定范围时发生的事件。</span><span class="sxs-lookup"><span data-stu-id="d1198-123">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="d1198-124">[**工作表**](https://dev.office.com/reference/add-ins/excel/worksheet)、[**表**](https://dev.office.com/reference/add-ins/excel/table)、[**捆绑**](https://dev.office.com/reference/add-ins/excel/binding)</span><span class="sxs-lookup"><span data-stu-id="d1198-124">**Worksheet**, [Table](https://dev.office.com/reference/add-ins/excel/worksheet), **Binding**</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="d1198-125">事件触发器</span><span class="sxs-lookup"><span data-stu-id="d1198-125">Event triggers</span></span>

<span data-ttu-id="d1198-126">Excel 工作簿内的事件可以通过下列方式触发：</span><span class="sxs-lookup"><span data-stu-id="d1198-126">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="d1198-127">更改工作簿的 Excel 用户界面 (UI) 用户交互</span><span class="sxs-lookup"><span data-stu-id="d1198-127">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="d1198-128">更改工作簿的 Office 加载项 (JavaScript) 代码</span><span class="sxs-lookup"><span data-stu-id="d1198-128">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="d1198-129">更改工作簿的 VBA 加载项（宏）代码</span><span class="sxs-lookup"><span data-stu-id="d1198-129">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="d1198-130">任何符合 Excel 默认行为的更改都会在工作簿中触发一个或多个相应事件。</span><span class="sxs-lookup"><span data-stu-id="d1198-130">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="d1198-131">事件处理程序的生命周期</span><span class="sxs-lookup"><span data-stu-id="d1198-131">Lifecycle of an event handler</span></span>

<span data-ttu-id="d1198-p102">事件处理程序在加载项注册事件处理程序时创建完成，并在加载项取消注册事件处理程序或加载项关闭时销毁。事件处理程序不会暂留为 Excel 文件的一部分。</span><span class="sxs-lookup"><span data-stu-id="d1198-p102">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="d1198-134">事件和共同创作</span><span class="sxs-lookup"><span data-stu-id="d1198-134">Events and coauthoring</span></span>

<span data-ttu-id="d1198-p103">借助[共同创作功能](co-authoring-in-excel-add-ins.md)，多个人可以共同协作，同时编辑同一个 Excel 工作簿。对于可由共同创作者触发的事件（如 `onChanged`），相应的 **Event** 对象会包含 **source** 属性，以指示事件是由当前用户在本地触发 (`event.source = Local`)，还是由远程共同创作者触发 (`event.source = Remote`)。</span><span class="sxs-lookup"><span data-stu-id="d1198-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="d1198-137">注册事件处理程序</span><span class="sxs-lookup"><span data-stu-id="d1198-137">Register an event handler</span></span>

<span data-ttu-id="d1198-138">下面的代码示例为 **Sample** 工作表中的 `onChanged` 事件注册事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="d1198-138">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="d1198-139">此代码指定 `handleDataChange` 函数应在工作表中的数据有变化时运行。</span><span class="sxs-lookup"><span data-stu-id="d1198-139">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="d1198-140">处理事件</span><span class="sxs-lookup"><span data-stu-id="d1198-140">Handle an event</span></span>

<span data-ttu-id="d1198-141">如上一示例所示，注册事件处理程序时，指定函数应在指定事件发生时运行。</span><span class="sxs-lookup"><span data-stu-id="d1198-141">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="d1198-142">可以将函数设计为执行方案所需的任何操作。</span><span class="sxs-lookup"><span data-stu-id="d1198-142">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="d1198-143">下面的代码示例展示了事件处理程序函数如何直接将事件信息写入控制台。</span><span class="sxs-lookup"><span data-stu-id="d1198-143">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="d1198-144">删除事件处理程序</span><span class="sxs-lookup"><span data-stu-id="d1198-144">Remove an event handler</span></span>

<span data-ttu-id="d1198-145">下面的代码示例为 **Sample** 工作表中的 `onSelectionChanged` 事件注册事件处理程序，并将 `handleSelectionChange` 函数定义为在事件发生时运行。</span><span class="sxs-lookup"><span data-stu-id="d1198-145">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="d1198-146">它还定义了随后可以调用的 `remove()` 函数，以删除相应事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="d1198-146">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="d1198-147">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d1198-147">See also</span></span>

- [<span data-ttu-id="d1198-148">Excel JavaScript API 核心概念</span><span class="sxs-lookup"><span data-stu-id="d1198-148">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="d1198-149">Excel JavaScript API 开放性规范</span><span class="sxs-lookup"><span data-stu-id="d1198-149">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)