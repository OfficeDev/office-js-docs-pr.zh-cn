---
title: ?? Excel JavaScript API ????
description: ''
ms.date: 01/29/2018
ms.openlocfilehash: 4e04b31e7a130f21d6a9c94d041dc2a122a5890e
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="910d6-102">?? Excel JavaScript API ????</span><span class="sxs-lookup"><span data-stu-id="910d6-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="910d6-103">???????? Excel ??????????????????????????? Excel JavaScript API ???????????????????????</span><span class="sxs-lookup"><span data-stu-id="910d6-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="910d6-104">?????? API ???????? (beta)??????????</span><span class="sxs-lookup"><span data-stu-id="910d6-104">The APIs described in this article are currently available only in public preview (beta) and are not intended for use in production environments.</span></span> <span data-ttu-id="910d6-105">???????????????????? Office???? Office.js CDN ? beta ??https://appsforoffice.microsoft.com/lib/beta/hosted/office.js?</span><span class="sxs-lookup"><span data-stu-id="910d6-105">To run the code samples that this article contains, you must use a sufficiently recent build of Office and reference the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="910d6-106">Excel ????</span><span class="sxs-lookup"><span data-stu-id="910d6-106">Events in Excel</span></span>

<span data-ttu-id="910d6-107">?? Excel ????????????????????????</span><span class="sxs-lookup"><span data-stu-id="910d6-107">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="910d6-108">?? Excel JavaScript API?????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="910d6-108">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="910d6-109">??????????</span><span class="sxs-lookup"><span data-stu-id="910d6-109">The following events are currently supported.</span></span>

| <span data-ttu-id="910d6-110">??</span><span class="sxs-lookup"><span data-stu-id="910d6-110">Event</span></span> | <span data-ttu-id="910d6-111">??</span><span class="sxs-lookup"><span data-stu-id="910d6-111">Description</span></span> | <span data-ttu-id="910d6-112">?????</span><span class="sxs-lookup"><span data-stu-id="910d6-112">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="910d6-113">???????????</span><span class="sxs-lookup"><span data-stu-id="910d6-113">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="910d6-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="910d6-114">**WorksheetCollection**</span></span>](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetaddedeventargs.md) |
| `onActivated` | <span data-ttu-id="910d6-115">???????????</span><span class="sxs-lookup"><span data-stu-id="910d6-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="910d6-116">[**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md)?[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="910d6-116">**WorksheetCollection**, [Worksheet](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md)</span></span> |
| `onDeactivated` | <span data-ttu-id="910d6-117">???????????</span><span class="sxs-lookup"><span data-stu-id="910d6-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="910d6-118">[**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md)?[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="910d6-118">**WorksheetCollection**, [Worksheet](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md)</span></span> |
| `onChanged` | <span data-ttu-id="910d6-119">???????????????</span><span class="sxs-lookup"><span data-stu-id="910d6-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="910d6-120">[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetchangedeventargs.md)?[**Table**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md)?[**TableCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md)?[**Binding**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingdatachangedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="910d6-120">**Worksheet**, [Table](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetchangedeventargs.md), **TableCollection**, [Binding](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="910d6-121">???????????????????</span><span class="sxs-lookup"><span data-stu-id="910d6-121">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="910d6-122">[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetselectionchangedeventargs.md)?[**Table**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tableselectionchangedeventargs.md)?[**Binding**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingselectionchangedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="910d6-122">**Worksheet**, [Table](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetselectionchangedeventargs.md), **Binding**</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="910d6-123">?????</span><span class="sxs-lookup"><span data-stu-id="910d6-123">Event triggers</span></span>

<span data-ttu-id="910d6-124">Excel ??????????????????</span><span class="sxs-lookup"><span data-stu-id="910d6-124">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="910d6-125">?????? Excel ???? (UI) ????</span><span class="sxs-lookup"><span data-stu-id="910d6-125">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="910d6-126">?????? Office ??? (JavaScript) ??</span><span class="sxs-lookup"><span data-stu-id="910d6-126">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="910d6-127">?????? VBA ????????</span><span class="sxs-lookup"><span data-stu-id="910d6-127">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="910d6-128">???? Excel ??????????????????????????</span><span class="sxs-lookup"><span data-stu-id="910d6-128">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="910d6-129">???????????</span><span class="sxs-lookup"><span data-stu-id="910d6-129">Lifecycle of an event handler</span></span>

<span data-ttu-id="910d6-p103">???????????????????????????????????????????????????????????? Excel ???????</span><span class="sxs-lookup"><span data-stu-id="910d6-p103">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="910d6-132">???????</span><span class="sxs-lookup"><span data-stu-id="910d6-132">Events and coauthoring</span></span>

<span data-ttu-id="910d6-p104">??[??????](co-authoring-in-excel-add-ins.md)?????????????????? Excel ???????????????????? `onChanged`????? **Event** ????? **source** ??????????????????? (`event.source = Local`)????????????? (`event.source = Remote`)?</span><span class="sxs-lookup"><span data-stu-id="910d6-p104">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="910d6-135">????????</span><span class="sxs-lookup"><span data-stu-id="910d6-135">Register an event handler</span></span>

<span data-ttu-id="910d6-136">???????? **Sample** ????? `onChanged` ???????????</span><span class="sxs-lookup"><span data-stu-id="910d6-136">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="910d6-137">????? `handleDataChange` ??????????????????</span><span class="sxs-lookup"><span data-stu-id="910d6-137">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="910d6-138">????</span><span class="sxs-lookup"><span data-stu-id="910d6-138">Handle an event</span></span>

<span data-ttu-id="910d6-139">??????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="910d6-139">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="910d6-140">????????????????????</span><span class="sxs-lookup"><span data-stu-id="910d6-140">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="910d6-141">?????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="910d6-141">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="910d6-142">????????</span><span class="sxs-lookup"><span data-stu-id="910d6-142">Remove an event handler</span></span>

<span data-ttu-id="910d6-143">???????? **Sample** ????? `onSelectionChanged` ????????????? `handleSelectionChange` ??????????????</span><span class="sxs-lookup"><span data-stu-id="910d6-143">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="910d6-144">???????????? `remove()` ???????????????</span><span class="sxs-lookup"><span data-stu-id="910d6-144">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="910d6-145">????</span><span class="sxs-lookup"><span data-stu-id="910d6-145">See also</span></span>

- [<span data-ttu-id="910d6-146">Excel JavaScript API ????</span><span class="sxs-lookup"><span data-stu-id="910d6-146">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="910d6-147">Excel JavaScript API ?????</span><span class="sxs-lookup"><span data-stu-id="910d6-147">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="910d6-148">Excel ??????????</span><span class="sxs-lookup"><span data-stu-id="910d6-148">Introduction to Excel Event Features (preview)</span></span>](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/Event_README.md)
