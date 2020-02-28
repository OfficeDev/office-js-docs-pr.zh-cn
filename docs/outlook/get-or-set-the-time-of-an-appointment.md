---
title: 获取或设置 Outlook 加载项中的约会时间
description: 了解如何在 Outlook 加载项中获取或设置约会开始和结束时间。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: d07d461b852e523626946a79a5c9c5e21c95fcdc
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324959"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a><span data-ttu-id="c23d2-103">在 Outlook 中撰写约会时获取或设置时间</span><span class="sxs-lookup"><span data-stu-id="c23d2-103">Get or set the time when composing an appointment in Outlook</span></span>

<span data-ttu-id="c23d2-104">Office JavaScript API 提供了异步方法（[getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-)和[setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)），以获取和设置用户正在撰写的约会的开始或结束时间。</span><span class="sxs-lookup"><span data-stu-id="c23d2-104">The Office JavaScript API provides asynchronous methods ([Time.getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-) and [Time.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)) to get and set the start or end time of an appointment that the user is composing.</span></span> <span data-ttu-id="c23d2-105">这些异步方法仅适用于撰写外接程序。若要使用这些方法，请确保已正确设置了 Outlook 以在撰写窗体中激活加载项的加载项清单，如[创建适用于撰写窗体的 Outlook 外接程序](compose-scenario.md)中所述。</span><span class="sxs-lookup"><span data-stu-id="c23d2-105">These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="c23d2-p102">[start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 和 [end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 属性对撰写和阅读窗体中的约会均适用。在阅读窗体中，您可以直接从父对象访问属性，类似于：</span><span class="sxs-lookup"><span data-stu-id="c23d2-p102">The [start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) properties are available for appointments in both compose and read forms. In a read form, you can access the properties directly from the parent object, as in:</span></span>

```js
item.start
```

<span data-ttu-id="c23d2-108">及：</span><span class="sxs-lookup"><span data-stu-id="c23d2-108">and in:</span></span>

```js
item.end
```

<span data-ttu-id="c23d2-109">但在撰写窗体中，由于用户和你的加载项可能同时插入或更改时间，因此必须使用异步方法 **getAsync** 来获取开始或结束时间，如下所示：</span><span class="sxs-lookup"><span data-stu-id="c23d2-109">But in a compose form, because both the user and your add-in can be inserting or changing the time at the same time, you must use the asynchronous method **getAsync** to get the start or end time, as shown below:</span></span>

```js
item.start.getAsync
```

<span data-ttu-id="c23d2-110">和：</span><span class="sxs-lookup"><span data-stu-id="c23d2-110">and:</span></span>

```js
item.end.getAsync
```

<span data-ttu-id="c23d2-111">与 Office JavaScript API 中的大多数异步方法一样， **getAsync**和**setAsync**采用可选的输入参数。</span><span class="sxs-lookup"><span data-stu-id="c23d2-111">As with most asynchronous methods in the Office JavaScript API, **getAsync** and **setAsync** take optional input parameters.</span></span> <span data-ttu-id="c23d2-112">有关指定这些可选输入参数的详细信息，请参阅 [Office 外接程序中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline)中的[向异步方法传递可选参数](../develop/asynchronous-programming-in-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="c23d2-112">For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-the-start-or-end-time"></a><span data-ttu-id="c23d2-113">获取开始或结束时间</span><span class="sxs-lookup"><span data-stu-id="c23d2-113">Get the start or end time</span></span>

<span data-ttu-id="c23d2-p104">本节演示一个代码示例，将获取用户正在撰写的约会的开始时间，并显示该时间。你可以使用相同的代码并将 **start** 属性替换为 **end** 属性来获取结束时间。此代码示例在加载项清单中假定了一个规则，将在撰写窗体中为约会激活加载项，如下所示。</span><span class="sxs-lookup"><span data-stu-id="c23d2-p104">This section shows a code sample that gets the start time of the appointment that the user is composing and displays the time. You can use the same code and replace the **start** property by the **end** property to get the end time. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment, as shown below.</span></span>


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

<span data-ttu-id="c23d2-p105">若要使用 **item.start.getAsync** 或 **item.end.getAsync**，请提供回调方法来检查异步调用的状态和结果。可以通过 _asyncContext_ 可选参数向回调方法提供任何需要的自变量。可以使用回调的输出参数 _asyncResult_ 来获取状态、结果和任何错误。如果异步调用成功，则可以使用 [AsyncResult.value](/javascript/api/office/office.asyncresult#value) 属性获取作为 **Date** 对象的 UTC 格式开始时间。</span><span class="sxs-lookup"><span data-stu-id="c23d2-p105">To use **item.start.getAsync** or **item.end.getAsync**, provide a callback method that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback method through the  _asyncContext_ optional parameter. You can obtain status, results and any error using the output parameter _asyncResult_ of the callback. If the asynchronous call is successful, you can get the start time as a **Date** object in UTC format using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>


```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the start time of the item being composed.
        getStartTime();
    });
}

// Get the start time of the item that the user is composing.
function getStartTime() {
    item.start.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the start time, display it, first in UTC and 
                // then convert the Date object to local time and display that.
                write ('The start time in UTC is: ' + asyncResult.value.toString());
                write ('The start time in local time is: ' + asyncResult.value.toLocaleString());
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="set-the-start-or-end-time"></a><span data-ttu-id="c23d2-121">设置开始或结束时间</span><span class="sxs-lookup"><span data-stu-id="c23d2-121">Set the start or end time</span></span>

<span data-ttu-id="c23d2-p106">本节演示一个代码示例，将设置用户正在撰写的约会或邮件的开始时间。你可以使用相同的代码并将 **start** 属性替换为 **end** 属性来设置结束时间。请注意，如果约会撰写窗体已有现有开始时间，随后设置开始时间将调整结束时间以保持约会的任何先前持续时间。如果约会撰写窗体已有现有结束时间，随后设置结束时间将同时调整持续时间和结束时间。如果已将约会设置为全天事件，那么设置开始时间会将结束时间调整为 24 小时后，并取消选中撰写窗体中全天事件的 UI。</span><span class="sxs-lookup"><span data-stu-id="c23d2-p106">This section shows a code sample that sets the start time of the appointment or message that the user is composing. You can use the same code and replace the **start** property by the **end** property to set the end time. Note that if the appointment compose form already has an existing start time, setting the start time subsequently will adjust the end time to maintain any previous duration for the appointment. If the appointment compose form already has an existing end time, setting the end time subsequently will adjust both the duration and end time. If the appointment has been set as an all-day event, setting the start time will adjust the end time to 24 hours later, and uncheck the UI for the all-day event in the compose form.</span></span>

<span data-ttu-id="c23d2-127">与上一示例类似，此代码示例假定外接程序清单中的某个规则将在撰写窗体中为约会激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="c23d2-127">Similar to the previous example, this code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment.</span></span>

<span data-ttu-id="c23d2-p107">若要使用 **item.start.setAsync** 或 **item.end.setAsync**，则在 _dateTime_ 参数中指定一个 UTC 格式的 **Date** 值。如果你根据用户在客户端的输入获取日期，则可以使用 [mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 将值转换为 UTC 格式的 **Date** 对象。你可以提供在 _asyncContext_ 参数中向回调方法提供可选回调方法和任何自变量。应在回调的 _asyncResult_ 输出参数中查看状态、结果和任何错误消息。如果异步调用成功，**setAsync** 会将指定的开始或结束时间字符串作为纯文本插入，覆盖该项的任何现有开始或结束时间。</span><span class="sxs-lookup"><span data-stu-id="c23d2-p107">To use **item.start.setAsync** or **item.end.setAsync**, specify a **Date** value in UTC in the _dateTime_ parameter. If you get a date based on an input by the user on the client, you can use [mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) to convert the value to a **Date** object in UTC. You can provide an optional callback method and any arguments for the callback method in the _asyncContext_ parameter. You should check the status, result and any error message in the _asyncResult_ output parameter of the callback. If the asynchronous call is successful, **setAsync** inserts the specified start or end time string as plain text, overwriting any existing start or end time for that item.</span></span>




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    var startDate = new Date("September 27, 2012 12:30:00");
    
    item.start.setAsync(
        startDate,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the start time.
                // Do whatever appropriate for your scenario
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="see-also"></a><span data-ttu-id="c23d2-133">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c23d2-133">See also</span></span>

- [<span data-ttu-id="c23d2-134">在 Outlook 撰写窗体中获取并设置项数据</span><span class="sxs-lookup"><span data-stu-id="c23d2-134">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)    
- [<span data-ttu-id="c23d2-135">在阅读或撰写窗体中获取并设置 Outlook 项目数据</span><span class="sxs-lookup"><span data-stu-id="c23d2-135">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)   
- [<span data-ttu-id="c23d2-136">创建适用于撰写窗体的 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="c23d2-136">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="c23d2-137">Office 外接程序中的异步编程</span><span class="sxs-lookup"><span data-stu-id="c23d2-137">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="c23d2-138">在 Outlook 中撰写约会或邮件时获取、设置或添加收件人</span><span class="sxs-lookup"><span data-stu-id="c23d2-138">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="c23d2-139">在 Outlook 中撰写约会或邮件时获取或设置主题</span><span class="sxs-lookup"><span data-stu-id="c23d2-139">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)   
- [<span data-ttu-id="c23d2-140">在 Outlook 中撰写约会或邮件时将数据插入到正文中</span><span class="sxs-lookup"><span data-stu-id="c23d2-140">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)   
- [<span data-ttu-id="c23d2-141">在 Outlook 中撰写约会时获取或设置位置</span><span class="sxs-lookup"><span data-stu-id="c23d2-141">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md)
    
