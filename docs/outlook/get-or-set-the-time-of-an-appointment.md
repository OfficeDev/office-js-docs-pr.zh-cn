---
title: 获取或设置 Outlook 加载项中的约会时间
description: 了解如何在 Outlook 加载项中获取或设置约会开始和结束时间。
ms.date: 10/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: c7aa40fda15c613aca869af8b277d4deb6fbf833
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541231"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a>在 Outlook 中撰写约会时获取或设置时间

Office JavaScript API ([Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1)) 和 [Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))) 提供异步方法，以获取和设置用户正在撰写约会的开始或结束时间。 这些异步方法仅可用于撰写加载项。若要使用这些方法，请确保为 Outlook 适当地设置了外接程序 XML 清单以激活撰写窗体中的外接程序，如 [创建撰写窗体的 Outlook 外](compose-scenario.md)接程序中所述。 使用 Office 外接程序的 Teams 清单的外接程序不支持激活规则 [ (预览) ](../develop/json-manifest-overview.md)。

The [start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) and [end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) properties are available for appointments in both compose and read forms. In a read form, you can access the properties directly from the parent object, as in:

```js
item.start
```

及：

```js
item.end
```

但在撰写窗体中，由于用户和你的加载项可能同时插入或更改时间，因此必须使用异步方法 **getAsync** 来获取开始或结束时间，如下所示：

```js
item.start.getAsync
```

和：

```js
item.end.getAsync
```

与 Office JavaScript API 中的大多数异步方法一样， **getAsync** 和 **setAsync** 采用可选输入参数。 有关指定这些可选输入参数的详细信息，请参阅 [Office 外接程序中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline)中的[向异步方法传递可选参数](../develop/asynchronous-programming-in-office-add-ins.md)。

## <a name="get-the-start-or-end-time"></a>获取开始或结束时间

This section shows a code sample that gets the start time of the appointment that the user is composing and displays the time. You can use the same code and replace the **start** property by the **end** property to get the end time. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment, as shown below.

```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
```

若要使用 **item.start.getAsync** 或 **item.end.getAsync**，请提供一个回调函数，用于检查异步调用的状态和结果。 可以通过  _asyncContext_ 可选参数向回调函数提供任何必要的参数。 您可以使用回调的输出形参 _asyncResult_ 来获取状态、结果和任何错误。 如果异步调用成功，您可以使用 **AsyncResult.value** 属性获取作为 [Date](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) 对象的 UTC 格式开始时间。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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

## <a name="set-the-start-or-end-time"></a>设置开始或结束时间

This section shows a code sample that sets the start time of the appointment or message that the user is composing. You can use the same code and replace the **start** property by the **end** property to set the end time. Note that if the appointment compose form already has an existing start time, setting the start time subsequently will adjust the end time to maintain any previous duration for the appointment. If the appointment compose form already has an existing end time, setting the end time subsequently will adjust both the duration and end time. If the appointment has been set as an all-day event, setting the start time will adjust the end time to 24 hours later, and uncheck the UI for the all-day event in the compose form.

与上一示例类似，此代码示例假定外接程序清单中的某个规则将在撰写窗体中为约会激活外接程序。

若要使用 **item.start.setAsync** 或 **item.end.setAsync**，请在 _dateTime_ 参数中指定 UTC 中的 **Date** 值。 如果您根据用户在客户端的输入获取日期，则可以使用 [mailbox.convertToUtcClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 将值转换为 UTC 格式的 **Date** 对象。 可以在 _asyncContext_ 参数中为回调函数提供可选回调函数和任何参数。 您应在回调的 _asyncResult_ 输出形参中查看状态、结果和任何错误消息。 如果异步调用成功， **setAsync** 会将指定的开始或结束时间字符串作为纯文本插入，覆盖该项的任何现有开始或结束时间。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    const startDate = new Date("September 27, 2012 12:30:00");
    
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

## <a name="see-also"></a>另请参阅

- [在 Outlook 撰写窗体中获取并设置项数据](get-and-set-item-data-in-a-compose-form.md)
- [在阅读或撰写窗体中获取并设置 Outlook 项目数据](item-data.md)
- [创建适用于撰写窗体的 Outlook 外接程序](compose-scenario.md)
- [Office 外接程序中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md)
- [在 Outlook 中撰写约会或邮件时获取、设置或添加收件人](get-set-or-add-recipients.md)  
- [在 Outlook 中撰写约会或邮件时获取或设置主题](get-or-set-the-subject.md)
- [在 Outlook 中撰写约会或邮件时将数据插入到正文中](insert-data-in-the-body.md)
- [在 Outlook 中撰写约会时获取或设置位置](get-or-set-the-location-of-an-appointment.md)
