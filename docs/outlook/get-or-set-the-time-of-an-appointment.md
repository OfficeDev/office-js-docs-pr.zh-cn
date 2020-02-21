---
title: 获取或设置 Outlook 加载项中的约会时间
description: 了解如何在 Outlook 加载项中获取或设置约会开始和结束时间。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: af4ec04c8f7af865c826a036b6670c0aec7341b4
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165926"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a>在 Outlook 中撰写约会时获取或设置时间

适用于 Office 的 JavaScript API 提供了异步方法（[Time.getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-) 和 [Time.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)）以获取和设置用户正在撰写的约会的开始和结束时间。这些异步方法仅对撰写外接程序可用。若要使用这些方法，请确保已将 Outlook 的外接程序清单相应地设置为在撰写窗体中激活外接程序，如[创建适用于撰写窗体的 Outlook 外接程序](compose-scenario.md)所述。

[start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 和 [end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 属性对撰写和阅读窗体中的约会均适用。在阅读窗体中，您可以直接从父对象访问属性，类似于：

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

与适用于 Office 的 JavaScript API 中的大多数异步方法相同，**getAsync** 和 **setAsync** 采用可选输入参数。有关指定这些可选输入参数的详细信息，请参阅 [Office 加载项中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md)中的[向异步方法传递可选参数](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline)。


## <a name="get-the-start-or-end-time"></a>获取开始或结束时间

本节演示一个代码示例，将获取用户正在撰写的约会的开始时间，并显示该时间。你可以使用相同的代码并将 **start** 属性替换为 **end** 属性来获取结束时间。此代码示例在加载项清单中假定了一个规则，将在撰写窗体中为约会激活加载项，如下所示。


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

若要使用 **item.start.getAsync** 或 **item.end.getAsync**，请提供回调方法来检查异步调用的状态和结果。可以通过 _asyncContext_ 可选参数向回调方法提供任何需要的自变量。可以使用回调的输出参数 _asyncResult_ 来获取状态、结果和任何错误。如果异步调用成功，则可以使用 [AsyncResult.value](/javascript/api/office/office.asyncresult#value) 属性获取作为 **Date** 对象的 UTC 格式开始时间。


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


## <a name="set-the-start-or-end-time"></a>设置开始或结束时间

本节演示一个代码示例，将设置用户正在撰写的约会或邮件的开始时间。你可以使用相同的代码并将 **start** 属性替换为 **end** 属性来设置结束时间。请注意，如果约会撰写窗体已有现有开始时间，随后设置开始时间将调整结束时间以保持约会的任何先前持续时间。如果约会撰写窗体已有现有结束时间，随后设置结束时间将同时调整持续时间和结束时间。如果已将约会设置为全天事件，那么设置开始时间会将结束时间调整为 24 小时后，并取消选中撰写窗体中全天事件的 UI。

与上一示例类似，此代码示例假定外接程序清单中的某个规则将在撰写窗体中为约会激活外接程序。

若要使用 **item.start.setAsync** 或 **item.end.setAsync**，则在 _dateTime_ 参数中指定一个 UTC 格式的 **Date** 值。如果你根据用户在客户端的输入获取日期，则可以使用 [mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 将值转换为 UTC 格式的 **Date** 对象。你可以提供在 _asyncContext_ 参数中向回调方法提供可选回调方法和任何自变量。应在回调的 _asyncResult_ 输出参数中查看状态、结果和任何错误消息。如果异步调用成功，**setAsync** 会将指定的开始或结束时间字符串作为纯文本插入，覆盖该项的任何现有开始或结束时间。




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


## <a name="see-also"></a>另请参阅

- [在 Outlook 撰写窗体中获取并设置项数据](get-and-set-item-data-in-a-compose-form.md)    
- [在阅读或撰写窗体中获取并设置 Outlook 项目数据](item-data.md)   
- [创建适用于撰写窗体的 Outlook 外接程序](compose-scenario.md)    
- [Office 外接程序中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md)
- [在 Outlook 中撰写约会或邮件时获取、设置或添加收件人](get-set-or-add-recipients.md)  
- [在 Outlook 中撰写约会或邮件时获取或设置主题](get-or-set-the-subject.md)   
- [在 Outlook 中撰写约会或邮件时将数据插入到正文中](insert-data-in-the-body.md)   
- [在 Outlook 中撰写约会时获取或设置位置](get-or-set-the-location-of-an-appointment.md)
    
