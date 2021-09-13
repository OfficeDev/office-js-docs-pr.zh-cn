---
title: 在 Outlook 加载项中获取或设置主题
description: 了解如何在 Outlook 加载项中获取或设置邮件或约会的主题。
ms.date: 04/15/2019
ms.localizationpriority: medium
ms.openlocfilehash: 534b0cff9d065873ddf566a0d82a3716529414df
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152400"
---
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a>在 Outlook 中撰写约会或邮件时获取或设置主题

JavaScript API Office提供用于 ([subject.getAsync](/javascript/api/outlook/office.Subject#getAsync_options__callback_)和[subject.setAsync](/javascript/api/outlook/office.subject#setAsync_subject__options__callback_)) 的异步方法，用于获取和设置用户正在撰写的约会或邮件的主题。 这些异步方法仅适用于撰写外接程序。若要使用这些方法，请确保已正确设置外接程序清单，Outlook在撰写窗体中激活外接程序。

**subject** 属性可用于约会和邮件的撰写和阅读窗体中的读取权限。在阅读窗体中，可以从父对象直接访问此属性，如：

```js
item.subject
```

但在撰写窗体中，由于用户和加载项可同时插入或更改主题，必须使用异步方法 **getAsync** 获取主题，如下所示：

```js
item.subject.getAsync
```

**subject** 属性仅适用于撰写窗体中（而不能用于阅读窗体中）的写入权限。

与 JavaScript API 中的大多数异步Office一样 **，getAsync** 和 **setAsync** 采用可选输入参数。 有关指定这些可选输入参数的详细信息，请参阅 [Office 外接程序中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md)中的“向异步方法传递可选参数”。


## <a name="get-the-subject"></a>获取主题

本节演示获取用户正在撰写的约会或邮件的主题并显示主题的代码示例。此代码示例假定外接程序清单中的某个规则将在撰写窗体中为约会或邮件激活外接程序，如下所述。


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```

若要使用 **item.subject.getAsync**，可提供一个检查异步调用状态和结果的回调方法。可以通过 _asyncContext_ 可选形参向回调方法提供任何必要实参。可以使用回调的输出形参 _asyncResult_ 获取状态、结果和任何错误。如果异步调用成功，则可以使用 [AsyncResult.value](/javascript/api/office/office.asyncresult#value) 属性获取纯文本字符串形式的主题。


```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the subject of the item being composed.
        getSubject();
    });
}

// Get the subject of the item that the user is composing.
function getSubject() {
    item.subject.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the subject, display it.
                write ('The subject is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="set-the-subject"></a>设置主题


本节演示设置用户正在撰写的约会或邮件的主题的代码示例。与上一示例类似，此代码示例假定外接程序清单中的某个规则将在撰写窗体中为约会或邮件激活外接程序。

若要使用 **item.subject.setAsync**，可在数据形参中指定一个最多 255 字符的字符串。也可以在 _asyncContext_ 形参中为回调方法提供一个回调方法和任何实参。应在回调的 _asyncResult_ 输出形参中检查状态、结果和所有错误消息。如果异步调用成功，**setAsync** 会将指定的主题字符串作为纯文本插入，并覆盖该项目的任何现有主题。

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    var today = new Date();
    var subject;

    // Customize the subject with today's date.
    subject = 'Summary for ' + today.toLocaleDateString();

    item.subject.setAsync(
        subject,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the subject.
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
- [在 Outlook 中撰写约会或邮件时将数据插入到正文中](insert-data-in-the-body.md)   
- [在 Outlook 中撰写约会时获取或设置位置](get-or-set-the-location-of-an-appointment.md) 
- [在 Outlook 中撰写约会时获取或设置时间](get-or-set-the-time-of-an-appointment.md)
    
