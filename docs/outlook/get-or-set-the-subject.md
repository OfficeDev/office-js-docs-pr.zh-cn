---
title: 在 Outlook 加载项中获取或设置主题
description: 了解如何在 Outlook 加载项中获取或设置邮件或约会的主题。
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 93864aee005af61d9648c39402a843d9105bb021
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325438"
---
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="6cab3-103">在 Outlook 中撰写约会或邮件时获取或设置主题</span><span class="sxs-lookup"><span data-stu-id="6cab3-103">Get or set the subject when composing an appointment or message in Outlook</span></span>

<span data-ttu-id="6cab3-104">Office JavaScript API 提供了异步方法（[getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-)和[setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)），以获取和设置用户正在撰写的约会或邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="6cab3-104">The Office JavaScript API provides asynchronous methods ([subject.getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-) and [subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)) to get and set the subject of an appointment or message that the user is composing.</span></span> <span data-ttu-id="6cab3-105">这些异步方法仅适用于撰写外接程序。若要使用这些方法，请确保已正确设置了 Outlook 以在撰写窗体中激活加载项的加载项清单。</span><span class="sxs-lookup"><span data-stu-id="6cab3-105">These asynchronous methods are available only to compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms.</span></span>

<span data-ttu-id="6cab3-p102">**subject** 属性可用于约会和邮件的撰写和阅读窗体中的读取权限。在阅读窗体中，可以从父对象直接访问此属性，如：</span><span class="sxs-lookup"><span data-stu-id="6cab3-p102">The **subject** property is available for read access in both compose and read forms of appointments and messages. In a read form, you can access the property directly from the parent object, as in:</span></span>

```js
item.subject
```

<span data-ttu-id="6cab3-108">但在撰写窗体中，由于用户和加载项可同时插入或更改主题，必须使用异步方法 **getAsync** 获取主题，如下所示：</span><span class="sxs-lookup"><span data-stu-id="6cab3-108">But in a compose form, because both the user and your add-in can be inserting or changing the subject at the same time, you must use the asynchronous method **getAsync** to get the subject, as shown below:</span></span>

```js
item.subject.getAsync
```

<span data-ttu-id="6cab3-109">**subject** 属性仅适用于撰写窗体中（而不能用于阅读窗体中）的写入权限。</span><span class="sxs-lookup"><span data-stu-id="6cab3-109">The **subject** property is available for write access in only compose forms and not in read forms.</span></span>

<span data-ttu-id="6cab3-110">与 Office JavaScript API 中的大多数异步方法一样， **getAsync**和**setAsync**采用可选的输入参数。</span><span class="sxs-lookup"><span data-stu-id="6cab3-110">As with most asynchronous methods in the Office JavaScript API, **getAsync** and **setAsync** take optional input parameters.</span></span> <span data-ttu-id="6cab3-111">有关指定这些可选输入参数的详细信息，请参阅 [Office 外接程序中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md)中的“向异步方法传递可选参数”。</span><span class="sxs-lookup"><span data-stu-id="6cab3-111">For more information about specifying these optional input parameters, see "Passing optional parameters to asynchronous methods" in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-the-subject"></a><span data-ttu-id="6cab3-112">获取主题</span><span class="sxs-lookup"><span data-stu-id="6cab3-112">Get the subject</span></span>

<span data-ttu-id="6cab3-p104">本节演示获取用户正在撰写的约会或邮件的主题并显示主题的代码示例。此代码示例假定外接程序清单中的某个规则将在撰写窗体中为约会或邮件激活外接程序，如下所述。</span><span class="sxs-lookup"><span data-stu-id="6cab3-p104">This section shows a code sample that gets the subject of the appointment or message that the user is composing, and displays the subject. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```

<span data-ttu-id="6cab3-p105">若要使用 **item.subject.getAsync**，可提供一个检查异步调用状态和结果的回调方法。可以通过 _asyncContext_ 可选形参向回调方法提供任何必要实参。可以使用回调的输出形参 _asyncResult_ 获取状态、结果和任何错误。如果异步调用成功，则可以使用 [AsyncResult.value](/javascript/api/office/office.asyncresult#value) 属性获取纯文本字符串形式的主题。</span><span class="sxs-lookup"><span data-stu-id="6cab3-p105">To use **item.subject.getAsync**, provide a callback method that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback method through the  _asyncContext_ optional parameter. You can obtain status, results and any error using the output parameter _asyncResult_ of the callback. If the asynchronous call is successful, you can get the subject as a plain text string using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>


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


## <a name="set-the-subject"></a><span data-ttu-id="6cab3-119">设置主题</span><span class="sxs-lookup"><span data-stu-id="6cab3-119">Set the subject</span></span>


<span data-ttu-id="6cab3-p106">本节演示设置用户正在撰写的约会或邮件的主题的代码示例。与上一示例类似，此代码示例假定外接程序清单中的某个规则将在撰写窗体中为约会或邮件激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="6cab3-p106">This section shows a code sample that sets the subject of the appointment or message that the user is composing. Similar to the previous example, this code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message.</span></span>

<span data-ttu-id="6cab3-p107">若要使用 **item.subject.setAsync**，可在数据形参中指定一个最多 255 字符的字符串。也可以在 _asyncContext_ 形参中为回调方法提供一个回调方法和任何实参。应在回调的 _asyncResult_ 输出形参中检查状态、结果和所有错误消息。如果异步调用成功，**setAsync** 会将指定的主题字符串作为纯文本插入，并覆盖该项目的任何现有主题。</span><span class="sxs-lookup"><span data-stu-id="6cab3-p107">To use **item.subject.setAsync**, specify a string of up to 255 characters in the data parameter. Optionally, you can provide a callback method and any arguments for the callback method in the  _asyncContext_ parameter. You should check the status, result and any error message in the _asyncResult_ output parameter of the callback. If the asynchronous call is successful, **setAsync** inserts the specified subject string as plain text, overwriting any existing subject for that item.</span></span>

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


## <a name="see-also"></a><span data-ttu-id="6cab3-126">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6cab3-126">See also</span></span>

- [<span data-ttu-id="6cab3-127">在 Outlook 撰写窗体中获取并设置项数据</span><span class="sxs-lookup"><span data-stu-id="6cab3-127">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)   
- [<span data-ttu-id="6cab3-128">在阅读或撰写窗体中获取并设置 Outlook 项目数据</span><span class="sxs-lookup"><span data-stu-id="6cab3-128">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)    
- [<span data-ttu-id="6cab3-129">创建适用于撰写窗体的 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="6cab3-129">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="6cab3-130">Office 外接程序中的异步编程</span><span class="sxs-lookup"><span data-stu-id="6cab3-130">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="6cab3-131">在 Outlook 中撰写约会或邮件时获取、设置或添加收件人</span><span class="sxs-lookup"><span data-stu-id="6cab3-131">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="6cab3-132">在 Outlook 中撰写约会或邮件时将数据插入到正文中</span><span class="sxs-lookup"><span data-stu-id="6cab3-132">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)   
- [<span data-ttu-id="6cab3-133">在 Outlook 中撰写约会时获取或设置位置</span><span class="sxs-lookup"><span data-stu-id="6cab3-133">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md) 
- [<span data-ttu-id="6cab3-134">在 Outlook 中撰写约会时获取或设置时间</span><span class="sxs-lookup"><span data-stu-id="6cab3-134">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
