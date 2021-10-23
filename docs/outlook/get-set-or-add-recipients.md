---
title: 在 Outlook 加载项中获取或修改收件人
description: 了解如何在 Outlook 加载项中获取、设置或添加邮件或约会的收件人。
ms.date: 10/15/2021
ms.localizationpriority: medium
ms.openlocfilehash: c85a49ea3c409b64e0bd62f3eae3aa79dd614568
ms.sourcegitcommit: e4d98eb90e516b9c90e3832f3212caf48691acf6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/22/2021
ms.locfileid: "60537455"
---
# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a>在 Outlook 中撰写约会或邮件时获取、设置或添加收件人


Office JavaScript API 提供了异步方法 ([Recipients.getAsync、Recipients.setAsync](/javascript/api/outlook/office.recipients#getAsync_options__callback_)或[Recipients.addAsync](/javascript/api/outlook/office.recipients#addAsync_recipients__options__callback_)) 分别获取、设置或添加约会或邮件撰写窗体中的收件人。 [](/javascript/api/outlook/office.recipients#setAsync_recipients__options__callback_) 这些异步方法仅适用于撰写加载项。若要使用这些方法，请确保为 Outlook 设置相应的外接程序清单以在撰写窗体中激活外接程序，如为撰写窗体创建[Outlook](compose-scenario.md)外接程序中所述。

部分表示约会或邮件中的收件人的属性在撰写窗体和阅读窗体中可以进行阅读访问。这些属性包括约会的 [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 和 [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)，以及邮件的 [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 和 [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)。

在阅读窗体中，你可以直接从父对象访问属性，例如：

```js
item.cc
```

但在撰写窗体中，由于用户和外接程序可以同时插入或更改收件人，因此您必须使用异步方法获取这些属性， `getAsync` 如以下示例所示。


```js
item.cc.getAsync
```

这些属性只在撰写窗体（而非阅读窗体）中可进行写入访问。

与 JavaScript API 中的大多数异步方法一样，Office、 、 和 `getAsync` `setAsync` `addAsync` 采用可选输入参数。 有关指定这些可选输入参数的详细信息，请参阅 [Office 外接程序中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline)中的[向异步方法传递可选参数](../develop/asynchronous-programming-in-office-add-ins.md)。


## <a name="get-recipients"></a>获取收件人


此部分显示的代码示例用于获取正在撰写的约会或邮件的收件人，并显示收件人的电子邮件地址。代码示例假设外接程序清单中有在撰写窗体中为约会或邮件激活外接程序的规则，如下所示。


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

在 Office JavaScript API 中，由于表示约会收件人的属性 ( **optionalAttendees** 和 **requiredAttendees**) 与邮件 ([bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)、 **cc** 和 **)** 的属性不同，因此应首先使用 [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)属性确定正在撰写的项目是约会还是邮件。 在撰写模式下，约会和邮件的所有这些属性都是 [Recipients](/javascript/api/outlook/office.Recipients) 对象，因此您可以应用异步方法 ， `Recipients.getAsync` 获取相应的收件人。

若要 `getAsync` 使用 提供回调方法，请检查异步调用返回的状态、结果和任何 `getAsync` 错误。 您可以使用可选 _asyncContext_ 形参为回调方法提供任意实参。 回调方法会返回 _asyncResult_ 输出形参。 可以使用 AsyncResult 参数对象的 和 属性检查异步调用的状态和任何错误消息，并使用 属性 `status` `error` [](/javascript/api/office/office.asyncresult) `value` 获取实际收件人。 收件人以 [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) 对象数组的形式表示。

请注意，由于 方法是异步的，因此，如果存在依赖成功获取收件人的后续操作，则应该仅在异步调用成功完成时，组织代码以仅在相应的回调方法中启动 `getAsync` 此类操作。

> [!IMPORTANT]
> 在 Outlook 网页版 中，如果用户通过从联系人或个人资料卡片激活联系人的电子邮件地址链接创建了一封新邮件，则外接程序的呼叫当前不会在关联对象的 属性中返回值。 `Recipients.getAsync` `displayName` `EmailAddressDetails`
> 有关更多详细信息，请参阅相关[GitHub问题](https://github.com/OfficeDev/office-js-docs-pr/issues/2962)。

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get all the recipients of the composed item.
        getAllRecipients();
    });
}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous 
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients. 
            write ('To-recipients of the item:');
            displayAddresses(asyncResult);
        }    
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            write ('Cc-recipients of the item:');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get bcc-recipients of the item completed.
            // Display the email addresses of the bcc-recipients.
            write ('Bcc-recipients of the item:');
            displayAddresses(asyncResult);
        }
                        
        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
    for (var i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="set-recipients"></a>设置收件人


此部分显示的代码示例会设置用户正在撰写的约会或邮件的收件人。 设置收件人将覆盖现有的全部收件人。 与之前获取撰写窗体中收件人的示例相似，此示例假设已在撰写窗体中为约会和邮件激活外接程序。 本示例首先验证撰写的项目是约会还是邮件，以便对表示约会或邮件收件人的适当属性应用异步 `Recipients.setAsync` 方法 。

调用 `setAsync` 时，以下列格式之一提供数组作为  _recipients_ 参数的输入参数。


- 为 SMTP 地址的字符串数组。
    
- 字典的数组，每个字典都包含显示名称和电子邮件地址，如下面的代码示例中所示。
    
- 对象数组 `EmailAddressDetails` ，类似于方法返回 `getAsync` 的对象数组。
    
您可以选择提供回调方法作为方法的输入参数，以确保依赖于成功设置收件人的任何代码仅在发生这种情况 `setAsync` 时执行。 还可以为使用可选 _asyncContext_ 形参的回调方法提供任意实参。 如果使用回调方法，可以访问 _asyncResult_ 输出参数，并使用参数对象的 **status** 和 **error** 属性检查异步调用的状态和 `AsyncResult` 任何错误消息。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set recipients of the composed item.
        setRecipients();
    });
}

// Set the display name and email addresses of the recipients of 
// the composed item.
function setRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;

    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    toRecipients.setAsync(
        [{
            "displayName":"Graham Durkin", 
            "emailAddress":"graham@contoso.com"
         },
         {
            "displayName" : "Donnie Weinberg",
            "emailAddress" : "donnie@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
             "displayName":"Perry Horning", 
             "emailAddress":"perry@contoso.com"
         },
         {
             "displayName" : "Guy Montenegro",
             "emailAddress" : "guy@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.


    // If the item has the bcc field, i.e., item is message,
    // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                 "displayName":"Lewis Cate", 
                 "emailAddress":"lewis@contoso.com"
             },
             {
                 "displayName" : "Francisco Stitt",
                 "emailAddress" : "francisco@contoso.com"
             }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to set bcc-recipients of the item completed.
                    // Do whatever appropriate for your scenario.
                }
        }); // End bcc setAsync.
    }
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```


## <a name="add-recipients"></a>添加收件人

如果不希望覆盖约会或邮件中任何现有收件人，而不是使用 ，可以使用异步 `Recipients.setAsync` `Recipients.addAsync` 方法追加收件人。 `addAsync` 其工作方式类似 `setAsync` ，因为它需要 _收件人_ 输入参数。 还可以选择使用 asyncContext 形参为回调提供回调方法和任意实参。 然后，可以使用回调方法的 `addAsync` _asyncResult_ 输出参数检查异步调用的状态、结果和任何错误。 以下示例检查正在撰写的项目是否是约会，并为该约会追加两个必需参与者。


```js
// Add specified recipients as required attendees of
// the composed appointment. 
function addAttendees() {
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName":"Kristie Jensen", 
            "emailAddress":"kristie@contoso.com"
         },
         {
            "displayName" : "Pansy Valenzuela",
            "emailAddress" : "pansy@contoso.com"
          }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to add attendees completed.
                // Do whatever appropriate for your scenario.
            }
        }); // End addAsync.
    }
}
```


## <a name="see-also"></a>另请参阅

- [在 Outlook 撰写窗体中获取并设置项数据](get-and-set-item-data-in-a-compose-form.md)
- [在阅读或撰写窗体中获取并设置 Outlook 项目数据](item-data.md)
- [创建适用于撰写窗体的 Outlook 外接程序](compose-scenario.md)
- [Office 外接程序中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md)
- [在 Outlook 中撰写约会或邮件时获取或设置主题](get-or-set-the-subject.md)
- [在 Outlook 中撰写约会或邮件时将数据插入到正文中](insert-data-in-the-body.md)
- [在 Outlook 中撰写约会时获取或设置位置](get-or-set-the-location-of-an-appointment.md)
- [在 Outlook 中撰写约会时获取或设置时间](get-or-set-the-time-of-an-appointment.md)
    
