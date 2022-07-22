---
title: 在 Outlook 加载项中获取或修改收件人
description: 了解如何在 Outlook 加载项中获取、设置或添加邮件或约会的收件人。
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: bcc4a76ef89e3bfaf7e884ad2fa4e1595782c62f
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958318"
---
# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a>在 Outlook 中撰写约会或邮件时获取、设置或添加收件人

Office JavaScript API 提供异步方法 ([Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1))、 [Recipients.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1)) 或 [Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))) 分别以约会或邮件的撰写形式获取、设置或添加收件人。 这些异步方法仅可用于撰写加载项。若要使用这些方法，请确保已适当地为 Outlook 设置外接程序清单以激活撰写窗体中的外接程序，如 [创建撰写窗体的 Outlook 外](compose-scenario.md)接程序中所述。

部分表示约会或邮件中的收件人的属性在撰写窗体和阅读窗体中可以进行阅读访问。这些属性包括约会的 [optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 和 [requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)，以及邮件的 [cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 和 [to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)。

在阅读窗体中，你可以直接从父对象访问属性，例如：

```js
item.cc
```

但在撰写窗体中，由于用户和外接程序可以同时插入或更改收件人，因此必须使用异步方法 `getAsync` 来获取这些属性，如以下示例所示。

```js
item.cc.getAsync
```

这些属性只在撰写窗体（而非阅读窗体）中可进行写入访问。

与适用于 Office 的 JavaScript API 中的大多数异步方法一样，`getAsync``setAsync`并`addAsync`采用可选输入参数。 有关指定这些可选输入参数的详细信息，请参阅 [Office 外接程序中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline)中的[向异步方法传递可选参数](../develop/asynchronous-programming-in-office-add-ins.md)。

## <a name="get-recipients"></a>获取收件人

此部分显示的代码示例用于获取正在撰写的约会或邮件的收件人，并显示收件人的电子邮件地址。代码示例假设外接程序清单中有在撰写窗体中为约会或邮件激活外接程序的规则，如下所示。

```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

在 Office JavaScript API 中，由于表示约会的收件人 ( **optionalAttendees** 和 **requiredAttendees**) 的属性不同于消息 ([bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)、 **cc** **和)** 的属性，因此应首先使用 [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 属性来标识正在撰写的项目是约会还是消息。 在撰写模式下，约会和邮件的所有这些属性都是 [收件人](/javascript/api/outlook/office.recipients) 对象，因此你可以应用异步方法 `Recipients.getAsync`来获取相应的收件人。

若要使用 `getAsync`，请提供回调函数来检查异步 `getAsync` 调用返回的状态、结果和任何错误。 可以使用可选 _的 asyncContext_ 参数向回调函数提供任何参数。 回调函数返回 _asyncResult_ 输出参数。 可以使用 `status` [AsyncResult](/javascript/api/office/office.asyncresult) 参数对象的和`error`属性来检查异步调用的状态和任何错误消息，以及`value`获取实际收件人的属性。 收件人以 [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) 对象数组的形式表示。

请注意，由于 `getAsync` 该方法是异步的，如果后续操作依赖于成功获取收件人，则应组织代码，以便仅在异步调用成功完成时在相应的回调函数中启动此类操作。

> [!IMPORTANT]
> 该 `getAsync` 方法仅返回 Outlook 客户端解析的收件人。 已解析的收件人具有以下特征。
>
> - 如果收件人通讯簿中包含已保存的条目，Outlook 会将电子邮件地址解析为收件人保存的显示名称。
> - Teams 会议状态图标显示在收件人的姓名或电子邮件地址之前。
> - 收件人的姓名或电子邮件地址后会显示分号。
> - 收件人的姓名或电子邮件地址带有下划线或括在一个框中。
>
> 若要在电子邮件地址添加到邮件项后解析电子邮件地址，发件人必须使用 **Tab** 键或从自动完成列表中选择建议的联系人或电子邮件地址。

> [!NOTE]
> 在 Outlook 网页版 和 Windows 中，如果用户通过从联系人或个人资料卡激活联系人的电子邮件地址链接来创建新邮件，则外接程序的`Recipients.getAsync`呼叫将返回关联`EmailAddressDetails`对象属性中`displayName`联系人的电子邮件地址，而不是联系人保存的名称。
>
> 有关详细信息，请参阅 [相关的 GitHub 问题](https://github.com/OfficeDev/office-js/issues/2201)。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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
    let toRecipients, ccRecipients, bccRecipients;
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
    for (let i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="set-recipients"></a>设置收件人

此部分显示的代码示例会设置用户正在撰写的约会或邮件的收件人。 设置收件人将覆盖现有的全部收件人。 与之前获取撰写窗体中收件人的示例相似，此示例假设已在撰写窗体中为约会和邮件激活外接程序。 本示例首先验证组合项目是约会还是消息，因此，若要对表示约会或邮件收件人的相应属性应用异步方法 `Recipients.setAsync`。

调用 `setAsync`时，以以下格式之一为  _收件人_ 参数提供数组作为输入参数。

- 为 SMTP 地址的字符串数组。
- 字典的数组，每个字典都包含显示名称和电子邮件地址，如下面的代码示例中所示。
- 对象数组 `EmailAddressDetails` ，类似于方法返回的 `getAsync` 对象数组。
  
可以选择提供回调函数作为方法的输入参数 `setAsync` ，以确保任何依赖于成功设置收件人的代码仅在发生这种情况时才执行。 还可以使用可选 _的 asyncContext_ 参数为回调函数提供任何参数。 如果使用回调函数，则可以访问 _asyncResult_ 输出参数，并使用参数对象的`AsyncResult`**状态** 和 **错误** 属性来检查异步调用的状态和任何错误消息。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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
    let toRecipients, ccRecipients, bccRecipients;

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

如果不想覆盖约会或邮件中的任何现有收件人，则可以使用异步方法来追加收件人，而不是使用`Recipients.setAsync``Recipients.addAsync`此方法。 `addAsync` 工作方式类似 `setAsync` ，因为它需要 _收件人_ 输入参数。 可以选择使用 asyncContext 参数为回调提供回调函数和任何参数。 然后，可以使用回调函数的 _asyncResult_ 输出参数检查异步`addAsync`调用的状态、结果和任何错误。 以下示例检查正在撰写的项目是否是约会，并为该约会追加两个必需参与者。

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
