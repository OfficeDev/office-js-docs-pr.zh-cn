---
title: 在 Outlook 加载项中获取或修改收件人
description: 了解如何在 Outlook 加载项中获取、设置或添加邮件或约会的收件人。
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 396f425f639c0d7043154ccfe1ddea16a236f993
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325431"
---
# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="335c2-103">在 Outlook 中撰写约会或邮件时获取、设置或添加收件人</span><span class="sxs-lookup"><span data-stu-id="335c2-103">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>


<span data-ttu-id="335c2-104">Office JavaScript API 提供了异步方法（[getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-)、 [setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)或[addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)），以在约会或邮件的撰写窗体中分别获取、设置或添加收件人。</span><span class="sxs-lookup"><span data-stu-id="335c2-104">The Office JavaScript API provides asynchronous methods ([Recipients.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-), [Recipients.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-), or [Recipients.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)) to respectively get, set, or add recipients in a compose form of an appointment or message.</span></span> <span data-ttu-id="335c2-105">这些异步方法仅适用于撰写外接程序。若要使用这些方法，请确保已正确设置了 Outlook 以在撰写窗体中激活加载项的加载项清单，如[创建适用于撰写窗体的 Outlook 外接程序](compose-scenario.md)中所述。</span><span class="sxs-lookup"><span data-stu-id="335c2-105">These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="335c2-p102">部分表示约会或邮件中的收件人的属性在撰写窗体和阅读窗体中可以进行阅读访问。这些属性包括约会的 [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 和 [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)，以及邮件的 [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 和 [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)。</span><span class="sxs-lookup"><span data-stu-id="335c2-p102">Some of the properties that represent recipients in an appointment or message are available for read access in a compose form and in a read form. These properties include  [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) for appointments, and [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), and  [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) for messages.</span></span>

<span data-ttu-id="335c2-108">在阅读窗体中，你可以直接从父对象访问属性，例如：</span><span class="sxs-lookup"><span data-stu-id="335c2-108">In a read form, you can access the property directly from the parent object, such as:</span></span>

```js
item.cc
```

<span data-ttu-id="335c2-109">但在撰写窗体中，由于用户和外接程序都可以同时插入或更改收件人，因此必须使用异步方法`getAsync`获取这些属性，如下面的示例所示：</span><span class="sxs-lookup"><span data-stu-id="335c2-109">But in a compose form, because both the user and your add-in can be inserting or changing a recipient at the same time, you must use the asynchronous method `getAsync` to get these properties, as in the following example:</span></span>


```js
item.cc.getAsync
```

<span data-ttu-id="335c2-110">这些属性只在撰写窗体（而非阅读窗体）中可进行写入访问。</span><span class="sxs-lookup"><span data-stu-id="335c2-110">These properties are available for write access in only compose forms and not read forms.</span></span>

<span data-ttu-id="335c2-111">与适用于 Office `getAsync` `setAsync`的 JavaScript API 中的大多数异步方法一样，和`addAsync`采用可选的输入参数。</span><span class="sxs-lookup"><span data-stu-id="335c2-111">As with most asynchronous methods in the JavaScript API for Office, `getAsync`, `setAsync`, and `addAsync` take optional input parameters.</span></span> <span data-ttu-id="335c2-112">有关指定这些可选输入参数的详细信息，请参阅 [Office 外接程序中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline)中的[向异步方法传递可选参数](../develop/asynchronous-programming-in-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="335c2-112">For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-recipients"></a><span data-ttu-id="335c2-113">获取收件人</span><span class="sxs-lookup"><span data-stu-id="335c2-113">Get recipients</span></span>


<span data-ttu-id="335c2-p104">此部分显示的代码示例用于获取正在撰写的约会或邮件的收件人，并显示收件人的电子邮件地址。代码示例假设外接程序清单中有在撰写窗体中为约会或邮件激活外接程序的规则，如下所示。</span><span class="sxs-lookup"><span data-stu-id="335c2-p104">This section shows a code sample that gets the recipients of the appointment or message that is being composed, and displays the email addresses of the recipients. The code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

<span data-ttu-id="335c2-116">在 Office JavaScript API 中，由于代表约会（ **optionalAttendees**和**requiredAttendees**）的收件人的属性不同于邮件的收件人（[bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)、 **cc**和**to**），因此应首先使用[item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)属性来标识正在撰写的项目是约会还是消息。</span><span class="sxs-lookup"><span data-stu-id="335c2-116">In the Office JavaScript API, because the properties that represent the recipients of an appointment ( **optionalAttendees** and **requiredAttendees**) are different from those of a message ([bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), **cc**, and **to**), you should first use the [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property to identify whether the item being composed is an appointment or message.</span></span> <span data-ttu-id="335c2-117">在撰写模式中，约会和邮件的所有属性都是[收件人](/javascript/api/outlook/office.Recipients)对象，因此您可以应用异步方法`Recipients.getAsync`来获取相应的收件人。</span><span class="sxs-lookup"><span data-stu-id="335c2-117">In compose mode, all these properties of appointments and messages are [Recipients](/javascript/api/outlook/office.Recipients) objects, so you can then apply the asynchronous method, `Recipients.getAsync`, to get the corresponding recipients.</span></span>

<span data-ttu-id="335c2-118">若要`getAsync`使用，请提供用于检查异步`getAsync`调用返回的状态、结果和任何错误的回调方法。</span><span class="sxs-lookup"><span data-stu-id="335c2-118">To use `getAsync` provide a callback method to check for the status, results, and any error returned by the asynchronous `getAsync` call.</span></span> <span data-ttu-id="335c2-119">您可以使用可选 _asyncContext_ 形参为回调方法提供任意实参。</span><span class="sxs-lookup"><span data-stu-id="335c2-119">You can provide any arguments to the callback method using the optional _asyncContext_ parameter.</span></span> <span data-ttu-id="335c2-120">回调方法会返回 _asyncResult_ 输出形参。</span><span class="sxs-lookup"><span data-stu-id="335c2-120">The callback method returns an _asyncResult_ output parameter.</span></span> <span data-ttu-id="335c2-121">您可以使用[AsyncResult](/javascript/api/office/office.asyncresult) parameter `error`对象的`status`和属性检查异步调用的状态和任何错误消息，以及获取实际收件人的`value`属性。</span><span class="sxs-lookup"><span data-stu-id="335c2-121">You can use the `status` and `error` properties of the [AsyncResult](/javascript/api/office/office.asyncresult) parameter object to check for status and any error messages of the asynchronous call, and the `value` property to get the actual recipients.</span></span> <span data-ttu-id="335c2-122">收件人以 [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) 对象数组的形式表示。</span><span class="sxs-lookup"><span data-stu-id="335c2-122">Recipients are represented as an array of [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) objects.</span></span>

<span data-ttu-id="335c2-123">请注意，由于`getAsync`方法是异步的，因此，如果后续操作依赖于成功获取收件人，则应将代码组织为在异步调用成功完成时，仅在相应的回调方法中启动此类操作。</span><span class="sxs-lookup"><span data-stu-id="335c2-123">Note that because the `getAsync` method is asynchronous, if there are subsequent actions that depend on successfully getting the recipients, you should organize your code to start such actions only in the corresponding callback method when the asynchronous call has successfully completed.</span></span>




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


## <a name="set-recipients"></a><span data-ttu-id="335c2-124">设置收件人</span><span class="sxs-lookup"><span data-stu-id="335c2-124">Set recipients</span></span>


<span data-ttu-id="335c2-125">此部分显示的代码示例会设置用户正在撰写的约会或邮件的收件人。</span><span class="sxs-lookup"><span data-stu-id="335c2-125">This section shows a code sample that sets the recipients of the appointment or message that is being composed by the user.</span></span> <span data-ttu-id="335c2-126">设置收件人将覆盖现有的全部收件人。</span><span class="sxs-lookup"><span data-stu-id="335c2-126">Setting recipients overwrites any existing recipients.</span></span> <span data-ttu-id="335c2-127">与之前获取撰写窗体中收件人的示例相似，此示例假设已在撰写窗体中为约会和邮件激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="335c2-127">Similar to the previous example that gets recipients in a compose form, this example assumes that the add-in is activated in compose forms for appointments and messages.</span></span> <span data-ttu-id="335c2-128">本示例首先验证组合项目是否为约会或邮件，以便在代表约会或邮件收件人的相应`Recipients.setAsync`属性上应用异步方法。</span><span class="sxs-lookup"><span data-stu-id="335c2-128">This example first verifies if the composed item is an appointment or message, so to apply the asynchronous method, `Recipients.setAsync`, on the appropriate properties that represent recipients of the appointment or message.</span></span>

<span data-ttu-id="335c2-129">调用`setAsync`时，请提供一个数组作为_收件人_参数的输入参数，采用以下格式之一：</span><span class="sxs-lookup"><span data-stu-id="335c2-129">When calling `setAsync`, provide an array as input argument for the  _recipients_ parameter, in one of the following formats:</span></span>


- <span data-ttu-id="335c2-130">为 SMTP 地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="335c2-130">An array of strings that are SMTP addresses.</span></span>
    
- <span data-ttu-id="335c2-131">字典的数组，每个字典都包含显示名称和电子邮件地址，如下面的代码示例中所示。</span><span class="sxs-lookup"><span data-stu-id="335c2-131">An array of dictionaries, each containing a display name and email address, as shown in the following code sample.</span></span>
    
- <span data-ttu-id="335c2-132">对象的`EmailAddressDetails`数组，与`getAsync`方法返回的数组相似。</span><span class="sxs-lookup"><span data-stu-id="335c2-132">An array of `EmailAddressDetails` objects, similar to the one returned by the `getAsync` method.</span></span>
    
<span data-ttu-id="335c2-133">您可以选择提供回调方法作为方法的`setAsync`输入参数，以确保仅当出现这种情况时，才会执行任何依赖于成功设置收件人的代码。</span><span class="sxs-lookup"><span data-stu-id="335c2-133">You can optionally provide a callback method as an input argument to the `setAsync` method, to make sure any code that depends on successfully setting the recipients would execute only when that happens.</span></span> <span data-ttu-id="335c2-134">还可以为使用可选 _asyncContext_ 形参的回调方法提供任意实参。</span><span class="sxs-lookup"><span data-stu-id="335c2-134">You can also provide any arguments for the callback method using the optional _asyncContext_ parameter.</span></span> <span data-ttu-id="335c2-135">如果使用的是回调方法，则可以访问_asyncResult_输出参数，并使用`AsyncResult` parameter 对象的**status**和**error**属性检查异步调用的状态和任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="335c2-135">If you use a callback method, you can access an _asyncResult_ output parameter, and use the **status** and **error** properties of the `AsyncResult` parameter object to check for status and any error messages of the asynchronous call.</span></span>




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


## <a name="add-recipients"></a><span data-ttu-id="335c2-136">添加收件人</span><span class="sxs-lookup"><span data-stu-id="335c2-136">Add recipients</span></span>

<span data-ttu-id="335c2-137">如果您不想覆盖约会或邮件中的任何现有收件人，而不是使用`Recipients.setAsync`，则可以使用`Recipients.addAsync`异步方法来追加收件人。</span><span class="sxs-lookup"><span data-stu-id="335c2-137">If you do not want to overwrite any existing recipients in an appointment or message, instead of using `Recipients.setAsync`, you can use the `Recipients.addAsync` asynchronous method to append recipients.</span></span> <span data-ttu-id="335c2-138">`addAsync`的工作方式`setAsync`类似于，它需要_收件人_输入参数。</span><span class="sxs-lookup"><span data-stu-id="335c2-138">`addAsync` works similarly as `setAsync` in that it requires a _recipients_ input argument.</span></span> <span data-ttu-id="335c2-139">还可以选择使用 asyncContext 形参为回调提供回调方法和任意实参。</span><span class="sxs-lookup"><span data-stu-id="335c2-139">You can optionally provide a callback method, and any arguments for the callback using the asyncContext parameter.</span></span> <span data-ttu-id="335c2-140">然后，可以使用回调方法的`addAsync` _asyncResult_ output 参数检查异步调用的状态、结果和任何错误。</span><span class="sxs-lookup"><span data-stu-id="335c2-140">You can then check the status, result, and any error of the asynchronous `addAsync` call by using the _asyncResult_ output parameter of the callback method.</span></span> <span data-ttu-id="335c2-141">以下示例检查正在撰写的项目是否是约会，并为该约会追加两个必需参与者。</span><span class="sxs-lookup"><span data-stu-id="335c2-141">The following example checks if the item being composed is an appointment, and appends two required attendees to the appointment.</span></span>


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


## <a name="see-also"></a><span data-ttu-id="335c2-142">另请参阅</span><span class="sxs-lookup"><span data-stu-id="335c2-142">See also</span></span>

- [<span data-ttu-id="335c2-143">在 Outlook 撰写窗体中获取并设置项数据</span><span class="sxs-lookup"><span data-stu-id="335c2-143">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)
- [<span data-ttu-id="335c2-144">在阅读或撰写窗体中获取并设置 Outlook 项目数据</span><span class="sxs-lookup"><span data-stu-id="335c2-144">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)
- [<span data-ttu-id="335c2-145">创建适用于撰写窗体的 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="335c2-145">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="335c2-146">Office 外接程序中的异步编程</span><span class="sxs-lookup"><span data-stu-id="335c2-146">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="335c2-147">在 Outlook 中撰写约会或邮件时获取或设置主题</span><span class="sxs-lookup"><span data-stu-id="335c2-147">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)
- [<span data-ttu-id="335c2-148">在 Outlook 中撰写约会或邮件时将数据插入到正文中</span><span class="sxs-lookup"><span data-stu-id="335c2-148">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)
- [<span data-ttu-id="335c2-149">在 Outlook 中撰写约会时获取或设置位置</span><span class="sxs-lookup"><span data-stu-id="335c2-149">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md)
- [<span data-ttu-id="335c2-150">在 Outlook 中撰写约会时获取或设置时间</span><span class="sxs-lookup"><span data-stu-id="335c2-150">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
