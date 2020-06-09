---
title: 在 Outlook 加载项的正文中插入数据
description: 了解如何将数据插入到 Outlook 加载项的邮件或约会的正文中。
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: e8100e036d29c13f12aedddd4436cf35569309cf
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609095"
---
# <a name="insert-data-in-the-body-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="7b1a6-103">在 Outlook 中撰写约会或邮件时将数据插入到正文中</span><span class="sxs-lookup"><span data-stu-id="7b1a6-103">Insert data in the body when composing an appointment or message in Outlook</span></span>

<span data-ttu-id="7b1a6-p101">您可以使用异步方法（[Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-)、[Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-)、[Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-)、[Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-) 和 [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)），以获取正文类型并在用户正在撰写的约会或邮件项目的正文中插入数据。这些异步方法仅适用于撰写外接程序。若要使用这些方法，请确保已正确设置外接程序清单，以便 Outlook 可以在撰写窗体中激活外接程序，如[创建适用于撰写窗体的 Outlook 外接程序](compose-scenario.md)中所述。</span><span class="sxs-lookup"><span data-stu-id="7b1a6-p101">You can use the asynchronous methods ([Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-), [Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-), [Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-), [Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-) and [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)) to get the body type and insert data in the body of an appointment or message item that the user is composing. These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately so that Outlook activates your add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="7b1a6-p102">在 Outlook 中，用户可以创建文本、HTML 或 RTF 格式的邮件，还可以创建 HTML 格式的约会。在插入之前，你应始终先通过调用 **getTypeAsync** 来验证支持的项格式。**getTypeAsync** 返回的值取决于原始项格式，以及对以 HTML 格式编辑的设备操作系统和主机的支持 (1)。然后相应地设置 **prependAsync** 或 **setSelectedDataAsync** 的 _coercionType_ 参数 (2) 以插入数据，如下表中所示。如果不指定自变量，**prependAsync** 和 **setSelectedDataAsync** 会假定要插入的数据为文本格式。</span><span class="sxs-lookup"><span data-stu-id="7b1a6-p102">In Outlook, a user can create a message in text, HTML, or Rich Text Format (RTF), and can create an appointment in HTML format. Before inserting, you should always first verify the supported item format by calling **getTypeAsync**, as you may need to take additional steps. The value that **getTypeAsync** returns depends on the original item format, as well as the support of the device operating system and host to editing in HTML format (1). Then set the  _coercionType_ parameter of **prependAsync** or **setSelectedDataAsync** accordingly (2) to insert the data, as shown in the following table. If you don't specify an argument, **prependAsync** and **setSelectedDataAsync** assume the data to insert is in text format.</span></span>

<br/>

|<span data-ttu-id="7b1a6-111">**要插入的数据**</span><span class="sxs-lookup"><span data-stu-id="7b1a6-111">**Data to insert**</span></span>|<span data-ttu-id="7b1a6-112">**getTypeAsync 返回的项目格式**</span><span class="sxs-lookup"><span data-stu-id="7b1a6-112">**Item format returned by getTypeAsync**</span></span>|<span data-ttu-id="7b1a6-113">**使用此 coercionType**</span><span class="sxs-lookup"><span data-stu-id="7b1a6-113">**Use this coercionType**</span></span>|
|:-----|:-----|:-----|
|<span data-ttu-id="7b1a6-114">文本</span><span class="sxs-lookup"><span data-stu-id="7b1a6-114">Text</span></span>|<span data-ttu-id="7b1a6-115">文本 (1)</span><span class="sxs-lookup"><span data-stu-id="7b1a6-115">Text (1)</span></span>|<span data-ttu-id="7b1a6-116">文本</span><span class="sxs-lookup"><span data-stu-id="7b1a6-116">Text</span></span>|
|<span data-ttu-id="7b1a6-117">HTML</span><span class="sxs-lookup"><span data-stu-id="7b1a6-117">HTML</span></span>|<span data-ttu-id="7b1a6-118">文本 (1)</span><span class="sxs-lookup"><span data-stu-id="7b1a6-118">Text (1)</span></span>|<span data-ttu-id="7b1a6-119">文本 (2)</span><span class="sxs-lookup"><span data-stu-id="7b1a6-119">Text (2)</span></span>|
|<span data-ttu-id="7b1a6-120">文本</span><span class="sxs-lookup"><span data-stu-id="7b1a6-120">Text</span></span>|<span data-ttu-id="7b1a6-121">HTML</span><span class="sxs-lookup"><span data-stu-id="7b1a6-121">HTML</span></span>|<span data-ttu-id="7b1a6-122">文本/HTML</span><span class="sxs-lookup"><span data-stu-id="7b1a6-122">Text/HTML</span></span>|
|<span data-ttu-id="7b1a6-123">HTML</span><span class="sxs-lookup"><span data-stu-id="7b1a6-123">HTML</span></span>|<span data-ttu-id="7b1a6-124">HTML</span><span class="sxs-lookup"><span data-stu-id="7b1a6-124">HTML</span></span> |<span data-ttu-id="7b1a6-125">HTML</span><span class="sxs-lookup"><span data-stu-id="7b1a6-125">HTML</span></span>|

1.  <span data-ttu-id="7b1a6-126">在平板电脑和智能手机上，如果操作系统或主机不支持编辑 HTML 格式的项（最初以 HTML 创建），**getTypeAsync** 将返回 **Office.MailboxEnums.BodyType.Text**。</span><span class="sxs-lookup"><span data-stu-id="7b1a6-126">On tablets and smartphones, **getTypeAsync** returns **Office.MailboxEnums.BodyType.Text** if the operating system or host does not support editing an item, which was originally created in HTML, in HTML format.</span></span>

2.  <span data-ttu-id="7b1a6-p103">如果要插入的数据是 HTML 但 **getTypeAsync** 返回该项的文本类型，请将你的数据重新组织为文本并插入，此时 **Office.MailboxEnums.BodyType.Text** 为 _coercionType_。如果仅插入具有文本强制类型的 HTML 数据，主机会将 HTML 标记显示为文本。如果尝试插入 HTML 数据（此时 **Office.MailboxEnums.BodyType.Html** 为 _coercionType_），将收到错误。</span><span class="sxs-lookup"><span data-stu-id="7b1a6-p103">If your data to insert is HTML and **getTypeAsync** returns a text type for that item, reorganize your data as text and insert it with **Office.MailboxEnums.BodyType.Text** as _coercionType_. If you simply insert the HTML data with a text coercion type, the host would display the HTML tags as text. If you attempt to insert the HTML data with **Office.MailboxEnums.BodyType.Html** as _coercionType_, you will get an error.</span></span>

<span data-ttu-id="7b1a6-p104">除了_coercionType_之外，与 OFFICE JavaScript API 中的大多数异步方法一样， **getTypeAsync**、 **prependAsync**和**document.setselecteddataasync**采用其他可选的可选输入参数。有关指定这些可选输入参数的详细信息，请参阅在[Office 外接程序](../develop/asynchronous-programming-in-office-add-ins.md)中将[可选参数传递给](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline)异步编程方法。</span><span class="sxs-lookup"><span data-stu-id="7b1a6-p104">In addition to  _coercionType_, as with most asynchronous methods in the Office JavaScript API, **getTypeAsync**, **prependAsync** and **setSelectedDataAsync** take other optional input parameters. For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="insert-data-at-the-current-cursor-position"></a><span data-ttu-id="7b1a6-132">在当前光标位置插入数据</span><span class="sxs-lookup"><span data-stu-id="7b1a6-132">Insert data at the current cursor position</span></span>


<span data-ttu-id="7b1a6-133">此部分显示的代码示例使用 **getTypeAsync** 验证正在撰写的项的正文类型，然后使用 **setSelectedDataAsync** 在当前光标位置插入数据。</span><span class="sxs-lookup"><span data-stu-id="7b1a6-133">This section shows a code sample that uses **getTypeAsync** to verify the body type of the item that is being composed, and then uses **setSelectedDataAsync** to insert data in the current cursor location.</span></span>

<span data-ttu-id="7b1a6-p105">可以将回调方法和可选输入参数传递到 **getTypeAsync**，并获取 _asyncResult_ 输出参数中的任意状态和结果。如果该方法成功，你可以获取 [AsyncResult.value](/javascript/api/office/office.asyncresult#value) 属性中项正文的类型，即“文本”或“html”。</span><span class="sxs-lookup"><span data-stu-id="7b1a6-p105">You can pass a callback method and optional input parameters to **getTypeAsync**, and get any status and results in the  _asyncResult_ output parameter. If the method succeeds, you can get the type of the item body in the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property, which is either "text" or "html".</span></span>

<span data-ttu-id="7b1a6-p106">必须将数据字符串传递到 **setSelectedDataAsync**，作为输入参数。根据项正文的类型，你可以相应地将此数据字符串指定为文本或 HTML 格式。如上所述，还可以选择指定要插入到 _coercionType_ 参数中的数据的类型。此外，你可以提供回调方法及其任意参数，作为可选输入参数。</span><span class="sxs-lookup"><span data-stu-id="7b1a6-p106">You must pass a data string as an input parameter to **setSelectedDataAsync**. Depending on the type of the item body, you can specify this data string in text or HTML format accordingly. As mentioned above, you can optionally specify the type of the data to be inserted in the  _coercionType_ parameter. In addition, you can provide a callback method and any of its parameters as optional input parameters.</span></span>

<span data-ttu-id="7b1a6-p107">如果用户尚未将光标放置在项正文中，**setSelectedDataAsync** 会将数据插入到正文的顶部。如果用户已经在项正文中选择了文本，**setSelectedDataAsync** 会用你指定的数据替换所选文本。请注意，如果用户在撰写项的同时更改光标位置，**setSelectedDataAsync** 可能会失败。一次最多可以插入 1,000,000 个字符。</span><span class="sxs-lookup"><span data-stu-id="7b1a6-p107">If the user hasn't placed the cursor in the item body, **setSelectedDataAsync** inserts the data at the top of the body. If the user has selected text in the item body, **setSelectedDataAsync** replaces the selected text by the data you specify. Note that **setSelectedDataAsync** can fail if the user is simultaneously changing the cursor position while composing the item. The maximum number of characters you can insert at one time is 1,000,000 characters.</span></span>

<span data-ttu-id="7b1a6-144">此代码示例假定外接程序清单中的某个规则将在撰写窗体中为约会或邮件激活外接程序，如下所述。</span><span class="sxs-lookup"><span data-stu-id="7b1a6-144">This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>




```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set data in the body of the composed item.
        setItemBody();
    });
}


// Get the body type of the composed item, and set data in 
// in the appropriate data type in the item body.
function setItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(result.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Set data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of setSelectedDataAsync.
                    item.body.setSelectedDataAsync(
                        '<b> Kindly note we now open 7 days a week.</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.setSelectedDataAsync(
                        ' Kindly note we now open 7 days a week.',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="insert-data-at-the-beginning-of-the-item-body"></a><span data-ttu-id="7b1a6-145">在项正文的开头插入数据</span><span class="sxs-lookup"><span data-stu-id="7b1a6-145">Insert data at the beginning of the item body</span></span>


<span data-ttu-id="7b1a6-p108">你也可以使用 **prependAsync** 在项正文的开头部分插入数据，无论当前光标位置如何均可插入。除了插入点不同之外，**prependAsync** 和 **setSelectedDataAsync** 的工作原理相似：</span><span class="sxs-lookup"><span data-stu-id="7b1a6-p108">Alternatively, you can use **prependAsync** to insert data at the beginning of the item body and disregard the current cursor location. Other than the point of insertion, **prependAsync** and **setSelectedDataAsync** behave in similar ways:</span></span>


- <span data-ttu-id="7b1a6-148">如果要将 HTML 数据预置到邮件正文中，应先检查邮件正文的类型，以免将 HTML 数据预置到文本格式的邮件中。</span><span class="sxs-lookup"><span data-stu-id="7b1a6-148">If you are prepending HTML data in a message body, you should first check for the type of the message body to avoid prepending HTML data to a message in text format.</span></span>
    
- <span data-ttu-id="7b1a6-149">提供以下内容，作为 **prependAsync** 的输入参数：文本格式或 HTML 格式的数据字符串、要插入的数据的格式（可选）、回调方法及其任意参数。</span><span class="sxs-lookup"><span data-stu-id="7b1a6-149">Provide the following as input parameters to **prependAsync**: a data string in either text or HTML format, and optionally the format of the data to be inserted, a callback method and any of its parameters.</span></span>
    
- <span data-ttu-id="7b1a6-150">一次最多可以预置 1,000,000 个字符。</span><span class="sxs-lookup"><span data-stu-id="7b1a6-150">The maximum number of characters you can prepend at one time is 1,000,000 characters.</span></span>
    
<span data-ttu-id="7b1a6-p109">以下 JavaScript 代码是在约会和邮件撰写窗体中激活的示例加载项的一部分。该示例调用 **getTypeAsync**，以验证项正文的类型，如果项是约会或 HTML 邮件，则将 HTML 数据插入到项正文的顶部，否则插入文本格式的数据。</span><span class="sxs-lookup"><span data-stu-id="7b1a6-p109">The following JavaScript code is part of a sample add-in that is activated in compose forms of appointments and messages. The sample calls **getTypeAsync** to verify the type of the item body, inserts HTML data to the top of the item body if the item is an appointment or HTML message, otherwise inserts the data in text format.</span></span>




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Insert data in the top of the body of the composed 
        // item.
        prependItemBody();
    });
}

// Get the body type of the composed item, and prepend data  
// in the appropriate data type in the item body.
function prependItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Prepend data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of prependAsync.
                    item.body.prependAsync(
                        '<b>Greetings!</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.prependAsync(
                        'Greetings!',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="see-also"></a><span data-ttu-id="7b1a6-153">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7b1a6-153">See also</span></span>

- [<span data-ttu-id="7b1a6-154">在 Outlook 撰写窗体中获取并设置项数据</span><span class="sxs-lookup"><span data-stu-id="7b1a6-154">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)    
- [<span data-ttu-id="7b1a6-155">在阅读或撰写窗体中获取并设置 Outlook 项目数据</span><span class="sxs-lookup"><span data-stu-id="7b1a6-155">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)    
- [<span data-ttu-id="7b1a6-156">创建适用于撰写窗体的 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="7b1a6-156">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="7b1a6-157">Office 外接程序中的异步编程</span><span class="sxs-lookup"><span data-stu-id="7b1a6-157">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)    
- [<span data-ttu-id="7b1a6-158">在 Outlook 中撰写约会或邮件时获取、设置或添加收件人</span><span class="sxs-lookup"><span data-stu-id="7b1a6-158">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="7b1a6-159">在 Outlook 中撰写约会或邮件时获取或设置主题</span><span class="sxs-lookup"><span data-stu-id="7b1a6-159">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)  
- [<span data-ttu-id="7b1a6-160">在 Outlook 中撰写约会时获取或设置位置</span><span class="sxs-lookup"><span data-stu-id="7b1a6-160">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md) 
- [<span data-ttu-id="7b1a6-161">在 Outlook 中撰写约会时获取或设置时间</span><span class="sxs-lookup"><span data-stu-id="7b1a6-161">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
