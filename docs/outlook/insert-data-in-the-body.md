---
title: 在 Outlook 加载项的正文中插入数据
description: 了解如何将数据插入到 Outlook 加载项的邮件或约会的正文中。
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 082b3c5ebf4f8c93a485d438d55a5587f51a405e
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166007"
---
# <a name="insert-data-in-the-body-when-composing-an-appointment-or-message-in-outlook"></a>在 Outlook 中撰写约会或邮件时将数据插入到正文中

您可以使用异步方法（[Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-)、[Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-)、[Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-)、[Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-) 和 [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)），以获取正文类型并在用户正在撰写的约会或邮件项目的正文中插入数据。这些异步方法仅适用于撰写外接程序。若要使用这些方法，请确保已正确设置外接程序清单，以便 Outlook 可以在撰写窗体中激活外接程序，如[创建适用于撰写窗体的 Outlook 外接程序](compose-scenario.md)中所述。

在 Outlook 中，用户可以创建文本、HTML 或 RTF 格式的邮件，还可以创建 HTML 格式的约会。在插入之前，你应始终先通过调用 **getTypeAsync** 来验证支持的项格式。**getTypeAsync** 返回的值取决于原始项格式，以及对以 HTML 格式编辑的设备操作系统和主机的支持 (1)。然后相应地设置 **prependAsync** 或 **setSelectedDataAsync** 的 _coercionType_ 参数 (2) 以插入数据，如下表中所示。如果不指定自变量，**prependAsync** 和 **setSelectedDataAsync** 会假定要插入的数据为文本格式。

<br/>

|**要插入的数据**|**getTypeAsync 返回的项目格式**|**使用此 coercionType**|
|:-----|:-----|:-----|
|文本|文本 (1)|文本|
|HTML|文本 (1)|文本 (2)|
|文本|HTML|文本/HTML|
|HTML|HTML |HTML|

1.  在平板电脑和智能手机上，如果操作系统或主机不支持编辑 HTML 格式的项（最初以 HTML 创建），**getTypeAsync** 将返回 **Office.MailboxEnums.BodyType.Text**。

2.  如果要插入的数据是 HTML 但 **getTypeAsync** 返回该项的文本类型，请将你的数据重新组织为文本并插入，此时 **Office.MailboxEnums.BodyType.Text** 为 _coercionType_。如果仅插入具有文本强制类型的 HTML 数据，主机会将 HTML 标记显示为文本。如果尝试插入 HTML 数据（此时 **Office.MailboxEnums.BodyType.Html** 为 _coercionType_），将收到错误。

除 _coercionType_ 以外，与适用于 Office 的 JavaScript API 中的大多数异步方法相似，**getTypeAsync**、**prependAsync** 和 **setSelectedDataAsync** 采用其他可选输入参数。有关指定这些可选输入参数的详细信息，请参阅 [Office 加载项中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md)中的[向异步方法传递可选参数](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline)。


## <a name="insert-data-at-the-current-cursor-position"></a>在当前光标位置插入数据


此部分显示的代码示例使用 **getTypeAsync** 验证正在撰写的项的正文类型，然后使用 **setSelectedDataAsync** 在当前光标位置插入数据。

可以将回调方法和可选输入参数传递到 **getTypeAsync**，并获取 _asyncResult_ 输出参数中的任意状态和结果。如果该方法成功，你可以获取 [AsyncResult.value](/javascript/api/office/office.asyncresult#value) 属性中项正文的类型，即“文本”或“html”。

必须将数据字符串传递到 **setSelectedDataAsync**，作为输入参数。根据项正文的类型，你可以相应地将此数据字符串指定为文本或 HTML 格式。如上所述，还可以选择指定要插入到 _coercionType_ 参数中的数据的类型。此外，你可以提供回调方法及其任意参数，作为可选输入参数。

如果用户尚未将光标放置在项正文中，**setSelectedDataAsync** 会将数据插入到正文的顶部。如果用户已经在项正文中选择了文本，**setSelectedDataAsync** 会用你指定的数据替换所选文本。请注意，如果用户在撰写项的同时更改光标位置，**setSelectedDataAsync** 可能会失败。一次最多可以插入 1,000,000 个字符。

此代码示例假定外接程序清单中的某个规则将在撰写窗体中为约会或邮件激活外接程序，如下所述。




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


## <a name="insert-data-at-the-beginning-of-the-item-body"></a>在项正文的开头插入数据


你也可以使用 **prependAsync** 在项正文的开头部分插入数据，无论当前光标位置如何均可插入。除了插入点不同之外，**prependAsync** 和 **setSelectedDataAsync** 的工作原理相似：


- 如果要将 HTML 数据预置到邮件正文中，应先检查邮件正文的类型，以免将 HTML 数据预置到文本格式的邮件中。
    
- 提供以下内容，作为 **prependAsync** 的输入参数：文本格式或 HTML 格式的数据字符串、要插入的数据的格式（可选）、回调方法及其任意参数。
    
- 一次最多可以预置 1,000,000 个字符。
    
以下 JavaScript 代码是在约会和邮件撰写窗体中激活的示例加载项的一部分。该示例调用 **getTypeAsync**，以验证项正文的类型，如果项是约会或 HTML 邮件，则将 HTML 数据插入到项正文的顶部，否则插入文本格式的数据。




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


## <a name="see-also"></a>另请参阅

- [在 Outlook 撰写窗体中获取并设置项数据](get-and-set-item-data-in-a-compose-form.md)    
- [在阅读或撰写窗体中获取并设置 Outlook 项目数据](item-data.md)    
- [创建适用于撰写窗体的 Outlook 外接程序](compose-scenario.md)    
- [Office 外接程序中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md)    
- [在 Outlook 中撰写约会或邮件时获取、设置或添加收件人](get-set-or-add-recipients.md)  
- [在 Outlook 中撰写约会或邮件时获取或设置主题](get-or-set-the-subject.md)  
- [在 Outlook 中撰写约会时获取或设置位置](get-or-set-the-location-of-an-appointment.md) 
- [在 Outlook 中撰写约会时获取或设置时间](get-or-set-the-time-of-an-appointment.md)
    
