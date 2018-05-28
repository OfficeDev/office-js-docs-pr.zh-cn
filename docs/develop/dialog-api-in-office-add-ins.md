---
title: '? Office ?????????? API'
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b026c3c5871372c52d0b44e36c01fc44a3d2bf04
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="use-the-dialog-api-in-your-office-add-ins"></a><span data-ttu-id="239bb-102">? Office ????????? API</span><span class="sxs-lookup"><span data-stu-id="239bb-102">Use the Dialog API in your Office Add-ins</span></span>

<span data-ttu-id="239bb-p101">??? Office ???????[??? API](https://dev.office.com/reference/add-ins/shared/officeui) ???????????????? Office ?????????? API ????</span><span class="sxs-lookup"><span data-stu-id="239bb-p101">You can use the [Dialog API](https://dev.office.com/reference/add-ins/shared/officeui) to open dialog boxes in your Office Add-in. This article provides guidance for using the Dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="239bb-p102">??????? API ????????????[??? API ???](https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets)????Word?Excel?PowerPoint ? Outlook ????? API?</span><span class="sxs-lookup"><span data-stu-id="239bb-p102">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.</span></span>

> <span data-ttu-id="239bb-107">??? API ????????? Google ? Facebook ??????????</span><span class="sxs-lookup"><span data-stu-id="239bb-107">A primary scenario for the Dialog APIs is to enable authentication with a resource such as Google or Facebook.</span></span> <span data-ttu-id="239bb-108">?????????? Microsoft Graph ?? Office ??????????? Office 365 ? OneDrive????????????? API?</span><span class="sxs-lookup"><span data-stu-id="239bb-108">If your add-in requires data about the Office user or their resources accessible through Microsoft Graph, such as Office 365 or OneDrive, we recommend that you use the single sign-on API whenever you can.</span></span> <span data-ttu-id="239bb-109">???????? API??????? Dialog API?</span><span class="sxs-lookup"><span data-stu-id="239bb-109">If you use the APIs for single sign-on, then you will not need the Dialog API.</span></span> <span data-ttu-id="239bb-110">??????????[? Office ?????????](sso-in-office-add-ins.md)?</span><span class="sxs-lookup"><span data-stu-id="239bb-110">For details, see [Enable single sign-on for Office Add-ins](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="239bb-111">????????/?????/[?????](../design/add-in-commands.md)???????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-111">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="239bb-112">???????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-112">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="239bb-113">???????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-113">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="239bb-114">????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-114">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="239bb-p104">??????? UI ??????????????????????????????????????????????????????????????????? [Excel ??? JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) ???</span><span class="sxs-lookup"><span data-stu-id="239bb-p104">Because overlapping UI elements are discouraged, avoid opening a dialog from a task pane unless your scenario requires it. When you consider how to use the surface area of a task pane, note that task panes can be tabbed. For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="239bb-118">???????????</span><span class="sxs-lookup"><span data-stu-id="239bb-118">The following image shows an example of a dialog box.</span></span>

![?????](../images/auth-o-dialog-open.png)

<span data-ttu-id="239bb-p105">???????????????????????????????????*???*????????????????? Office ????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p105">Note that the dialog box always opens in the center of the screen. The user can move and resize it. The window is *nonmodal*--a user can continue to interact with both the document in the host Office application and with the host page in the task pane, if there is one.</span></span>

## <a name="dialog-api-scenarios"></a><span data-ttu-id="239bb-123">Dialog API ????</span><span class="sxs-lookup"><span data-stu-id="239bb-123">Dialog API scenarios</span></span>

<span data-ttu-id="239bb-124">Office JavaScript API ??????????? [Office.context.ui ????](https://dev.office.com/reference/add-ins/shared/officeui)??? [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) ????????</span><span class="sxs-lookup"><span data-stu-id="239bb-124">The Office JavaScript APIs support the following scenarios with a [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) object and two functions in the [Office.context.ui namespace](https://dev.office.com/reference/add-ins/shared/officeui).</span></span>

### <a name="open-a-dialog-box"></a><span data-ttu-id="239bb-125">?????</span><span class="sxs-lookup"><span data-stu-id="239bb-125">Open a dialog box</span></span>

<span data-ttu-id="239bb-p106">?????????????????? [displayDialogAsync](https://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) ??????????? URL ??????????????????? MVC ?????????????Web ??????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p106">To open a dialog box, your code in the task pane calls the [displayDialogAsync](https://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) method and passes to it the URL of the resource that you want to open. This is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource. In this article, 'page' or 'website' refers to the resource in the dialog. The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="239bb-p107">URL ?? HTTP**S** ?????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p107">The URL uses the HTTP**S** protocol. This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="239bb-p108">????????????????????????????????????[????](https://dev.office.com/reference/add-ins/manifest/functionfile)???????? `displayDialogAsync` ?????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p108">The domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](https://dev.office.com/reference/add-ins/manifest/functionfile) of an add-in command. This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

<span data-ttu-id="239bb-p109">????????????????????????? HTTPS ?????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p109">After the first page (or other resource) is loaded, a user can go to any website (or other resource) that uses HTTPS. You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="239bb-136">????????????????????? 80%???????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-136">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="239bb-137">????????????????? [Office ??? Dialog API ??](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)?</span><span class="sxs-lookup"><span data-stu-id="239bb-137">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="239bb-p110">???????? 100% ????????????????? 99.5%??????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p110">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="239bb-p111">????????????????????????????????????????????????????????????????????????????????????????????[?????](../design/add-in-commands.md)?????????????????? HTML ??????????????????????????????????????????????????????? [displayDialogAsync ?????](#errors-from-displaydialogasync)?</span><span class="sxs-lookup"><span data-stu-id="239bb-p111">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-online"></a><span data-ttu-id="239bb-146">?? Office Online ??????</span><span class="sxs-lookup"><span data-stu-id="239bb-146">Take advantage of a performance option in Office Online</span></span>

<span data-ttu-id="239bb-p112">???????? `displayDialogAsync` ????????????????????? `true`?????? Office Online ????????????????? iframe?????????????????????????????????`displayInIframe`</span><span class="sxs-lookup"><span data-stu-id="239bb-p112">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`. When this property is set to `true`, and the add-in is running in a document opened in Office Online, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster. The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="239bb-150">???? `false`?????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-150">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="239bb-151">???????? Office Online ????`displayInIframe` ?????</span><span class="sxs-lookup"><span data-stu-id="239bb-151">If the add-in is not running in Office Online, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="239bb-p114">?????????????? iframe ???????**?**??? `displayInIframe: true`???????? Web ???? Google ? Microsoft ??????????? iframe ????</span><span class="sxs-lookup"><span data-stu-id="239bb-p114">You should **not** use `displayInIframe: true` if the dialog will at any point redirect to a page that cannot be opened in an iframe. For example, the sign in pages of many popular web services, such as Google and Microsoft Account, cannot be opened in an iframe.</span></span>

### <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="239bb-154">?????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-154">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="239bb-155">???????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-155">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="239bb-156">????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-156">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="239bb-p115">Office JavaScript ????????????? Office JavaScript ???????????????? `Office.initialize` ???????????????????????????[???????](understanding-the-javascript-api-for-office.md#initializing-your-add-in)??</span><span class="sxs-lookup"><span data-stu-id="239bb-p115">The Office JavaScript library is loaded in the page. (Like any page that uses the Office JavaScript library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initializing your add-in](understanding-the-javascript-api-for-office.md#initializing-your-add-in).)</span></span>

<span data-ttu-id="239bb-p116">?????????? `messageParent` ???????????????????????????????XML blob????? JSON ???????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p116">Code in the dialog page uses the `messageParent` function to send either a Boolean value or a string message to the host page. The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string. The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - <span data-ttu-id="239bb-p117">???*??*???????????? Office API ??????? `Office.context.requirements.isSetSupported`???????????[?? Office ??? API ??](specify-office-hosts-and-api-requirements.md)?`messageParent`</span><span class="sxs-lookup"><span data-stu-id="239bb-p117">The `messageParent` function is one of *only* two Office APIs that can be called in the dialog box. The other is `Office.context.requirements.isSetSupported`. For information about it, see [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md).</span></span>
> - <span data-ttu-id="239bb-166">??????????????????????????????`messageParent`</span><span class="sxs-lookup"><span data-stu-id="239bb-166">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>

<span data-ttu-id="239bb-167">????????`googleProfile` ??? Google ????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-167">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="239bb-p118">???????????????????? `displayDialogAsync` ??????????????? `DialogMessageReceived` ??????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p118">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20},
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);
```

> [!NOTE]
> - <span data-ttu-id="239bb-p119">Office ? [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) ??????????????????????????????????????????????????[???????](#handle-errors-and-events)???</span><span class="sxs-lookup"><span data-stu-id="239bb-p119">Office passes an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object to the callback. It represents the result of the attempt to open the dialog box. It does not represent the outcome of any events in the dialog box. For more on this distinction, see the section [Handle errors and events](#handle-errors-and-events).</span></span>
> - <span data-ttu-id="239bb-176">? `value` ????? [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) ???????????????????????????`asyncResult`</span><span class="sxs-lookup"><span data-stu-id="239bb-176">The `value` property of the `asyncResult` is set to a [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="239bb-p120">??????????????????????`processMessage`</span><span class="sxs-lookup"><span data-stu-id="239bb-p120">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="239bb-179">??????????????? `processMessage` ?????????`dialog`</span><span class="sxs-lookup"><span data-stu-id="239bb-179">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="239bb-180">????? `DialogMessageReceived` ????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-180">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="239bb-p121">Office ? `arg` ???????????? `message` ???????? `messageParent` ????????????????????? Microsoft ??? Google ??????????????????????? `JSON.parse` ??????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p121">Office passes the `arg` object to the handler. Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog. In this example, it is a stringified representation of a user's profile from a service such as Microsoft Account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="239bb-p122">??? `showUserName` ??????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p122">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="239bb-186">????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-186">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="239bb-187">????? `displayDialogAsync` ????????`dialog`</span><span class="sxs-lookup"><span data-stu-id="239bb-187">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="239bb-188">???? Office ????????`dialog.close`</span><span class="sxs-lookup"><span data-stu-id="239bb-188">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="239bb-189">?????????????????? [Office ?????? API ??](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)?</span><span class="sxs-lookup"><span data-stu-id="239bb-189">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="239bb-p123">????????????????????????????? `window.location.replace` ???? `window.location.href`??????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p123">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="239bb-192">?????????????????[? PowerPoint ?????? Microsoft Graph ?? Excel ??](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)???</span><span class="sxs-lookup"><span data-stu-id="239bb-192">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint Add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

#### <a name="conditional-messaging"></a><span data-ttu-id="239bb-193">????</span><span class="sxs-lookup"><span data-stu-id="239bb-193">Conditional messaging</span></span>
<span data-ttu-id="239bb-p124">???????????? `messageParent` ????????????? `DialogMessageReceived` ????????????????????????????????????????????????????? Microsoft ??? Google????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p124">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages. For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft Account or Google, it sends the user's profile as a message. If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

```js
if (loginSuccess) {
    var userProfile = getProfile();
    var messageObject = {messageType: "signinSuccess", profile: userProfile};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    var errorDetails = getError();
    var messageObject = {messageType: "signinFailure", error: errorDetails};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
> - <span data-ttu-id="239bb-197">??????????????? HTTP ????????`loginSuccess`</span><span class="sxs-lookup"><span data-stu-id="239bb-197">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="239bb-p125">??? `getProfile` ? `getError` ?????????????????? HTTP ??????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p125">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="239bb-p126">????????????????????????? `messageType` ????????????? `profile` ??????? `error` ???</span><span class="sxs-lookup"><span data-stu-id="239bb-p126">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="239bb-202">????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-202">For samples that use conditional messaging, see:</span></span>
- [<span data-ttu-id="239bb-203">?? Auth0 ????????? Office ???</span><span class="sxs-lookup"><span data-stu-id="239bb-203">Office Add-in that uses the Auth0 Service to Simplify Social Login</span></span>](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [<span data-ttu-id="239bb-204">?? OAuth.io ????????????? Office ???</span><span class="sxs-lookup"><span data-stu-id="239bb-204">Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services</span></span>](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

<span data-ttu-id="239bb-p127">????????????? `messageType` ??????????????????????`showUserName` ??????????????`showNotification` ??????? UI ??????</span><span class="sxs-lookup"><span data-stu-id="239bb-p127">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "signinSuccess") {
        dialog.close();
        showUserName(messageFromDialog.profile.name);
        window.location.replace("/newPage.html");
    } else {
        dialog.close();
        showNotification("Unable to authenticate user: " + messageFromDialog.error);
    }
}
```

### <a name="closing-the-dialog-box"></a><span data-ttu-id="239bb-207">?????</span><span class="sxs-lookup"><span data-stu-id="239bb-207">Closing the dialog box</span></span>

<span data-ttu-id="239bb-p128">???????????????????????????????????? `messageParent` ????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p128">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="239bb-p129">??????????? `dialog.close`??????????????????????????????????????`DialogMessageReceived`</span><span class="sxs-lookup"><span data-stu-id="239bb-p129">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example. (See previous examples that show how the dialog object is initialized.)</span></span>


```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="239bb-213">?????????????? [Office ????????????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)?????[?????????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation)?</span><span class="sxs-lookup"><span data-stu-id="239bb-213">For a sample that uses this technique, see the [dialog navigation design pattern](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation) in the [UX design patterns for Office Add-ins](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) repo.</span></span>

<span data-ttu-id="239bb-p130">????????????? UI???????????????? **X** ???????????? `DialogEventReceived` ?????????????????????????????????????????????[????????????](#errors-and-events-in-the-dialog-window)???</span><span class="sxs-lookup"><span data-stu-id="239bb-p130">Even when you don't have your own close dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner. This action triggers the `DialogEventReceived` event. If your host pane needs to know when this happens, it should declare a handler for this event. See the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window) for details.</span></span>

## <a name="handle-errors-and-events"></a><span data-ttu-id="239bb-218">???????</span><span class="sxs-lookup"><span data-stu-id="239bb-218">Handle errors and events</span></span>

<span data-ttu-id="239bb-219">??????????</span><span class="sxs-lookup"><span data-stu-id="239bb-219">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="239bb-220">??????????????????`displayDialogAsync`</span><span class="sxs-lookup"><span data-stu-id="239bb-220">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="239bb-221">???????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-221">Errors, and other events, in the dialog window.</span></span>

### <a name="errors-from-displaydialogasync"></a><span data-ttu-id="239bb-222">DisplayDialogAsync ?????</span><span class="sxs-lookup"><span data-stu-id="239bb-222">Errors from displayDialogAsync</span></span>

<span data-ttu-id="239bb-223">??????????????? `displayDialogAsync` ????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-223">In addition to general platform and system errors, three errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="239bb-224">????</span><span class="sxs-lookup"><span data-stu-id="239bb-224">Code number</span></span>|<span data-ttu-id="239bb-225">??</span><span class="sxs-lookup"><span data-stu-id="239bb-225">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="239bb-226">12004</span><span class="sxs-lookup"><span data-stu-id="239bb-226">12004</span></span>|<span data-ttu-id="239bb-p131">??? `displayDialogAsync` ? URL ??????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p131">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="239bb-229">12005</span><span class="sxs-lookup"><span data-stu-id="239bb-229">12005</span></span>|<span data-ttu-id="239bb-p132">??? `displayDialogAsync` ? URL ?? HTTP ??????? HTTPS??? Office ????????? 12005 ???????? 12004 ??????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p132">The URL passed to `displayDialogAsync` uses the HTTP protocol. HTTPS is required. (In some versions of Office, the error message returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="239bb-233"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="239bb-233"><span id="12007">12007</span></span></span>|<span data-ttu-id="239bb-p133">???????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p133">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|

<span data-ttu-id="239bb-p134">?? `displayDialogAsync` ?????? [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) ??????????????????????????????`AsyncResult` ??? `value` ??? [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) ???????????[?????????????](#send-information-from-the-dialog-box-to-the-host-page)??????? `displayDialogAsync` ??????????`AsyncResult` ??? `status` ??????failed?????????? `error` ?????????????? `status`???????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p134">When `displayDialogAsync` is called, it always passes an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object to its callback function. When the call is successful - that is, the dialog window is opened - the `value` property of the `AsyncResult` object is a [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) object. An example of this is in the section [Send information from the dialog box to the host page](#send-information-from-the-dialog-box-to-the-host-page). When the call to `displayDialogAsync` fails, the window is not created, the `status` property of the `AsyncResult` object is set to "failed", and the `error` property of the object is populated. You should always have a callback that tests the `status` and responds when it's an error. For an example that simply reports the error message regardless of its code number, see the following code:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === "failed") {
        showNotification(asynceResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

### <a name="errors-and-events-in-the-dialog-window"></a><span data-ttu-id="239bb-242">????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-242">Errors and events in the dialog window</span></span>

<span data-ttu-id="239bb-243">???????????????????????????? `DialogEventReceived` ???</span><span class="sxs-lookup"><span data-stu-id="239bb-243">Three errors and events, known by their code numbers, in the dialog box will trigger a `DialogEventReceived` event in the host page.</span></span>

|<span data-ttu-id="239bb-244">????</span><span class="sxs-lookup"><span data-stu-id="239bb-244">Code number</span></span>|<span data-ttu-id="239bb-245">??</span><span class="sxs-lookup"><span data-stu-id="239bb-245">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="239bb-246">12002</span><span class="sxs-lookup"><span data-stu-id="239bb-246">12002</span></span>|<span data-ttu-id="239bb-247">???????</span><span class="sxs-lookup"><span data-stu-id="239bb-247">One of the following:</span></span><br> <span data-ttu-id="239bb-248">- ??? `displayDialogAsync` ? URL ????????</span><span class="sxs-lookup"><span data-stu-id="239bb-248">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="239bb-249">- ??? `displayDialogAsync` ??????????????????????????????????????? URL?</span><span class="sxs-lookup"><span data-stu-id="239bb-249">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was directed to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="239bb-250">12003</span><span class="sxs-lookup"><span data-stu-id="239bb-250">12003</span></span>|<span data-ttu-id="239bb-p135">???????? HTTP ??? URL????? HTTPS?</span><span class="sxs-lookup"><span data-stu-id="239bb-p135">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="239bb-253">12006</span><span class="sxs-lookup"><span data-stu-id="239bb-253">12006</span></span>|<span data-ttu-id="239bb-254">????????????????? **X** ???</span><span class="sxs-lookup"><span data-stu-id="239bb-254">The dialog box was closed, usually because the user chooses the **X** button.</span></span>|

<span data-ttu-id="239bb-p136">??????? `displayDialogAsync` ??? `DialogEventReceived` ???????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p136">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="239bb-257">??????????????????? `DialogEventReceived` ??????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-257">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example:</span></span>

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
        default:
            showNotification("Unknown error in dialog box.");
            break;
    }
}
```

<span data-ttu-id="239bb-258">?????????????????? [Office ??? Dialog API ??](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)?</span><span class="sxs-lookup"><span data-stu-id="239bb-258">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>


## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="239bb-259">????????</span><span class="sxs-lookup"><span data-stu-id="239bb-259">Pass information to the dialog box</span></span>

<span data-ttu-id="239bb-p137">????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p137">Sometimes the host page needs to pass information to the dialog box. You can do this in two primary ways:</span></span>

- <span data-ttu-id="239bb-262">???? `displayDialogAsync` ? URL ???????</span><span class="sxs-lookup"><span data-stu-id="239bb-262">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="239bb-p138">??????????????????????????????????????*??????????*????????????????[????](http://www.w3schools.com/html/html5_webstorage.asp)?</span><span class="sxs-lookup"><span data-stu-id="239bb-p138">Store the information somewhere that is accessible to both the host window and dialog box. The two windows do not share a common session storage, but *if they have the same domain* (including port number, if any),  they share a common [local storage](http://www.w3schools.com/html/html5_webstorage.asp).</span></span>

### <a name="use-local-storage"></a><span data-ttu-id="239bb-265">??????</span><span class="sxs-lookup"><span data-stu-id="239bb-265">Use local storage</span></span>

<span data-ttu-id="239bb-266">???????????????????? `window.localStorage` ??? `setItem` ???????? `displayDialogAsync`?????????</span><span class="sxs-lookup"><span data-stu-id="239bb-266">To use local storage, your code calls the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="239bb-267">??????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-267">Code in the dialog window reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

<span data-ttu-id="239bb-268">?????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-268">For sample add-ins that uses local storage in this way, see:</span></span>

- [<span data-ttu-id="239bb-269">?? Auth0 ????????? Office ???</span><span class="sxs-lookup"><span data-stu-id="239bb-269">Office Add-in that uses the Auth0 Service to Simplify Social Login</span></span>](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [<span data-ttu-id="239bb-270">?? OAuth.io ????????????? Office ???</span><span class="sxs-lookup"><span data-stu-id="239bb-270">Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services</span></span>](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

### <a name="use-query-parameters"></a><span data-ttu-id="239bb-271">??????</span><span class="sxs-lookup"><span data-stu-id="239bb-271">Use query parameters</span></span>

<span data-ttu-id="239bb-272">?????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-272">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="239bb-273">??????????????[? PowerPoint ?????? Microsoft Graph ?? Excel ??](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)?</span><span class="sxs-lookup"><span data-stu-id="239bb-273">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint Add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="239bb-274">????????????? URL????????</span><span class="sxs-lookup"><span data-stu-id="239bb-274">Code in your dialog window can parse the URL and read the parameter value.</span></span>

> [!NOTE]
> <span data-ttu-id="239bb-p139">Office ??????? `displayDialogAsync` ? URL ?????? `_host_info`??????????????????????????????????? URL??Microsoft ????????????????????????????????????????????????????????*????????????????*?</span><span class="sxs-lookup"><span data-stu-id="239bb-p139">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.</span></span>

## <a name="use-the-dialog-apis-to-show-a-video"></a><span data-ttu-id="239bb-280">????? API ????</span><span class="sxs-lookup"><span data-stu-id="239bb-280">Use the Dialog APIs to show a video</span></span>

<span data-ttu-id="239bb-281">????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-281">To show a video in a dialog box:</span></span>

1.  <span data-ttu-id="239bb-p140">?????? iframe ????iframe ? `src` ??????????? URL ???? HTTP**S** ????????????video.dialogbox.html????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p140">Create a page whose only content is an iframe. The `src` attribute of the iframe points to an online video. The protocol of the video's URL must be HTTP**S**. In this article we'll call this page "video.dialogbox.html". The following is an example of the markup:</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2.  <span data-ttu-id="239bb-287">video.dialogbox.html ???????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-287">The video.dialogbox.html page must be in the same domain as the host page.</span></span>
3.  <span data-ttu-id="239bb-288">??????? `displayDialogAsync`??? video.dialogbox.html?</span><span class="sxs-lookup"><span data-stu-id="239bb-288">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
4.  <span data-ttu-id="239bb-p141">?????????????????????? `DialogEventReceived` ???????????? 12006 ?????????????[????????????](#errors-and-events-in-the-dialog-window)???</span><span class="sxs-lookup"><span data-stu-id="239bb-p141">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event. For details, see the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window).</span></span>

<span data-ttu-id="239bb-291">?????????????????? [Office ?????????????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)?????[??????????](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat)?</span><span class="sxs-lookup"><span data-stu-id="239bb-291">For a sample that shows a video in a dialog box, see the [video placemat design pattern](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat) in the [UX design patterns for Office Add-ins](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) repo.</span></span>

![??????????????????](../images/video-placemats-dialog-open.png)

## <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="239bb-293">???????????? API</span><span class="sxs-lookup"><span data-stu-id="239bb-293">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="239bb-294">??? API ????????????? Iframe ?????????????????? Microsoft ???Office 365?Google ? Facebook????????</span><span class="sxs-lookup"><span data-stu-id="239bb-294">A primary scenario for the Dialog APIs is to enable authentication with a resource or identity provider that does not allow its sign-in page to open in an Iframe, such as Microsoft Account, Office 365, Google, and Facebook.</span></span>

> [!NOTE]
> <span data-ttu-id="239bb-p142">?????? API ?????????*?*??? `displayDialogAsync` ??? `displayInIframe: true` ???????????[?? Office Online ??????](#take-advantage-of-a-performance-option-in-office-online)?????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p142">When you are using the Dialog APIs for this scenario, do *not* use the `displayInIframe: true` option in the call to `displayDialogAsync`. See [Take advantage of a performance option in Office Online](#take-advantage-of-a-performance-option-in-office-online) previously in this article for details about this option.</span></span>

<span data-ttu-id="239bb-297">????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-297">The following is a simple and typical authentication flow:</span></span>

1. <span data-ttu-id="239bb-p143">??????????????????????????????????????????????????? UI???????????????????????? *NAME-OF-PROVIDER* ???????????????????????????????????? URL??[????????](#pass-information-to-the-dialog-box)????</span><span class="sxs-lookup"><span data-stu-id="239bb-p143">The first page that opens in the dialog box is a local page (or other resource) that is hosted in the add-in's domain; that is, the host window's domain. This page can have a simple UI that says "Please wait, we are redirecting you to the page where you can sign in to *NAME-OF-PROVIDER*." Code in this page constructs the URL of the identity provider's sign-in page by using information that is passed to the dialog box as described in [Pass information to the dialog box](#pass-information-to-the-dialog-box).</span></span>
2. <span data-ttu-id="239bb-p144">????????????????URL ??????????????????????????????????????????????????? "redirectPage.html"??*????????????????*????????????????????????? `messageParent`????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p144">The dialog window then redirects to the sign-in page. The URL includes a query parameter that tells the identity provider to redirect the dialog window, after the user signs in, to a specific page. In this article, we'll call this page "redirectPage.html". (*This must be a page in the same domain as the host window*, because the only way for the dialog window to pass the results of the sign-in attempt is with a call of `messageParent`, which can only be called on a page with the same domain as the host window.)</span></span>
2. <span data-ttu-id="239bb-p145">????????????????????? GET ??????????????????????? redirectPage.html??????????????????????????????????????????????????????????????????????????????????????????????? redirectPage.html????????????? **X** ?????????????????????????? redirectPage.html?????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p145">The identity provider's service processes the incoming GET request from the dialog window. If the user is already logged on, it immediately redirects the window to redirectPage.html and includes user data as a query parameter. If the user is not already signed in, the provider's sign-in page appears in the window, and the user signs in. For most providers, if the user cannot sign in successfully, the provider shows an error page in the dialog window and does not redirect to redirectPage.html. The user must close the window by selecting the **X** in the corner. If the user successfully signs in, the dialog window is redirected to redirectPage.html and user data is included as a query parameter.</span></span>
3. <span data-ttu-id="239bb-311">? redirectPage.html ?????????? `messageParent` ????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-311">When the redirectPage.html page opens, it calls `messageParent` to report the success or failure to the host page and optionally also report user data or error data.</span></span>
4. <span data-ttu-id="239bb-312">?????????????????????????????????????`DialogMessageReceived`</span><span class="sxs-lookup"><span data-stu-id="239bb-312">The `DialogMessageReceived` event fires in the host page and its handler closes the dialog window and optionally does other processing of the message.</span></span>

<span data-ttu-id="239bb-313">???????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-313">For sample add-ins that use this pattern, see:</span></span>

- <span data-ttu-id="239bb-p146">[? PowerPoint ??????? Microsoft Graph ?? Excel ??](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)??????????????????????????????????? Office 365 ????</span><span class="sxs-lookup"><span data-stu-id="239bb-p146">[Insert Excel charts using Microsoft Graph in a PowerPoint Add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart): The resource that is initially opened in the dialog window is a controller method that has no view of its own. It redirects to the Office 365 sign in page.</span></span>
- <span data-ttu-id="239bb-316">[Office ???? Office 365 ??? AngularJS ????](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)???????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-316">[Office Add-in Office 365 Client Authentication for AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth): The resource that is initially opened in the dialog window is a page.</span></span>

#### <a name="support-multiple-identity-providers"></a><span data-ttu-id="239bb-317">??????????</span><span class="sxs-lookup"><span data-stu-id="239bb-317">Support multiple identity providers</span></span>

<span data-ttu-id="239bb-p147">?????????????????? Microsoft ???Google ? Facebook???????????????????????????????????? UI??????????? URL ????????? URL?</span><span class="sxs-lookup"><span data-stu-id="239bb-p147">If your add-in gives the user a choice of providers, such as Microsoft Account, Google, or Facebook, you need a local first page (see preceding section) that provides a UI for the user to select a provider. Selection triggers the construction of the sign-in URL and redirection to it.</span></span>

<span data-ttu-id="239bb-320">??????????????[?? Auth0 ????????? Office ????](https://github.com/OfficeDev/Office-Add-in-Auth0)?</span><span class="sxs-lookup"><span data-stu-id="239bb-320">For a sample that uses this pattern, see [Office Add-in that uses the Auth0 Service to Simplify Social Login](https://github.com/OfficeDev/Office-Add-in-Auth0).</span></span>

#### <a name="authorization-of-the-add-in-to-an-external-resource"></a><span data-ttu-id="239bb-321">????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-321">Authorization of the add-in to an external resource</span></span>

<span data-ttu-id="239bb-p148">???????Web ?????????????????????????????????? Office 365?Google Plus?Facebook ? LinkedIn??????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p148">In the modern web, web applications are security principals just as users are, and the application has its own identity and permissions to an online resource such as Office 365, Google Plus, Facebook, or LinkedIn. The application is registered with the resource provider before it is deployed. The registration includes:</span></span>

- <span data-ttu-id="239bb-325">???????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-325">A list of the permissions that the application needs to a user's resources.</span></span>
- <span data-ttu-id="239bb-326">??????????????????????? URL?</span><span class="sxs-lookup"><span data-stu-id="239bb-326">A URL to which the resource service should return an access token when the application accesses the service.</span></span>  

<span data-ttu-id="239bb-p149">????????????????????????????????????????????????????????????????????????????????? URL????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p149">When a user invokes a function in the application that accesses the user's data in the resource service, they are prompted to sign in to the service and then prompted to grant the application the permissions it needs to the user's resources. The service then redirects the sign-in window to the previously registered URL and passes the access token. The application uses the access token to access the user's resources.</span></span>

<span data-ttu-id="239bb-p150">??????? API ??????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p150">You can use the Dialog APIs to manage this process by using a flow that is similar to the one described for users to sign in. The only differences are:</span></span>

- <span data-ttu-id="239bb-332">???????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-332">If the user hasn't previously granted the application the permissions it needs, she is prompted to do so in the dialog box after signing in.</span></span>
- <span data-ttu-id="239bb-p151">??????? `messageParent` ???????????????????????????????????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p151">The dialog window sends the access token to the host window either by using `messageParent` to send the stringified access token or by storing the access token where the host window can retrieve it. The token has a time limit, but while it lasts, the host window can use it to directly access the user's resources without any further prompting.</span></span>

<span data-ttu-id="239bb-335">?????????? API ??????</span><span class="sxs-lookup"><span data-stu-id="239bb-335">The following samples use the Dialog APIs for this purpose:</span></span>
- <span data-ttu-id="239bb-336">[? PowerPoint ??????? Microsoft Graph ?? Excel ??](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) - ?????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-336">[Insert Excel charts using Microsoft Graph in a PowerPoint Add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) - Stores the access token in a database.</span></span>
- [<span data-ttu-id="239bb-337">?? OAuth.io ????????????? Office ???</span><span class="sxs-lookup"><span data-stu-id="239bb-337">Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services</span></span>](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

<span data-ttu-id="239bb-338">???????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-338">For more information about authentication and authorization in add-ins, see:</span></span>
- [<span data-ttu-id="239bb-339">? Office ??????????</span><span class="sxs-lookup"><span data-stu-id="239bb-339">Authorize external services in your Office Add-in</span></span>](auth-external-add-ins.md)
- [<span data-ttu-id="239bb-340">Office JavaScript API ?????</span><span class="sxs-lookup"><span data-stu-id="239bb-340">Office JavaScript API Helpers library</span></span>](https://github.com/OfficeDev/office-js-helpers)


## <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="239bb-341">? Office Dialog API ?????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-341">Use the Office Dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="239bb-342">??????????????????????????????????? URL ??? [ displayDialogAsync ](http://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) ???????????? HTML ??? URL?</span><span class="sxs-lookup"><span data-stu-id="239bb-342">If your add-in uses client-side routing, as single-page applications typically do, you have the option to pass the URL of a route to the [displayDialogAsync](http://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) method, instead of the URL of a complete and separate HTML page.</span></span>

> [!IMPORTANT]
><span data-ttu-id="239bb-p152">??????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="239bb-p152">The dialog box is in a new window with its own execution context. If you pass a route, your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog window. So this technique launches a second instance of your application in the dialog window. Code that changes variables in the dialog window does not change the task pane version of the same variables. Similarly, the dialog window has its own session storage, which is not accessible from code in the task pane.</span></span>
