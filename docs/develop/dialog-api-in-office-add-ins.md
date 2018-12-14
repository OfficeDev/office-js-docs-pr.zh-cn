---
title: 在 Office 加载项中使用对话框 API
description: ''
ms.date: 11/28/2018
ms.openlocfilehash: b19d56d3f4fb831eb8c0ca16af53ee309989d223
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270955"
---
# <a name="use-the-dialog-api-in-your-office-add-ins"></a><span data-ttu-id="cfc70-102">在 Office 加载项中使用对话框 API</span><span class="sxs-lookup"><span data-stu-id="cfc70-102">Use the Dialog API in your Office Add-ins</span></span>

<span data-ttu-id="cfc70-p101">可以在 Office 外接程序中使用[对话框 API](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) 打开对话框。本文提供了有关如何在 Office 外接程序中使用对话框 API 的指南。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p101">You can use the [Dialog API](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) to open dialog boxes in your Office Add-in. This article provides guidance for using the Dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="cfc70-p102">若要了解对话框 API 目前的受支持情况，请参阅[对话框 API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js)。目前，Word、Excel、PowerPoint 和 Outlook 支持对话框 API。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p102">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.</span></span>

> <span data-ttu-id="cfc70-107">对话框 API 的主要应用场景是为 Google 或 Facebook 等资源启用身份验证。</span><span class="sxs-lookup"><span data-stu-id="cfc70-107">A primary scenario for the Dialog APIs is to enable authentication with a resource such as Google or Facebook.</span></span>

<span data-ttu-id="cfc70-108">不妨通过任务窗格/内容加载项/[加载项命令](../design/add-in-commands.md)打开对话框，以便执行下列操作：</span><span class="sxs-lookup"><span data-stu-id="cfc70-108">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="cfc70-109">显示无法直接在任务窗格中打开的登录页。</span><span class="sxs-lookup"><span data-stu-id="cfc70-109">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="cfc70-110">为加载项中的某些任务提供更多屏幕空间，或甚至整个屏幕。</span><span class="sxs-lookup"><span data-stu-id="cfc70-110">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="cfc70-111">托管在任务窗格中显得太小的视频。</span><span class="sxs-lookup"><span data-stu-id="cfc70-111">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="cfc70-p103">由于不赞成重叠 UI 元素，因此除非应用场景需要，否则请勿从任务窗格打开对话框。考虑如何使用任务窗格的区域时，请注意任务窗格可以是选项卡式。有关示例，请参阅 [Excel 加载项 JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 示例。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p103">Because overlapping UI elements are discouraged, avoid opening a dialog from a task pane unless your scenario requires it. When you consider how to use the surface area of a task pane, note that task panes can be tabbed. For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="cfc70-115">下图展示了对话框示例。</span><span class="sxs-lookup"><span data-stu-id="cfc70-115">The following image shows an example of a dialog box.</span></span>

![加载项命令](../images/auth-o-dialog-open.png)

<span data-ttu-id="cfc70-p104">请注意，对话框总是在屏幕中央打开。用户可以移动它，并重设大小。对话框是*非模式*窗口。也就是说，用户可以继续与主机 Office 应用中的文档，以及与任务窗格中的主机页（若有）进行交互。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p104">Note that the dialog box always opens in the center of the screen. The user can move and resize it. The window is *nonmodal*--a user can continue to interact with both the document in the host Office application and with the host page in the task pane, if there is one.</span></span>

## <a name="dialog-api-scenarios"></a><span data-ttu-id="cfc70-120">对话框 API 应用场景</span><span class="sxs-lookup"><span data-stu-id="cfc70-120">Dialog API scenarios</span></span>

<span data-ttu-id="cfc70-121">Office JavaScript API 支持以下应用场景，其在 [Office.context.ui 命名空间](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js)中使用 [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) 对象和两个函数。</span><span class="sxs-lookup"><span data-stu-id="cfc70-121">The Office JavaScript APIs support the following scenarios with a [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) object and two functions in the [Office.context.ui namespace](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js).</span></span>

### <a name="open-a-dialog-box"></a><span data-ttu-id="cfc70-122">打开对话框</span><span class="sxs-lookup"><span data-stu-id="cfc70-122">Open a dialog box</span></span>

<span data-ttu-id="cfc70-p105">为了打开对话框，任务窗格中的代码调用 [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) 方法，并将要打开的资源 URL 传递到此方法。这通常是页面，但也可以是 MVC 应用中的控制器方法、路由、Web 服务方法或其他任何资源。在本文中，“页面”或“网站”是指对话框中的资源。下面的代码就是一个简单示例：</span><span class="sxs-lookup"><span data-stu-id="cfc70-p105">To open a dialog box, your code in the task pane calls the [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) method and passes to it the URL of the resource that you want to open. This is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource. In this article, 'page' or 'website' refers to the resource in the dialog. The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="cfc70-p106">URL 使用 **HTTPS** 协议。对话框中加载的所有页面都必须遵循此要求，而不仅仅是加载的第一个页面。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p106">The URL uses the HTTP**S** protocol. This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="cfc70-129">对话框资源域与宿主页的域相同，宿主页可以是任务窗格中的页面，也可以是加载项命令的[函数文件](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/functionfile?view=office-js)。</span><span class="sxs-lookup"><span data-stu-id="cfc70-129">The dialog resource's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/functionfile?view=office-js) of an add-in command.</span></span> <span data-ttu-id="cfc70-130">这要求：传递到 `displayDialogAsync` 方法的页面、控制器方法或其他资源必须与主机页位于相同的域。</span><span class="sxs-lookup"><span data-stu-id="cfc70-130">This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cfc70-131">宿主页面和对话框资源必须具有相同的完整域。</span><span class="sxs-lookup"><span data-stu-id="cfc70-131">The host page and the resources of the dialog must have the same full domain.</span></span> <span data-ttu-id="cfc70-132">如果尝试传递 `displayDialogAsync` 加载项域的子域，则不会起作用。</span><span class="sxs-lookup"><span data-stu-id="cfc70-132">If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work.</span></span> <span data-ttu-id="cfc70-133">完整域（包括任何子域）必须匹配。</span><span class="sxs-lookup"><span data-stu-id="cfc70-133">The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="cfc70-p109">在第一个页面（或其他资源）加载后，用户可以转到使用 HTTPS 的任意网站（或其他资源）。还可以将第一个页面设计为直接重定向到另一个站点。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p109">After the first page (or other resource) is loaded, a user can go to any website (or other resource) that uses HTTPS. You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="cfc70-136">默认情况下，对话框的高度和宽度占设备屏幕的 80%。不过，也可以设置不同的百分比，只需将配置对象传递给方法即可，如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="cfc70-136">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="cfc70-137">有关实现这一点的样本加载项，请参阅 [Office 加载项 Dialog API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。</span><span class="sxs-lookup"><span data-stu-id="cfc70-137">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="cfc70-p110">将两个值均设置为 100% 可有效提供全屏体验。（有效最大值为 99.5%，窗口仍可移动和调整大小。）</span><span class="sxs-lookup"><span data-stu-id="cfc70-p110">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="cfc70-p111">只能从主机窗口打开一个对话框。如果尝试再打开一个对话框，就会生成错误。比方说，如果用户从任务窗格打开一个对话框，她就无法再从任务窗格中的其他页面打开第二个对话框。不过，如果对话框是通过[加载项命令](../design/add-in-commands.md)打开，那么只要选择此命令，就会打开新 HTML 文件（但不可见）。这会新建（不可见的）主机窗口，所以每个这样的窗口都可以启动自己的对话框。有关详细信息，请参阅 [displayDialogAsync 返回的错误](#errors-from-displaydialogasync)。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p111">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-online"></a><span data-ttu-id="cfc70-146">利用 Office Online 中的性能选项</span><span class="sxs-lookup"><span data-stu-id="cfc70-146">Take advantage of a performance option in Office Online</span></span>

<span data-ttu-id="cfc70-p112">`displayInIframe` 属性是可以传递到 `displayDialogAsync` 的配置对象中的附加属性。如果将此属性设置为 `true`，且加载项在 Office Online 打开的文档中运行，对话框就会以浮动 iframe（而不是独立窗口）的形式打开，从而加快对话框的打开速度。示例如下：</span><span class="sxs-lookup"><span data-stu-id="cfc70-p112">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`. When this property is set to `true`, and the add-in is running in a document opened in Office Online, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster. The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="cfc70-150">默认值为 `false`，与完全省略此属性时相同。</span><span class="sxs-lookup"><span data-stu-id="cfc70-150">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="cfc70-151">如果加载项没有在 Office Online 中运行，`displayInIframe` 将被忽略。</span><span class="sxs-lookup"><span data-stu-id="cfc70-151">If the add-in is not running in Office Online, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="cfc70-p114">如果对话框始终重定向到无法在 iframe 中打开的页面，**不**得使用 `displayInIframe: true`。例如，许多热门 Web 服务（如 Google 和 Microsoft 帐户）的登录页都无法在 iframe 中打开。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p114">You should **not** use `displayInIframe: true` if the dialog will at any point redirect to a page that cannot be opened in an iframe. For example, the sign in pages of many popular web services, such as Google and Microsoft Account, cannot be opened in an iframe.</span></span>

### <a name="handling-pop-up-blockers-with-office-online"></a><span data-ttu-id="cfc70-154">使用 Office Online 处理弹出窗口阻止程序</span><span class="sxs-lookup"><span data-stu-id="cfc70-154">Handling pop-up blockers with Office Online</span></span>

<span data-ttu-id="cfc70-155">如果尝试在使用 Office Online 时显示对话框，可能会导致浏览器的弹出窗口阻止程序阻止对话框。</span><span class="sxs-lookup"><span data-stu-id="cfc70-155">Attempting to display a dialog while using Office Online may cause the browser's pop-up blocker to block the dialog.</span></span> <span data-ttu-id="cfc70-156">如果加载项用户先同意加载项发出的提示，可以避开浏览器的弹出窗口阻止程序。</span><span class="sxs-lookup"><span data-stu-id="cfc70-156">The browser's pop-up blocker can be circumvented if the user of your add-in first agrees to a prompt from the add-in.</span></span> <span data-ttu-id="cfc70-157">`displayDialogAsync` 的 [DialogOptions](/javascript/api/office/office.dialogoptions) 包含可触发此类弹出窗口的 `promptBeforeOpen` 属性。</span><span class="sxs-lookup"><span data-stu-id="cfc70-157">`displayDialogAsync`'s [DialogOptions](/javascript/api/office/office.dialogoptions) has the `promptBeforeOpen` property to trigger such a pop-up.</span></span> <span data-ttu-id="cfc70-158">`promptBeforeOpen` 是提供以下行为的布尔值：</span><span class="sxs-lookup"><span data-stu-id="cfc70-158">`promptBeforeOpen` is a boolean value which provides the following behavior:</span></span>
 
 - <span data-ttu-id="cfc70-159">`true` - 框架显示用于触发导航的弹出窗口，并避开浏览器的弹出窗口阻止程序。</span><span class="sxs-lookup"><span data-stu-id="cfc70-159">`true` - The framework displays a pop-up to trigger the navigation and avoid the browser's pop-up blocker.</span></span> 
 - <span data-ttu-id="cfc70-160">`false` - 对话框不会显示，开发人员必须处理弹出窗口（通过提供用户界面项目来触发导航）。</span><span class="sxs-lookup"><span data-stu-id="cfc70-160">`false` - The dialog will not be shown and the developer must handle pop-ups (by providing a user interface artifact to trigger the navigation).</span></span> 
 
<span data-ttu-id="cfc70-161">弹出窗口如以下屏幕截图中所示：</span><span class="sxs-lookup"><span data-stu-id="cfc70-161">The pop-up looks similiar to that in the following screenshot:</span></span>

![加载项对话框可以生成提示，以避免浏览器内的弹出窗口阻止程序。](../images/dialog-prompt-before-open.png)
 
### <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="cfc70-163">将信息从对话框发送到主机页</span><span class="sxs-lookup"><span data-stu-id="cfc70-163">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="cfc70-164">对话框无法与任务窗格中的主机页进行通信，除非：</span><span class="sxs-lookup"><span data-stu-id="cfc70-164">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="cfc70-165">对话框中的当前页面与主机页在同一个域中。</span><span class="sxs-lookup"><span data-stu-id="cfc70-165">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="cfc70-p116">Office JavaScript 库已在页面中加载。（与使用 Office JavaScript 库的所有页面一样，页面脚本必须为 `Office.initialize` 属性分配方法，尽管方法可以是空的。有关详细信息，请参阅[初始化外接程序](understanding-the-javascript-api-for-office.md#initializing-your-add-in)。）</span><span class="sxs-lookup"><span data-stu-id="cfc70-p116">The Office JavaScript library is loaded in the page. (Like any page that uses the Office JavaScript library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initializing your add-in](understanding-the-javascript-api-for-office.md#initializing-your-add-in).)</span></span>

<span data-ttu-id="cfc70-p117">对话框页中的代码使用 `messageParent` 函数，向主机页发送布尔值或字符串消息。字符串可以是字词、句子、XML blob、字符串化 JSON 或其他任何能够串行化为字符串的内容。示例如下：</span><span class="sxs-lookup"><span data-stu-id="cfc70-p117">Code in the dialog page uses the `messageParent` function to send either a Boolean value or a string message to the host page. The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string. The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - <span data-ttu-id="cfc70-p118">`messageParent` 函数是*唯一*可以在对话框中调用的两个 Office API 之一。另一个是 `Office.context.requirements.isSetSupported`。有关详细信息，请参阅[指定 Office 主机和 API 要求](specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p118">The `messageParent` function is one of *only* two Office APIs that can be called in the dialog box. The other is `Office.context.requirements.isSetSupported`. For information about it, see [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md).</span></span>
> - <span data-ttu-id="cfc70-175">`messageParent` 函数只能在与主机页位于同一域（包括协议和端口）的页面上调用。</span><span class="sxs-lookup"><span data-stu-id="cfc70-175">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>

<span data-ttu-id="cfc70-176">在下一个示例中，`googleProfile` 是用户 Google 配置文件的字符串化版本。</span><span class="sxs-lookup"><span data-stu-id="cfc70-176">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="cfc70-p119">必须将主机页配置为接收消息。为此，可以向 `displayDialogAsync` 的原始调用添加回调参数。回调向 `DialogMessageReceived` 事件分配处理程序。示例如下：</span><span class="sxs-lookup"><span data-stu-id="cfc70-p119">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

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
> - <span data-ttu-id="cfc70-p120">Office 将 [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js) 对象传递给回调。它表示尝试打开对话框的结果，不表示对话框中任何事件的结果。若要详细了解此区别，请参阅[处理错误和事件](#handle-errors-and-events)部分。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p120">Office passes an [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js) object to the callback. It represents the result of the attempt to open the dialog box. It does not represent the outcome of any events in the dialog box. For more on this distinction, see the section [Handle errors and events](#handle-errors-and-events).</span></span>
> - <span data-ttu-id="cfc70-185">`asyncResult` 的 `value` 属性设置为 [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) 对象，此对象位于主机页（而不是对话框的执行上下文）中。</span><span class="sxs-lookup"><span data-stu-id="cfc70-185">The `value` property of the `asyncResult` is set to a [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="cfc70-p121">`processMessage` 是用于处理事件的函数。可以根据需要任意命名。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p121">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="cfc70-188">`dialog` 变量的声明范围比回调更广，因为 `processMessage` 中也会引用此变量。</span><span class="sxs-lookup"><span data-stu-id="cfc70-188">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="cfc70-189">下面展示了 `DialogMessageReceived` 事件处理程序的简单示例：</span><span class="sxs-lookup"><span data-stu-id="cfc70-189">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="cfc70-p122">Office 将 `arg` 对象传递给处理程序。它的 `message` 属性是对话框中的 `messageParent` 调用发送的布尔值或字符串。在此示例中，它是 Microsoft 帐户或 Google 等服务的用户配置文件的字符串化表示。因此，使用 `JSON.parse` 将其反序列化回对象。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p122">Office passes the `arg` object to the handler. Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog. In this example, it is a stringified representation of a user's profile from a service such as Microsoft Account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="cfc70-p123">未显示 `showUserName` 实现。它可能在任务窗格上显示定制的欢迎消息。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p123">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="cfc70-195">在用户完成与对话框的交互后，消息处理程序应关闭对话框，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="cfc70-195">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="cfc70-196">`dialog` 对象必须是 `displayDialogAsync` 调用返回的对象。</span><span class="sxs-lookup"><span data-stu-id="cfc70-196">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="cfc70-197">`dialog.close` 调用指示 Office 立即关闭对话框。</span><span class="sxs-lookup"><span data-stu-id="cfc70-197">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="cfc70-198">有关使用这些技术的示例加载项，请参阅 [Office 加载项对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。</span><span class="sxs-lookup"><span data-stu-id="cfc70-198">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="cfc70-p124">如果加载项在收到消息后需要打开任务窗格的其他页面，可以使用 `window.location.replace` 方法（或 `window.location.href`）作为处理程序的最后一行。示例如下：</span><span class="sxs-lookup"><span data-stu-id="cfc70-p124">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="cfc70-201">有关具有此用途的加载项示例，请参阅[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)（在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表）示例。</span><span class="sxs-lookup"><span data-stu-id="cfc70-201">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

#### <a name="conditional-messaging"></a><span data-ttu-id="cfc70-202">条件消息</span><span class="sxs-lookup"><span data-stu-id="cfc70-202">Conditional messaging</span></span>
<span data-ttu-id="cfc70-p125">由于可以从对话框发送多个 `messageParent` 调用，但在主机页中只有一个 `DialogMessageReceived` 事件处理程序，因此处理程序必须使用条件逻辑来区分不同的消息。比方说，如果对话框提示用户登录标识提供程序（如 Microsoft 帐户或 Google），则会以消息形式发送用户配置文件。如果身份验证失败，对话框会将错误消息发送到主机页，如下面的示例所示：</span><span class="sxs-lookup"><span data-stu-id="cfc70-p125">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages. For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft Account or Google, it sends the user's profile as a message. If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

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
> - <span data-ttu-id="cfc70-206">`loginSuccess` 变量通过读取标识提供程序返回的 HTTP 响应进行初始化。</span><span class="sxs-lookup"><span data-stu-id="cfc70-206">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="cfc70-p126">未显示 `getProfile` 和 `getError` 函数的实现。这两个函数均从查询参数或 HTTP 响应的正文获取数据。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p126">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="cfc70-p127">根据登录是否成功，发送不同类型的匿名对象。两者都有 `messageType` 属性。不同之处在于，一个有 `profile` 属性，另一个有 `error` 属性。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p127">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="cfc70-211">有关使用条件消息的样本，请参阅：</span><span class="sxs-lookup"><span data-stu-id="cfc70-211">For samples that use conditional messaging, see:</span></span>
- [<span data-ttu-id="cfc70-212">使用 Auth0 服务简化社交登录的 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="cfc70-212">Office Add-in that uses the Auth0 Service to Simplify Social Login</span></span>](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [<span data-ttu-id="cfc70-213">使用 OAuth.io 服务简化热门在线服务访问的 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="cfc70-213">Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services</span></span>](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

<span data-ttu-id="cfc70-p128">主机页中的处理程序代码使用 `messageType` 属性的值设置分支，如下面的示例所示。请注意，`showUserName` 函数的用法与之前的示例相同，`showNotification` 函数在主机页的 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p128">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

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

### <a name="closing-the-dialog-box"></a><span data-ttu-id="cfc70-216">关闭对话框</span><span class="sxs-lookup"><span data-stu-id="cfc70-216">Closing the dialog box</span></span>

<span data-ttu-id="cfc70-p129">可以在对话框中实现对话框关闭按钮。为此，关闭按钮的单击事件处理程序应使用 `messageParent` 通知主机页，关闭按钮已获单击。示例如下：</span><span class="sxs-lookup"><span data-stu-id="cfc70-p129">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="cfc70-p130">`DialogMessageReceived` 的主机页处理程序会调用 `dialog.close`，如下面的示例所示。（请参阅之前的示例，其中展示了对话框对象的初始化方式。）</span><span class="sxs-lookup"><span data-stu-id="cfc70-p130">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example. (See previous examples that show how the dialog object is initialized.)</span></span>


```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="cfc70-222">有关使用此技术的示例，请参阅 [Office 加载项的用户体验设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)存储库中的[对话框导航设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation)。</span><span class="sxs-lookup"><span data-stu-id="cfc70-222">For a sample that uses this technique, see the [dialog navigation design pattern](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation) in the [UX design patterns for Office Add-ins](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) repo.</span></span>

<span data-ttu-id="cfc70-p131">即使你没有自己的关闭对话框 UI，最终用户也可以通过选择右上角的 **X** 关闭对话框。此操作将触发 `DialogEventReceived` 事件。如果主机窗格需要知道此事件何时发生，应为此事件声明一个处理程序。有关详细信息，请参阅[对话框窗口中的错误和事件](#errors-and-events-in-the-dialog-window)部分。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p131">Even when you don't have your own close dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner. This action triggers the `DialogEventReceived` event. If your host pane needs to know when this happens, it should declare a handler for this event. See the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window) for details.</span></span>

## <a name="handle-errors-and-events"></a><span data-ttu-id="cfc70-227">处理错误和事件</span><span class="sxs-lookup"><span data-stu-id="cfc70-227">Handle errors and events</span></span>

<span data-ttu-id="cfc70-228">代码应处理两类事件：</span><span class="sxs-lookup"><span data-stu-id="cfc70-228">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="cfc70-229">`displayDialogAsync` 调用返回的错误，因为无法创建对话框。</span><span class="sxs-lookup"><span data-stu-id="cfc70-229">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="cfc70-230">对话框窗口中的错误和其他事件。</span><span class="sxs-lookup"><span data-stu-id="cfc70-230">Errors, and other events, in the dialog window.</span></span>

### <a name="errors-from-displaydialogasync"></a><span data-ttu-id="cfc70-231">DisplayDialogAsync 返回的错误</span><span class="sxs-lookup"><span data-stu-id="cfc70-231">Errors from displayDialogAsync</span></span>

<span data-ttu-id="cfc70-232">除常规的平台和系统错误外，调用 `displayDialogAsync` 会返回以下三个特定错误。</span><span class="sxs-lookup"><span data-stu-id="cfc70-232">In addition to general platform and system errors, three errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="cfc70-233">代码编号</span><span class="sxs-lookup"><span data-stu-id="cfc70-233">Code number</span></span>|<span data-ttu-id="cfc70-234">含义</span><span class="sxs-lookup"><span data-stu-id="cfc70-234">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="cfc70-235">12004</span><span class="sxs-lookup"><span data-stu-id="cfc70-235">12004</span></span>|<span data-ttu-id="cfc70-p132">传递给 `displayDialogAsync` 的 URL 的域不受信任。此域必须与主机页的域相同（包括协议和端口号）。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p132">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="cfc70-238">12005</span><span class="sxs-lookup"><span data-stu-id="cfc70-238">12005</span></span>|<span data-ttu-id="cfc70-p133">传递给 `displayDialogAsync` 的 URL 使用 HTTP 协议。需要使用 HTTPS。（在 Office 的某些版本中，返回 12005 的错误消息与返回 12004 错误消息是相同的。）</span><span class="sxs-lookup"><span data-stu-id="cfc70-p133">The URL passed to `displayDialogAsync` uses the HTTP protocol. HTTPS is required. (In some versions of Office, the error message returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="cfc70-242"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="cfc70-242"><span id="12007">12007</span></span></span>|<span data-ttu-id="cfc70-p134">已从此主机窗口打开了一个对话框。主机窗口（如任务窗格）一次只能打开一个对话框。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p134">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|
|<span data-ttu-id="cfc70-245">12009</span><span class="sxs-lookup"><span data-stu-id="cfc70-245">12009</span></span>|<span data-ttu-id="cfc70-246">用户已选择忽略对话框。</span><span class="sxs-lookup"><span data-stu-id="cfc70-246">The user chose to ignore the dialog box.</span></span> <span data-ttu-id="cfc70-247">联机版本的 Office 中可能会发生此错误，用户可能会选择不允许加载项显示对话框。</span><span class="sxs-lookup"><span data-stu-id="cfc70-247">This error can occur in online versions of Office, where users may choose not to allow an add-in to present a dialog.</span></span>|

<span data-ttu-id="cfc70-248">调用 `displayDialogAsync` 时，总是将 [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js) 对象传递给它的回调函数。</span><span class="sxs-lookup"><span data-stu-id="cfc70-248">When `displayDialogAsync` is called, it always passes an [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js) object to its callback function.</span></span> <span data-ttu-id="cfc70-249">如果调用成功（即对话框窗口已打开），`AsyncResult` 对象的 `value` 属性是 [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) 对象。</span><span class="sxs-lookup"><span data-stu-id="cfc70-249">When the call is successful - that is, the dialog window is opened - the `value` property of the `AsyncResult` object is a [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) object.</span></span> <span data-ttu-id="cfc70-250">有关示例，请参阅[将信息从对话框发送到宿主页](#send-information-from-the-dialog-box-to-the-host-page)部分。</span><span class="sxs-lookup"><span data-stu-id="cfc70-250">An example of this is in the section [Send information from the dialog box to the host page](#send-information-from-the-dialog-box-to-the-host-page).</span></span> <span data-ttu-id="cfc70-251">如果调用 `displayDialogAsync` 失败，不会创建窗口，`AsyncResult` 对象的 `status` 属性设置为 `Office.AsyncResultStatus.Failed`，并且会填充对象的 `error` 属性。</span><span class="sxs-lookup"><span data-stu-id="cfc70-251">When the call to `displayDialogAsync` fails, the window is not created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated.</span></span> <span data-ttu-id="cfc70-252">应始终有用于测试 `status` 并在出错时进行响应的回调。</span><span class="sxs-lookup"><span data-stu-id="cfc70-252">You should always have a callback that tests the `status` and responds when it's an error.</span></span> <span data-ttu-id="cfc70-253">有关仅报告错误消息（无论代码编号是什么）的示例，请参阅下面的代码：</span><span class="sxs-lookup"><span data-stu-id="cfc70-253">For an example that simply reports the error message regardless of its code number, see the following code:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

### <a name="errors-and-events-in-the-dialog-window"></a><span data-ttu-id="cfc70-254">对话框窗口中的错误和事件</span><span class="sxs-lookup"><span data-stu-id="cfc70-254">Errors and events in the dialog window</span></span>

<span data-ttu-id="cfc70-255">对话框中的三个错误和事件（具有代码编码）会在主机页中触发 `DialogEventReceived` 事件。</span><span class="sxs-lookup"><span data-stu-id="cfc70-255">Three errors and events, known by their code numbers, in the dialog box will trigger a `DialogEventReceived` event in the host page.</span></span>

|<span data-ttu-id="cfc70-256">代码编号</span><span class="sxs-lookup"><span data-stu-id="cfc70-256">Code number</span></span>|<span data-ttu-id="cfc70-257">含义</span><span class="sxs-lookup"><span data-stu-id="cfc70-257">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="cfc70-258">12002</span><span class="sxs-lookup"><span data-stu-id="cfc70-258">12002</span></span>|<span data-ttu-id="cfc70-259">下列一种含义：</span><span class="sxs-lookup"><span data-stu-id="cfc70-259">One of the following:</span></span><br> <span data-ttu-id="cfc70-260">- 传递给 `displayDialogAsync` 的 URL 没有对应的页面。</span><span class="sxs-lookup"><span data-stu-id="cfc70-260">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="cfc70-261">- 传递给 `displayDialogAsync` 的页面已加载，但对话框定向到找不到或无法加载的页面，或者已定向到使用无效语法的 URL。</span><span class="sxs-lookup"><span data-stu-id="cfc70-261">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was directed to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="cfc70-262">12003</span><span class="sxs-lookup"><span data-stu-id="cfc70-262">12003</span></span>|<span data-ttu-id="cfc70-p137">对话框定向到使用 HTTP 协议的 URL。必须使用 HTTPS。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p137">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="cfc70-265">12006</span><span class="sxs-lookup"><span data-stu-id="cfc70-265">12006</span></span>|<span data-ttu-id="cfc70-266">对话框已关闭，通常是因为用户选择了 **X** 按钮。</span><span class="sxs-lookup"><span data-stu-id="cfc70-266">The dialog box was closed, usually because the user chooses the **X** button.</span></span>|

<span data-ttu-id="cfc70-p138">代码可以在调用 `displayDialogAsync` 时分配 `DialogEventReceived` 事件处理程序。下面展示了一个简单示例：</span><span class="sxs-lookup"><span data-stu-id="cfc70-p138">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="cfc70-269">有关为每个错误代码创建自定义错误消息的 `DialogEventReceived` 事件处理程序示例，请参阅下面的示例：</span><span class="sxs-lookup"><span data-stu-id="cfc70-269">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example:</span></span>

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

<span data-ttu-id="cfc70-270">有关这样处理错误的样本加载项，请参阅 [Office 加载项 Dialog API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。</span><span class="sxs-lookup"><span data-stu-id="cfc70-270">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>


## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="cfc70-271">向对话框传递信息</span><span class="sxs-lookup"><span data-stu-id="cfc70-271">Pass information to the dialog box</span></span>

<span data-ttu-id="cfc70-p139">有时，主机页需要向对话框传递信息。完成此操作的方式主要分为两种：</span><span class="sxs-lookup"><span data-stu-id="cfc70-p139">Sometimes the host page needs to pass information to the dialog box. You can do this in two primary ways:</span></span>

- <span data-ttu-id="cfc70-274">向传递给 `displayDialogAsync` 的 URL 添加查询参数。</span><span class="sxs-lookup"><span data-stu-id="cfc70-274">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="cfc70-p140">将信息存储在主机窗口和对话框都可访问的位置。这两个窗口不共享通用会话存储，但*如果它们具有相同的域*（包括端口号，若有），则共享通用[本地存储](https://www.w3schools.com/html/html5_webstorage.asp)。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p140">Store the information somewhere that is accessible to both the host window and dialog box. The two windows do not share a common session storage, but *if they have the same domain* (including port number, if any),  they share a common [local storage](https://www.w3schools.com/html/html5_webstorage.asp).</span></span>

### <a name="use-local-storage"></a><span data-ttu-id="cfc70-277">使用本地存储</span><span class="sxs-lookup"><span data-stu-id="cfc70-277">Use local storage</span></span>

<span data-ttu-id="cfc70-278">为了使用本地存储，代码会先在主机页中调用 `window.localStorage` 对象的 `setItem` 方法，然后再调用 `displayDialogAsync`，如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="cfc70-278">To use local storage, your code calls the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="cfc70-279">对话框窗口中的代码会在需要时读取项，如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="cfc70-279">Code in the dialog window reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

<span data-ttu-id="cfc70-280">有关通过这种方式使用本地存储的样本加载项，请参阅：</span><span class="sxs-lookup"><span data-stu-id="cfc70-280">For sample add-ins that uses local storage in this way, see:</span></span>

- [<span data-ttu-id="cfc70-281">使用 Auth0 服务简化社交登录的 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="cfc70-281">Office Add-in that uses the Auth0 Service to Simplify Social Login</span></span>](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [<span data-ttu-id="cfc70-282">使用 OAuth.io 服务简化热门在线服务访问的 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="cfc70-282">Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services</span></span>](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

### <a name="use-query-parameters"></a><span data-ttu-id="cfc70-283">使用查询参数</span><span class="sxs-lookup"><span data-stu-id="cfc70-283">Use query parameters</span></span>

<span data-ttu-id="cfc70-284">下面的示例展示了如何使用查询参数传递数据：</span><span class="sxs-lookup"><span data-stu-id="cfc70-284">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="cfc70-285">有关使用此技术的示例，请参阅[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)（在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表）。</span><span class="sxs-lookup"><span data-stu-id="cfc70-285">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="cfc70-286">对话框窗口中的代码可以分析 URL，并读取参数值。</span><span class="sxs-lookup"><span data-stu-id="cfc70-286">Code in your dialog window can parse the URL and read the parameter value.</span></span>

> [!NOTE]
> <span data-ttu-id="cfc70-p141">Office 会自动向传递给 `displayDialogAsync` 的 URL 添加查询参数 `_host_info`。（附加在自定义查询参数（若有）后面，不会附加到对话框导航到的任何后续 URL。）Microsoft 可能会更改此值的内容，或者将来会将其全部删除，因此代码不得读取此值。相同的值会被添加到对话框的会话存储中。同样，*代码不得对此值执行读取和写入操作*。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p141">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.</span></span>

## <a name="use-the-dialog-apis-to-show-a-video"></a><span data-ttu-id="cfc70-292">使用对话框 API 显示视频</span><span class="sxs-lookup"><span data-stu-id="cfc70-292">Use the Dialog APIs to show a video</span></span>

<span data-ttu-id="cfc70-293">若要在对话框中显示视频，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="cfc70-293">To show a video in a dialog box:</span></span>

1.  <span data-ttu-id="cfc70-p142">创建内容仅有 iframe 的页面。iframe 的 `src` 属性指向联机视频。视频 URL 必须使用 HTTP**S** 协议。本文将此页面称为“video.dialogbox.html”。下面展示了标记示例：</span><span class="sxs-lookup"><span data-stu-id="cfc70-p142">Create a page whose only content is an iframe. The `src` attribute of the iframe points to an online video. The protocol of the video's URL must be HTTP**S**. In this article we'll call this page "video.dialogbox.html". The following is an example of the markup:</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2.  <span data-ttu-id="cfc70-299">video.dialogbox.html 页面必须与主机页位于同一域中。</span><span class="sxs-lookup"><span data-stu-id="cfc70-299">The video.dialogbox.html page must be in the same domain as the host page.</span></span>
3.  <span data-ttu-id="cfc70-300">在主机页中调用 `displayDialogAsync`，打开 video.dialogbox.html。</span><span class="sxs-lookup"><span data-stu-id="cfc70-300">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
4.  <span data-ttu-id="cfc70-p143">如果外接程序需要知道用户何时关闭对话框，请为 `DialogEventReceived` 事件注册处理程序，并处理 12006 事件。有关详细信息，请参阅[对话框窗口中的错误和事件](#errors-and-events-in-the-dialog-window)部分。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p143">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event. For details, see the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window).</span></span>

<span data-ttu-id="cfc70-303">有关在对话框中显示视频的示例，请参阅 [Office 外接程序的用户体验设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)存储库中的[视频展示位置设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat)。</span><span class="sxs-lookup"><span data-stu-id="cfc70-303">For a sample that shows a video in a dialog box, see the [video placemat design pattern](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat) in the [UX design patterns for Office Add-ins](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) repo.</span></span>

![在加载项对话框中显示的视频的屏幕截图](../images/video-placemats-dialog-open.png)

## <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="cfc70-305">在身份验证流中使用对话框 API</span><span class="sxs-lookup"><span data-stu-id="cfc70-305">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="cfc70-306">对话框 API 的主要应用场景是为不允许在 Iframe 中打开登录页的资源或标识提供程序（如 Microsoft 帐户、Office 365、Google 和 Facebook）启用身份验证。</span><span class="sxs-lookup"><span data-stu-id="cfc70-306">A primary scenario for the Dialog APIs is to enable authentication with a resource or identity provider that does not allow its sign-in page to open in an Iframe, such as Microsoft Account, Office 365, Google, and Facebook.</span></span>

> [!NOTE]
> <span data-ttu-id="cfc70-p144">若要将对话框 API 用于此应用场景，请*勿*在调用 `displayDialogAsync` 时使用 `displayInIframe: true` 选项。请参阅本文前面的[使用 Office Online 中的性能选项](#take-advantage-of-a-performance-option-in-office-online)，详细了解此选项。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p144">When you are using the Dialog APIs for this scenario, do *not* use the `displayInIframe: true` option in the call to `displayDialogAsync`. See [Take advantage of a performance option in Office Online](#take-advantage-of-a-performance-option-in-office-online) previously in this article for details about this option.</span></span>

<span data-ttu-id="cfc70-309">下面展示了简单的典型身份验证流：</span><span class="sxs-lookup"><span data-stu-id="cfc70-309">The following is a simple and typical authentication flow:</span></span>

1. <span data-ttu-id="cfc70-p145">对话框中打开的第一个页面是加载项域（即主机窗口域）中托管的本地页面（或其他资源）。此页面可以显示简单的 UI，提示用户“请稍候，我们正在将你重定向到可以登录 *NAME-OF-PROVIDER* 的页面。”此页面中的代码使用传递给对话框的信息，构造标识提供程序的登录页 URL，如[向对话框传递信息](#pass-information-to-the-dialog-box)中所述。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p145">The first page that opens in the dialog box is a local page (or other resource) that is hosted in the add-in's domain; that is, the host window's domain. This page can have a simple UI that says "Please wait, we are redirecting you to the page where you can sign in to *NAME-OF-PROVIDER*." Code in this page constructs the URL of the identity provider's sign-in page by using information that is passed to the dialog box as described in [Pass information to the dialog box](#pass-information-to-the-dialog-box).</span></span>
2. <span data-ttu-id="cfc70-p146">然后，对话框窗口重定向到登录页。URL 包含一个查询参数，用于提示标识提供程序在用户登录特定页面后重定向对话框窗口。在本文中，我们将此页面称为 "redirectPage.html"。（*此页面必须与主机窗口位于相同域中*，因为对话框窗口传递登录尝试结果的唯一方法就是调用 `messageParent`，而它只能在与主机窗口位于同一域的页面上调用）。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p146">The dialog window then redirects to the sign-in page. The URL includes a query parameter that tells the identity provider to redirect the dialog window, after the user signs in, to a specific page. In this article, we'll call this page "redirectPage.html". (*This must be a page in the same domain as the host window*, because the only way for the dialog window to pass the results of the sign-in attempt is with a call of `messageParent`, which can only be called on a page with the same domain as the host window.)</span></span>
2. <span data-ttu-id="cfc70-p147">标识提供程序的服务处理来自对话框窗口的传入 GET 请求。如果用户已经登录，它会立即将窗口重定向到 redirectPage.html，并将用户数据作为查询参数添加。如果用户尚未登录，提供程序的登录页会显示在窗口中，以便用户登录。对于大多数提供程序，如果用户无法成功登录，提供程序会在对话框窗口中显示错误页面，而不会重定向到 redirectPage.html。用户必须通过选择右上角的 **X** 来关闭窗口。如果用户成功登录，则对话框窗口会重定向到 redirectPage.html，并且用户数据会作为查询参数添加。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p147">The identity provider's service processes the incoming GET request from the dialog window. If the user is already logged on, it immediately redirects the window to redirectPage.html and includes user data as a query parameter. If the user is not already signed in, the provider's sign-in page appears in the window, and the user signs in. For most providers, if the user cannot sign in successfully, the provider shows an error page in the dialog window and does not redirect to redirectPage.html. The user must close the window by selecting the **X** in the corner. If the user successfully signs in, the dialog window is redirected to redirectPage.html and user data is included as a query parameter.</span></span>
3. <span data-ttu-id="cfc70-323">当 redirectPage.html 页面打开时，它会调用 `messageParent` 向主机页报告登录是否成功，而且还会视情况报告用户数据或错误数据。</span><span class="sxs-lookup"><span data-stu-id="cfc70-323">When the redirectPage.html page opens, it calls `messageParent` to report the success or failure to the host page and optionally also report user data or error data.</span></span>
4. <span data-ttu-id="cfc70-324">`DialogMessageReceived` 事件在主机页中触发，其处理程序关闭对话框窗口，并视情况对消息进行其他处理。</span><span class="sxs-lookup"><span data-stu-id="cfc70-324">The `DialogMessageReceived` event fires in the host page and its handler closes the dialog window and optionally does other processing of the message.</span></span>

<span data-ttu-id="cfc70-325">有关使用此模式的示例加载项，请参阅：</span><span class="sxs-lookup"><span data-stu-id="cfc70-325">For sample add-ins that use this pattern, see:</span></span>

- <span data-ttu-id="cfc70-326">[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)（在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表）：对话框窗口最初打开的资源是没有自己视图的控制器方法。</span><span class="sxs-lookup"><span data-stu-id="cfc70-326">[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart): The resource that is initially opened in the dialog window is a controller method that has no view of its own.</span></span> <span data-ttu-id="cfc70-327">然后，其重定向到 Office 365 登录页。</span><span class="sxs-lookup"><span data-stu-id="cfc70-327">It redirects to the Office 365 sign in page.</span></span>
- <span data-ttu-id="cfc70-328">[Office Add-in Office 365 Client Authentication for AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)（Office 加载项 Office 365 客户端 AngularJS 身份验证）：对话框窗口最初打开的资源是一个页面。</span><span class="sxs-lookup"><span data-stu-id="cfc70-328">[Office Add-in Office 365 Client Authentication for AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth): The resource that is initially opened in the dialog window is a page.</span></span>

#### <a name="support-multiple-identity-providers"></a><span data-ttu-id="cfc70-329">支持多个标识提供程序</span><span class="sxs-lookup"><span data-stu-id="cfc70-329">Support multiple identity providers</span></span>

<span data-ttu-id="cfc70-p149">如果外接程序允许用户选择提供程序（如 Microsoft 帐户、Google 或 Facebook），你需要使用本地第一个页面（见前一部分），为用户提供用于选择提供程序的 UI。用户的选择会触发登录 URL 的构建并重定向到该 URL。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p149">If your add-in gives the user a choice of providers, such as Microsoft Account, Google, or Facebook, you need a local first page (see preceding section) that provides a UI for the user to select a provider. Selection triggers the construction of the sign-in URL and redirection to it.</span></span>

<span data-ttu-id="cfc70-332">有关使用此模式的示例，请参阅[使用 Auth0 服务简化社交登录的 Office 外接程序](https://github.com/OfficeDev/Office-Add-in-Auth0)。</span><span class="sxs-lookup"><span data-stu-id="cfc70-332">For a sample that uses this pattern, see [Office Add-in that uses the Auth0 Service to Simplify Social Login](https://github.com/OfficeDev/Office-Add-in-Auth0).</span></span>

#### <a name="authorization-of-the-add-in-to-an-external-resource"></a><span data-ttu-id="cfc70-333">在外接程序中授权外部资源</span><span class="sxs-lookup"><span data-stu-id="cfc70-333">Authorization of the add-in to an external resource</span></span>

<span data-ttu-id="cfc70-p150">在现代网络中，Web 应用程序是安全主体（就像用户一样），拥有自己的标识以及对联机资源（如 Office 365、Google Plus、Facebook 或 LinkedIn）的权限。在部署前，需要先向资源提供程序注册应用程序。注册内容包括：</span><span class="sxs-lookup"><span data-stu-id="cfc70-p150">In the modern web, web applications are security principals just as users are, and the application has its own identity and permissions to an online resource such as Office 365, Google Plus, Facebook, or LinkedIn. The application is registered with the resource provider before it is deployed. The registration includes:</span></span>

- <span data-ttu-id="cfc70-337">应用程序访问用户资源所需的权限的列表。</span><span class="sxs-lookup"><span data-stu-id="cfc70-337">A list of the permissions that the application needs to a user's resources.</span></span>
- <span data-ttu-id="cfc70-338">当应用访问服务时，资源服务应向其返回访问令牌的 URL。</span><span class="sxs-lookup"><span data-stu-id="cfc70-338">A URL to which the resource service should return an access token when the application accesses the service.</span></span>  

<span data-ttu-id="cfc70-p151">如果用户在应用中调用访问资源服务中用户数据的函数，系统会先提示用户登录相应服务，再提示用户向应用授予访问用户资源所需的权限。然后，服务将登录窗口重定向到先前注册的 URL，并传递访问令牌。应用使用访问令牌访问用户资源。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p151">When a user invokes a function in the application that accesses the user's data in the resource service, they are prompted to sign in to the service and then prompted to grant the application the permissions it needs to the user's resources. The service then redirects the sign-in window to the previously registered URL and passes the access token. The application uses the access token to access the user's resources.</span></span>

<span data-ttu-id="cfc70-p152">可以使用对话框 API 来管理此过程，具体方法是使用与用户登录流类似的流。只有下面两处不同：</span><span class="sxs-lookup"><span data-stu-id="cfc70-p152">You can use the Dialog APIs to manage this process by using a flow that is similar to the one described for users to sign in. The only differences are:</span></span>

- <span data-ttu-id="cfc70-344">如果用户先前未向应用程序授予所需的权限，则登录后会在对话框中看到这样做的提示。</span><span class="sxs-lookup"><span data-stu-id="cfc70-344">If the user hasn't previously granted the application the permissions it needs, she is prompted to do so in the dialog box after signing in.</span></span>
- <span data-ttu-id="cfc70-p153">对话框窗口使用 `messageParent` 发送字符串化访问令牌，或将访问令牌存储在主机窗口可以检索到的位置，从而将访问令牌发送给主机窗口。令牌具有时间限制，但在持续期间，主机窗口可以使用它直接访问用户资源，而无需进一步提示。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p153">The dialog window sends the access token to the host window either by using `messageParent` to send the stringified access token or by storing the access token where the host window can retrieve it. The token has a time limit, but while it lasts, the host window can use it to directly access the user's resources without any further prompting.</span></span>

<span data-ttu-id="cfc70-347">下面的示例使用对话框 API 实现此目的：</span><span class="sxs-lookup"><span data-stu-id="cfc70-347">The following samples use the Dialog APIs for this purpose:</span></span>
- <span data-ttu-id="cfc70-348">[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)（在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表） - 将访问令牌存储在数据库中。</span><span class="sxs-lookup"><span data-stu-id="cfc70-348">[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) - Stores the access token in a database.</span></span>
- <span data-ttu-id="cfc70-349">[Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services](https://github.com/OfficeDev/Office-Add-in-OAuth.io)（使用 OAuth.io 服务简化热门在线服务访问的 Office 加载项）</span><span class="sxs-lookup"><span data-stu-id="cfc70-349">[Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services](https://github.com/OfficeDev/Office-Add-in-OAuth.io)</span></span>

<span data-ttu-id="cfc70-350">若要详细了解加载项中的身份验证和授权，请参阅：</span><span class="sxs-lookup"><span data-stu-id="cfc70-350">For more information about authentication and authorization in add-ins, see:</span></span>
- [<span data-ttu-id="cfc70-351">在 Office 加载项中授权外部服务</span><span class="sxs-lookup"><span data-stu-id="cfc70-351">Authorize external services in your Office Add-in</span></span>](auth-external-add-ins.md)
- [<span data-ttu-id="cfc70-352">Office JavaScript API 帮助程序库</span><span class="sxs-lookup"><span data-stu-id="cfc70-352">Office JavaScript API Helpers library</span></span>](https://github.com/OfficeDev/office-js-helpers)


## <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="cfc70-353">将 Office Dialog API 与单页应用程序和客户端路由结合使用</span><span class="sxs-lookup"><span data-stu-id="cfc70-353">Use the Office Dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="cfc70-354">如果外接程序使用客户端路由（单页应用程序通常这样做），则可以选择将路由 URL 传递给 [ displayDialogAsync ](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) 方法，而不是传递各个完整 HTML 页面的 URL。</span><span class="sxs-lookup"><span data-stu-id="cfc70-354">If your add-in uses client-side routing, as single-page applications typically do, you have the option to pass the URL of a route to the [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) method, instead of the URL of a complete and separate HTML page.</span></span>

> [!IMPORTANT]
><span data-ttu-id="cfc70-p154">对话框位于有自己执行上下文的新窗口中。如果传递路由，基页面及其所有初始化和启动代码都会在这个新上下文中再次运行，且所有变量都会在对话框中设置为各自的初始值。所以，此技术会在对话框窗口中启动应用的第二个实例。更改对话框窗口中变量的代码不会更改任务窗格版本的相同变量。同样，对话框窗口有自己的会话存储，任务窗格中的代码无法访问此类存储。</span><span class="sxs-lookup"><span data-stu-id="cfc70-p154">The dialog box is in a new window with its own execution context. If you pass a route, your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog window. So this technique launches a second instance of your application in the dialog window. Code that changes variables in the dialog window does not change the task pane version of the same variables. Similarly, the dialog window has its own session storage, which is not accessible from code in the task pane.</span></span>
