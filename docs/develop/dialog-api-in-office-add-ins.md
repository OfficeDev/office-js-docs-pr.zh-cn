---
title: 在 Office 加载项中使用对话框 API
description: ''
ms.date: 08/07/2019
localization_priority: Priority
ms.openlocfilehash: 5cafb2396c92576bd5ac6d6d52105e0bb5ee579d
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/13/2019
ms.locfileid: "36302579"
---
# <a name="use-the-dialog-api-in-your-office-add-ins"></a><span data-ttu-id="80966-102">在 Office 加载项中使用对话框 API</span><span class="sxs-lookup"><span data-stu-id="80966-102">Use the Dialog API in your Office Add-ins</span></span>

<span data-ttu-id="80966-p101">可以在 Office 外接程序中使用[对话框 API](/javascript/api/office/office.ui) 打开对话框。本文提供了有关如何在 Office 外接程序中使用对话框 API 的指南。</span><span class="sxs-lookup"><span data-stu-id="80966-p101">You can use the [Dialog API](/javascript/api/office/office.ui) to open dialog boxes in your Office Add-in. This article provides guidance for using the Dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="80966-p102">若要了解对话框 API 目前的受支持情况，请参阅[对话框 API 要求集](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets)。目前，Word、Excel、PowerPoint 和 Outlook 支持对话框 API。</span><span class="sxs-lookup"><span data-stu-id="80966-p102">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.</span></span>

<span data-ttu-id="80966-107">对话框 API 的主要应用场景是为 Google、Facebook 或 Microsoft Graph 等资源启用身份验证。</span><span class="sxs-lookup"><span data-stu-id="80966-107">A primary scenario for the Dialog APIs is to enable authentication with a resource such as Google or Facebook.</span></span> <span data-ttu-id="80966-108">有关详细信息，请在熟悉本文*之后*，参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。</span><span class="sxs-lookup"><span data-stu-id="80966-108">For more information, see [Authenticate with the Office Dialog API](auth-with-office-dialog-api.md) *after* you are familiar with this article.</span></span>

<span data-ttu-id="80966-109">不妨通过任务窗格/内容加载项/[加载项命令](../design/add-in-commands.md)打开对话框，以便执行下列操作：</span><span class="sxs-lookup"><span data-stu-id="80966-109">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="80966-110">显示无法直接在任务窗格中打开的登录页。</span><span class="sxs-lookup"><span data-stu-id="80966-110">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="80966-111">为加载项中的某些任务提供更多屏幕空间，或甚至整个屏幕。</span><span class="sxs-lookup"><span data-stu-id="80966-111">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="80966-112">托管在任务窗格中显得太小的视频。</span><span class="sxs-lookup"><span data-stu-id="80966-112">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="80966-p104">由于不赞成重叠 UI 元素，因此除非应用场景需要，否则请勿从任务窗格打开对话框。考虑如何使用任务窗格的区域时，请注意任务窗格可以是选项卡式。有关示例，请参阅 [Excel 加载项 JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 示例。</span><span class="sxs-lookup"><span data-stu-id="80966-p104">Because overlapping UI elements are discouraged, avoid opening a dialog from a task pane unless your scenario requires it. When you consider how to use the surface area of a task pane, note that task panes can be tabbed. For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="80966-116">下图展示了对话框示例。</span><span class="sxs-lookup"><span data-stu-id="80966-116">The following image shows an example of a dialog box.</span></span>

![加载项命令](../images/auth-o-dialog-open.png)

<span data-ttu-id="80966-p105">请注意，对话框总是在屏幕中央打开。用户可以移动它，并重设大小。对话框是*非模式*窗口。也就是说，用户可以继续与主机 Office 应用中的文档，以及与任务窗格中的主机页（若有）进行交互。</span><span class="sxs-lookup"><span data-stu-id="80966-p105">Note that the dialog box always opens in the center of the screen. The user can move and resize it. The window is *nonmodal*--a user can continue to interact with both the document in the host Office application and with the host page in the task pane, if there is one.</span></span>

## <a name="dialog-api-scenarios"></a><span data-ttu-id="80966-121">对话框 API 应用场景</span><span class="sxs-lookup"><span data-stu-id="80966-121">Dialog API scenarios</span></span>

<span data-ttu-id="80966-122">Office JavaScript API 支持以下应用场景，其在 [Office.context.ui 命名空间](/javascript/api/office/office.ui)中使用 [Dialog](/javascript/api/office/office.dialog) 对象和两个函数。</span><span class="sxs-lookup"><span data-stu-id="80966-122">The Office JavaScript APIs support the following scenarios with a [Dialog](/javascript/api/office/office.dialog) object and two functions in the [Office.context.ui namespace](/javascript/api/office/office.ui).</span></span>

### <a name="open-a-dialog-box"></a><span data-ttu-id="80966-123">打开对话框</span><span class="sxs-lookup"><span data-stu-id="80966-123">Open a dialog box</span></span>

<span data-ttu-id="80966-p106">为了打开对话框，任务窗格中的代码调用 [displayDialogAsync](/javascript/api/office/office.ui) 方法，并将要打开的资源 URL 传递到此方法。这通常是页面，但也可以是 MVC 应用中的控制器方法、路由、Web 服务方法或其他任何资源。在本文中，“页面”或“网站”是指对话框中的资源。下面的代码就是一个简单示例：</span><span class="sxs-lookup"><span data-stu-id="80966-p106">To open a dialog box, your code in the task pane calls the [displayDialogAsync](/javascript/api/office/office.ui) method and passes to it the URL of the resource that you want to open. This is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource. In this article, 'page' or 'website' refers to the resource in the dialog. The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="80966-p107">URL 使用 **HTTPS** 协议。对话框中加载的所有页面都必须遵循此要求，而不仅仅是加载的第一个页面。</span><span class="sxs-lookup"><span data-stu-id="80966-p107">The URL uses the HTTP**S** protocol. This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="80966-130">对话框资源域与宿主页的域相同，宿主页可以是任务窗格中的页面，也可以是加载项命令的[函数文件](/office/dev/add-ins/reference/manifest/functionfile)。</span><span class="sxs-lookup"><span data-stu-id="80966-130">The dialog resource's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](/office/dev/add-ins/reference/manifest/functionfile) of an add-in command.</span></span> <span data-ttu-id="80966-131">这要求：传递到 `displayDialogAsync` 方法的页面、控制器方法或其他资源必须与主机页位于相同的域。</span><span class="sxs-lookup"><span data-stu-id="80966-131">This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="80966-132">宿主页面和对话框资源必须具有相同的完整域。</span><span class="sxs-lookup"><span data-stu-id="80966-132">The host page and the resources of the dialog must have the same full domain.</span></span> <span data-ttu-id="80966-133">如果尝试传递 `displayDialogAsync` 加载项域的子域，则不会起作用。</span><span class="sxs-lookup"><span data-stu-id="80966-133">If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work.</span></span> <span data-ttu-id="80966-134">完整域（包括任何子域）必须匹配。</span><span class="sxs-lookup"><span data-stu-id="80966-134">The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="80966-p110">在第一个页面（或其他资源）加载后，用户可以转到使用 HTTPS 的任意网站（或其他资源）。还可以将第一个页面设计为直接重定向到另一个站点。</span><span class="sxs-lookup"><span data-stu-id="80966-p110">After the first page (or other resource) is loaded, a user can go to any website (or other resource) that uses HTTPS. You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="80966-137">默认情况下，对话框的高度和宽度占设备屏幕的 80%。不过，也可以设置不同的百分比，只需将配置对象传递给方法即可，如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="80966-137">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="80966-138">有关实现这一点的样本加载项，请参阅 [Office 加载项 Dialog API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。</span><span class="sxs-lookup"><span data-stu-id="80966-138">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="80966-p111">将两个值均设置为 100% 可有效提供全屏体验。（有效最大值为 99.5%，窗口仍可移动和调整大小。）</span><span class="sxs-lookup"><span data-stu-id="80966-p111">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="80966-p112">只能从主机窗口打开一个对话框。如果尝试再打开一个对话框，就会生成错误。比方说，如果用户从任务窗格打开一个对话框，她就无法再从任务窗格中的其他页面打开第二个对话框。不过，如果对话框是通过[加载项命令](../design/add-in-commands.md)打开，那么只要选择此命令，就会打开新 HTML 文件（但不可见）。这会新建（不可见的）主机窗口，所以每个这样的窗口都可以启动自己的对话框。有关详细信息，请参阅 [displayDialogAsync 返回的错误](#errors-from-displaydialogasync)。</span><span class="sxs-lookup"><span data-stu-id="80966-p112">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a><span data-ttu-id="80966-147">利用 Office 网页版中的性能选项</span><span class="sxs-lookup"><span data-stu-id="80966-147">Take advantage of a performance option in Office Online</span></span>

<span data-ttu-id="80966-148">`displayInIframe` 属性是配置对象中另一个可以传递到 `displayDialogAsync` 的属性。</span><span class="sxs-lookup"><span data-stu-id="80966-148">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`.</span></span> <span data-ttu-id="80966-149">如果将此属性设置为 `true`，且加载项在 Office 网页版打开的文档中运行，对话框就会以浮动 iframe（而不是独立窗口）的形式打开，从而加快对话框的打开速度。</span><span class="sxs-lookup"><span data-stu-id="80966-149">When this property is set to `true`, and the add-in is running in a document opened in Office Online, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster.</span></span> <span data-ttu-id="80966-150">示例如下：</span><span class="sxs-lookup"><span data-stu-id="80966-150">The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="80966-151">默认值为 `false`，与完全省略此属性时相同。</span><span class="sxs-lookup"><span data-stu-id="80966-151">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="80966-152">如果加载项没有在 Office 网页版中运行，`displayInIframe` 将被忽略。</span><span class="sxs-lookup"><span data-stu-id="80966-152">If the add-in is not running in Office Online, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="80966-p115">如果对话框始终重定向到无法在 iframe 中打开的页面，**不**得使用 `displayInIframe: true`。例如，许多热门 Web 服务（如 Google 和 Microsoft 帐户）的登录页都无法在 iframe 中打开。</span><span class="sxs-lookup"><span data-stu-id="80966-p115">You should **not** use `displayInIframe: true` if the dialog will at any point redirect to a page that cannot be opened in an iframe. For example, the sign in pages of many popular web services, such as Google and Microsoft Account, cannot be opened in an iframe.</span></span>

### <a name="handling-pop-up-blockers-with-office-on-the-web"></a><span data-ttu-id="80966-155">使用 Office 网页版处理弹出窗口阻止程序</span><span class="sxs-lookup"><span data-stu-id="80966-155">Handling pop-up blockers with Office on the web</span></span>

<span data-ttu-id="80966-156">如果尝试在使用 Office 网页版时显示对话框，可能会导致浏览器的弹出窗口阻止程序阻止对话框。</span><span class="sxs-lookup"><span data-stu-id="80966-156">Attempting to display a dialog while using Office Online may cause the browser's pop-up blocker to block the dialog.</span></span> <span data-ttu-id="80966-157">如果加载项用户先同意加载项发出的提示，可以避开浏览器的弹出窗口阻止程序。</span><span class="sxs-lookup"><span data-stu-id="80966-157">The browser's pop-up blocker can be circumvented if the user of your add-in first agrees to a prompt from the add-in.</span></span> <span data-ttu-id="80966-158">`displayDialogAsync` 的 [DialogOptions](/javascript/api/office/office.dialogoptions) 包含可触发此类弹出窗口的 `promptBeforeOpen` 属性。</span><span class="sxs-lookup"><span data-stu-id="80966-158">`displayDialogAsync`'s [DialogOptions](/javascript/api/office/office.dialogoptions) has the `promptBeforeOpen` property to trigger such a pop-up.</span></span> <span data-ttu-id="80966-159">`promptBeforeOpen` 是提供以下行为的布尔值：</span><span class="sxs-lookup"><span data-stu-id="80966-159">`promptBeforeOpen` is a boolean value which provides the following behavior:</span></span>

 - <span data-ttu-id="80966-160">`true` - 框架显示用于触发导航的弹出窗口，并避开浏览器的弹出窗口阻止程序。</span><span class="sxs-lookup"><span data-stu-id="80966-160">`true` - The framework displays a pop-up to trigger the navigation and avoid the browser's pop-up blocker.</span></span> 
 - <span data-ttu-id="80966-161">`false` - 对话框不会显示，开发人员必须处理弹出窗口（通过提供用户界面项目来触发导航）。</span><span class="sxs-lookup"><span data-stu-id="80966-161">`false` - The dialog will not be shown and the developer must handle pop-ups (by providing a user interface artifact to trigger the navigation).</span></span> 
 
<span data-ttu-id="80966-162">弹出窗口如以下屏幕截图中所示：</span><span class="sxs-lookup"><span data-stu-id="80966-162">The pop-up looks similiar to that in the following screenshot:</span></span>

![加载项对话框可以生成提示，以避免浏览器内的弹出窗口阻止程序。](../images/dialog-prompt-before-open.png)
 
### <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="80966-164">将信息从对话框发送到主机页</span><span class="sxs-lookup"><span data-stu-id="80966-164">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="80966-165">对话框无法与任务窗格中的主机页进行通信，除非：</span><span class="sxs-lookup"><span data-stu-id="80966-165">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="80966-166">对话框中的当前页面与主机页在同一个域中。</span><span class="sxs-lookup"><span data-stu-id="80966-166">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="80966-p117">Office JavaScript 库已在页面中加载。（与使用 Office JavaScript 库的所有页面一样，页面脚本必须为 `Office.initialize` 属性分配方法，尽管方法可以是空的。有关详细信息，请参阅[初始化外接程序](understanding-the-javascript-api-for-office.md#initializing-your-add-in)。）</span><span class="sxs-lookup"><span data-stu-id="80966-p117">The Office JavaScript library is loaded in the page. (Like any page that uses the Office JavaScript library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initializing your add-in](understanding-the-javascript-api-for-office.md#initializing-your-add-in).)</span></span>

<span data-ttu-id="80966-p118">对话框页中的代码使用 `messageParent` 函数，向主机页发送布尔值或字符串消息。字符串可以是字词、句子、XML blob、字符串化 JSON 或其他任何能够串行化为字符串的内容。示例如下：</span><span class="sxs-lookup"><span data-stu-id="80966-p118">Code in the dialog page uses the `messageParent` function to send either a Boolean value or a string message to the host page. The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string. The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - <span data-ttu-id="80966-p119">`messageParent` 函数是*唯一*可以在对话框中调用的两个 Office API 之一。另一个是 `Office.context.requirements.isSetSupported`。有关详细信息，请参阅[指定 Office 主机和 API 要求](specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="80966-p119">The `messageParent` function is one of *only* two Office APIs that can be called in the dialog box. The other is `Office.context.requirements.isSetSupported`. For information about it, see [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md).</span></span>
> - <span data-ttu-id="80966-176">`messageParent` 函数只能在与主机页位于同一域（包括协议和端口）的页面上调用。</span><span class="sxs-lookup"><span data-stu-id="80966-176">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>

<span data-ttu-id="80966-177">在下一个示例中，`googleProfile` 是用户 Google 配置文件的字符串化版本。</span><span class="sxs-lookup"><span data-stu-id="80966-177">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="80966-p120">必须将主机页配置为接收消息。为此，可以向 `displayDialogAsync` 的原始调用添加回调参数。回调向 `DialogMessageReceived` 事件分配处理程序。示例如下：</span><span class="sxs-lookup"><span data-stu-id="80966-p120">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

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
> - <span data-ttu-id="80966-p121">Office 将 [AsyncResult](/javascript/api/office/office.asyncresult) 对象传递给回调。它表示尝试打开对话框的结果，不表示对话框中任何事件的结果。若要详细了解此区别，请参阅[处理错误和事件](#handle-errors-and-events)部分。</span><span class="sxs-lookup"><span data-stu-id="80966-p121">Office passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to the callback. It represents the result of the attempt to open the dialog box. It does not represent the outcome of any events in the dialog box. For more on this distinction, see the section [Handle errors and events](#handle-errors-and-events).</span></span>
> - <span data-ttu-id="80966-186">`asyncResult` 的 `value` 属性设置为 [Dialog](/javascript/api/office/office.dialog) 对象，此对象位于主机页（而不是对话框的执行上下文）中。</span><span class="sxs-lookup"><span data-stu-id="80966-186">The `value` property of the `asyncResult` is set to a [Dialog](/javascript/api/office/office.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="80966-p122">`processMessage` 是用于处理事件的函数。可以根据需要任意命名。</span><span class="sxs-lookup"><span data-stu-id="80966-p122">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="80966-189">`dialog` 变量的声明范围比回调更广，因为 `processMessage` 中也会引用此变量。</span><span class="sxs-lookup"><span data-stu-id="80966-189">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="80966-190">下面展示了 `DialogMessageReceived` 事件处理程序的简单示例：</span><span class="sxs-lookup"><span data-stu-id="80966-190">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="80966-p123">Office 将 `arg` 对象传递给处理程序。它的 `message` 属性是对话框中的 `messageParent` 调用发送的布尔值或字符串。在此示例中，它是 Microsoft 帐户或 Google 等服务的用户配置文件的字符串化表示。因此，使用 `JSON.parse` 将其反序列化回对象。</span><span class="sxs-lookup"><span data-stu-id="80966-p123">Office passes the `arg` object to the handler. Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog. In this example, it is a stringified representation of a user's profile from a service such as Microsoft Account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="80966-p124">未显示 `showUserName` 实现。它可能在任务窗格上显示定制的欢迎消息。</span><span class="sxs-lookup"><span data-stu-id="80966-p124">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="80966-196">在用户完成与对话框的交互后，消息处理程序应关闭对话框，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="80966-196">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="80966-197">`dialog` 对象必须是 `displayDialogAsync` 调用返回的对象。</span><span class="sxs-lookup"><span data-stu-id="80966-197">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="80966-198">`dialog.close` 调用指示 Office 立即关闭对话框。</span><span class="sxs-lookup"><span data-stu-id="80966-198">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="80966-199">有关使用这些技术的示例加载项，请参阅 [Office 加载项对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。</span><span class="sxs-lookup"><span data-stu-id="80966-199">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="80966-p125">如果加载项在收到消息后需要打开任务窗格的其他页面，可以使用 `window.location.replace` 方法（或 `window.location.href`）作为处理程序的最后一行。示例如下：</span><span class="sxs-lookup"><span data-stu-id="80966-p125">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="80966-202">有关具有此用途的加载项示例，请参阅[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)（在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表）示例。</span><span class="sxs-lookup"><span data-stu-id="80966-202">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

#### <a name="conditional-messaging"></a><span data-ttu-id="80966-203">条件消息</span><span class="sxs-lookup"><span data-stu-id="80966-203">Conditional messaging</span></span>

<span data-ttu-id="80966-p126">由于可以从对话框发送多个 `messageParent` 调用，但在主机页中只有一个 `DialogMessageReceived` 事件处理程序，因此处理程序必须使用条件逻辑来区分不同的消息。比方说，如果对话框提示用户登录标识提供程序（如 Microsoft 帐户或 Google），则会以消息形式发送用户配置文件。如果身份验证失败，对话框会将错误消息发送到主机页，如下面的示例所示：</span><span class="sxs-lookup"><span data-stu-id="80966-p126">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages. For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft Account or Google, it sends the user's profile as a message. If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

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
> - <span data-ttu-id="80966-207">`loginSuccess` 变量通过读取标识提供程序返回的 HTTP 响应进行初始化。</span><span class="sxs-lookup"><span data-stu-id="80966-207">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="80966-p127">未显示 `getProfile` 和 `getError` 函数的实现。这两个函数均从查询参数或 HTTP 响应的正文获取数据。</span><span class="sxs-lookup"><span data-stu-id="80966-p127">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="80966-p128">根据登录是否成功，发送不同类型的匿名对象。两者都有 `messageType` 属性。不同之处在于，一个有 `profile` 属性，另一个有 `error` 属性。</span><span class="sxs-lookup"><span data-stu-id="80966-p128">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="80966-p129">主机页中的处理程序代码使用 `messageType` 属性的值设置分支，如下面的示例所示。请注意，`showUserName` 函数的用法与之前的示例相同，`showNotification` 函数在主机页的 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="80966-p129">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

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

> [!NOTE]
> <span data-ttu-id="80966-214">`showNotification` 实施未在本文提供的示例代码中显示。</span><span class="sxs-lookup"><span data-stu-id="80966-214">The `showNotification` implementation is not shown in the sample code provided by this article.</span></span> <span data-ttu-id="80966-215">有关如何在外接程序中实施此函数的示例，请参阅 [Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。</span><span class="sxs-lookup"><span data-stu-id="80966-215">For an example of how you might implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

### <a name="closing-the-dialog-box"></a><span data-ttu-id="80966-216">关闭对话框</span><span class="sxs-lookup"><span data-stu-id="80966-216">Closing the dialog box</span></span>

<span data-ttu-id="80966-p131">可以在对话框中实现对话框关闭按钮。为此，关闭按钮的单击事件处理程序应使用 `messageParent` 通知主机页，关闭按钮已获单击。示例如下：</span><span class="sxs-lookup"><span data-stu-id="80966-p131">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="80966-p132">`DialogMessageReceived` 的主机页处理程序会调用 `dialog.close`，如下面的示例所示。（请参阅之前的示例，其中展示了对话框对象的初始化方式。）</span><span class="sxs-lookup"><span data-stu-id="80966-p132">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example. (See previous examples that show how the dialog object is initialized.)</span></span>


```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="80966-p133">即使你没有自己的关闭对话框 UI，最终用户也可以通过选择右上角的 **X** 关闭对话框。此操作将触发 `DialogEventReceived` 事件。如果主机窗格需要知道此事件何时发生，应为此事件声明一个处理程序。有关详细信息，请参阅[对话框窗口中的错误和事件](#errors-and-events-in-the-dialog-window)部分。</span><span class="sxs-lookup"><span data-stu-id="80966-p133">Even when you don't have your own close dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner. This action triggers the `DialogEventReceived` event. If your host pane needs to know when this happens, it should declare a handler for this event. See the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window) for details.</span></span>

## <a name="handle-errors-and-events"></a><span data-ttu-id="80966-226">处理错误和事件</span><span class="sxs-lookup"><span data-stu-id="80966-226">Handle errors and events</span></span>

<span data-ttu-id="80966-227">代码应处理两类事件：</span><span class="sxs-lookup"><span data-stu-id="80966-227">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="80966-228">`displayDialogAsync` 调用返回的错误，因为无法创建对话框。</span><span class="sxs-lookup"><span data-stu-id="80966-228">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="80966-229">对话框窗口中的错误和其他事件。</span><span class="sxs-lookup"><span data-stu-id="80966-229">Errors, and other events, in the dialog window.</span></span>

### <a name="errors-from-displaydialogasync"></a><span data-ttu-id="80966-230">DisplayDialogAsync 返回的错误</span><span class="sxs-lookup"><span data-stu-id="80966-230">Errors from displayDialogAsync</span></span>

<span data-ttu-id="80966-231">除常规的平台和系统错误外，调用 `displayDialogAsync` 会返回以下三个特定错误。</span><span class="sxs-lookup"><span data-stu-id="80966-231">In addition to general platform and system errors, three errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="80966-232">代码编号</span><span class="sxs-lookup"><span data-stu-id="80966-232">Code number</span></span>|<span data-ttu-id="80966-233">含义</span><span class="sxs-lookup"><span data-stu-id="80966-233">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="80966-234">12004</span><span class="sxs-lookup"><span data-stu-id="80966-234">12004</span></span>|<span data-ttu-id="80966-p134">传递给 `displayDialogAsync` 的 URL 的域不受信任。此域必须与主机页的域相同（包括协议和端口号）。</span><span class="sxs-lookup"><span data-stu-id="80966-p134">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="80966-237">12005</span><span class="sxs-lookup"><span data-stu-id="80966-237">12005</span></span>|<span data-ttu-id="80966-p135">传递给 `displayDialogAsync` 的 URL 使用 HTTP 协议。需要使用 HTTPS。（在 Office 的某些版本中，返回 12005 的错误消息与返回 12004 错误消息是相同的。）</span><span class="sxs-lookup"><span data-stu-id="80966-p135">The URL passed to `displayDialogAsync` uses the HTTP protocol. HTTPS is required. (In some versions of Office, the error message returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="80966-241"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="80966-241"><span id="12007">12007</span></span></span>|<span data-ttu-id="80966-p136">已从此主机窗口打开了一个对话框。主机窗口（如任务窗格）一次只能打开一个对话框。</span><span class="sxs-lookup"><span data-stu-id="80966-p136">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|
|<span data-ttu-id="80966-244">12009</span><span class="sxs-lookup"><span data-stu-id="80966-244">12009</span></span>|<span data-ttu-id="80966-245">用户已选择忽略对话框。</span><span class="sxs-lookup"><span data-stu-id="80966-245">The user chose to ignore the dialog box.</span></span> <span data-ttu-id="80966-246">联机版本的 Office 中可能会发生此错误，用户可能会选择不允许加载项显示对话框。</span><span class="sxs-lookup"><span data-stu-id="80966-246">This error can occur in online versions of Office, where users may choose not to allow an add-in to present a dialog.</span></span>|

<span data-ttu-id="80966-247">调用 `displayDialogAsync` 时，总是将 [AsyncResult](/javascript/api/office/office.asyncresult) 对象传递给它的回调函数。</span><span class="sxs-lookup"><span data-stu-id="80966-247">When `displayDialogAsync` is called, it always passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to its callback function.</span></span> <span data-ttu-id="80966-248">如果调用成功（即对话框窗口已打开），`AsyncResult` 对象的 `value` 属性是 [Dialog](/javascript/api/office/office.dialog) 对象。</span><span class="sxs-lookup"><span data-stu-id="80966-248">When the call is successful - that is, the dialog window is opened - the `value` property of the `AsyncResult` object is a [Dialog](/javascript/api/office/office.dialog) object.</span></span> <span data-ttu-id="80966-249">有关示例，请参阅[将信息从对话框发送到宿主页](#send-information-from-the-dialog-box-to-the-host-page)部分。</span><span class="sxs-lookup"><span data-stu-id="80966-249">An example of this is in the section [Send information from the dialog box to the host page](#send-information-from-the-dialog-box-to-the-host-page).</span></span> <span data-ttu-id="80966-250">如果调用 `displayDialogAsync` 失败，不会创建窗口，`AsyncResult` 对象的 `status` 属性设置为 `Office.AsyncResultStatus.Failed`，并且会填充对象的 `error` 属性。</span><span class="sxs-lookup"><span data-stu-id="80966-250">When the call to `displayDialogAsync` fails, the window is not created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated.</span></span> <span data-ttu-id="80966-251">应始终有用于测试 `status` 并在出错时进行响应的回调。</span><span class="sxs-lookup"><span data-stu-id="80966-251">You should always have a callback that tests the `status` and responds when it's an error.</span></span> <span data-ttu-id="80966-252">有关仅报告错误消息（无论代码编号是什么）的示例，请参阅下面的代码：</span><span class="sxs-lookup"><span data-stu-id="80966-252">For an example that simply reports the error message regardless of its code number, see the following code:</span></span>

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

### <a name="errors-and-events-in-the-dialog-window"></a><span data-ttu-id="80966-253">对话框窗口中的错误和事件</span><span class="sxs-lookup"><span data-stu-id="80966-253">Errors and events in the dialog window</span></span>

<span data-ttu-id="80966-254">对话框中的三个错误和事件（具有代码编码）会在主机页中触发 `DialogEventReceived` 事件。</span><span class="sxs-lookup"><span data-stu-id="80966-254">Three errors and events, known by their code numbers, in the dialog box will trigger a `DialogEventReceived` event in the host page.</span></span>

|<span data-ttu-id="80966-255">代码编号</span><span class="sxs-lookup"><span data-stu-id="80966-255">Code number</span></span>|<span data-ttu-id="80966-256">含义</span><span class="sxs-lookup"><span data-stu-id="80966-256">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="80966-257">12002</span><span class="sxs-lookup"><span data-stu-id="80966-257">12002</span></span>|<span data-ttu-id="80966-258">下列一种含义：</span><span class="sxs-lookup"><span data-stu-id="80966-258">One of the following:</span></span><br> <span data-ttu-id="80966-259">- 传递给 `displayDialogAsync` 的 URL 没有对应的页面。</span><span class="sxs-lookup"><span data-stu-id="80966-259">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="80966-260">- 传递给 `displayDialogAsync` 的页面已加载，但对话框定向到找不到或无法加载的页面，或者已定向到使用无效语法的 URL。</span><span class="sxs-lookup"><span data-stu-id="80966-260">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was directed to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="80966-261">12003</span><span class="sxs-lookup"><span data-stu-id="80966-261">12003</span></span>|<span data-ttu-id="80966-p139">对话框定向到使用 HTTP 协议的 URL。必须使用 HTTPS。</span><span class="sxs-lookup"><span data-stu-id="80966-p139">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="80966-264">12006</span><span class="sxs-lookup"><span data-stu-id="80966-264">12006</span></span>|<span data-ttu-id="80966-265">对话框已关闭，通常是因为用户选择了 **X** 按钮。</span><span class="sxs-lookup"><span data-stu-id="80966-265">The dialog box was closed, usually because the user chooses the **X** button.</span></span>|

<span data-ttu-id="80966-p140">代码可以在调用 `displayDialogAsync` 时分配 `DialogEventReceived` 事件处理程序。下面展示了一个简单示例：</span><span class="sxs-lookup"><span data-stu-id="80966-p140">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="80966-268">有关为每个错误代码创建自定义错误消息的 `DialogEventReceived` 事件处理程序示例，请参阅下面的示例：</span><span class="sxs-lookup"><span data-stu-id="80966-268">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example:</span></span>

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

<span data-ttu-id="80966-269">有关这样处理错误的样本加载项，请参阅 [Office 加载项 Dialog API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。</span><span class="sxs-lookup"><span data-stu-id="80966-269">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>


## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="80966-270">向对话框传递信息</span><span class="sxs-lookup"><span data-stu-id="80966-270">Pass information to the dialog box</span></span>

<span data-ttu-id="80966-p141">有时，主机页需要向对话框传递信息。完成此操作的方式主要分为两种：</span><span class="sxs-lookup"><span data-stu-id="80966-p141">Sometimes the host page needs to pass information to the dialog box. You can do this in two primary ways:</span></span>

- <span data-ttu-id="80966-273">向传递给 `displayDialogAsync` 的 URL 添加查询参数。</span><span class="sxs-lookup"><span data-stu-id="80966-273">Add query parameters to the URL that is passed to `displayDialogAsync`.</span></span>
- <span data-ttu-id="80966-p142">将信息存储在主机窗口和对话框都可访问的位置。这两个窗口不共享通用会话存储，但*如果它们具有相同的域*（包括端口号，若有），则共享通用[本地存储](https://www.w3schools.com/html/html5_webstorage.asp)。</span><span class="sxs-lookup"><span data-stu-id="80966-p142">Store the information somewhere that is accessible to both the host window and dialog box. The two windows do not share a common session storage, but *if they have the same domain* (including port number, if any),  they share a common [local storage](https://www.w3schools.com/html/html5_webstorage.asp).</span></span>

### <a name="use-local-storage"></a><span data-ttu-id="80966-276">使用本地存储</span><span class="sxs-lookup"><span data-stu-id="80966-276">Use local storage</span></span>

<span data-ttu-id="80966-277">为了使用本地存储，代码会先在主机页中调用 `window.localStorage` 对象的 `setItem` 方法，然后再调用 `displayDialogAsync`，如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="80966-277">To use local storage, your code calls the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:</span></span>

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

<span data-ttu-id="80966-278">对话框窗口中的代码会在需要时读取项，如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="80966-278">Code in the dialog window reads the item when it's needed, as in the following example:</span></span>

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

### <a name="use-query-parameters"></a><span data-ttu-id="80966-279">使用查询参数</span><span class="sxs-lookup"><span data-stu-id="80966-279">Use query parameters</span></span>

<span data-ttu-id="80966-280">下面的示例展示了如何使用查询参数传递数据：</span><span class="sxs-lookup"><span data-stu-id="80966-280">The following example shows how to pass data with a query parameter:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

<span data-ttu-id="80966-281">有关使用此技术的示例，请参阅[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)（在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表）。</span><span class="sxs-lookup"><span data-stu-id="80966-281">For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).</span></span>

<span data-ttu-id="80966-282">对话框窗口中的代码可以分析 URL，并读取参数值。</span><span class="sxs-lookup"><span data-stu-id="80966-282">Code in your dialog window can parse the URL and read the parameter value.</span></span>

> [!NOTE]
> <span data-ttu-id="80966-p143">Office 会自动向传递给 `displayDialogAsync` 的 URL 添加查询参数 `_host_info`。（附加在自定义查询参数（若有）后面，不会附加到对话框导航到的任何后续 URL。）Microsoft 可能会更改此值的内容，或者将来会将其全部删除，因此代码不得读取此值。相同的值会被添加到对话框的会话存储中。同样，*代码不得对此值执行读取和写入操作*。</span><span class="sxs-lookup"><span data-stu-id="80966-p143">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.</span></span>

## <a name="use-the-dialog-apis-to-show-a-video"></a><span data-ttu-id="80966-288">使用对话框 API 显示视频</span><span class="sxs-lookup"><span data-stu-id="80966-288">Use the Dialog APIs to show a video</span></span>

<span data-ttu-id="80966-289">若要在对话框中显示视频，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="80966-289">To show a video in a dialog box:</span></span>

1.  <span data-ttu-id="80966-p144">创建内容仅有 iframe 的页面。iframe 的 `src` 属性指向联机视频。视频 URL 必须使用 HTTP**S** 协议。本文将此页面称为“video.dialogbox.html”。下面展示了标记示例：</span><span class="sxs-lookup"><span data-stu-id="80966-p144">Create a page whose only content is an iframe. The `src` attribute of the iframe points to an online video. The protocol of the video's URL must be HTTP**S**. In this article we'll call this page "video.dialogbox.html". The following is an example of the markup:</span></span>

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2.  <span data-ttu-id="80966-295">video.dialogbox.html 页面必须与主机页位于同一域中。</span><span class="sxs-lookup"><span data-stu-id="80966-295">The video.dialogbox.html page must be in the same domain as the host page.</span></span>
3.  <span data-ttu-id="80966-296">在主机页中调用 `displayDialogAsync`，打开 video.dialogbox.html。</span><span class="sxs-lookup"><span data-stu-id="80966-296">Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.</span></span>
4.  <span data-ttu-id="80966-p145">如果外接程序需要知道用户何时关闭对话框，请为 `DialogEventReceived` 事件注册处理程序，并处理 12006 事件。有关详细信息，请参阅[对话框窗口中的错误和事件](#errors-and-events-in-the-dialog-window)部分。</span><span class="sxs-lookup"><span data-stu-id="80966-p145">If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event. For details, see the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window).</span></span>

<span data-ttu-id="80966-299">有关在对话框中显示视频的示例，请参阅[视频展示位置设计模式](/office/dev/add-ins/design/first-run-experience-patterns#video-placemat)。</span><span class="sxs-lookup"><span data-stu-id="80966-299">For a sample that shows a video in a dialog box, see the [video placemat design pattern](/office/dev/add-ins/design/first-run-experience-patterns#video-placemat).</span></span>

![在加载项对话框中显示的视频的屏幕截图](../images/video-placemats-dialog-open.png)

## <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="80966-301">在身份验证流中使用对话框 API</span><span class="sxs-lookup"><span data-stu-id="80966-301">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="80966-302">请参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。</span><span class="sxs-lookup"><span data-stu-id="80966-302">See [Authenticate with the Office Dialog API](auth-with-office-dialog-api.md).</span></span>

## <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="80966-303">将 Office 对话框 API 与单页应用程序和客户端路由结合使用</span><span class="sxs-lookup"><span data-stu-id="80966-303">Using the Office Dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="80966-304">如果加载项使用客户端路由（单页应用程序 (SPA) 通常这样做），则可以选择将路由 URL 传递给 [ displayDialogAsync ](/javascript/api/office/office.ui) 方法（*不建议这样做*），而不是传递各个完整 HTML 页面的 URL。</span><span class="sxs-lookup"><span data-stu-id="80966-304">If your add-in uses client-side routing, as single-page applications typically do, you have the option to pass the URL of a route to the [displayDialogAsync](/javascript/api/office/office.ui) method, instead of the URL of a complete and separate HTML page.</span></span>

<span data-ttu-id="80966-305">对话框位于有自己执行上下文的新窗口中。</span><span class="sxs-lookup"><span data-stu-id="80966-305">The dialog box is in a new window with its own execution context.</span></span> <span data-ttu-id="80966-306">如果你传递路由，则基本页及其所有初始化和引导代码会在这个新的上下文中再次运行，且所有变量都会在对话框中设置为各自的初始值。</span><span class="sxs-lookup"><span data-stu-id="80966-306">If you pass a route, your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog window.</span></span> <span data-ttu-id="80966-307">因此，此技术将在对话窗口中加载和启动另一个应用程序实例，这将部分破坏 SPA 的作用。</span><span class="sxs-lookup"><span data-stu-id="80966-307">So this technique downloads and launches a second instance of your application in the dialog window, which partially defeats the purpose of an SPA.</span></span> <span data-ttu-id="80966-308">此外，更改对话框窗口中变量的代码不会更改任务窗格版本的相同变量。</span><span class="sxs-lookup"><span data-stu-id="80966-308">Code that changes variables in the dialog window does not change the task pane version of the same variables.</span></span> <span data-ttu-id="80966-309">同样，对话框窗口有其自己的会话存储，任务窗格中的代码无法访问此类存储。</span><span class="sxs-lookup"><span data-stu-id="80966-309">Similarly, the dialog window has its own session storage, which is not accessible from code in the task pane.</span></span>

<span data-ttu-id="80966-310">因此，如果将路由传递给 `displayDialogAsync` 方法，则你并非真正拥有 SPA；你拥有的是相同 SPA 的两个实例。</span><span class="sxs-lookup"><span data-stu-id="80966-310">So, if you passed a route to the `displayDialogAsync` method, you wouldn't really have an SPA; you'd have two instances of the same SPA.</span></span> <span data-ttu-id="80966-311">此外，任务窗格实例中的大部分代码将永远不会用于该实例中，并且对话框实例中的大部分代码也永远不会用于该实例中。</span><span class="sxs-lookup"><span data-stu-id="80966-311">Moreover, much of the code in the task pane instance would never be used in that instance and much of the code in the dialog instance would never be used in that instance.</span></span> <span data-ttu-id="80966-312">这相当于相同捆绑包中拥有两个 SPA。</span><span class="sxs-lookup"><span data-stu-id="80966-312">It would be like having two SPAs in the same bundle.</span></span> <span data-ttu-id="80966-313">如果想要在对话框中运行的代码非常复杂，则可能想要显式执行此操作；也就是说，在相同域的不同文件夹中包含两个 SPA。</span><span class="sxs-lookup"><span data-stu-id="80966-313">If the code that you want to run in the dialog is sufficiently complex, you might want to do this explicitly; that is, have two SPAs in different folders of the same domain.</span></span> <span data-ttu-id="80966-314">但是在大多数情况下，对话框中只需要一个简单的逻辑。</span><span class="sxs-lookup"><span data-stu-id="80966-314">But in most scenarios, only simple logic is needed in the dialog.</span></span> <span data-ttu-id="80966-315">在这种情况下，只需将嵌入式或引用的 JavaScript 的简单 HTML 页面托管到 SPA 域中即可显著简化你的项目。</span><span class="sxs-lookup"><span data-stu-id="80966-315">In such cases, your project will be greatly simplified by simply hosting a simple HTML page, with embedded or referenced JavaScript, in the domain of your SPA.</span></span> <span data-ttu-id="80966-316">将页面的 URL 传递给 `displayDialogAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="80966-316">Pass the URL of the page to the `displayDialogAsync` method.</span></span> <span data-ttu-id="80966-317">这可能意味着，你将偏移单页应用程序的本意；但正如上面所提到的，在使用对话框时，你并没有真正拥有单个 SPA 实例。</span><span class="sxs-lookup"><span data-stu-id="80966-317">This might mean that you are deviating from the literal idea of a single-page app; but as noted above you don't really have a single instance of an SPA anyway when you are using the dialog.</span></span>
