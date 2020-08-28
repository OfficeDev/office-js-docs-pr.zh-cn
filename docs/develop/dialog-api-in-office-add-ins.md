---
title: 在 Office 加载项中使用 Office 对话框 API
description: 了解在 Office 加载项中创建对话框的基础知识
ms.date: 08/20/2020
localization_priority: Normal
ms.openlocfilehash: 9d333c12d629232ece39bc30948318fbcafa3aa0
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292789"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a><span data-ttu-id="53647-103">在 Office 加载项中使用 Office 对话框 API</span><span class="sxs-lookup"><span data-stu-id="53647-103">Use the Office dialog API in Office Add-ins</span></span>

<span data-ttu-id="53647-104">可以在 Office 加载项中使用 [Office 对话框 API](/javascript/api/office/office.ui) 打开对话框。</span><span class="sxs-lookup"><span data-stu-id="53647-104">You can use the [Office dialog API](/javascript/api/office/office.ui) to open dialog boxes in your Office Add-in.</span></span> <span data-ttu-id="53647-105">本文提供了有关如何在 Office 加载项中使用对话框 API 的指南。</span><span class="sxs-lookup"><span data-stu-id="53647-105">This article provides guidance for using the dialog API in your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="53647-p102">若要了解对话框 API 目前的受支持情况，请参阅[对话框 API 要求集](../reference/requirement-sets/dialog-api-requirement-sets.md)。目前，Word、Excel、PowerPoint 和 Outlook 支持对话框 API。</span><span class="sxs-lookup"><span data-stu-id="53647-p102">For information about where the Dialog API is currently supported, see [Dialog API requirement sets](../reference/requirement-sets/dialog-api-requirement-sets.md). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.</span></span>

<span data-ttu-id="53647-108">对话框 API 的主要应用场景是为 Google、Facebook 或 Microsoft Graph 等资源启用身份验证。</span><span class="sxs-lookup"><span data-stu-id="53647-108">A primary scenario for the Dialog API is to enable authentication with a resource such as Google, Facebook, or Microsoft Graph.</span></span> <span data-ttu-id="53647-109">有关详细信息，请在熟悉本文*之后*，参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。</span><span class="sxs-lookup"><span data-stu-id="53647-109">For more information, see [Authenticate with the Office dialog API](auth-with-office-dialog-api.md) *after* you are familiar with this article.</span></span>

<span data-ttu-id="53647-110">不妨通过任务窗格/内容加载项/[加载项命令](../design/add-in-commands.md)打开对话框，以便执行下列操作：</span><span class="sxs-lookup"><span data-stu-id="53647-110">Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:</span></span>

- <span data-ttu-id="53647-111">显示无法直接在任务窗格中打开的登录页。</span><span class="sxs-lookup"><span data-stu-id="53647-111">Display sign in pages that cannot be opened directly in a task pane.</span></span>
- <span data-ttu-id="53647-112">为加载项中的某些任务提供更多屏幕空间，或甚至整个屏幕。</span><span class="sxs-lookup"><span data-stu-id="53647-112">Provide more screen space, or even a full screen, for some tasks in your add-in.</span></span>
- <span data-ttu-id="53647-113">托管在任务窗格中显得太小的视频。</span><span class="sxs-lookup"><span data-stu-id="53647-113">Host a video that would be too small if confined to a task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="53647-114">由于不赞成重叠 UI 元素，因此除非应用场景需要，否则请勿从任务窗格打开对话框。</span><span class="sxs-lookup"><span data-stu-id="53647-114">Because overlapping UI elements are discouraged, avoid opening a dialog box from a task pane unless your scenario requires it.</span></span> <span data-ttu-id="53647-115">考虑如何使用任务窗格区域时，请注意任务窗格中可以有选项卡。</span><span class="sxs-lookup"><span data-stu-id="53647-115">When you consider how to use the surface area of a task pane, note that task panes can be tabbed.</span></span> <span data-ttu-id="53647-116">有关示例，请参阅 [Excel 加载项 JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 样本。</span><span class="sxs-lookup"><span data-stu-id="53647-116">For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

<span data-ttu-id="53647-117">下图展示了对话框示例。</span><span class="sxs-lookup"><span data-stu-id="53647-117">The following image shows an example of a dialog box.</span></span>

![外接程序命令](../images/auth-o-dialog-open.png)

<span data-ttu-id="53647-119">请注意，对话框总是在屏幕的中心打开。</span><span class="sxs-lookup"><span data-stu-id="53647-119">Note that the dialog box always opens in the center of the screen.</span></span> <span data-ttu-id="53647-120">用户可以移动并重设对话框的大小。</span><span class="sxs-lookup"><span data-stu-id="53647-120">The user can move and resize it.</span></span> <span data-ttu-id="53647-121">窗口为 *是非*--用户可以继续与 Office 应用程序中的文档和任务窗格中的页面进行交互（如果有的话）。</span><span class="sxs-lookup"><span data-stu-id="53647-121">The window is *nonmodal*--a user can continue to interact with both the document in the Office application and with the page in the task pane, if there is one.</span></span>

## <a name="open-a-dialog-box-from-a-host-page"></a><span data-ttu-id="53647-122">从主机页面打开对话框</span><span class="sxs-lookup"><span data-stu-id="53647-122">Open a dialog box from a host page</span></span>

<span data-ttu-id="53647-123">Office JavaScript API 在 [Office.context.ui 命名空间](/javascript/api/office/office.ui)中包含一个 [Dialog](/javascript/api/office/office.dialog) 对象和两个函数。</span><span class="sxs-lookup"><span data-stu-id="53647-123">The Office JavaScript APIs include a [Dialog](/javascript/api/office/office.dialog) object and two functions in the [Office.context.ui namespace](/javascript/api/office/office.ui).</span></span>

<span data-ttu-id="53647-124">为了打开对话框，代码（通常是任务窗格中的一页）调用 [displayDialogAsync](/javascript/api/office/office.ui) 方法，并将要打开的资源 URL 传递到此方法。</span><span class="sxs-lookup"><span data-stu-id="53647-124">To open a dialog box, your code, typically a page in a task pane, calls the [displayDialogAsync](/javascript/api/office/office.ui) method and passes to it the URL of the resource that you want to open.</span></span> <span data-ttu-id="53647-125">调用方法的页面称为“主机页”。</span><span class="sxs-lookup"><span data-stu-id="53647-125">The page on which this method is called is known as the "host page".</span></span> <span data-ttu-id="53647-126">例如在任务窗格中的 index.html 页面上使用脚本调用此方法，随后 index.html 是打开此方法对话框的主机页。</span><span class="sxs-lookup"><span data-stu-id="53647-126">For example, if you call this method in script on index.html in a task pane, then index.html is the host page of the dialog box that the method opens.</span></span>

<span data-ttu-id="53647-127">对话框中打开的资源通常是页面，，但也可以是 MVC 应用中的控制器方法、路由、Web 服务方法或其他任何资源。</span><span class="sxs-lookup"><span data-stu-id="53647-127">The resource that is opened in the dialog box is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource.</span></span> <span data-ttu-id="53647-128">在本文中，“页面”或“网站”是指对话框中的资源。</span><span class="sxs-lookup"><span data-stu-id="53647-128">In this article, 'page' or 'website' refers to the resource in the dialog box.</span></span> <span data-ttu-id="53647-129">下面的代码就是一个非常简单的示例：</span><span class="sxs-lookup"><span data-stu-id="53647-129">The following code is a simple example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - <span data-ttu-id="53647-130">URL 使用 HTTP**S** 协议。</span><span class="sxs-lookup"><span data-stu-id="53647-130">The URL uses the HTTP**S** protocol.</span></span> <span data-ttu-id="53647-131">对话框中加载的所有页面都必须要遵循此要求，而不仅仅是加载的第一个页面。</span><span class="sxs-lookup"><span data-stu-id="53647-131">This is mandatory for all pages loaded in a dialog box, not just the first page loaded.</span></span>
> - <span data-ttu-id="53647-132">对话框域与宿主页的域相同，宿主页可以是任务窗格中的页面，也可以是加载项命令的[函数文件](../reference/manifest/functionfile.md)。</span><span class="sxs-lookup"><span data-stu-id="53647-132">The dialog box's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](../reference/manifest/functionfile.md) of an add-in command.</span></span> <span data-ttu-id="53647-133">这要求：传递到 `displayDialogAsync` 方法的页面、控制器方法或其他资源必须与主机页位于相同的域。</span><span class="sxs-lookup"><span data-stu-id="53647-133">This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="53647-134">对话框中打开的主机页面和资源必须具有相同的完整域。</span><span class="sxs-lookup"><span data-stu-id="53647-134">The host page and the resource that opens in the dialog box must have the same full domain.</span></span> <span data-ttu-id="53647-135">如果尝试传递 `displayDialogAsync` 加载项域的子域，则不会起作用。</span><span class="sxs-lookup"><span data-stu-id="53647-135">If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work.</span></span> <span data-ttu-id="53647-136">完整域（包括任何子域）必须匹配。</span><span class="sxs-lookup"><span data-stu-id="53647-136">The full domain, including any subdomain, must match.</span></span>

<span data-ttu-id="53647-137">加载第一个页面（或其它资源）后，用户可使用链接或其它用户界面来导航至任何使用 HTTPS 的网站（或其他资源）。</span><span class="sxs-lookup"><span data-stu-id="53647-137">After the first page (or other resource) is loaded, a user can use links or other UI to navigate to any website (or other resource) that uses HTTPS.</span></span> <span data-ttu-id="53647-138">还可以将第一个页面设计为直接重定向到另一个站点。</span><span class="sxs-lookup"><span data-stu-id="53647-138">You can also design the first page to immediately redirect to another site.</span></span>

<span data-ttu-id="53647-139">默认情况下，对话框的高度和宽度占设备屏幕的 80%。不过，也可以设置不同的百分比，只需将配置对象传递给方法即可，如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="53647-139">By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

<span data-ttu-id="53647-140">有关实现这一点的样本加载项，请参阅 [Office 加载项 Dialog API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。</span><span class="sxs-lookup"><span data-stu-id="53647-140">For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="53647-p112">将两个值均设置为 100% 可有效提供全屏体验。（有效最大值为 99.5%，窗口仍可移动和调整大小。）</span><span class="sxs-lookup"><span data-stu-id="53647-p112">Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)</span></span>

> [!NOTE]
> <span data-ttu-id="53647-p113">只能从主机窗口打开一个对话框。如果尝试再打开一个对话框，就会生成错误。比方说，如果用户从任务窗格打开一个对话框，她就无法再从任务窗格中的其他页面打开第二个对话框。不过，如果对话框是通过[加载项命令](../design/add-in-commands.md)打开，那么只要选择此命令，就会打开新 HTML 文件（但不可见）。这会新建（不可见的）主机窗口，所以每个这样的窗口都可以启动自己的对话框。有关详细信息，请参阅 [displayDialogAsync 返回的错误](dialog-handle-errors-events.md#errors-from-displaydialogasync)。</span><span class="sxs-lookup"><span data-stu-id="53647-p113">You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).</span></span>

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a><span data-ttu-id="53647-149">利用 Office 网页版中的性能选项</span><span class="sxs-lookup"><span data-stu-id="53647-149">Take advantage of a performance option in Office on the web</span></span>

<span data-ttu-id="53647-150">`displayInIframe` 属性是配置对象中另一个可以传递到 `displayDialogAsync` 的属性。</span><span class="sxs-lookup"><span data-stu-id="53647-150">The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`.</span></span> <span data-ttu-id="53647-151">如果将此属性设置为 `true`，且加载项在 Office 网页版打开的文档中运行，对话框就会以浮动 iframe（而不是独立窗口）的形式打开，从而加快对话框的打开速度。</span><span class="sxs-lookup"><span data-stu-id="53647-151">When this property is set to `true`, and the add-in is running in a document opened in Office on the web, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster.</span></span> <span data-ttu-id="53647-152">示例如下：</span><span class="sxs-lookup"><span data-stu-id="53647-152">The following is an example:</span></span>

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

<span data-ttu-id="53647-153">默认值为 `false`，与完全省略此属性时相同。</span><span class="sxs-lookup"><span data-stu-id="53647-153">The default value is `false`, which is the same as omitting the property entirely.</span></span> <span data-ttu-id="53647-154">如果加载项没有在 Office 网页版中运行，`displayInIframe` 将被忽略。</span><span class="sxs-lookup"><span data-stu-id="53647-154">If the add-in is not running in Office on the web, the `displayInIframe` is ignored.</span></span>

> [!NOTE]
> <span data-ttu-id="53647-155">如果对话框始终重定向到无法在 iframe 中打开的页面，**不**得使用 `displayInIframe: true`。</span><span class="sxs-lookup"><span data-stu-id="53647-155">You should **not** use `displayInIframe: true` if the dialog box will at any point redirect to a page that cannot be opened in an iframe.</span></span> <span data-ttu-id="53647-156">例如，不能在 iframe 中打开许多常用 web 服务（如 Google 和 Microsoft 帐户）的 "登录" 页面。</span><span class="sxs-lookup"><span data-stu-id="53647-156">For example, the sign in pages of many popular web services, such as Google and Microsoft account, cannot be opened in an iframe.</span></span>

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a><span data-ttu-id="53647-157">将信息从对话框发送到主机页</span><span class="sxs-lookup"><span data-stu-id="53647-157">Send information from the dialog box to the host page</span></span>

<span data-ttu-id="53647-158">对话框无法与任务窗格中的主机页进行通信，除非：</span><span class="sxs-lookup"><span data-stu-id="53647-158">The dialog box cannot communicate with the host page in the task pane unless:</span></span>

- <span data-ttu-id="53647-159">对话框中的当前页面与主机页在同一个域中。</span><span class="sxs-lookup"><span data-stu-id="53647-159">The current page in the dialog box is in the same domain as the host page.</span></span>
- <span data-ttu-id="53647-p117">在页面中加载 Office JavaScript API 库。 (与使用 Office JavaScript API 库的任何页面一样，页面的脚本必须为属性分配方法 `Office.initialize` ，尽管它可以是空方法。有关详细信息，请参阅 [初始化 Office 外接程序](initialize-add-in.md)。 ) </span><span class="sxs-lookup"><span data-stu-id="53647-p117">The Office JavaScript API library is loaded in the page. (Like any page that uses the Office JavaScript API library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initialize your Office Add-in](initialize-add-in.md).)</span></span>

<span data-ttu-id="53647-163">对话框中的代码使用 [messageParent](/javascript/api/office/office.ui#messageparent-message-) 函数，向主机页发送布尔值或字符串消息。</span><span class="sxs-lookup"><span data-stu-id="53647-163">Code in the dialog box uses the [messageParent](/javascript/api/office/office.ui#messageparent-message-) function to send either a Boolean value or a string message to the host page.</span></span> <span data-ttu-id="53647-164">字符串可以是单词、句子、XML blob、字符串化 JSON 或其他任何能够序列化成字符串的内容。</span><span class="sxs-lookup"><span data-stu-id="53647-164">The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string.</span></span> <span data-ttu-id="53647-165">示例如下：</span><span class="sxs-lookup"><span data-stu-id="53647-165">The following is an example:</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!IMPORTANT]
> - <span data-ttu-id="53647-166">`messageParent` 函数只能在与主机页位于同一域（包括协议和端口）的页面上调用。</span><span class="sxs-lookup"><span data-stu-id="53647-166">The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.</span></span>
> - <span data-ttu-id="53647-167">`messageParent`函数*只*是可在对话框中调用的两个 Office JS api 之一。</span><span class="sxs-lookup"><span data-stu-id="53647-167">The `messageParent` function is one of *only* two Office JS APIs that can be called in the dialog box.</span></span> 
> - <span data-ttu-id="53647-168">可以在对话框中调用的其他 JS API 为 `Office.context.requirements.isSetSupported` 。</span><span class="sxs-lookup"><span data-stu-id="53647-168">The other JS API that can be called in the dialog box is `Office.context.requirements.isSetSupported`.</span></span> <span data-ttu-id="53647-169">有关它的信息，请参阅 [指定 Office 应用程序和 API 要求](specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="53647-169">For information about it, see [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md).</span></span> <span data-ttu-id="53647-170">但是，在该对话框中，Outlook 2016 1-time purchase (中不支持此 API，即 MSI 版本) 。</span><span class="sxs-lookup"><span data-stu-id="53647-170">However, in the dialog box, this API isn't supported in Outlook 2016 one-time purchase (that is, the MSI version).</span></span>


<span data-ttu-id="53647-171">在下一个示例中，`googleProfile` 是用户 Google 配置文件的字符串化版本。</span><span class="sxs-lookup"><span data-stu-id="53647-171">In the next example, `googleProfile` is a stringified version of the user's Google profile.</span></span>

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

<span data-ttu-id="53647-p120">必须将主机页配置为接收消息。为此，可以向 `displayDialogAsync` 的原始调用添加回调参数。回调向 `DialogMessageReceived` 事件分配处理程序。示例如下：</span><span class="sxs-lookup"><span data-stu-id="53647-p120">The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:</span></span>

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
> - <span data-ttu-id="53647-176">Office 将 [AsyncResult ](/javascript/api/office/office.asyncresult) 对象传递给回叫。</span><span class="sxs-lookup"><span data-stu-id="53647-176">Office passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to the callback.</span></span> <span data-ttu-id="53647-177">表示尝试打开对话框的结果，</span><span class="sxs-lookup"><span data-stu-id="53647-177">It represents the result of the attempt to open the dialog box.</span></span> <span data-ttu-id="53647-178">不表示对话框中任何事件的结果。</span><span class="sxs-lookup"><span data-stu-id="53647-178">It does not represent the outcome of any events in the dialog box.</span></span> <span data-ttu-id="53647-179">若要详细了解此区别，请参阅[处理错误和事件](dialog-handle-errors-events.md)。</span><span class="sxs-lookup"><span data-stu-id="53647-179">For more on this distinction, see [Handle errors and events](dialog-handle-errors-events.md).</span></span>
> - <span data-ttu-id="53647-180">`asyncResult` 的 `value` 属性设置为 [Dialog](/javascript/api/office/office.dialog) 对象，此对象位于主机页（而不是对话框的执行上下文）中。</span><span class="sxs-lookup"><span data-stu-id="53647-180">The `value` property of the `asyncResult` is set to a [Dialog](/javascript/api/office/office.dialog) object, which exists in the host page, not in the dialog box's execution context.</span></span>
> - <span data-ttu-id="53647-p122">`processMessage` 是用于处理事件的函数。可以根据需要任意命名。</span><span class="sxs-lookup"><span data-stu-id="53647-p122">The `processMessage` is the function that handles the event. You can give it any name you want.</span></span>
> - <span data-ttu-id="53647-183">`dialog` 变量的声明范围比回调更广，因为 `processMessage` 中也会引用此变量。</span><span class="sxs-lookup"><span data-stu-id="53647-183">The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.</span></span>

<span data-ttu-id="53647-184">下面展示了 `DialogMessageReceived` 事件处理程序的简单示例：</span><span class="sxs-lookup"><span data-stu-id="53647-184">The following is a simple example of a handler for the `DialogMessageReceived` event:</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - <span data-ttu-id="53647-185">Office 将 `arg` 对象传递给处理程序。</span><span class="sxs-lookup"><span data-stu-id="53647-185">Office passes the `arg` object to the handler.</span></span> <span data-ttu-id="53647-186">它的 `message` 属性是对话框中的 `messageParent` 调用发送的布尔值或字符串。</span><span class="sxs-lookup"><span data-stu-id="53647-186">Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog box.</span></span> <span data-ttu-id="53647-187">在此示例中，它是用户配置文件从 Microsoft 帐户或 Google 等服务的字符串化表示形式，因此它将被反序列化为使用的对象 `JSON.parse` 。</span><span class="sxs-lookup"><span data-stu-id="53647-187">In this example, it is a stringified representation of a user's profile from a service such as Microsoft account or Google, so it is deserialized back to an object with `JSON.parse`.</span></span>
> - <span data-ttu-id="53647-p124">未显示 `showUserName` 实现。它可能在任务窗格上显示定制的欢迎消息。</span><span class="sxs-lookup"><span data-stu-id="53647-p124">The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.</span></span>

<span data-ttu-id="53647-190">在用户完成与对话框的交互后，消息处理程序应关闭对话框，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="53647-190">When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.</span></span>

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
> - <span data-ttu-id="53647-191">`dialog` 对象必须是 `displayDialogAsync` 调用返回的对象。</span><span class="sxs-lookup"><span data-stu-id="53647-191">The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.</span></span>
> - <span data-ttu-id="53647-192">`dialog.close` 调用指示 Office 立即关闭对话框。</span><span class="sxs-lookup"><span data-stu-id="53647-192">The call of `dialog.close` tells Office to immediately close the dialog box.</span></span>

<span data-ttu-id="53647-193">有关使用这些技术的示例加载项，请参阅 [Office 加载项对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。</span><span class="sxs-lookup"><span data-stu-id="53647-193">For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

<span data-ttu-id="53647-p125">如果加载项在收到消息后需要打开任务窗格的其他页面，可以使用 `window.location.replace` 方法（或 `window.location.href`）作为处理程序的最后一行。示例如下：</span><span class="sxs-lookup"><span data-stu-id="53647-p125">If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:</span></span>

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

<span data-ttu-id="53647-196">有关具有此用途的加载项示例，请参阅[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)（在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表）示例。</span><span class="sxs-lookup"><span data-stu-id="53647-196">For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.</span></span>

### <a name="conditional-messaging"></a><span data-ttu-id="53647-197">条件消息</span><span class="sxs-lookup"><span data-stu-id="53647-197">Conditional messaging</span></span>

<span data-ttu-id="53647-198">由于可以从对话框发送多个 `messageParent` 调用，但在主机页中只有一个 `DialogMessageReceived` 事件处理程序，因此处理程序必须使用条件逻辑来区分不同的消息。</span><span class="sxs-lookup"><span data-stu-id="53647-198">Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="53647-199">例如，如果对话框提示用户登录到标识提供程序（如 Microsoft 帐户或 Google），则它会将用户的配置文件作为邮件发送。</span><span class="sxs-lookup"><span data-stu-id="53647-199">For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft account or Google, it sends the user's profile as a message.</span></span> <span data-ttu-id="53647-200">如果身份验证失败，对话框会将错误消息发送到主机页，如下面的示例所示：</span><span class="sxs-lookup"><span data-stu-id="53647-200">If authentication fails, the dialog box sends error information to the host page, as in the following example:</span></span>

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
> - <span data-ttu-id="53647-201">`loginSuccess` 变量通过读取标识提供程序返回的 HTTP 响应进行初始化。</span><span class="sxs-lookup"><span data-stu-id="53647-201">The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.</span></span>
> - <span data-ttu-id="53647-p127">未显示 `getProfile` 和 `getError` 函数的实现。这两个函数均从查询参数或 HTTP 响应的正文获取数据。</span><span class="sxs-lookup"><span data-stu-id="53647-p127">The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.</span></span>
> - <span data-ttu-id="53647-p128">根据登录是否成功，发送不同类型的匿名对象。两者都有 `messageType` 属性。不同之处在于，一个有 `profile` 属性，另一个有 `error` 属性。</span><span class="sxs-lookup"><span data-stu-id="53647-p128">Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.</span></span>

<span data-ttu-id="53647-p129">主机页中的处理程序代码使用 `messageType` 属性的值设置分支，如下面的示例所示。请注意，`showUserName` 函数的用法与之前的示例相同，`showNotification` 函数在主机页的 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="53647-p129">The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.</span></span>

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
> <span data-ttu-id="53647-208">`showNotification` 实施未在本文提供的示例代码中显示。</span><span class="sxs-lookup"><span data-stu-id="53647-208">The `showNotification` implementation is not shown in the sample code provided by this article.</span></span> <span data-ttu-id="53647-209">有关如何在外接程序中实施此函数的示例，请参阅 [Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。</span><span class="sxs-lookup"><span data-stu-id="53647-209">For an example of how you might implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>

## <a name="pass-information-to-the-dialog-box"></a><span data-ttu-id="53647-210">向对话框传递信息</span><span class="sxs-lookup"><span data-stu-id="53647-210">Pass information to the dialog box</span></span>

<span data-ttu-id="53647-211">您的外接程序可以使用[messageChild](/javascript/api/office/office.dialog#messagechild-message-)将邮件从[主机页面](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)发送到对话框。</span><span class="sxs-lookup"><span data-stu-id="53647-211">Your add-in can send messages from the [host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) to a dialog box using [Dialog.messageChild](/javascript/api/office/office.dialog#messagechild-message-).</span></span>

> [!NOTE]
> <span data-ttu-id="53647-212">仅 Excel、PowerPoint 和 Word 支持这些对话框 Api。</span><span class="sxs-lookup"><span data-stu-id="53647-212">These dialog APIs are supported in only Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="53647-213">对 Outlook 的支持正在开发中。</span><span class="sxs-lookup"><span data-stu-id="53647-213">Support for Outlook is under development.</span></span>

### <a name="use-messagechild-from-the-host-page"></a><span data-ttu-id="53647-214">`messageChild()`从主机页使用</span><span class="sxs-lookup"><span data-stu-id="53647-214">Use `messageChild()` from the host page</span></span>

<span data-ttu-id="53647-215">调用 Office 对话框 API 打开对话框时，将返回 [dialog](/javascript/api/office/office.dialog) 对象。</span><span class="sxs-lookup"><span data-stu-id="53647-215">When you call the Office dialog API to open a dialog box, a [Dialog](/javascript/api/office/office.dialog) object is returned.</span></span> <span data-ttu-id="53647-216">应将其分配给具有大于 [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) 方法的作用域的变量，因为该对象将被其他方法引用。</span><span class="sxs-lookup"><span data-stu-id="53647-216">It should be assigned to a variable that has greater scope than the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) method because the object will be referenced by other methods.</span></span> <span data-ttu-id="53647-217">示例如下：</span><span class="sxs-lookup"><span data-stu-id="53647-217">The following is an example:</span></span>

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

<span data-ttu-id="53647-218">此 `Dialog` 对象具有一个 [messageChild](/javascript/api/office/office.dialog#messagechild-message-) 方法，该方法将任何字符串（包括字符串化数据）发送到对话框。</span><span class="sxs-lookup"><span data-stu-id="53647-218">This `Dialog` object has a [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method that sends any string, including stringified data, to the dialog box.</span></span> <span data-ttu-id="53647-219">这 `DialogParentMessageReceived` 将在对话框中引发事件。</span><span class="sxs-lookup"><span data-stu-id="53647-219">This raises a `DialogParentMessageReceived` event in the dialog box.</span></span> <span data-ttu-id="53647-220">您的代码应处理此事件，如下一节中所示。</span><span class="sxs-lookup"><span data-stu-id="53647-220">Your code should handle this event, as shown in the next section.</span></span>

<span data-ttu-id="53647-221">假设对话框的 UI 与当前活动的工作表相关，并且该工作表相对于其他工作表的位置。</span><span class="sxs-lookup"><span data-stu-id="53647-221">Consider a scenario in which the UI of the dialog is related to the currently active worksheet and that worksheet's position relative to the other worksheets.</span></span> <span data-ttu-id="53647-222">在下面的示例中， `sheetPropertiesChanged` 将 Excel 工作表属性发送到对话框。</span><span class="sxs-lookup"><span data-stu-id="53647-222">In the following example, `sheetPropertiesChanged` sends Excel worksheet properties to the dialog box.</span></span> <span data-ttu-id="53647-223">在这种情况下，当前工作表名为 "我的工作表"，并且它是工作簿中的第二个工作表。</span><span class="sxs-lookup"><span data-stu-id="53647-223">In this case, the current worksheet is named "My Sheet" and it's the second sheet in the workbook.</span></span> <span data-ttu-id="53647-224">数据封装在对象和字符串化中，以便可以将其传递给 `messageChild` 。</span><span class="sxs-lookup"><span data-stu-id="53647-224">The data is encapsulated in an object and stringified so that it can be passed to `messageChild`.</span></span>

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a><span data-ttu-id="53647-225">在对话框中处理 DialogParentMessageReceived</span><span class="sxs-lookup"><span data-stu-id="53647-225">Handle DialogParentMessageReceived in the dialog box</span></span>

<span data-ttu-id="53647-226">在对话框的 JavaScript 中， `DialogParentMessageReceived` 使用 [addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) 方法为事件注册处理程序。</span><span class="sxs-lookup"><span data-stu-id="53647-226">In the dialog box's JavaScript, register a handler for the `DialogParentMessageReceived` event with the [UI.addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) method.</span></span> <span data-ttu-id="53647-227">通常在 [onReady 或 Office.initialize 方法](initialize-add-in.md)中执行此操作，如下所示。</span><span class="sxs-lookup"><span data-stu-id="53647-227">This is typically done in the [Office.onReady or Office.initialize methods](initialize-add-in.md), as shown in the following.</span></span> <span data-ttu-id="53647-228"> (更强健的示例如下所示。 ) </span><span class="sxs-lookup"><span data-stu-id="53647-228">(A more robust example is below.)</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

<span data-ttu-id="53647-229">然后，定义该 `onMessageFromParent` 处理程序。</span><span class="sxs-lookup"><span data-stu-id="53647-229">Then, define the `onMessageFromParent` handler.</span></span> <span data-ttu-id="53647-230">下面的代码将继续上一节中的示例。</span><span class="sxs-lookup"><span data-stu-id="53647-230">The following code continues the example from the preceding section.</span></span> <span data-ttu-id="53647-231">请注意，Office 会将参数传递给处理程序，并确保 `message` argument 对象的属性包含主机页中的字符串。</span><span class="sxs-lookup"><span data-stu-id="53647-231">Note that Office passes an argument to the handler and that the `message` property of the argument object contains the string from the host page.</span></span> <span data-ttu-id="53647-232">在此示例中，邮件被 reconverted 到对象，jQuery 用于将对话框的顶部标题设置为与新工作表名称相匹配。</span><span class="sxs-lookup"><span data-stu-id="53647-232">In this example, the message is reconverted to an object and jQuery is used to set the top heading of the dialog to match the new worksheet name.</span></span>

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

<span data-ttu-id="53647-233">最佳做法是验证是否正确注册了处理程序。</span><span class="sxs-lookup"><span data-stu-id="53647-233">It is a best practice to verify that your handler is properly registered.</span></span> <span data-ttu-id="53647-234">为此，可以将回调传递给 `addHandlerAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="53647-234">You can do this by passing a callback to the `addHandlerAsync` method.</span></span> <span data-ttu-id="53647-235">注册处理程序的尝试完成时，将运行此过程。</span><span class="sxs-lookup"><span data-stu-id="53647-235">This runs when the attempt to register the handler completes.</span></span> <span data-ttu-id="53647-236">如果未成功注册处理程序，请使用该处理程序记录或显示错误。</span><span class="sxs-lookup"><span data-stu-id="53647-236">Use the handler to log or show an error if the handler was not successfully registered.</span></span> <span data-ttu-id="53647-237">示例如下。</span><span class="sxs-lookup"><span data-stu-id="53647-237">The following is an example.</span></span> <span data-ttu-id="53647-238">请注意，这 `reportError` 是未在此处定义的函数，它会记录或显示错误。</span><span class="sxs-lookup"><span data-stu-id="53647-238">Note that `reportError` is a function, not defined here, that logs or displays the error.</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a><span data-ttu-id="53647-239">"将父页的条件消息传递到" 对话框</span><span class="sxs-lookup"><span data-stu-id="53647-239">Conditional messaging from parent page to dialog box</span></span>

<span data-ttu-id="53647-240">由于可以 `messageChild` 从主机页进行多次调用，但在该事件的对话框中只有一个处理程序 `DialogParentMessageReceived` ，因此处理程序必须使用条件逻辑来区分不同的消息。</span><span class="sxs-lookup"><span data-stu-id="53647-240">Because you can make multiple `messageChild` calls from the host page, but you have only one handler in the dialog box for the `DialogParentMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="53647-241">您可以按照与 [条件消息](#conditional-messaging)中所述的方式将消息发送到主机页时，精确地与构造条件消息传递的方式完全并行。</span><span class="sxs-lookup"><span data-stu-id="53647-241">You can do this in a way that is precisely parallel to how you would structure conditional messaging when the dialog box is sending a message to the host page as described in [Conditional messaging](#conditional-messaging).</span></span>

> [!NOTE]
> <span data-ttu-id="53647-242">在某些情况下， `messageChild` 可能不支持作为 [DialogApi 1.2 要求集](../reference/requirement-sets/dialog-api-requirement-sets.md)的一部分的 API。</span><span class="sxs-lookup"><span data-stu-id="53647-242">In some situations, the `messageChild` API, which is a part of the [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md),  may not be supported.</span></span> <span data-ttu-id="53647-243">以 [其他方式将邮件从其主机页传递到对话框，以其他方式](parent-to-dialog.md)对 "父对话" 对话消息进行描述。</span><span class="sxs-lookup"><span data-stu-id="53647-243">Some alternative ways for parent-to-dialog-box messaging are described in [Alternative ways of passing messages to a dialog box from its host page](parent-to-dialog.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="53647-244">不能在外接程序清单的部分中指定 [DialogApi 1.2 要求集](../reference/requirement-sets/dialog-api-requirement-sets.md) `<Requirements>` 。</span><span class="sxs-lookup"><span data-stu-id="53647-244">The [DialogApi 1.2 requirement set](../reference/requirement-sets/dialog-api-requirement-sets.md) cannot be specified in the `<Requirements>` section of an add-in manifest.</span></span> <span data-ttu-id="53647-245">您必须在运行时使用 [isSetSupported](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) 方法检查是否支持 DialogApi 1.2。</span><span class="sxs-lookup"><span data-stu-id="53647-245">You will have to check for support for DialogApi 1.2 at runtime using the [isSetSupported](specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code) method.</span></span> <span data-ttu-id="53647-246">对清单要求的支持正在开发中。</span><span class="sxs-lookup"><span data-stu-id="53647-246">Support for manifest requirements is under development.</span></span>

## <a name="closing-the-dialog-box"></a><span data-ttu-id="53647-247">关闭对话框</span><span class="sxs-lookup"><span data-stu-id="53647-247">Closing the dialog box</span></span>

<span data-ttu-id="53647-p141">可以在对话框中实现对话框关闭按钮。为此，关闭按钮的单击事件处理程序应使用 `messageParent` 通知主机页，关闭按钮已获单击。示例如下：</span><span class="sxs-lookup"><span data-stu-id="53647-p141">You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:</span></span>

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

<span data-ttu-id="53647-251">`DialogMessageReceived` 的主机页处理程序将调用 `dialog.close`，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="53647-251">The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example.</span></span> <span data-ttu-id="53647-252">（请参阅前面的示例，其中展示了 `dialog` 对象的初始化方式。）</span><span class="sxs-lookup"><span data-stu-id="53647-252">(See previous examples that show how the `dialog` object is initialized.)</span></span>

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

<span data-ttu-id="53647-253">即使你没有自己的关闭对话框 UI，最终用户也可以通过选择右上角的 **X** 关闭对话框。</span><span class="sxs-lookup"><span data-stu-id="53647-253">Even when you don't have your own close-dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner.</span></span> <span data-ttu-id="53647-254">此操作将触发 `DialogEventReceived` 事件。</span><span class="sxs-lookup"><span data-stu-id="53647-254">This action triggers the `DialogEventReceived` event.</span></span> <span data-ttu-id="53647-255">如果主机窗格需要知道此事件何时发生，应为此事件声明一个处理程序。</span><span class="sxs-lookup"><span data-stu-id="53647-255">If your host pane needs to know when this happens, it should declare a handler for this event.</span></span> <span data-ttu-id="53647-256">有关详细信息，请参阅[对话框中的错误和事件](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box)部分。</span><span class="sxs-lookup"><span data-stu-id="53647-256">See the section [Errors and events in the dialog box](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box) for details.</span></span>

## <a name="advanced-topics-and-special-scenarios"></a><span data-ttu-id="53647-257">高级主题和特殊情景</span><span class="sxs-lookup"><span data-stu-id="53647-257">Advanced topics and special scenarios</span></span>

### <a name="use-the-dialog-api-to-show-a-video"></a><span data-ttu-id="53647-258">使用对话框 API 显示视频</span><span class="sxs-lookup"><span data-stu-id="53647-258">Use the Dialog API to show a video</span></span>

<span data-ttu-id="53647-259">参见“[使用 Office 对话框显示视频](dialog-video.md)”。</span><span class="sxs-lookup"><span data-stu-id="53647-259">See [Use the Office dialog box to show a video](dialog-video.md).</span></span>

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a><span data-ttu-id="53647-260">在身份验证流中使用对话框 API</span><span class="sxs-lookup"><span data-stu-id="53647-260">Use the Dialog APIs in an authentication flow</span></span>

<span data-ttu-id="53647-261">请参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。</span><span class="sxs-lookup"><span data-stu-id="53647-261">See [Authenticate with the Office dialog API](auth-with-office-dialog-api.md).</span></span>

### <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a><span data-ttu-id="53647-262">将 Office 对话框 API 与单页应用程序和客户端路由结合使用</span><span class="sxs-lookup"><span data-stu-id="53647-262">Using the Office dialog API with single-page applications and client-side routing</span></span>

<span data-ttu-id="53647-263">使用 Office 对话框 API 时，需要小心处理 SPA 和客户端路由。</span><span class="sxs-lookup"><span data-stu-id="53647-263">SPAs and client-side routing need to be handled with care when you are using the Office dialog API.</span></span> <span data-ttu-id="53647-264">请参阅“[在 SPA 中使用 Office 对话框 API 的最佳做法](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa)”。</span><span class="sxs-lookup"><span data-stu-id="53647-264">Please see [Best practices for using the Office dialog API in an SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).</span></span>

### <a name="error-and-event-handling"></a><span data-ttu-id="53647-265">错误和事件处理</span><span class="sxs-lookup"><span data-stu-id="53647-265">Error and event handling</span></span>

<span data-ttu-id="53647-266">参见“[处理 Office 对话框中的错误和事件](dialog-handle-errors-events.md)。</span><span class="sxs-lookup"><span data-stu-id="53647-266">See [Handling errors and events in the Office dialog box](dialog-handle-errors-events.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="53647-267">后续步骤</span><span class="sxs-lookup"><span data-stu-id="53647-267">Next steps</span></span>

<span data-ttu-id="53647-268">在“[Office 对话框 API 最佳做法和规则](dialog-best-practices.md)”中了解 Office 对话框 API 的陷阱和最佳做法。</span><span class="sxs-lookup"><span data-stu-id="53647-268">Learn about gotchas and best practices for the Office dialog API in [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>
