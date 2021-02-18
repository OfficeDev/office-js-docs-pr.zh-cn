---
title: Office 对话框 API 最佳做法和规则
description: '提供适用于 Office 对话框 API 的规则和最佳做法，例如适用于 SPA 应用程序的单页 (最佳实践) '
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 4359d116e9720255278c5b3f543b135013c7e76c
ms.sourcegitcommit: 7cd501d0fdbbd4636bd08647b638dd5ca4c7c630
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50282980"
---
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a><span data-ttu-id="4e34b-103">Office 对话框 API 最佳做法和规则</span><span class="sxs-lookup"><span data-stu-id="4e34b-103">Best practices and rules for the Office dialog API</span></span>

<span data-ttu-id="4e34b-104">本文提供 Office 对话框 API 的规则、链链和最佳做法，包括设计对话框 UI 的最佳实践，以及将 API 与在单页应用程序 (SPA) </span><span class="sxs-lookup"><span data-stu-id="4e34b-104">This article provides rules, gotchas, and best practices for the Office dialog API, including best practices for designing the UI of a dialog and using the API with in a single-page application (SPA)</span></span>

> [!NOTE]
> <span data-ttu-id="4e34b-105">本文假定你熟悉使用 Office 对话框 API 的基础知识，如在 Office 外接程序中使用 [Office](dialog-api-in-office-add-ins.md)对话框 API 中所述。</span><span class="sxs-lookup"><span data-stu-id="4e34b-105">This article presupposes that you are familiar with the basics of using the Office dialog API as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>
> 
> <span data-ttu-id="4e34b-106">另请参阅 [使用 Office 对话框处理错误和事件](dialog-handle-errors-events.md)。</span><span class="sxs-lookup"><span data-stu-id="4e34b-106">See also [Handling errors and events with the Office dialog box](dialog-handle-errors-events.md).</span></span>

## <a name="rules-and-gotchas"></a><span data-ttu-id="4e34b-107">规则和陷阱</span><span class="sxs-lookup"><span data-stu-id="4e34b-107">Rules and gotchas</span></span>

- <span data-ttu-id="4e34b-108">对话框只能导航到 HTTPS URL，不能导航到 HTTP。</span><span class="sxs-lookup"><span data-stu-id="4e34b-108">The dialog box can only navigate to HTTPS URLs, not HTTP.</span></span>
- <span data-ttu-id="4e34b-109">传递给 [displayDialogAsync](/javascript/api/office/office.ui) 方法的 URL 必须与加载项本身完全相同的域中。</span><span class="sxs-lookup"><span data-stu-id="4e34b-109">The URL passed to the [displayDialogAsync](/javascript/api/office/office.ui) method must be in the exact same domain as the add-in itself.</span></span> <span data-ttu-id="4e34b-110">它不能是子域。</span><span class="sxs-lookup"><span data-stu-id="4e34b-110">It cannot be a subdomain.</span></span> <span data-ttu-id="4e34b-111">但是，传递给它的页面可以重定向到另一个域中的页面。</span><span class="sxs-lookup"><span data-stu-id="4e34b-111">But the page that is passed to it can redirect to a page in another domain.</span></span>
- <span data-ttu-id="4e34b-112">主机窗口可以是任务窗格或加载项命令的无 UI 函数文件[](../reference/manifest/functionfile.md)，一次只能打开一个对话框。</span><span class="sxs-lookup"><span data-stu-id="4e34b-112">A host window, which can be a task pane or the UI-less [function file](../reference/manifest/functionfile.md) of an add-in command, can have only one dialog box open at a time.</span></span>
- <span data-ttu-id="4e34b-113">对话框中只能调用两个 Office API：</span><span class="sxs-lookup"><span data-stu-id="4e34b-113">Only two Office APIs can be called in the dialog box:</span></span>
  - <span data-ttu-id="4e34b-114">[messageParent](/javascript/api/office/office.ui#messageparent-message-)函数。</span><span class="sxs-lookup"><span data-stu-id="4e34b-114">The [messageParent](/javascript/api/office/office.ui#messageparent-message-) function.</span></span>
  - <span data-ttu-id="4e34b-115">`Office.context.requirements.isSetSupported` (有关详细信息，请参阅指定 [Office 应用程序和 API](specify-office-hosts-and-api-requirements.md)要求 .) </span><span class="sxs-lookup"><span data-stu-id="4e34b-115">`Office.context.requirements.isSetSupported` (For more information, see [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md).)</span></span>
- <span data-ttu-id="4e34b-116">[messageParent](/javascript/api/office/office.ui#messageparent-message-)函数只能从与外接程序本身完全相同的域中的页面调用。</span><span class="sxs-lookup"><span data-stu-id="4e34b-116">The [messageParent](/javascript/api/office/office.ui#messageparent-message-) function can only be called from a page in the exact same domain as the add-in itself.</span></span>

## <a name="best-practices"></a><span data-ttu-id="4e34b-117">最佳做法</span><span class="sxs-lookup"><span data-stu-id="4e34b-117">Best practices</span></span>

### <a name="avoid-overusing-dialog-boxes"></a><span data-ttu-id="4e34b-118">避免过度使用对话框</span><span class="sxs-lookup"><span data-stu-id="4e34b-118">Avoid overusing dialog boxes</span></span>

<span data-ttu-id="4e34b-119">由于不赞成重叠 UI 元素，因此除非应用场景需要，否则请勿从任务窗格打开对话框。</span><span class="sxs-lookup"><span data-stu-id="4e34b-119">Because overlapping UI elements are discouraged, avoid opening a dialog box from a task pane unless your scenario requires it.</span></span> <span data-ttu-id="4e34b-120">考虑如何使用任务窗格区域时，请注意任务窗格中可以有选项卡。</span><span class="sxs-lookup"><span data-stu-id="4e34b-120">When you consider how to use the surface area of a task pane, note that task panes can be tabbed.</span></span> <span data-ttu-id="4e34b-121">有关选项卡式任务窗格的示例，请参阅 [Excel 外接程序 JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 示例。</span><span class="sxs-lookup"><span data-stu-id="4e34b-121">For an example of a tabbed task pane, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.</span></span>

### <a name="designing-a-dialog-box-ui"></a><span data-ttu-id="4e34b-122">设计对话框 UI</span><span class="sxs-lookup"><span data-stu-id="4e34b-122">Designing a dialog box UI</span></span>

<span data-ttu-id="4e34b-123">有关对话框设计中的最佳方案，请参阅 Office [加载项中的对话框](../design/dialog-boxes.md)。</span><span class="sxs-lookup"><span data-stu-id="4e34b-123">For best practices in dialog box design, see [Dialog boxes in Office Add-ins](../design/dialog-boxes.md).</span></span>

### <a name="handling-pop-up-blockers-with-office-on-the-web"></a><span data-ttu-id="4e34b-124">使用 Office 网页版处理弹出窗口阻止程序</span><span class="sxs-lookup"><span data-stu-id="4e34b-124">Handling pop-up blockers with Office on the web</span></span>

<span data-ttu-id="4e34b-125">尝试在 Web 上使用 Office 时显示对话框可能会导致浏览器的弹出窗口阻止程序阻止对话框。</span><span class="sxs-lookup"><span data-stu-id="4e34b-125">Attempting to display a dialog box while using Office on the web may cause the browser's pop-up blocker to block the dialog box.</span></span> <span data-ttu-id="4e34b-126">Office 网页 Office 具有一项功能，该功能可使外接程序的对话框成为浏览器弹出窗口阻止程序例外。</span><span class="sxs-lookup"><span data-stu-id="4e34b-126">Office on the web has a feature that enables your add-in's dialog boxes to be an exception to the browser's pop-up blocker.</span></span> <span data-ttu-id="4e34b-127">当代码调用该方法 `displayDialogAsync` 时，Office 网页发布将打开类似于以下内容的提示。</span><span class="sxs-lookup"><span data-stu-id="4e34b-127">When your code calls the `displayDialogAsync` method, then Office on the web will open a prompt similar to the following.</span></span>

![Screenshot showing the prompt with a brief description and Allow and Ignore buttons that an add-in can generate to avoid in-browser pop-up blockers](../images/dialog-prompt-before-open.png)

<span data-ttu-id="4e34b-129">如果用户选择"允许 **"，** 将打开 Office 对话框。</span><span class="sxs-lookup"><span data-stu-id="4e34b-129">If the user chooses **Allow**, the Office dialog box opens.</span></span> <span data-ttu-id="4e34b-130">如果用户选择"忽略 **"，** 则提示将关闭，并且 Office 对话框不会打开。</span><span class="sxs-lookup"><span data-stu-id="4e34b-130">If the user chooses **Ignore**, the prompt closes and the Office dialog box does not open.</span></span> <span data-ttu-id="4e34b-131">相反， `displayDialogAsync` 此方法返回错误 12009。</span><span class="sxs-lookup"><span data-stu-id="4e34b-131">Instead, the `displayDialogAsync` method returns error 12009.</span></span> <span data-ttu-id="4e34b-132">代码应捕获此错误，并提供不需要对话框的备用体验，或向用户显示一条消息，提示加载项要求他们允许对话框。</span><span class="sxs-lookup"><span data-stu-id="4e34b-132">Your code should catch this error and either provide an alternate experience that does not require a dialog, or display a message to the user advising that the add-in requires them to allow the dialog.</span></span> <span data-ttu-id="4e34b-133"> (有关 12009 的详细信息，请参阅 [displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).) </span><span class="sxs-lookup"><span data-stu-id="4e34b-133">(For more about 12009, see [Errors from displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).)</span></span>

<span data-ttu-id="4e34b-134">如果出于任何原因要关闭此功能，则你的代码必须选择退出。它使用传递给该方法的 [DialogOptions](/javascript/api/office/office.dialogoptions) 对象进行 `displayDialogAsync` 此请求。</span><span class="sxs-lookup"><span data-stu-id="4e34b-134">If, for any reason, you want to turn off this feature, then your code must opt out. It makes this request with the [DialogOptions](/javascript/api/office/office.dialogoptions) object that is passed to the `displayDialogAsync` method.</span></span> <span data-ttu-id="4e34b-135">具体来说，对象应包括 `promptBeforeOpen: false` 。</span><span class="sxs-lookup"><span data-stu-id="4e34b-135">Specifically, the object should include `promptBeforeOpen: false`.</span></span> <span data-ttu-id="4e34b-136">当此选项设置为 false 时，Web 上的 Office 不会提示用户允许外接程序打开对话框，并且 Office 对话框将不会打开。</span><span class="sxs-lookup"><span data-stu-id="4e34b-136">When this option is set to false, Office on the web will not prompt the user to allow the add-in open a dialog, and the Office dialog will not open.</span></span>

### <a name="do-not-use-the-_host_info-value"></a><span data-ttu-id="4e34b-137">请勿使用 \_ 主机 \_ 信息值</span><span class="sxs-lookup"><span data-stu-id="4e34b-137">Do not use the \_host\_info value</span></span>

<span data-ttu-id="4e34b-138">Office 会自动向传递给 `_host_info` 的 URL 添加查询参数 `displayDialogAsync`。</span><span class="sxs-lookup"><span data-stu-id="4e34b-138">Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`.</span></span> <span data-ttu-id="4e34b-139">它将追加到自定义查询参数（如果有）之后。</span><span class="sxs-lookup"><span data-stu-id="4e34b-139">It is appended after your custom query parameters, if any.</span></span> <span data-ttu-id="4e34b-140">它未追加到对话框导航到的任何后续 URL。</span><span class="sxs-lookup"><span data-stu-id="4e34b-140">It is not appended to any subsequent URLs that the dialog box navigates to.</span></span> <span data-ttu-id="4e34b-141">Microsoft 可能会更改此值的内容或完全删除它，因此代码不应读取它。</span><span class="sxs-lookup"><span data-stu-id="4e34b-141">Microsoft may change the content of this value, or remove it entirely, so your code should not read it.</span></span> <span data-ttu-id="4e34b-142">相同的值将添加到对话框的会话存储 (，即 [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) 属性) 。</span><span class="sxs-lookup"><span data-stu-id="4e34b-142">The same value is added to the dialog box's session storage (that is, the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property).</span></span> <span data-ttu-id="4e34b-143">同样，*代码不得对此值执行读取和写入操作*。</span><span class="sxs-lookup"><span data-stu-id="4e34b-143">Again, *your code should neither read nor write to this value*.</span></span>

### <a name="opening-another-dialog-immediately-after-closing-one"></a><span data-ttu-id="4e34b-144">关闭另一个对话框后立即打开另一个对话框</span><span class="sxs-lookup"><span data-stu-id="4e34b-144">Opening another dialog immediately after closing one</span></span>

<span data-ttu-id="4e34b-145">不能从给定的主机页打开多个对话框，因此代码应在打开的对话框中调用 [Dialog.close，](/javascript/api/office/office.dialog#close__) 然后再调用它以 `displayDialogAsync` 打开另一个对话框。</span><span class="sxs-lookup"><span data-stu-id="4e34b-145">You can't have more than one dialog open from a given host page, so your code should call [Dialog.close](/javascript/api/office/office.dialog#close__) on an open dialog before it calls `displayDialogAsync` to open another dialog.</span></span> <span data-ttu-id="4e34b-146">该方法 `close` 是异步的。</span><span class="sxs-lookup"><span data-stu-id="4e34b-146">The `close` method is asynchronous.</span></span> <span data-ttu-id="4e34b-147">因此，如果在调用后立即调用，则当 Office 尝试打开第二个对话框时，第一个对话框 `displayDialogAsync` `close` 可能尚未完全关闭。</span><span class="sxs-lookup"><span data-stu-id="4e34b-147">For this reason, if you call `displayDialogAsync` immediately after a call of `close`, the first dialog may not have completely closed when Office attempts to open the second.</span></span> <span data-ttu-id="4e34b-148">如果发生这种情况，Office 将返回 [12007](dialog-handle-errors-events.md#12007) 错误："操作失败，因为此外接程序已有活动对话框。"</span><span class="sxs-lookup"><span data-stu-id="4e34b-148">If that happens, Office will return a [12007](dialog-handle-errors-events.md#12007) error: "The operation failed because this add-in already has an active dialog."</span></span>

<span data-ttu-id="4e34b-149">该方法不接受回调参数，并且不会返回 Promise 对象，因此无法使用关键字或 `close` `await` 方法等待 `then` 该对象。</span><span class="sxs-lookup"><span data-stu-id="4e34b-149">The `close` method doesn't accept a callback parameter, and it doesn't return a Promise object so it cannot be awaited with either the `await` keyword or with a `then` method.</span></span> <span data-ttu-id="4e34b-150">出于此原因，建议在关闭对话框后立即打开新对话框时采用以下技术：封装代码以在方法中打开新对话框，并设计方法以递归方式调用自身（如果调用 `displayDialogAsync` 返回 `12007` ）。</span><span class="sxs-lookup"><span data-stu-id="4e34b-150">For this reason, we suggest the following technique when you need to open a new dialog immediately after closing a dialog: encapsulate the code to open the new dialog in a method and design the method to recursively call itself if the call of `displayDialogAsync` returns `12007`.</span></span> <span data-ttu-id="4e34b-151">示例如下。</span><span class="sxs-lookup"><span data-stu-id="4e34b-151">The following is an example.</span></span>

```javascript
function openFirstDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/firstDialog.html", { width: 50, height: 50},
     (result) => {
      if(result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.close();
        openSecondDialog();
      }
      else {
         // Handle errors
      }
    }
  );
}
 
function openSecondDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/secondDialog.html", { width: 50, height: 50},
    (result) => {
      if(result.status === Office.AsyncResultStatus.Failed) {
        if (result.error.code === 12007) {
          openSecondDialog(); // Recursive call
        }
        else {
         // Handle other errors
        }
      }
    }
  );
}
```

<span data-ttu-id="4e34b-152">或者，在代码尝试使用 [setTimeout](https://www.w3schools.com/jsref/met_win_settimeout.asp) 方法打开第二个对话框之前，可以强制代码暂停。</span><span class="sxs-lookup"><span data-stu-id="4e34b-152">Alternatively, you could force the code to pause before it tries to open the second dialog by using the [setTimeout](https://www.w3schools.com/jsref/met_win_settimeout.asp) method.</span></span> <span data-ttu-id="4e34b-153">示例如下。</span><span class="sxs-lookup"><span data-stu-id="4e34b-153">The following is an example.</span></span>

```javascript
function openFirstDialog() {
  Office.context.ui.displayDialogAsync("https://MyDomain/firstDialog.html", { width: 50, height: 50},
     (result) => {
      if(result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.close();
        setTimeout(() => { 
          Office.context.ui.displayDialogAsync("https://MyDomain/secondDialog.html", { width: 50, height: 50},
             (result) => { /* callback body */ }
          );
        }, 1000);
      }
      else {
         // Handle errors
      }
    }
  );
}
```

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a><span data-ttu-id="4e34b-154">在 SPA 中使用 Office 对话框 API 的最佳实践</span><span class="sxs-lookup"><span data-stu-id="4e34b-154">Best practices for using the Office dialog API in an SPA</span></span>

<span data-ttu-id="4e34b-155">如果您的外接程序使用客户端路由，就像单页应用程序 (SBA) 通常一样，您可以选择将路由的 URL 传递给 [displayDialogAsync](/javascript/api/office/office.ui) 方法，而不是单独的 HTML 页面的 URL。</span><span class="sxs-lookup"><span data-stu-id="4e34b-155">If your add-in uses client-side routing, as single-page applications (SPAs) typically do, you have the option to pass the URL of a route to the [displayDialogAsync](/javascript/api/office/office.ui) method instead of the URL of a separate HTML page.</span></span> <span data-ttu-id="4e34b-156">*出于以下给定原因，建议不要这样做。*</span><span class="sxs-lookup"><span data-stu-id="4e34b-156">*We recommend against doing so for the reasons given below.*</span></span>

> [!NOTE]
> <span data-ttu-id="4e34b-157">本文与服务器端路由不相关，例如，在基于 Express 的 Web 应用程序中。</span><span class="sxs-lookup"><span data-stu-id="4e34b-157">This article is not relevant to *server-side* routing, such as in an Express-based web application.</span></span>

#### <a name="problems-with-spas-and-the-office-dialog-api"></a><span data-ttu-id="4e34b-158">SBA 和 Office 对话框 API 的问题</span><span class="sxs-lookup"><span data-stu-id="4e34b-158">Problems with SPAs and the Office dialog API</span></span>

<span data-ttu-id="4e34b-159">Office 对话框位于具有自己的 JavaScript 引擎实例的新窗口中，因此它是自己的完整执行上下文。</span><span class="sxs-lookup"><span data-stu-id="4e34b-159">The Office dialog box is in a new window with its own instance of the JavaScript engine, and hence it's own complete execution context.</span></span> <span data-ttu-id="4e34b-160">如果传递路由，基页及其所有初始化和引导代码将在此新上下文中再次运行，并且任何变量都设置为对话框中的初始值。</span><span class="sxs-lookup"><span data-stu-id="4e34b-160">If you pass a route, your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog box.</span></span> <span data-ttu-id="4e34b-161">因此，此技术在框窗口中下载并启动应用程序的第二个实例，这部分抵消了 SPA 的用途。</span><span class="sxs-lookup"><span data-stu-id="4e34b-161">So this technique downloads and launches a second instance of your application in the  box window, which partially defeats the purpose of an SPA.</span></span> <span data-ttu-id="4e34b-162">此外，在对话框窗口中更改变量的代码不会更改相同变量的任务窗格版本。</span><span class="sxs-lookup"><span data-stu-id="4e34b-162">In addition, code that changes variables in the dialog box window does not change the task pane version of the same variables.</span></span> <span data-ttu-id="4e34b-163">同样，对话框窗口具有其自己的会话存储 ([Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) 属性) ，无法从任务窗格中的代码访问。</span><span class="sxs-lookup"><span data-stu-id="4e34b-163">Similarly, the dialog box window has its own session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property), which is not accessible from code in the task pane.</span></span> <span data-ttu-id="4e34b-164">对话框和被调用的主机页与服务器的 `displayDialogAsync` 两个不同的客户端类似。</span><span class="sxs-lookup"><span data-stu-id="4e34b-164">The dialog box and the host page on which `displayDialogAsync` was called look like two different clients to your server.</span></span> <span data-ttu-id="4e34b-165"> (有关主机页的提醒，请参阅从主机页[.) ](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)</span><span class="sxs-lookup"><span data-stu-id="4e34b-165">(For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).)</span></span>

<span data-ttu-id="4e34b-166">因此，如果向该方法传递了路由，则实际上没有 SPA;同一 SPA 有两 `displayDialogAsync` *个实例*。</span><span class="sxs-lookup"><span data-stu-id="4e34b-166">So, if you passed a route to the `displayDialogAsync` method, you wouldn't really have an SPA; you'd have *two instances of the same SPA*.</span></span> <span data-ttu-id="4e34b-167">此外，任务窗格实例中的大部分代码绝不会用于该实例，并且对话框实例中的大部分代码绝不会用于该实例。</span><span class="sxs-lookup"><span data-stu-id="4e34b-167">Moreover, much of the code in the task pane instance would never be used in that instance and much of the code in the dialog box instance would never be used in that instance.</span></span> <span data-ttu-id="4e34b-168">这相当于相同捆绑包中拥有两个 SPA。</span><span class="sxs-lookup"><span data-stu-id="4e34b-168">It would be like having two SPAs in the same bundle.</span></span>

#### <a name="microsoft-recommendations"></a><span data-ttu-id="4e34b-169">Microsoft 建议</span><span class="sxs-lookup"><span data-stu-id="4e34b-169">Microsoft recommendations</span></span>

<span data-ttu-id="4e34b-170">建议执行下列操作之一，而不是将客户端路由传递给 `displayDialogAsync` 该方法：</span><span class="sxs-lookup"><span data-stu-id="4e34b-170">Instead of passing a client-side route to the `displayDialogAsync` method, we recommend that you do one of the following:</span></span>

* <span data-ttu-id="4e34b-171">如果要在对话框中运行的代码非常复杂，请显式创建两个不同的 SBA;也就是说，在同一域的不同文件夹中具有两个 SBA。</span><span class="sxs-lookup"><span data-stu-id="4e34b-171">If the code that you want to run in the dialog box is sufficiently complex, create two different SPAs explicitly; that is, have two SPAs in different folders of the same domain.</span></span> <span data-ttu-id="4e34b-172">一个 SPA 在对话框中运行，另一个 SPA 在对话框的主机页中运行，其中一个 SPA 在调用 `displayDialogAsync` 该对话框的主机页中运行。</span><span class="sxs-lookup"><span data-stu-id="4e34b-172">One SPA runs in the dialog box and the other in the dialog box's host page where `displayDialogAsync` was called.</span></span> 
* <span data-ttu-id="4e34b-173">在大多数情况下，对话框中只需要简单逻辑。</span><span class="sxs-lookup"><span data-stu-id="4e34b-173">In most scenarios, only simple logic is needed in the dialog box.</span></span> <span data-ttu-id="4e34b-174">在这种情况下，您的项目将在 SPA 域中承载单个 HTML 页面（包含嵌入或引用的 JavaScript）大大简化。</span><span class="sxs-lookup"><span data-stu-id="4e34b-174">In such cases, your project will be greatly simplified by hosting a single HTML page, with embedded or referenced JavaScript, in the domain of your SPA.</span></span> <span data-ttu-id="4e34b-175">将页面的 URL 传递给 `displayDialogAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="4e34b-175">Pass the URL of the page to the `displayDialogAsync` method.</span></span> <span data-ttu-id="4e34b-176">虽然这意味着你正在从单页应用文字概念中脱除;使用 Office 对话框 API 时，实际上没有 SPA 的单个实例。</span><span class="sxs-lookup"><span data-stu-id="4e34b-176">While this means that you are deviating from the literal idea of a single-page app; you don't really have a single instance of an SPA when you are using the Office dialog API.</span></span>
