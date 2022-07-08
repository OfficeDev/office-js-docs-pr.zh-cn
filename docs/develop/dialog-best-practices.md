---
title: Office 对话框 API 最佳做法和规则
description: 为 Office 对话框 API 提供规则和最佳做法，例如单页应用程序 (SPA) 的最佳做法。
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: dfe9841d12865c488a86a203026684e0b3570352
ms.sourcegitcommit: c62d087c27422db51f99ed7b14216c1acfda7fba
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/08/2022
ms.locfileid: "66689395"
---
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a>Office 对话框 API 最佳做法和规则

本文提供 Office 对话 API 的规则、权限和最佳做法，包括设计对话框的 UI 以及在单页应用程序中使用 API 的最佳做法 (SPA) 

> [!NOTE]
> 本文假设你熟悉使用 Office 对话 API 的基础知识，如 Office 加载项中的 [“使用 Office”对话框 API 中](dialog-api-in-office-add-ins.md)所述。
> 
> 另请参阅 [Office 对话框中的“处理错误和事件](dialog-handle-errors-events.md)”。

## <a name="rules-and-gotchas"></a>规则和陷阱

- 对话框只能导航到 HTTPS URL，而不能导航到 HTTP。
- 传递给 [displayDialogAsync 方法的](/javascript/api/office/office.ui) URL 必须与外接程序本身位于完全相同的域中。 它不能是子域。 但是传递给它的页面可以重定向到另一个域中的页面。
- 主机页一次只能打开一个对话框。 主机页可以是任务窗格，也可以是[函数命令的函](../design/add-in-commands.md#types-of-add-in-commands)数[文件](/javascript/api/manifest/functionfile)。
- 对话框中只能调用两个 Office API：
  - [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) 函数。
  - `Office.context.requirements.isSetSupported` (有关详细信息，请参阅 [指定 Office 应用程序和 API 要求](specify-office-hosts-and-api-requirements.md)。) 
- [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) 函数通常应从与加载项本身完全相同的域中的页面调用，但这不是必需的。 有关详细信息，请参阅[向主机运行时间跨域消息传递](dialog-api-in-office-add-ins.md#cross-domain-messaging-to-the-host-runtime)。

## <a name="best-practices"></a>最佳做法

### <a name="avoid-overusing-dialog-boxes"></a>避免过度使用对话框

由于不赞成重叠 UI 元素，因此除非应用场景需要，否则请勿从任务窗格打开对话框。 考虑如何使用任务窗格区域时，请注意任务窗格中可以有选项卡。 有关选项卡式任务窗格的示例，请参阅 [Excel 加载项 JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 示例。

### <a name="design-a-dialog-box-ui"></a>设计对话框 UI

有关对话框设计中的最佳做法，请参阅 [Office 加载项中的对话框](../develop/dialog-api-in-office-add-ins.md)。

### <a name="handle-pop-up-blockers-with-office-on-the-web"></a>使用Office web 版处理弹出窗口阻止程序

尝试在使用Office web 版时显示对话框可能会导致浏览器的弹出窗口阻止程序阻止该对话框。 如果发生这种情况，Office web 版将打开如下所示的提示。

![显示提示的屏幕截图，其中包含简短说明以及加载项可以生成的“允许和忽略”按钮，以避免浏览器中弹出的阻止程序](../images/dialog-prompt-before-open.png)

如果用户选择 **“允许**”，则会打开“Office”对话框。 如果用户选择 **“忽略**”，则提示将关闭，并且“Office”对话框不会打开。 相反，该 `displayDialogAsync` 方法返回错误 12009。 代码应捕获此错误，并提供不需要对话的备用体验，或向用户显示一条消息，告知加载项要求他们允许对话。  (有关 12009 的详细信息，请参阅 [displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).) 

如果出于任何原因想要关闭此功能，则代码必须选择退出。它使用传递给`displayDialogAsync`方法的 [DialogOptions](/javascript/api/office/office.dialogoptions) 对象发出此请求。 具体而言，该对象应包括在内 `promptBeforeOpen: false`。 当此选项设置为 false 时，Office web 版不会提示用户允许外接程序打开对话框，并且 Office 对话框将不会打开。

### <a name="do-not-use-the-_host_info-value"></a>请勿使用 \_主机\_信息值

Office 会自动向传递给 `_host_info` 的 URL 添加查询参数 `displayDialogAsync`。 它将追加到自定义查询参数（如果有）之后。 它不会追加到对话框导航到的任何后续 URL。 Microsoft 可能会更改此值的内容，或将其完全删除，因此代码不应读取它。 同一值添加到对话框的会话存储 (即 [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) 属性) 。 同样，*代码不得对此值执行读取和写入操作*。

### <a name="open-another-dialog-immediately-after-closing-one"></a>关闭一个对话框后立即打开另一个对话框

不能从给定的主机页打开多个对话框，因此代码应在打开的对话框中调用 [Dialog.close](/javascript/api/office/office.dialog#office-office-dialog-close-member(1)) ，然后再调用 `displayDialogAsync` 该对话框以打开另一个对话框。 该 `close` 方法是异步的。 因此，如果在呼叫`close`后立即呼叫`displayDialogAsync`，则当 Office 尝试打开第二个对话框时，第一个对话可能尚未完全关闭。 如果发生这种情况，Office 将返回 [12007](dialog-handle-errors-events.md#12007) 错误：“操作失败，因为此加载项已具有活动对话框。

该`close`方法不接受回调参数，也不会返回 Promise 对象，因此无法使用关键字或`then`方法等待`await`它。 因此，当需要在关闭对话框后立即打开新对话框时，我们建议使用以下技术：封装代码以在方法中打开新对话框，并在调用 `displayDialogAsync` 返回 `12007`时设计以递归方式调用自己的方法。 示例如下。

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

或者，可以使用 [setTimeout](https://www.w3schools.com/jsref/met_win_settimeout.asp) 方法强制代码暂停，然后再尝试打开第二个对话框。 示例如下。

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

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a>在 SPA 中使用 Office 对话 API 的最佳做法

如果加载项使用客户端路由，就像单页应用程序 (SPA) 通常一样，则可以选择将路由的 URL 传递到 [displayDialogAsync](/javascript/api/office/office.ui) 方法，而不是单独的 HTML 页面的 URL。 *出于下面给出的原因，建议不要这样做。*

> [!NOTE]
> 本文与 *服务器端* 路由（例如基于 Express 的 Web 应用程序）无关。

#### <a name="problems-with-spas-and-the-office-dialog-api"></a>SPA 和 Office 对话框 API 出现问题

Office 对话框位于一个新窗口中，其中包含自己的 JavaScript 引擎实例，因此它是自己的完整执行上下文。 如果传递路由，则基页及其所有初始化和启动代码在此新上下文中再次运行，并且任何变量都设置为对话框中的初始值。 因此，此技术在框窗口中下载并启动应用程序的第二个实例，这部分地违背了 SPA 的用途。 此外，在对话框窗口中更改变量的代码不会更改相同变量的任务窗格版本。 同样，对话框窗口具有自己的会话存储 ([Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) 属性) ，该属性无法从任务窗格中的代码访问。 调用的对话框和主机页类似于服务器的 `displayDialogAsync` 两个不同的客户端。  (有关主机页的提醒，请参阅 [主机页中的“打开”对话框](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。) 

因此，如果将路由传递到 `displayDialogAsync` 该方法，则不会真正具有 SPA;你将拥有 *两个相同 SPA 的* 实例。 此外，任务窗格实例中的大部分代码永远不会在该实例中使用，对话框实例中的许多代码永远不会在该实例中使用。 这相当于相同捆绑包中拥有两个 SPA。

#### <a name="microsoft-recommendations"></a>Microsoft 建议

我们建议你执行以下操作之一，而不是将客户端路由传递给 `displayDialogAsync` 方法：

* 如果要在对话框中运行的代码足够复杂，请显式创建两个不同的 SPA;也就是说，在同一域的不同文件夹中具有两个 SPA。 一个 SPA 在对话框中运行，另一个在调用的对话框的主机页 `displayDialogAsync` 中运行。 
* 在大多数情况下，对话框中只需要简单的逻辑。 在这种情况下，通过在 SPA 域中托管包含嵌入或引用 JavaScript 的单个 HTML 页面，将大大简化项目。 将页面的 URL 传递给 `displayDialogAsync` 方法。 虽然这意味着你偏离了单页应用的文本概念;使用 Office 对话 API 时，实际上没有 SPA 的单个实例。
