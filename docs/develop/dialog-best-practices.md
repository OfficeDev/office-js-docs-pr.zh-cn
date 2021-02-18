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
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a>Office 对话框 API 最佳做法和规则

本文提供 Office 对话框 API 的规则、链链和最佳做法，包括设计对话框 UI 的最佳实践，以及将 API 与在单页应用程序 (SPA) 

> [!NOTE]
> 本文假定你熟悉使用 Office 对话框 API 的基础知识，如在 Office 外接程序中使用 [Office](dialog-api-in-office-add-ins.md)对话框 API 中所述。
> 
> 另请参阅 [使用 Office 对话框处理错误和事件](dialog-handle-errors-events.md)。

## <a name="rules-and-gotchas"></a>规则和陷阱

- 对话框只能导航到 HTTPS URL，不能导航到 HTTP。
- 传递给 [displayDialogAsync](/javascript/api/office/office.ui) 方法的 URL 必须与加载项本身完全相同的域中。 它不能是子域。 但是，传递给它的页面可以重定向到另一个域中的页面。
- 主机窗口可以是任务窗格或加载项命令的无 UI 函数文件[](../reference/manifest/functionfile.md)，一次只能打开一个对话框。
- 对话框中只能调用两个 Office API：
  - [messageParent](/javascript/api/office/office.ui#messageparent-message-)函数。
  - `Office.context.requirements.isSetSupported` (有关详细信息，请参阅指定 [Office 应用程序和 API](specify-office-hosts-and-api-requirements.md)要求 .) 
- [messageParent](/javascript/api/office/office.ui#messageparent-message-)函数只能从与外接程序本身完全相同的域中的页面调用。

## <a name="best-practices"></a>最佳做法

### <a name="avoid-overusing-dialog-boxes"></a>避免过度使用对话框

由于不赞成重叠 UI 元素，因此除非应用场景需要，否则请勿从任务窗格打开对话框。 考虑如何使用任务窗格区域时，请注意任务窗格中可以有选项卡。 有关选项卡式任务窗格的示例，请参阅 [Excel 外接程序 JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 示例。

### <a name="designing-a-dialog-box-ui"></a>设计对话框 UI

有关对话框设计中的最佳方案，请参阅 Office [加载项中的对话框](../design/dialog-boxes.md)。

### <a name="handling-pop-up-blockers-with-office-on-the-web"></a>使用 Office 网页版处理弹出窗口阻止程序

尝试在 Web 上使用 Office 时显示对话框可能会导致浏览器的弹出窗口阻止程序阻止对话框。 Office 网页 Office 具有一项功能，该功能可使外接程序的对话框成为浏览器弹出窗口阻止程序例外。 当代码调用该方法 `displayDialogAsync` 时，Office 网页发布将打开类似于以下内容的提示。

![Screenshot showing the prompt with a brief description and Allow and Ignore buttons that an add-in can generate to avoid in-browser pop-up blockers](../images/dialog-prompt-before-open.png)

如果用户选择"允许 **"，** 将打开 Office 对话框。 如果用户选择"忽略 **"，** 则提示将关闭，并且 Office 对话框不会打开。 相反， `displayDialogAsync` 此方法返回错误 12009。 代码应捕获此错误，并提供不需要对话框的备用体验，或向用户显示一条消息，提示加载项要求他们允许对话框。  (有关 12009 的详细信息，请参阅 [displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).) 

如果出于任何原因要关闭此功能，则你的代码必须选择退出。它使用传递给该方法的 [DialogOptions](/javascript/api/office/office.dialogoptions) 对象进行 `displayDialogAsync` 此请求。 具体来说，对象应包括 `promptBeforeOpen: false` 。 当此选项设置为 false 时，Web 上的 Office 不会提示用户允许外接程序打开对话框，并且 Office 对话框将不会打开。

### <a name="do-not-use-the-_host_info-value"></a>请勿使用 \_ 主机 \_ 信息值

Office 会自动向传递给 `_host_info` 的 URL 添加查询参数 `displayDialogAsync`。 它将追加到自定义查询参数（如果有）之后。 它未追加到对话框导航到的任何后续 URL。 Microsoft 可能会更改此值的内容或完全删除它，因此代码不应读取它。 相同的值将添加到对话框的会话存储 (，即 [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) 属性) 。 同样，*代码不得对此值执行读取和写入操作*。

### <a name="opening-another-dialog-immediately-after-closing-one"></a>关闭另一个对话框后立即打开另一个对话框

不能从给定的主机页打开多个对话框，因此代码应在打开的对话框中调用 [Dialog.close，](/javascript/api/office/office.dialog#close__) 然后再调用它以 `displayDialogAsync` 打开另一个对话框。 该方法 `close` 是异步的。 因此，如果在调用后立即调用，则当 Office 尝试打开第二个对话框时，第一个对话框 `displayDialogAsync` `close` 可能尚未完全关闭。 如果发生这种情况，Office 将返回 [12007](dialog-handle-errors-events.md#12007) 错误："操作失败，因为此外接程序已有活动对话框。"

该方法不接受回调参数，并且不会返回 Promise 对象，因此无法使用关键字或 `close` `await` 方法等待 `then` 该对象。 出于此原因，建议在关闭对话框后立即打开新对话框时采用以下技术：封装代码以在方法中打开新对话框，并设计方法以递归方式调用自身（如果调用 `displayDialogAsync` 返回 `12007` ）。 示例如下。

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

或者，在代码尝试使用 [setTimeout](https://www.w3schools.com/jsref/met_win_settimeout.asp) 方法打开第二个对话框之前，可以强制代码暂停。 示例如下。

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

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a>在 SPA 中使用 Office 对话框 API 的最佳实践

如果您的外接程序使用客户端路由，就像单页应用程序 (SBA) 通常一样，您可以选择将路由的 URL 传递给 [displayDialogAsync](/javascript/api/office/office.ui) 方法，而不是单独的 HTML 页面的 URL。 *出于以下给定原因，建议不要这样做。*

> [!NOTE]
> 本文与服务器端路由不相关，例如，在基于 Express 的 Web 应用程序中。

#### <a name="problems-with-spas-and-the-office-dialog-api"></a>SBA 和 Office 对话框 API 的问题

Office 对话框位于具有自己的 JavaScript 引擎实例的新窗口中，因此它是自己的完整执行上下文。 如果传递路由，基页及其所有初始化和引导代码将在此新上下文中再次运行，并且任何变量都设置为对话框中的初始值。 因此，此技术在框窗口中下载并启动应用程序的第二个实例，这部分抵消了 SPA 的用途。 此外，在对话框窗口中更改变量的代码不会更改相同变量的任务窗格版本。 同样，对话框窗口具有其自己的会话存储 ([Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) 属性) ，无法从任务窗格中的代码访问。 对话框和被调用的主机页与服务器的 `displayDialogAsync` 两个不同的客户端类似。  (有关主机页的提醒，请参阅从主机页[.) ](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)

因此，如果向该方法传递了路由，则实际上没有 SPA;同一 SPA 有两 `displayDialogAsync` *个实例*。 此外，任务窗格实例中的大部分代码绝不会用于该实例，并且对话框实例中的大部分代码绝不会用于该实例。 这相当于相同捆绑包中拥有两个 SPA。

#### <a name="microsoft-recommendations"></a>Microsoft 建议

建议执行下列操作之一，而不是将客户端路由传递给 `displayDialogAsync` 该方法：

* 如果要在对话框中运行的代码非常复杂，请显式创建两个不同的 SBA;也就是说，在同一域的不同文件夹中具有两个 SBA。 一个 SPA 在对话框中运行，另一个 SPA 在对话框的主机页中运行，其中一个 SPA 在调用 `displayDialogAsync` 该对话框的主机页中运行。 
* 在大多数情况下，对话框中只需要简单逻辑。 在这种情况下，您的项目将在 SPA 域中承载单个 HTML 页面（包含嵌入或引用的 JavaScript）大大简化。 将页面的 URL 传递给 `displayDialogAsync` 方法。 虽然这意味着你正在从单页应用文字概念中脱除;使用 Office 对话框 API 时，实际上没有 SPA 的单个实例。
