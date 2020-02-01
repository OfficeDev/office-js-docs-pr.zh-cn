---
title: Office 对话框 API 的最佳实践和规则
description: 提供 Office 对话框 API 的规则和最佳做法，例如单页应用程序的最佳实践（SPA）
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 7a38337ca9a263df1f8405f2883fa4481c342e6b
ms.sourcegitcommit: 4c9e02dac6f8030efc7415e699370753ec9415c8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/01/2020
ms.locfileid: "41650071"
---
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a>Office 对话框 API 的最佳实践和规则

本文提供 Office dialog API 的规则、陷阱和最佳实践，包括设计对话的 UI 和在单页面应用程序（SPA）中使用 API 的最佳实践

> [!NOTE]
> 本文 presupposes 您熟悉使用 Office 对话框 API 的基础知识，如在[Office 外接程序中使用 office 对话框 api](dialog-api-in-office-add-ins.md)中所述。
> 
> 另请参阅[处理 Office 对话框中的错误和事件](dialog-handle-errors-events.md)。

## <a name="rules-and-gotchas"></a>规则和陷阱

- 对话框只能导航到 HTTPS Url，而不能导航到 HTTP。
- 传递给[displayDialogAsync](/javascript/api/office/office.ui)方法的 URL 必须与加载项本身在完全相同的域中。 它不能是子域。 但传递给它的页面可以重定向到另一个域中的页面。
- 主机窗口（可以是任务窗格或外接程序命令的无用户 UI 的[函数文件](/office/dev/add-ins/reference/manifest/functionfile)）一次只能打开一个对话框。
- 在该对话框中仅可调用两个 Office Api：
  - [Office.context.ui.messageparent](/javascript/api/office/office.ui#messageparent-message-)函数。
  - `Office.context.requirements.isSetSupported`（有关详细信息，请参阅[指定 Office 主机和 API 要求](specify-office-hosts-and-api-requirements.md)。）
- [Office.context.ui.messageparent](/javascript/api/office/office.ui#messageparent-message-)函数只能从与加载项本身完全相同的域中的页面进行调用。

## <a name="best-practices"></a>最佳做法

### <a name="avoid-overusing-dialog-boxes"></a>避免过度使用对话框

由于不鼓励重叠的 UI 元素，因此除非您的方案需要，否则请避免从任务窗格中打开对话框。 考虑如何使用任务窗格的区域时，请注意任务窗格可以是选项卡式。 有关示例，请参阅 [Excel 加载项 JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 示例。

### <a name="designing-a-dialog-box-ui"></a>设计对话框 UI

有关对话框设计中的最佳做法，请参阅[Office 外接程序中的对话框](../design/dialog-boxes.md)。

### <a name="handling-pop-up-blockers-with-office-on-the-web"></a>使用 Office 网页版处理弹出窗口阻止程序

在使用 web 上的 Office 时尝试显示对话框可能会导致浏览器的弹出窗口阻止器阻止对话框。 网站上的 Office 具有一项功能，可使您的外接程序的对话框成为浏览器的弹出窗口阻止程序的例外。 当代码调用`displayDialogAsync`方法时，网站上的 Office 将打开一个类似于以下的提示。

![外接程序可以生成的提示，以避免在浏览器中弹出窗口阻止程序。](../images/dialog-prompt-before-open.png)

如果用户选择 "**允许**"，则会打开 "Office" 对话框。 如果用户选择 "**忽略**"，则会关闭提示，并且 Office 对话框不会打开。 相反，该`displayDialogAsync`方法将返回错误12009。 您的代码应捕获此错误，并提供不需要对话框的备用体验，或向用户显示一条消息，提示外接程序要求其允许对话框。 （有关12009的详细信息，请参阅[displayDialogAsync 中的错误](dialog-handle-errors-events.md#errors-from-displaydialogasync)。）

如果出于任何原因需要关闭此功能，则代码必须选择退出。它向传递给该`displayDialogAsync`方法的[DialogOptions](/javascript/api/office/office.dialogoptions)对象发出此请求。 具体来说，该对象应`promptBeforeOpen: false`包含。 当此选项设置为 false 时，web 上的 Office 将不会提示用户允许加载项打开对话框，并且 Office 对话框不会打开。

### <a name="do-not-use-the-_host_info-value"></a>不使用\_主机\_信息值

Office 会自动向传递给 `_host_info` 的 URL 添加查询参数 `displayDialogAsync`。 将其追加到自定义查询参数（如果有）后面。 它不会追加到对话框导航到的任何后续 Url 中。 Microsoft 可能会更改此值的内容或将其全部删除，因此您的代码不应读取它。 相同的值会被添加到对话框的会话存储中。 同样，*代码不得对此值执行读取和写入操作*。

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a>在 SPA 中使用 Office 对话框 API 的最佳实践

如果外接程序使用客户端路由，作为单页应用程序（Spa）通常会选择将路由的 URL 传递到[displayDialogAsync](/javascript/api/office/office.ui)方法，而不是将其 url 传递到单独的 HTML 页面的 url。 *出于以下给出的原因，我们建议您不要这样做。*

> [!NOTE]
> 本文与*服务器端*路由不相关，例如在基于 Express 的 web 应用程序中。

#### <a name="problems-with-spas-and-the-office-dialog-api"></a>Spa 和 Office 对话框 API 存在的问题

Office 对话框位于具有其自己的 JavaScript 引擎实例的新窗口中，因此它拥有完整的执行上下文。 如果传递路由，基本页面及其所有初始化和引导代码将在此新上下文中再次运行，并且所有变量在对话框中都设置为其初始值。 因此，此技术会在 box 窗口中下载并启动应用程序的第二个实例，这部分将导致 SPA 的目的不是一个。 此外，更改对话框窗口中的变量的代码不会更改相同变量的任务窗格版本。 同样，对话框窗口具有自己的会话存储，该存储无法从任务窗格中的代码访问。 对话框和主页面在上面称为与您`displayDialogAsync`的服务器的两个不同的客户端。 （有关主机页的提示，请参阅[从主机页打开对话框](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。）

因此，如果您将路由传递给`displayDialogAsync`方法，则不会真正有 SPA;您有*两个相同 SPA 的实例*。 此外，任务窗格实例中的很多代码永远不会在该实例中使用，并且对话框实例中的大部分代码也不会在该实例中使用。 这相当于相同捆绑包中拥有两个 SPA。

#### <a name="microsoft-recommendations"></a>Microsoft 建议

建议您执行以下操作之一，而不是`displayDialogAsync`将客户端路由传递给方法：

* 如果要在对话框中运行的代码足够复杂，请显式创建两个不同的 Spa;也就是说，在同一域的不同文件夹中有两个 Spa。 一个 SPA 在对话框中运行，而另一个 SPA 在对话框的主机页`displayDialogAsync`中调用。 
* 在大多数情况下，只有简单逻辑在对话框中是必需的。 在这种情况下，您的项目将通过在您的 SPA 的域中托管单个 HTML 页面（包括嵌入或引用的 JavaScript）而大大简化。 将页面的 URL 传递给 `displayDialogAsync` 方法。 这意味着您可以从单页应用程序的原义想法中 deviating;如果使用的是 Office 对话框 API，则实际上不具有 SPA 的单个实例。
