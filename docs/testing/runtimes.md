---
title: Office 加载项中的运行时
description: 了解 Office 外接程序使用的运行时。
ms.date: 08/29/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8d28f6db028d2f4c7036db51ccc5dbcc2144bdf3
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616040"
---
# <a name="runtimes-in-office-add-ins"></a>Office 加载项中的运行时

Office 加载项在 Office 中嵌入的运行时中执行。 作为解释语言，JavaScript 必须在 JavaScript 引擎中运行。 作为单线程同步语言，JavaScript 没有用于并发执行的固有容量;但新式 JavaScript 引擎可以请求并发操作 (包括来自主机操作系统的网络通信) 并接收来自 OS 的数据作为响应。 这种引擎使 JavaScript *能够有效地* 实现异步。 在本文中，此类引擎称为 *运行时*。 [Node.js](https://nodejs.org) 和新式浏览器是此类运行时的示例。 

## <a name="types-of-runtimes"></a>运行时类型

Office 外接程序使用两种类型的运行时：

- **仅限 JavaScript 的运行时**：JavaScript 引擎补充了对 [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API)、 [完整 CORS (跨源资源共享)](https://developer.mozilla.org/docs/Web/HTTP/CORS)和客户端存储数据的支持。  (它不支持 [本地存储](https://developer.mozilla.org/docs/Web/API/Window/localStorage) 或 Cookie.)  
- **浏览器运行时**：包括仅限 JavaScript 的运行时的所有功能，并添加对 [本地存储](https://developer.mozilla.org/docs/Web/API/Window/localStorage)、呈现 HTML 的 [呈现引擎](https://developer.mozilla.org/docs/Glossary/Rendering_engine) 和 Cookie 的支持。

本文稍后在 [仅限 JavaScript 的运行时](#javascript-only-runtime) 和 [浏览器运行时](#browser-runtime)中介绍了这些类型的详细信息。

下表显示了加载项的哪些可能功能使用每种类型的运行时。 

> [!NOTE]
> 选择要使用的运行时类型是 Microsoft 随时可以更改的实现详细信息。 Office JavaScript 库不假定同一类型的运行时将始终用于给定功能，外接程序体系结构也不应假定这一点。

| 运行时类型 | 加载项功能 |
|:-----|:-----|
| 仅 JavaScript | Excel [自定义函数](../excel/custom-functions-overview.md)</br> (，除非[共享](#shared-runtime)运行时或外接程序在Office web 版) 中运行</br></br>[基于 Outlook 事件的任务](../outlook/autolaunch.md)</br>仅当外接程序在 Outlook on Windows) 中运行时才 (|
| 浏览器 | [任务窗格](../design/task-pane-add-ins.md)</br></br>[对话框](../develop/dialog-api-in-office-add-ins.md)</br></br>[function 命令](../design/add-in-commands.md#types-of-add-in-commands)</br></br>Excel [自定义函数](../excel/custom-functions-overview.md)</br> ([共享](#shared-runtime)运行时或外接程序在Office web 版) 中运行时</br></br>[基于 Outlook 事件的任务](../outlook/autolaunch.md)</br> (加载项在 Outlook on Mac 中运行或Outlook 网页版) |

下表显示了由哪种类型的运行时用于加载项的各种可能功能的相同信息。

| 加载项功能 | Windows 上的运行时类型 | Mac 上的运行时类型 | Web 上的运行时类型 |
|:-----|:-----|:-----|:-----|
|Excel 自定义函数 | 仅 JavaScript</br> (但在共享运行时时 *浏览器*) |仅 JavaScript</br> (但在共享运行时时 *浏览器*) | 浏览器 |
|基于 Outlook 事件的任务 | 仅 JavaScript | 浏览器 | 浏览器 |
|任务窗格 | 浏览器 | 浏览器 | 浏览器 |
|对话框 | 浏览器 | 浏览器 | 浏览器 |
|function 命令 | 浏览器 | 浏览器 | 浏览器 |


在Office web 版中，所有内容始终在浏览器类型运行时中运行。 事实上，除了一个例外，Web 加载项中的所有内容都 *在同* 一浏览器进程中运行：用户在其中打开的浏览器进程Office web 版。 例外情况是打开对话框时，会调用 [Office.ui.displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1))，*并且未* 传递并设置为 `true`[DialogOptions.displayInIFrame](/javascript/api/office/office.dialogoptions#office-office-dialogoptions-displayiniframe-member) 选项。 如果选项未 (传递，因此它具有默认 `false` 值) ，则对话框将在其自己的进程中打开。 相同的原则适用于 [OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) 方法和 [OfficeRuntime.DisplayWebDialogOptions.displayInIFrame](/javascript/api/office-runtime/officeruntime.displaywebdialogoptions#office-runtime-officeruntime-displaywebdialogoptions-displayiniframe-member) 选项。

当加载项在 Web 以外的平台上运行时，将应用以下原则。

- 对话框在其自己的运行时进程中运行。 
- 基于 Outlook 事件的任务在其自己的运行时进程中运行。 
- 默认情况下，任务窗格、函数命令和 Excel 自定义函数分别在其自己的运行时进程中运行。 但是，对于某些 Office 主机应用程序，可以配置外接程序清单，以便任何两个或全部三个都可以在同一运行时运行。 请参阅 [共享运行时](#shared-runtime)。

根据主机 Office 应用程序和外接程序中使用的功能，外接程序中可能会有许多运行时。 每个操作通常都在自己的进程中运行，但不一定同时运行。 示例如下。

- 不共享任何运行时并包含以下功能的 PowerPoint 或 Word 加载项具有多达三个运行时。

  - 任务窗格
  - 函数命令
  - 对话 (可以从任务窗格或函数命令启动对话框。)  
  
      > [!NOTE]
      > 同时打开多个对话框不是一个好的做法，但如果加载项允许用户同时从任务窗格打开一个对话框，而从函数命令打开另一个对话框，则此加载项将有四个运行时。 任务窗格和给定函数命令调用一次只能有一个打开的对话框：但是，如果多次调用函数命令，则会在其前置对话框的顶部打开一个新对话框，每次调用，因此可能会有多个运行时。 此列表的其余部分忽略多个打开对话框的可能性。

- 不共享任何运行时并包含以下功能的 Excel 加载项具有多达 *四* 个运行时。

  - 任务窗格
  - 函数命令
  - 自定义函数
  - 可以从任务窗格、函数命令或自定义函数启动对话 (对话框。) 

- 具有相同功能并配置为在任务窗格、函数命令和自定义函数之间共享相同运行时的 Excel 外接程序具有 *两* 个运行时。 共享运行时一次只能打开一个对话框。
- 具有相同功能的 Excel 外接程序（只不过它没有对话框，并且配置为在任务窗格、函数命令和自定义函数之间共享相同的运行时）有 *一个* 运行时。
- 具有以下功能的 Outlook 加载项具有多达 *四* 个运行时。  (运行时不能在 Outlook.) 中共享

  - 任务窗格
  - 函数命令
  - 基于事件的任务
  - 对话 (对话可以从任务窗格或函数命令启动，但不能从基于事件的任务启动。) 

## <a name="share-data-across-runtimes"></a>跨运行时共享数据

> [!NOTE]
> - 如果知道外接程序将仅在Office web 版中使用，并且它不会打开任何对话框`displayInIFrame`，并且选项设置为`true`，则可以忽略此部分。 由于加载项中的所有内容都在同一运行时进程中运行，因此只需使用全局变量在功能之间共享数据即可。
> - 如上所述 [，在运行时类型](#types-of-runtimes)中，功能使用的运行时类型因平台而异。 最好避免使用基于平台的加载项代码，因此本部分中的指南建议跨平台工作的技术。 只有一种情况，如下所述，其中需要分支代码。 

对于 Excel、PowerPoint 和 Word 加载项，当需要共享数据的两个或多个功能（对话除外）时，请使用 [共享运行时](#shared-runtime) 。 在 Outlook 中，或者在共享运行时不可行的情况下，需要替代方法。 独立运行时进程中的外接程序部分不会自动共享全局数据，外接程序的 Web 应用程序服务器会将它们视为单独的会话，因此 [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) 不能用于在它们之间共享数据。 *以下指南假定你未使用共享运行时。*

- 使用 [Office.ui.messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) 和 [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) 方法在对话框及其父任务窗格、函数命令或自定义函数之间传递数据。 

    > [!NOTE]
    > `OfficeRuntime.storage`无法在对话框中调用这些方法，因此这不是在对话框与另一个运行时之间共享数据的选项。 

- 若要在任务窗格和函数命令之间共享数据，请将数据存储在 [Window.localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) 中，该数据在访问同一特定 [源](https://developer.mozilla.org/docs/Glossary/Origin)的所有运行时之间共享。 
    > [!NOTE]
    > LocalStorage 在仅限 JavaScript 的运行时中不可访问，因此在 Excel 自定义函数中不可用。 它也不能用于与基于 Outlook 事件的任务共享数据 (，因为这些任务在某些平台上使用仅限 JavaScript 的运行时) 。

    > [!TIP]
    > `Window.localStorage`加载项会话之间保留数据，由具有相同源的加载项共享。 对于加载项来说，这两个特征通常都是不可取的。 
    >
    > - 若要确保给定加载项的每个会话在加载项启动时开始重新调用 [Window.localStorage.clear](https://developer.mozilla.org/docs/Web/API/Storage/clear) 方法。 
    > - 若要允许某些存储值保留，但重新初始化其他值，请在加载项启动时使用 [Window.localStorage.setItem](https://developer.mozilla.org/docs/Web/API/Storage/setItem) ，该加载项应重置为初始值。 
    > - 若要完全删除项，请调用 [Window.localStorage.removeItem](https://developer.mozilla.org/docs/Web/API/Storage/removeItem)。

- 若要在 Excel 自定义函数与任何其他运行时之间共享数据，请使用 [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage)。
- 若要在基于 Outlook 事件的任务和任务窗格或函数命令之间共享数据，必须按 [Office.context.platform](/javascript/api/office/office.context#office-office-context-platform-member) 属性的值对代码进行分支。 

    - 当值 `PC` (Windows) 时，使用 [Office.sessionData](/javascript/api/outlook/office.sessiondata) API 存储和检索数据。
    - 值为 `Mac`值时，请按此列表前面所述使用 `Window.localStorage` 。

共享数据的其他方法包括：

- 将共享数据存储在可供所有运行时访问的联机数据库中。
- 将共享数据存储在加载项域的 Cookie 中，以便在浏览器运行时之间共享。 仅限 JavaScript 的运行时不支持 Cookie。

有关详细信息，请参阅 [保留加载项状态和设置](../develop/persisting-add-in-state-and-settings.md) ， [以及管理 Outlook 加载项的状态和设置](../outlook/manage-state-and-settings-outlook.md)。

## <a name="javascript-only-runtime"></a>仅限 JavaScript 的运行时

Office 外接程序中使用的仅限 JavaScript 的运行时是对最初为[React Native](https://reactnative.dev/)创建的开放源代码运行时的修改。 它包含一个 JavaScript 引擎，该引擎补充了对 [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API)、 [完整 CORS (跨源资源共享) ](https://developer.mozilla.org/docs/Web/HTTP/CORS)和 [OfficeRuntime.storage 的支持](/javascript/api/office-runtime/officeruntime.storage)。 它没有呈现引擎，也不支持 Cookie 或 [本地存储](https://developer.mozilla.org/docs/Web/API/Window/localStorage)。

这种类型的运行时仅在 Windows 上的 Office 和 Excel 自定义函数中用于基于 Outlook 事件的任务， *除非* 自定义函数 [共享运行时](#shared-runtime)。 

- 当用于 Excel 自定义函数时，运行时在工作表重新计算或自定义函数计算时启动。 在工作簿关闭之前，它不会关闭。  
- 在基于 Outlook 事件的任务中使用时，运行时会在事件发生时启动。 当发生以下第一个操作时，它将结束。

  - 事件处理程序调用 `completed` 其事件参数的方法。
  - 触发事件已过去 5 分钟。
  - 用户从触发事件的窗口更改焦点，例如消息撰写窗口。

JavaScript 运行时使用的内存更少，启动速度比浏览器运行时快，但功能较少。

## <a name="browser-runtime"></a>浏览器运行时

Office 外接程序使用不同的浏览器类型运行时，具体取决于 Office 运行 (Web、Mac 或 Windows) 的平台，以及 Windows 和 Office 的版本和版本。 例如，如果用户在 FireFox 浏览器中运行Office web 版，则使用 Firefox 运行时。 如果用户在 Mac 上运行 Office，则使用 Safari 运行时。 如果用户在 Windows 上运行 Office，则 Edge 或 Internet Explorer 将提供运行时，具体取决于 Windows 和 Office 的版本。 可在 [Office 加载项使用的浏览器中](../concepts/browsers-used-by-office-web-add-ins.md)找到详细信息。

所有这些运行时都包括 HTML 呈现引擎，并支持 [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API)、 [完整 CORS (跨源资源共享) ](https://developer.mozilla.org/docs/Web/HTTP/CORS)、 [本地存储](https://developer.mozilla.org/docs/Web/API/Window/localStorage)和 Cookie。 

浏览器运行时寿命因其实现的功能以及是否共享而异。

- 启动包含任务窗格的外接程序时，浏览器运行时将启动，除非它是已运行的共享运行时。 如果是共享运行时，则文档关闭时会关闭。 如果它不是共享运行时，则在任务窗格关闭时关闭。
- 打开对话框后，浏览器运行时将启动。 当对话关闭时，它会关闭。
- 当执行函数命令 (当用户选择其按钮或菜单项) 时，浏览器运行时将启动，除非它是已运行的共享运行时。 如果是共享运行时，则文档关闭时会关闭。 如果不是共享运行时，它会在发生以下第一个运行时关闭。
 
  - 函数命令调用 `completed` 其事件参数的方法。
  - 触发事件已过去 5 分钟。  (如果在函数命令中打开了对话框，并且在父运行时超时时仍处于打开状态，则对话框运行时将保持运行状态，直到对话关闭。) 

- 当 Excel 自定义函数使用共享运行时时，当自定义函数计算共享运行时是否出于其他某种原因尚未启动时，将启动浏览器类型运行时。 文档关闭时，它将关闭。

> [!NOTE]
> [共享](#shared-runtime)运行时时，代码可以在不关闭加载项的情况下关闭任务窗格。 有关详细信息，请参阅 [“显示或隐藏 Office 外接程序”的任务窗格](../develop/show-hide-add-in.md) 。

浏览器运行时比仅限 JavaScript 的运行时具有更多的功能，但启动速度较慢，使用内存较多。

### <a name="shared-runtime"></a>共享运行时

“共享运行时”不是一种运行时类型。 它指的是由外接程序的功能共享的 [浏览器类型运行时](#browser-runtime) ，否则每个运行时都有自己的运行时。 具体而言，可以选择将加载项的任务窗格和函数命令配置为共享运行时。 在 Excel 外接程序中，还可以配置自定义函数以共享任务窗格或函数命令的运行时，或者同时共享这两个函数。 执行此操作时，自定义函数在浏览器类型运行时（而不是 [仅限 JavaScript 的运行时](#javascript-only-runtime) ）中运行，否则会运行。 请参阅 [配置外接程序以使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md) ，了解共享运行时的优点和限制，以及将外接程序配置为使用共享运行时的说明。 简言之，仅 JavaScript 运行时使用的内存更少，启动速度更快，但功能较少。

> [!NOTE]
> - 只能在 Excel、PowerPoint 和 Word 中共享运行时。 
> - 无法将对话配置为共享运行时。 每个对话始终有自己的对话框，除非在Office web 版`displayInIFrame`中启动对话，并且选项设置为`true`”
> - 共享运行时从不使用原始 Microsoft Edge WebView (EdgeHTML) 运行时。 如果在 Office 外接程序) 使用的[浏览器](../concepts/browsers-used-by-office-web-add-ins.md)中指定 (满足将 Microsoft Edge 与基于 WebView2 (Chromium) 的条件，则使用该运行时。 否则，将使用 Internet Explorer 11 运行时。