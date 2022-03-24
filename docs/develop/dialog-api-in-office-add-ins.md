---
title: 在 Office 加载项中使用 Office 对话框 API
description: 了解在加载项中Office对话框的基础知识。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: c84a5cd9079b1af754375dfc803284165ccea6e9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743830"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>在 Office 加载项中使用 Office 对话框 API

可以在 Office 加载项中使用 [Office 对话框 API](/javascript/api/office/office.ui) 打开对话框。 本文提供了有关如何在 Office 加载项中使用对话框 API 的指南。

> [!NOTE]
> 若要了解对话框 API 目前的受支持情况，请参阅[对话框 API 要求集](../reference/requirement-sets/dialog-api-requirement-sets.md)。 对话框 API 当前受 Excel、PowerPoint 和 Word 支持。 Outlook不同邮箱要求&mdash;集中包含的支持，请参阅 API 参考了解更多详细信息。

对话框 API 的主要应用场景是为 Google、Facebook 或 Microsoft Graph 等资源启用身份验证。 有关详细信息，请在熟悉本文 *之后*，参阅 [使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。

不妨通过任务窗格/内容加载项/[加载项命令](../design/add-in-commands.md)打开对话框，以便执行下列操作：

- 显示无法直接在任务窗格中打开的登录页。
- 为加载项中的某些任务提供更多屏幕空间，或甚至整个屏幕。
- 托管在任务窗格中显得太小的视频。

> [!NOTE]
> 由于不赞成重叠 UI 元素，因此除非应用场景需要，否则请勿从任务窗格打开对话框。 考虑如何使用任务窗格区域时，请注意任务窗格中可以有选项卡。 有关选项卡式任务窗格的示例，请参阅 Excel [外接程序 JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 示例。

下图展示了对话框示例。

![Screenshot showing dialog with 3 sign-in options displayed in front of Word.](../images/auth-o-dialog-open.png)

请注意，对话框总是在屏幕的中心打开。 用户可以移动并重设对话框的大小。 该窗口 *是非* 模式窗口-用户可以继续与 Office 应用程序中的文档以及任务窗格中的页面（如果有）进行交互。

## <a name="open-a-dialog-box-from-a-host-page"></a>从主机页面打开对话框

Office JavaScript API 在 [Office.context.ui 命名空间](/javascript/api/office/office.ui)中包含一个 [Dialog](/javascript/api/office/office.dialog) 对象和两个函数。

为了打开对话框，代码（通常是任务窗格中的一页）调用 [displayDialogAsync](/javascript/api/office/office.ui) 方法，并将要打开的资源 URL 传递到此方法。 调用方法的页面称为“主机页”。 例如在任务窗格中的 index.html 页面上使用脚本调用此方法，随后 index.html 是打开此方法对话框的主机页。

对话框中打开的资源通常是页面，，但也可以是 MVC 应用中的控制器方法、路由、Web 服务方法或其他任何资源。 在本文中，“页面”或“网站”是指对话框中的资源。 以下代码是一个简单示例。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - URL 使用 **HTTPS** 协议。对话框中加载的所有页面都必须遵循此要求，而不仅仅是加载的第一个页面。
> - 对话框域与宿主页的域相同，宿主页可以是任务窗格中的页面，也可以是加载项命令的[函数文件](../reference/manifest/functionfile.md)。 这要求：传递到 `displayDialogAsync` 方法的页面、控制器方法或其他资源必须与主机页位于相同的域。

> [!IMPORTANT]
> 对话框中打开的主机页面和资源必须具有相同的完整域。 如果尝试传递 `displayDialogAsync` 加载项域的子域，则不会起作用。 完整域（包括任何子域）必须匹配。

加载第一个页面（或其它资源）后，用户可使用链接或其它用户界面来导航至任何使用 HTTPS 的网站（或其他资源）。 还可以将第一个页面设计为直接重定向到另一个站点。

默认情况下，对话框的高度和宽度占设备屏幕的 80%。不过，你也可以设置不同的百分比，只需将配置对象传递给方法即可，如下面的示例所示。

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

有关实现这一点的示例外接程序，请参阅 [Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。 有关使用 的更多示例， `displayDialogAsync`请参阅 [Samples](#samples)。

将两个值均设置为 100% 可有效提供全屏体验。（有效最大值为 99.5%，窗口仍可移动和调整大小。）

> [!NOTE]
> 只能从主机窗口打开一个对话框。 如果尝试再打开一个对话框，就会生成错误。 例如，如果用户从任务窗格打开对话框，则她无法从任务窗格中其他页面打开另一个对话框。 不过，如果对话框是通过[加载项命令](../design/add-in-commands.md)打开，那么只要选择此命令，就会打开新 HTML 文件（但不可见）。 这会新建（不可见的）主机窗口，所以每个这样的窗口都可以启动自己的对话框。 有关详细信息，请参阅 [displayDialogAsync 返回的错误](dialog-handle-errors-events.md#errors-from-displaydialogasync)。

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a>利用 Office 网页版中的性能选项

`displayInIframe` 属性是配置对象中另一个可以传递到 `displayDialogAsync` 的属性。 如果将此属性设置为 `true`，且加载项在 Office 网页版打开的文档中运行，对话框就会以浮动 iframe（而不是独立窗口）的形式打开，从而加快对话框的打开速度。 示例如下。

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

默认值为 `false`，与完全省略此属性时相同。 如果加载项未在加载项Office web 版，将`displayInIframe`忽略 。

> [!NOTE]
> 如果 **对话框** 将 `displayInIframe: true` 随时重定向到在 iframe 中无法打开的页面，则不应使用 。 例如，许多热门 Web 服务的登录页面（如 Google 和 Microsoft 帐户）无法通过 iframe 打开。

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>将信息从对话框发送到主机页

> [!NOTE]
>
> - 为清楚起见，在此部分中，我们将消息称为面向主机页，但严格来说，消息将进入任务窗格 (中的 *JavaScript* 运行时或托管函数文件) 的运行时。[](../reference/manifest/functionfile.md) 这种区别仅在跨域邮件的情况下十分明显。 有关详细信息，请参阅[向主机运行时间跨域消息传递](#cross-domain-messaging-to-the-host-runtime)。
> - 对话框无法与任务窗格中的主机页通信，除非Office JavaScript API 库。  (与使用 JavaScript API Office的任何页面一样，页面的脚本必须初始化外接程序。 有关详细信息，请参阅 [Initialize your Office Add-in](initialize-add-in.md).) 

对话框中的代码使用 [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) 函数向主机页发送字符串消息。 该字符串可以是单词、句子、XML blob、字符串化 JSON 或其他任何可以序列化为字符串或转换为字符串的字符串。 示例如下。

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true.toString());
}
```

> [!IMPORTANT]
> - 函数`messageParent`是唯一 *可以在Office* 调用的两个 JS API 之一。
> - 可以在对话框中调用的其他 JS API 是 `Office.context.requirements.isSetSupported`。 有关它的信息，请参阅指定[Office应用程序和 API 要求](specify-office-hosts-and-api-requirements.md)。 但是，在对话框中，此 API 在一Outlook 2016购买中不受 (，即 MSI 版本) 。

在下一个示例中，`googleProfile` 是用户 Google 配置文件的字符串化版本。

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

必须将主机页配置为接收消息。 为此，可以向 `displayDialogAsync` 的原始调用添加回调参数。 回叫会为 `DialogMessageReceived` 事件分配处理程序。 示例如下。

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
>
> - Office 将 [AsyncResult ](/javascript/api/office/office.asyncresult) 对象传递给回叫。 表示尝试打开对话框的结果， 不表示对话框中任何事件的结果。 若要详细了解此区别，请参阅[处理错误和事件](dialog-handle-errors-events.md)。
> - `asyncResult` 的 `value` 属性设置为 [Dialog](/javascript/api/office/office.dialog) 对象，此对象位于主机页（而不是对话框的执行上下文）中。
> - `processMessage` 是用于处理事件的函数。可以根据需要任意命名。
> - `dialog` 变量的声明范围比回调更广，因为 `processMessage` 中也会引用此变量。

下面是一个非常简单的示例，展示了 `DialogMessageReceived` 事件的处理程序。

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
>
> - Office 将 `arg` 对象传递给处理程序。 其 `message` 属性是对话框中的 调用 `messageParent` 发送的字符串。 本示例中，它是 Microsoft 帐户或 Google 等服务中用户配置文件的字符串化表示形式，因此使用 反初始化回 对象 `JSON.parse`。
> - 未 `showUserName` 显示实现。 它可能在任务窗格上显示定制的欢迎消息。

在用户完成与对话框的交互后，消息处理程序应关闭对话框，如下面的示例所示。

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
>
> - `dialog` 对象必须是 `displayDialogAsync` 调用返回的对象。
> - `dialog.close` 调用指示 Office 立即关闭对话框。

有关使用这些技术的示例加载项，请参阅 [Office 加载项对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。

如果外接程序在收到消息后需要打开任务窗格的其他页面，可以使用 `window.location.replace` 方法（或 `window.location.href`）作为处理程序的最后一行。示例如下。

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

有关具有此用途的加载项示例，请参阅[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)（在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表）示例。

### <a name="conditional-messaging"></a>条件消息

由于可以从对话框发送多个 `messageParent` 调用，但在主机页中只有一个 `DialogMessageReceived` 事件处理程序，因此处理程序必须使用条件逻辑来区分不同的消息。 例如，如果对话框提示用户登录标识提供程序（如 Microsoft 帐户或 Google），则它会以消息身份发送用户配置文件。 如果身份验证失败，对话框将向主机页发送错误信息，如以下示例所示。

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
>
> - `loginSuccess` 变量通过读取标识提供程序返回的 HTTP 响应进行初始化。
> - 未显示 `getProfile` 和 `getError` 函数的实现。这两个函数均从查询参数或 HTTP 响应的正文获取数据。
> - 根据登录是否成功，发送不同类型的匿名对象。两者都有 `messageType` 属性。不同之处在于，一个有 `profile` 属性，另一个有 `error` 属性。

主机页中的处理程序代码使用 `messageType` 属性的值设置分支，如下面的示例所示。请注意，`showUserName` 函数的用法与之前的示例相同，`showNotification` 函数在主机页的 UI 中显示错误。

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
> 本文 `showNotification` 提供的示例代码中未显示实现。 有关如何在外接程序中实施此函数的示例，请参阅 [Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。

### <a name="cross-domain-messaging-to-the-host-runtime"></a>到主机运行时的跨域消息传送

对话框或父 JavaScript 运行时 (在任务窗格中或托管函数文件) 的无 UI 运行时中，在对话框打开后，可能会导航离开外接程序的域。 如果发生以上任 `messageParent` 一情况，则调用 将失败，除非你的代码指定父运行时的域。 为此，将 [DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) 参数添加到 的调用中 `messageParent`。 此对象具有 `targetOrigin` 一个属性，该属性指定邮件应发送到的域。 如果未使用 参数，则Office假定目标与对话框当前托管的域相同。

> [!NOTE]
> 使用 `messageParent` 发送跨域邮件需要 [Dialog Origin 1.1 要求集](../reference/requirement-sets/dialog-origin-requirement-sets.md)。 在`DialogMessageOptions`不支持要求集的早期版本Office参数将被忽略，因此，如果传递方法，此方法的行为将不受影响。

下面是使用 发送 `messageParent` 跨域邮件的示例。

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "https://resource.contoso.com" });
```

> [!NOTE]
> 参数 `DialogMessageOptions` 大约在 2021 年 7 月 19 日发布。 在此日期后大约 30 天内，在 Office web 版 `messageParent` `DialogMessageOptions` 中，首次调用不带 参数且父级与对话框不同的域时，将提示用户批准向目标域发送数据。 如果用户批准，用户的答案将缓存 24 小时。 在此期间，如果使用相同的目标域调用 `messageParent` 用户，则系统不会再次提示用户。

如果邮件不包含敏感数据， `targetOrigin` 可以将 设置为"\*"，以允许发送到任何域。 示例如下。

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "*" });
```

> [!TIP]
> `messageParent`在 `DialogMessageOptions` 2021 年中，参数已作为必需参数添加到方法中。 使用 方法发送跨域邮件的旧版加载项在更新为使用新参数之前不再有效。 在更新外接程序之前，仅在 *Office for Windows* 上，用户和系统管理员可以通过使用注册表设置指定受信任的域 () ，使这些外接程序继续工作： **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**。 为此，请创建一个扩展`.reg`名文件，将其保存到Windows计算机，然后双击它以运行该文件。 以下是此类文件的内容示例。
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="pass-information-to-the-dialog-box"></a>向对话框传递信息

加载项可以使用 [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) 将消息[](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)从主机页发送到对话框。

### <a name="use-messagechild-from-the-host-page"></a>从 `messageChild()` 主机页使用

调用对话框 API Office对话框时，将返回 [Dialog](/javascript/api/office/office.dialog) 对象。 它应分配给比 [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) 方法范围更大的变量，因为对象将被其他方法引用。 示例如下。

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

此 `Dialog` 对象具有 [messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) 方法，该方法将任何字符串（包括字符串化数据）发送到对话框。 这将在对话框中 `DialogParentMessageReceived` 引发事件。 代码应处理此事件，如下一节所示。

请考虑以下方案：对话框的 UI 与当前活动的工作表相关，并且该工作表相对于其他工作表的位置。 在下面的示例中，将 `sheetPropertiesChanged` Excel工作表属性发送到对话框。 在这种情况下，当前工作表名为"My Sheet"，它是工作簿中的第二个工作表。 数据封装在对象中并字符串化，以便可以传递给 `messageChild`。

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>处理对话框中的 DialogParentMessageReceived

在对话框的 JavaScript 中，使用 `DialogParentMessageReceived` [UI.addHandlerAsync](/javascript/api/office/office.ui#office-office-ui-addhandlerasync-member(1)) 方法为事件注册处理程序。 这通常在 [Office.onReady 或 Office.initialize](initialize-add-in.md) 方法中完成，如下所示。  (下面是一个更可靠的示例。) 

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

然后，定义 `onMessageFromParent` 处理程序。 以下代码继续上一节中的示例。 请注意Office参数传递给处理程序`message`，并且 argument 对象的 属性包含主机页中的字符串。 本示例将消息重新转换到对象，jQuery 用于设置对话框的顶部标题，以匹配新的工作表名称。

```javascript
function onMessageFromParent(arg) {
    var messageFromParent = JSON.parse(arg.message);
    $('h1').text(messageFromParent.name);
}
```

最佳做法是验证处理程序是否正确注册。 为此，可以将回调传递给 `addHandlerAsync` 方法。 此操作在注册处理程序的尝试完成时运行。 如果未成功注册处理程序，请使用处理程序记录或显示错误。 示例如下。 请注意， `reportError` 这是一个记录或显示错误的函数（此处未定义）。

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a>从父页面到对话框的条件消息

由于可以从主机页`messageChild``DialogParentMessageReceived`进行多次调用，但在事件的对话框中只有一个处理程序，因此处理程序必须使用条件逻辑来区分不同的消息。 您可以以与对话框向主机页发送邮件时如何构造条件消息的方式完全一样的方式完成此操作，如条件邮件 [中所述](#conditional-messaging)。

> [!NOTE]
> 在某些情况下， `messageChild` API（即 [DialogApi 1.2](../reference/requirement-sets/dialog-api-requirement-sets.md) 要求集的一部分）可能不受支持。 有关父级到对话框消息传递的一些替代方法，在将邮件从对话框的主机页传递到对话框 [的替代方法中进行了介绍](parent-to-dialog.md)。

> [!IMPORTANT]
> [DialogApi 1.2](../reference/requirement-sets/dialog-api-requirement-sets.md) 要求集无法指定在外接程序清单的 **"** 要求"部分。 必须在运行时使用 方法检查对 DialogApi 1.2 `isSetSupported` 的支持，如运行时检查方法和要求集支持 [中所述](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support)。 对清单要求的支持正在开发中。

### <a name="cross-domain-messaging-to-the-dialog-runtime"></a>到对话框运行时的跨域消息传送

对话框或父 JavaScript 运行时 (在任务窗格中或托管函数文件) 的无 UI 运行时中，在对话框打开后，可能会导航离开外接程序的域。 如果发生上述 `messageChild` 任一情况，则调用 将失败，除非你的代码指定对话框运行时的域。 为此，将 [DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) 参数添加到 的调用中 `messageChild`。 此对象具有 `targetOrigin` 一个属性，该属性指定邮件应发送到的域。 如果未使用 参数，则Office假定目标与父运行时当前托管的域相同。 

> [!NOTE]
> 使用 `messageChild` 发送跨域邮件需要 [Dialog Origin 1.1 要求集](../reference/requirement-sets/dialog-origin-requirement-sets.md)。 在`DialogMessageOptions`不支持要求集的早期版本Office参数将被忽略，因此，如果传递方法，此方法的行为将不受影响。

下面是使用 发送 `messageChild` 跨域邮件的示例。

```js
dialog.messageChild(messageToDialog, { targetOrigin: "https://resource.contoso.com" });
```

如果邮件不包含敏感数据， `targetOrigin` 可以将 设置为"\*"，以允许 *发送到任何域* 。 示例如下。

```js
dialog.messageChild(messageToDialog, { targetOrigin: "*" });
```

由于托管对话框的 JavaScript 运行时无法访问清单的 **AppDomains**  `DialogParentMessageReceived` 部分，因此确定邮件来自的域是否受信任，因此必须使用处理程序来确定这一点。 传递给处理程序的对象包含当前作为属性托管在父级中的 `origin` 域。 下面是如何使用 属性的示例。

```javascript
function onMessageFromParent(arg) {
    if (arg.origin === "https://addin.fabrikam.com") {
        // process message
    } else {
        dialog.close();
        showNotification("Messages from " + arg.origin + " are not accepted.");
    }
}
```

例如，您的代码可以使用 [Office.onReady 或 Office.initialize](initialize-add-in.md) 方法将受信任域的数组存储在全局变量中。 然后 `arg.origin` ，可以在处理程序中针对该列表检查该属性。

> [!TIP]
> `messageChild`在 `DialogMessageOptions` 2021 年中，参数已作为必需参数添加到方法中。 使用 方法发送跨域邮件的旧版加载项在更新为使用新参数之前不再有效。 在更新外接程序之前，仅在 *Office for Windows* 上，用户和系统管理员可以通过使用注册表设置指定受信任的域 () ，使这些外接程序继续工作： **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**。 为此，请创建一个扩展`.reg`名文件，将其保存到Windows计算机，然后双击它以运行该文件。 以下是此类文件的内容示例。
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="close-the-dialog-box"></a>关闭对话框

可以在对话框中实现一个用于关闭对话框的按钮。为此，该按钮的单击事件处理程序应使用 `messageParent` 通知主机页该按钮已获得单击。示例如下。

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

`DialogMessageReceived` 的主机页处理程序将调用 `dialog.close`，如以下示例所示。 （请参阅前面的示例，其中展示了 `dialog` 对象的初始化方式。）

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

即使你没有自己的关闭对话框 UI，最终用户也可以通过选择右上角的 **X** 关闭对话框。 此操作将触发 `DialogEventReceived` 事件。 如果主机窗格需要知道此事件何时发生，应为此事件声明一个处理程序。 有关详细信息，请参阅[对话框中的错误和事件](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box)部分。

## <a name="advanced-topics-and-special-scenarios"></a>高级主题和特殊情景

### <a name="use-the-dialog-api-to-show-a-video"></a>使用对话框 API 显示视频

参见“[使用 Office 对话框显示视频](dialog-video.md)”。

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a>在身份验证流中使用对话框 API

请参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。

### <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>将 Office 对话框 API 与单页应用程序和客户端路由一同使用

使用 Office 对话框 API 时，需要小心处理 SPA 和客户端路由。 请参阅“[在 SPA 中使用 Office 对话框 API 的最佳做法](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa)”。

### <a name="error-and-event-handling"></a>错误和事件处理

参见“[处理 Office 对话框中的错误和事件](dialog-handle-errors-events.md)。

## <a name="next-steps"></a>后续步骤

在“[Office 对话框 API 最佳做法和规则](dialog-best-practices.md)”中了解 Office 对话框 API 的陷阱和最佳做法。

## <a name="samples"></a>示例

以下所有示例都使用 `displayDialogAsync`。 一些服务器基于 NodeJS，其他服务器具有基于 ASP.NET/IIS 的服务器，但无论外接程序的服务器端如何实现，使用此方法的逻辑都是相同的。

**基础知识：**

- [Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [培训内容/构建外接程序 (示例) ](https://github.com/OfficeDev/TrainingContent/tree/2db14a16774e1539a3eebae7dada4798142b8493/OfficeAddin)

**更复杂的示例：**

- [Office Microsoft 外接程序 Graph ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Office 加载项 Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
- [Office 加载项 NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)
- [Office外接程序 ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)
- [Office加载项 SAAS 盈利示例](https://github.com/OfficeDev/office-add-in-saas-monetization-sample)
- [Outlook Microsoft 外接程序 Graph ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Outlook加载项 SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)
- [Outlook加载项令牌查看器](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Outlook加载项可操作邮件](https://github.com/OfficeDev/Outlook-Add-In-Actionable-Message)
- [Outlook外接程序共享到OneDrive](https://github.com/OfficeDev/Outlook-Add-in-Sharing-to-OneDrive)
- [PowerPoint Microsoft 外接程序 Graph ASPNET 插入图](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Excel共享运行时方案](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario)
- [Excel外接程序 ASPNET QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Word 外接程序 JS 修订](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word 外接程序 JS SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
- [Word 外接程序 AngularJS 客户端 OAuth](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)
- [Office 外接程序 Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Office加载项 OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io)
- [Office外接程序 UX 设计模式代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
