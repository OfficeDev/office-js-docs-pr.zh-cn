---
title: 在 Office 加载项中使用 Office 对话框 API
description: 了解在 Office 加载项中创建对话框的基础知识。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 947b08575d100c639a440c1ca25d45199b4507ad
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/05/2022
ms.locfileid: "68466955"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>在 Office 加载项中使用 Office 对话框 API

可以在 Office 加载项中使用 [Office 对话框 API](/javascript/api/office/office.ui) 打开对话框。 本文提供了有关如何在 Office 加载项中使用对话框 API 的指南。

> [!NOTE]
> 若要了解对话框 API 目前的受支持情况，请参阅[对话框 API 要求集](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)。 Excel、PowerPoint 和 Word 目前支持该对话框 API。 Outlook 支持包含在各种邮箱要求集中&mdash;，请参阅 API 参考，了解更多详细信息。

对话框 API 的主要应用场景是为 Google、Facebook 或 Microsoft Graph 等资源启用身份验证。 有关详细信息，请在熟悉本文 *之后*，参阅 [使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。

不妨通过任务窗格/内容加载项/[加载项命令](../design/add-in-commands.md)打开对话框，以便执行下列操作：

- 显示无法直接在任务窗格中打开的登录页。
- 为加载项中的某些任务提供更多屏幕空间，或甚至整个屏幕。
- 托管在任务窗格中显得太小的视频。

> [!NOTE]
> 由于不赞成重叠 UI 元素，因此除非应用场景需要，否则请勿从任务窗格打开对话框。 考虑如何使用任务窗格区域时，请注意任务窗格中可以有选项卡。 有关选项卡式任务窗格的示例，请参阅 [Excel 加载项 JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 示例。

下图展示了对话框示例。

![Word 前面显示 3 个登录选项的对话框。](../images/auth-o-dialog-open.png)

请注意，对话框总是在屏幕的中心打开。 用户可以移动并重设对话框的大小。 窗口是 *非模式的*-如果有，用户可以继续与 Office 应用程序中的文档和任务窗格中的页面进行交互。

## <a name="open-a-dialog-box-from-a-host-page"></a>从主机页面打开对话框

Office JavaScript API 在 [Office.context.ui 命名空间](/javascript/api/office/office.ui)中包含一个 [Dialog](/javascript/api/office/office.dialog) 对象和两个函数。

为了打开对话框，代码（通常是任务窗格中的一页）调用 [displayDialogAsync](/javascript/api/office/office.ui) 方法，并将要打开的资源 URL 传递到此方法。 调用方法的页面称为“主机页”。 例如在任务窗格中的 index.html 页面上使用脚本调用此方法，随后 index.html 是打开此方法对话框的主机页。

对话框中打开的资源通常是页面，，但也可以是 MVC 应用中的控制器方法、路由、Web 服务方法或其他任何资源。 在本文中，“页面”或“网站”是指对话框中的资源。 下面的代码是一个简单的示例。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
>
> - URL 使用 HTTP **S** 协议。 对话框中加载的所有页面都必须要遵循此要求，而不仅仅是加载的第一个页面。
> - 对话框域与宿主页的域相同，宿主页可以是任务窗格中的页面，也可以是加载项命令的[函数文件](/javascript/api/manifest/functionfile)。 这要求：传递到 `displayDialogAsync` 方法的页面、控制器方法或其他资源必须与主机页位于相同的域。

> [!IMPORTANT]
> 对话框中打开的主机页面和资源必须具有相同的完整域。 如果尝试传递 `displayDialogAsync` 加载项域的子域，则不会起作用。 完整域（包括任何子域）必须匹配。

加载第一个页面（或其它资源）后，用户可使用链接或其它用户界面来导航至任何使用 HTTPS 的网站（或其他资源）。 还可以将第一个页面设计为直接重定向到另一个站点。

默认情况下，对话框的高度和宽度占设备屏幕的 80%。不过，你也可以设置不同的百分比，只需将配置对象传递给方法即可，如下面的示例所示。

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

有关实现这一点的样本加载项，请参阅 [Office 加载项 Dialog API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。 有关使用的 `displayDialogAsync`更多示例，请参阅 [示例](#samples)。

Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)

> [!NOTE]
> 只能从主机窗口打开一个对话框。 如果尝试再打开一个对话框，就会生成错误。 例如，如果用户从任务窗格打开对话框，则无法从任务窗格中其他页面打开第二个对话框。 不过，如果对话框是通过[加载项命令](../design/add-in-commands.md)打开，那么只要选择此命令，就会打开新 HTML 文件（但不可见）。 这会新建（不可见的）主机窗口，所以每个这样的窗口都可以启动自己的对话框。 有关详细信息，请参阅 [displayDialogAsync 返回的错误](dialog-handle-errors-events.md#errors-from-displaydialogasync)。

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a>利用 Office 网页版中的性能选项

`displayInIframe` 属性是配置对象中另一个可以传递到 `displayDialogAsync` 的属性。 如果将此属性设置为 `true`，且加载项在 Office 网页版打开的文档中运行，对话框就会以浮动 iframe（而不是独立窗口）的形式打开，从而加快对话框的打开速度。 示例如下。

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

默认值为 `false`，与完全省略此属性时相同。 如果加载项未在Office web 版中运行，则忽略该加载项`displayInIframe`。

> [!NOTE]
> 如果对话框会在任何时候重定向到无法在 iframe 中打开的页面，**则不应使用**`displayInIframe: true`。 例如，许多常用 Web 服务（如 Google 和 Microsoft 帐户）的登录页面无法在 iframe 中打开。

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>将信息从对话框发送到主机页

> [!NOTE]
>
> - 为了清楚起见，在本部分中，我们将消息调用主机 *页*，但严格地说，消息将转到任务窗格中的 [运行时](../testing/runtimes.md) (或托管函数 [文件](/javascript/api/manifest/functionfile) 的运行时) 。 只有在跨域消息传送的情况下，这种区别才很重要。 有关详细信息，请参阅[向主机运行时间跨域消息传递](#cross-domain-messaging-to-the-host-runtime)。
> - 除非在页面中加载了 Office JavaScript API 库，否则对话框无法与任务窗格中的主机页通信。  (与使用 Office JavaScript API 库的任何页面一样，该页面的脚本必须初始化加载项。 有关详细信息，请参阅 [“初始化 Office 加载项](initialize-add-in.md)) 

对话框中的代码使用 [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) 函数将字符串消息发送到主机页。 该字符串可以是字词、句子、XML Blob、字符串化 JSON，也可以是可序列化为字符串或强制转换为字符串的任何其他内容。 示例如下。

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true.toString());
}
```

> [!IMPORTANT]
>
> - 该 `messageParent` 函数是对话框中 *只能* 调用的两个 Office JS API 之一。
> - 可在对话框中调用的其他 JS API 是 `Office.context.requirements.isSetSupported`。 有关它的信息，请参阅 [指定 Office 应用程序和 API 要求](specify-office-hosts-and-api-requirements.md)。 但是，在对话框中，批量许可的永久Outlook 2016 (（即 MSI 版本) ）不支持此 API。

在下一个示例中，`googleProfile` 是用户 Google 配置文件的字符串化版本。

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

必须将主机页配置为接收消息。 为此，可以向 `displayDialogAsync` 的原始调用添加回调参数。 回叫会为 `DialogMessageReceived` 事件分配处理程序。 示例如下。

```js
let dialog;
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
> - The `processMessage` is the function that handles the event. You can give it any name you want.
> - `dialog` 变量的声明范围比回调更广，因为 `processMessage` 中也会引用此变量。

下面是一个非常简单的示例，展示了 `DialogMessageReceived` 事件的处理程序。

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
>
> - Office 将 `arg` 对象传递给处理程序。 它的 `message` 属性是由对话框中的调用发送的 `messageParent` 字符串。 在此示例中，它是 Microsoft 帐户或 Google 等服务中用户配置文件的字符串化表示形式，因此会反序列化回对象 `JSON.parse`。
> - `showUserName`未显示实现。 它可能在任务窗格上显示定制的欢迎消息。

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

If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example.

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

由于可以从对话框发送多个 `messageParent` 调用，但在主机页中只有一个 `DialogMessageReceived` 事件处理程序，因此处理程序必须使用条件逻辑来区分不同的消息。 例如，如果对话框提示用户登录到标识提供者（如 Microsoft 帐户或 Google），则会将用户的个人资料作为消息发送。 如果身份验证失败，对话框会将错误信息发送到主机页，如以下示例所示。

```js
if (loginSuccess) {
    const userProfile = getProfile();
    const messageObject = {messageType: "signinSuccess", profile: userProfile};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    const errorDetails = getError();
    const messageObject = {messageType: "signinFailure", error: errorDetails};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
>
> - `loginSuccess` 变量通过读取标识提供程序返回的 HTTP 响应进行初始化。
> - 不显示这些函数和`getError`函数的`getProfile`实现。 这两个函数均从查询参数或 HTTP 响应的正文获取数据。
> - Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.

The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
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
> 本文 `showNotification` 提供的示例代码中未显示此实现。 有关如何在外接程序中实施此函数的示例，请参阅 [Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。

### <a name="cross-domain-messaging-to-the-host-runtime"></a>到主机运行时的跨域消息传送

对话框打开后，对话框或父运行时可能会从外接程序的域中导航。 如果发生上述任一情况，则调用 `messageParent` 将失败，除非代码指定父运行时的域。 为此，请将 [DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) 参数添加到调用 。`messageParent` 此对象具有一个 `targetOrigin` 属性，该属性指定应向其发送消息的域。 如果未使用该参数，Office 假定目标与对话当前托管的域相同。

> [!NOTE]
> 用于 `messageParent` 发送跨域消息需要 [对话框源 1.1 要求集](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets)。 在 `DialogMessageOptions` 不支持要求集的较旧版本的 Office 上忽略该参数，因此，如果传递该参数，该方法的行为将不受影响。

下面是用于 `messageParent` 发送跨域消息的示例。

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "https://resource.contoso.com" });
```

> [!NOTE]
> 该 `DialogMessageOptions` 参数于 2021 年 7 月 19 日左右发布。 在该日期之后的大约 30 天内，在Office web 版中，第一次`messageParent`在没有参数的情况下`DialogMessageOptions`调用，并且父域与对话框不同，系统会提示用户批准将数据发送到目标域。 如果用户批准，则用户的答案将缓存 24 小时。 在此期间，当使用相同的目标域调用用户时 `messageParent` ，不会再次提示用户。

如果消息不包含敏感数据，则可以将消息设置为 `targetOrigin` “\*”，允许将其发送到任何域。 示例如下。

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "*" });
```

> [!TIP]
> 该 `DialogMessageOptions` 参数在 2021 年年中作为必需参数添加到 `messageParent` 该方法。 使用该方法发送跨域消息的较旧加载项在更新为使用新参数之前不再工作。 在加载项更新之前， *仅在 Office for Windows 上*，用户和系统管理员可以通过使用注册表设置指定受信任域 () 来使这些加载项继续工作： **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**。 为此，请创建具有扩展名的文件，将其保存到 Windows 计算机，然后双击该文件 `.reg` 以运行它。 下面是此类文件的内容示例。
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="pass-information-to-the-dialog-box"></a>向对话框传递信息

外接程序可以使用 [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) 将[消息从主机页](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)发送到对话框。

### <a name="use-messagechild-from-the-host-page"></a>从主机页使用`messageChild()`

调用 Office 对话 API 以打开对话框时，将返回 [Dialog](/javascript/api/office/office.dialog) 对象。 应将其分配给范围大于 [displayDialogAsync 方法的](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) 变量，因为其他方法将引用该对象。 示例如下。

```javascript
let dialog;
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

此 `Dialog` 对象具有 [messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) 方法，可将任何字符串（包括字符串化数据）发送到对话框。 这会在对话框中引发 `DialogParentMessageReceived` 事件。 代码应处理此事件，如下一部分所示。

假设对话框的 UI 与当前活动的工作表以及工作表相对于其他工作表的位置相关的场景。 在以下示例中， `sheetPropertiesChanged` 将 Excel 工作表属性发送到对话框。 在这种情况下，当前工作表名为“My Sheet”，它是工作簿中的第二个工作表。 数据封装在对象中并进行字符串化，以便可以传递到 `messageChild`该对象。

```javascript
function sheetPropertiesChanged() {
    const messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>在对话框中处理 DialogParentMessageReceived

在对话框的 JavaScript 中，使用 [UI.addHandlerAsync](/javascript/api/office/office.ui#office-office-ui-addhandlerasync-member(1)) 方法注册`DialogParentMessageReceived`事件的处理程序。 这通常在 [Office.onReady 或 Office.initialize 函](initialize-add-in.md)数中完成，如下所示。  (本文稍后将包含一个更可靠的示例。) 

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

然后，定义 `onMessageFromParent` 处理程序。 下面的代码延续了上一部分中的示例。 请注意，Office 将参数传递给处理程序， `message` 并且参数对象的属性包含主机页中的字符串。 在此示例中，消息将重新转换为对象，jQuery 用于设置对话框的顶部标题以匹配新的工作表名称。

```javascript
function onMessageFromParent(arg) {
    const messageFromParent = JSON.parse(arg.message);
    $('h1').text(messageFromParent.name);
}
```

最好验证处理程序是否已正确注册。 可以通过将回调传递给方法来 `addHandlerAsync` 执行此操作。 当注册处理程序的尝试完成时，此操作会运行。 如果处理程序未成功注册，请使用处理程序记录或显示错误。 示例如下。 请注意， `reportError` 此处未定义的函数记录或显示错误。

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

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a>从父页到对话框的条件消息传送

由于可以从主机页进行多次 `messageChild` 调用，但该事件的对话框 `DialogParentMessageReceived` 中只有一个处理程序，因此处理程序必须使用条件逻辑来区分不同的消息。 执行此操作的方式与在对话框向主机页发送消息时如何构建条件消息的方式完全并行，如 [条件消息传送](#conditional-messaging)中所述。

> [!NOTE]
> 在某些情况下， `messageChild` 可能不支持 API（这是 [DialogApi 1.2 要求集](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)的一部分）。 从 [主机页将消息传递到对话框的替代方法](parent-to-dialog.md)介绍了父到对话框消息的一些替代方法。

> [!IMPORTANT]
> 无法在加载项清单部分中 **\<Requirements\>** 指定 [DialogApi 1.2 要求集](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)。 必须使用 `isSetSupported` 运行时检查 [方法和要求集支持](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support)时所述的方法检查对 DialogApi 1.2 的支持。 正在开发对清单要求的支持。

### <a name="cross-domain-messaging-to-the-dialog-runtime"></a>跨域消息传送到对话框运行时

对话框打开后，对话框或父运行时可能会从外接程序的域中导航。 如果发生上述任一情况，则除非代码指定对话运行时的域，否则调 `messageChild` 用将失败。 为此，请将 [DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) 参数添加到调用 。`messageChild` 此对象具有一个 `targetOrigin` 属性，该属性指定应向其发送消息的域。 如果未使用该参数，Office 假定目标与父运行时当前托管的域相同。

> [!NOTE]
> 用于 `messageChild` 发送跨域消息需要 [对话框源 1.1 要求集](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets)。 在 `DialogMessageOptions` 不支持要求集的较旧版本的 Office 上忽略该参数，因此，如果传递该参数，该方法的行为将不受影响。

下面是用于 `messageChild` 发送跨域消息的示例。

```js
dialog.messageChild(messageToDialog, { targetOrigin: "https://resource.contoso.com" });
```

如果消息不包含敏感数据，则可以将消息设置为 `targetOrigin` “\*”，允许将其 *发送* 到任何域。 示例如下。

```js
dialog.messageChild(messageToDialog, { targetOrigin: "*" });
```

由于托管对话框的运行时无法访问 **\<AppDomains\>** 清单部分，从而确定 *消息来自的* 域是否受信任，因此必须使用 `DialogParentMessageReceived` 处理程序来确定这一点。 传递给处理程序的对象包含当前托管在父级中作为其 `origin` 属性的域。 下面是如何使用该属性的示例。

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

例如，代码可以使用 [Office.onReady 或 Office.initialize 函](initialize-add-in.md) 数将受信任域数组存储在全局变量中。 `arg.origin`然后，可以针对处理程序中的该列表检查该属性。

> [!TIP]
> 该 `DialogMessageOptions` 参数在 2021 年年中作为必需参数添加到 `messageChild` 该方法。 使用该方法发送跨域消息的较旧加载项在更新为使用新参数之前不再工作。 在加载项更新之前， *仅在 Office for Windows 上*，用户和系统管理员可以通过使用注册表设置指定受信任域 () 来使这些加载项继续工作： **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**。 为此，请创建具有扩展名的文件，将其保存到 Windows 计算机，然后双击该文件 `.reg` 以运行它。 下面是此类文件的内容示例。
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="close-the-dialog-box"></a>关闭对话框

You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example.

```js
function closeButtonClick() {
    const messageObject = {messageType: "dialogClosed"};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

`DialogMessageReceived` 的主机页处理程序将调用 `dialog.close`，如以下示例所示。 （请参阅前面的示例，其中展示了 `dialog` 对象的初始化方式。）

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
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

### <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>将 Office 对话框 API 与单页应用程序和客户端路由配合使用

使用 Office 对话框 API 时，需要小心处理 SPA 和客户端路由。 请参阅“[在 SPA 中使用 Office 对话框 API 的最佳做法](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa)”。

### <a name="error-and-event-handling"></a>错误和事件处理

参见“[处理 Office 对话框中的错误和事件](dialog-handle-errors-events.md)。

## <a name="next-steps"></a>后续步骤

在“[Office 对话框 API 最佳做法和规则](dialog-best-practices.md)”中了解 Office 对话框 API 的陷阱和最佳做法。

## <a name="samples"></a>示例

以下所有示例均使用 `displayDialogAsync`。 有些服务器基于 NodeJS，而另一些服务器具有 ASP.NET/IIS-based 服务器，但无论如何实现外接程序的服务器端，使用该方法的逻辑都是相同的。

**基础：**

- [Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [训练内容/生成加载项 (多个示例) ](https://github.com/OfficeDev/TrainingContent/tree/2db14a16774e1539a3eebae7dada4798142b8493/OfficeAddin)

**更复杂的示例：**

- [Office 加载项 Microsoft Graph ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Office 加载项 Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
- [Office 加载项 NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)
- [Office 加载项 ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)
- [Office 加载项 SAAS 盈利示例](https://github.com/OfficeDev/office-add-in-saas-monetization-sample)
- [Outlook 加载项 Microsoft Graph ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Outlook 外接程序 SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)
- [Outlook 外接程序令牌查看器](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Outlook 加载项可操作消息](https://github.com/OfficeDev/Outlook-Add-In-Actionable-Message)
- [Outlook 外接程序共享到 OneDrive](https://github.com/OfficeDev/Outlook-Add-in-Sharing-to-OneDrive)
- [PowerPoint 外接程序 Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Excel 共享运行时方案](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario)
- [Excel 加载项 ASPNET QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Word 加载项 JS 修订](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word 加载项 JS SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
- [Word 加载项 AngularJS 客户端 OAuth](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)
- [Office 外接程序 Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Office 加载项 OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io)
- [Office 加载项 UX 设计模式代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

** 另请参阅**

- [Office 加载项中的运行时](../testing/runtimes.md)