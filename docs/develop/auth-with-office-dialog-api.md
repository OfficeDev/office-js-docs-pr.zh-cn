---
title: 使用 Office 对话框 API 进行身份验证和授权
description: 了解如何使用 Office 对话框 API 使用户能够登录到 Google、Facebook、Microsoft 365 以及受 Microsoft 标识平台保护的其他服务。
ms.date: 01/25/2022
ms.localizationpriority: high
ms.openlocfilehash: 90a8bed04a5f563de1bdbb509def39d96c732b11
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/02/2022
ms.locfileid: "62320184"
---
# <a name="authenticate-and-authorize-with-the-office-dialog-api"></a>使用 Office 对话框 API 进行身份验证和授权

始终使用 Office 对话框 API 通过 Office 加载项对用户进行身份验证和授权。 如果在无法使用单一登录 (SSO) 时实现回退身份验证，则还应使用 Office 对话框 API。

许多身份验证机构（也称为安全令牌服务 (STS)）会阻止其登录页面在 Iframe 中打开。 这包括 Google、Facebook 以及由 Microsoft 标识平台（以前称为 Azure AD V 2.0）保护的服务，例如 Microsoft 帐户、Microsoft 365 教育或工作帐户以及其他常用帐户。 这会导致 Office 加载项出现问题，因为当此加载项在 **Office 网页版** 上运行时，任务窗格是一个 Iframe。 如果加载项可以打开完全独立的浏览器实例,则加载项的用户只能登录到其中一个服务。 这就是为什么 Office 提供 [Office 对话框 API](dialog-api-in-office-add-ins.md)（尤其是[displayDialogAsync](/javascript/api/office/office.ui) 方法）的原因。

> [!NOTE]
> 本文假设你熟悉[在 Office 加载项中使用 Office 对话框 API](dialog-api-in-office-add-ins.md)。

使用此 API 打开的对话框具有以下特征。

- 它是[非模态](https://en.wikipedia.org/wiki/Dialog_box)。
- 它是完全独立于任务窗格的浏览器实例，这意味着：
  - 它拥有自己的 JavaScript 运行时环境和窗口对象及全局变量。
  - 没有与任务窗格共享的执行环境。
  - 它没有与任务窗格共享相同的会话存储（[Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) 属性）。
- 对话框中打开的第一个页面必须与任务窗格位于同一域中，包括协议、子域和端口（如果有）。
- 该对话框可以使用 [messageParent](/javascript/api/office/office.ui#messageParent_message__messageOptions_) 方法将信息发送回任务窗格。 建议仅从与任务窗格托管在同一域中的页面调用此方法，包括协议、子域、端口。 否则，调用方法和处理消息的方式会出现复杂情况。 有关详细信息，请参阅[向主机运行时间跨域消息传递](dialog-api-in-office-add-ins.md#cross-domain-messaging-to-the-host-runtime)。

默认情况下，对话框在新的 Web 视图控件打开，而不是 iframe 中打开。 这可确保它可以打开标识提供程序的登录页面。 如下所示，该 Office 对话框的特征对如何使用身份验证或授权库（例如 Microsoft 身份验证库 (MSAL) 和护照）有一定影响。

> [!NOTE]
> 可通过以下方式配置要在浮动 iframe 中打开的对话框：只需在对 `displayInIframe: true` 的调用中传递 `displayDialogAsync` 选项。 使用对话框 API 登录时, 请 *不要* 这样做。

## <a name="authentication-flow-with-the-office-dialog-box"></a>使用 Office 对话框的身份验证流程

下面是一个典型的身份验证流程。

![显示任务窗格与对话框浏览器进程的关系的图示。](../images/taskpane-dialog-processes.gif)

1. 对话框中打开的第一个页面托管在加载项域（即与任务窗格相同的域）中的一个页面（或其他资源）。 此页面可以显示 UI，提示用户“请稍候，正在重定向到可以登录 *NAME-OF-PROVIDER* 的页面。” 此页面中的代码使用传递给对话框的信息（如[向对话框传递信息](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)中所述）构造身份提供程序的登录页 URL，或者硬编码到加载项的配置文件中，例如 web.config 文件。
2. 然后，对话框窗口重定向到登录页。 URL 包含一个查询参数，用于告知身份提供程序在用户登录后将对话框窗口重定向到特定页面。 在本文中，我们将此页面称为 **redirectPage.html**。 在此页上，登录尝试的结果可以通过调用 `messageParent` 传递到任务窗格。 *建议此页与主机窗口位于同一域中*。
3. 身份提供程序的服务处理来自对话框窗口的传入 GET 请求。 如果用户已经登录，它会立即将窗口重定向到 **redirectPage.html**，并包括用户数据作为查询参数。 如果用户尚未登录，提供程序的登录页会显示在窗口中，以便用户登录。 对于大多数提供程序，如果用户无法成功登录，提供程序会在对话框窗口中显示错误页面，而不会重定向到 **redirectPage.html**。 用户必须通过选择右上角的 **X** 来关闭窗口。 如果用户成功登录，则对话框窗口会重定向到 **redirectPage.html**，并包括用户数据会作为查询参数。
4. 当 **redirectPage.html** 页面打开时，它会调用 `messageParent` 向任务窗格页报告登录是否成功，而且还会视情况报告用户数据或错误数据。 其他可能的消息包括传递访问令牌或告知任务窗格信息位于存储中。
5. `DialogMessageReceived` 事件在任务窗格页中触发，其处理程序关闭对话框窗口，并可能对消息进行进一步处理。

#### <a name="support-multiple-identity-providers"></a>支持多个标识提供程序

如果加载项允许用户选择提供程序（如 Microsoft 帐户、Google 或 Facebook），你需要使用本地第一个页面（见前一部分），为用户提供用于选择提供程序的 UI。用户的选择会触发登录 URL 的构建并重定向到该 URL。

#### <a name="authorization-of-the-add-in-to-an-external-resource"></a>加载项中对外部资源的授权

在现代网络中，用户和 Web 应用程序是安全主体。应用程序拥有自己的身份以及对联机资源（如 Microsoft 365、Google Plus、Facebook 或 LinkedIn）拥有相应权限。在部署前，需要先向资源提供程序注册应用程序。注册内容包括：

- 应用程序访所需的权限的列表。
- 当应用访问服务时，资源服务应向其返回访问令牌的 URL。  

如果用户在应用中调用访问资源服务中用户数据的函数，系统会先提示用户登录相应服务，再提示用户向应用授予访问用户资源所需的权限。然后，服务将登录窗口重定向到先前注册的 URL，并传递访问令牌。应用使用访问令牌访问用户资源。

可以使用 Office 对话框 API 来管理此过程，具体方法是使用与用户登录流程类似的流程。唯一的区别是：

- 如果用户先前未向应用程序授予所需的权限，则登录后会在对话框中看到这样做的提示。
- 对话框窗口中的代码使用 `messageParent` 发送字符串化访问令牌，或将访问令牌存储在主机窗口可以检索到的位置（并使用 `messageParent` 告知主机窗口令牌可用），从而将访问令牌发送到主机窗口。 令牌具有时间限制，但在持续期间，主机窗口可以使用它直接访问用户资源，而无需进一步提示。

[示例](#samples)中列出了使用 Office 对话框 API 来实现此目的的一些身份验证示例加载项。

## <a name="use-authentication-libraries-with-the-dialog-box"></a>将身份验证库与对话框结合使用

Office 对话框和任务窗格在不同的浏览器、JavaScript 运行时实例中运行，你必须使用多个身份验证/授权库，必须与在同一窗口中进行身份验证和授权时使用它们的方式不同。 以下部分介绍了可以使用和不能使用这些库的方法。

### <a name="you-usually-cannot-use-the-librarys-internal-cache-to-store-tokens"></a>通常无法使用库的内部缓存来存储令牌

通常，身份验证相关库提供内存缓存来存储访问令牌。 如果对资源提供程序（例如 Google、Microsoft Graph、Facebook 等）进行了后续调用，则库将首先检查以确定其缓存中的令牌是否已过期。 如果未过期，库将返回缓存的令牌，而不是为新令牌再执行一次到 STS 的往返行程。 但 Office 加载项中无法使用此模式。由于登录发生在 Office 对话框的浏览器实例中，因此令牌缓存处于该实例中。

与此非常密切相关的是，库通常会同时提供用于获取令牌的交互式和“无提示”方法。 如果你既可以进行身份验证，也可以在同一浏览器实例中对资源进行数据调用，则代码会调用无提示方法来获取令牌，然后马上将该令牌添加到数据调用。 无提示方法会检查缓存是否有中未过期的令牌，并将其返回（如果有）。 否则，无提示方法将调用重定向到 STS 登录的交互式方法。 登录完成后，交互式方法将返回令牌，但同时会将其缓存在内存中。 但是，在使用 Office 对话框 API 时，对资源的数据调用（它将调用无提示方法）位于任务窗格的浏览器实例中。 库的令牌缓存在该实例中不存在。

或者，加载项的对话框浏览器实例可以直接调用库的交互式方法。 该方法返回令牌时，代码必须将令牌显式存储在任务窗格的浏览器可检索到的位置，例如本地存储\*或服务器端数据库。 另一种选择是使用 `messageParent` 方法将令牌传递到任务窗格。 仅当交互式方法将访问令牌存储在代码可以读取的位置时，才可以使用此替代选项。 有时，库的交互式方法设计为将令牌存储到代码无法访问的对象的私有属性中。

> [!NOTE]
> \*有一个 bug 将影响你的令牌处理策略。 如果加载项正使用 Safari 或 Edge 浏览器在 **Office 网页版** 上运行，则对话框和任务窗格不共享同一本地存储，因此该存储无法在它们之间通信。

### <a name="you-usually-cannot-use-the-librarys-auth-context-object"></a>通常无法使用库的“身份验证上下文”对象

通常情况下，与身份验证相关的库有一种方法，该方法既能够以交互方式获取令牌，也会创建方法返回的“身份验证上下文”对象。 令牌是对象的一个属性（可能是私有属性，并且无法直接从代码中访问）。 该对象具有从资源中获取数据的方法。 这些方法将令牌包括在其对资源提供程序（例如 Google、Microsoft Graph、Facebook 等）进行的 HTTP 请求中。

这些身份验证上下文对象和创建它们的方法在 Office 加载项中不可用。由于登录发生在 Office 对话框的浏览器实例中，因此必须在该处创建对象。 但对资源的数据调用位于任务窗格浏览器实例中，因此无法将对象从一个实例获取到另一个实例。 例如，无法通过 `messageParent` 传递对象，因为 `messageParent` 只能传递字符串值。 无法可靠地将包含方法的 JavaScript 对象字符串化。

### <a name="how-you-can-use-libraries-with-the-office-dialog-api"></a>如何将库与 Office 对话框 API 结合使用

大多数库提供了更低抽象级别的 API 作为单一“身份验证相关”对象的补充（或取代这些对象），可让代码创建不太单一的整体帮助程序对象。 例如，[MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation) v. 3. x.x 有一个用于构造登录 URL 的 API，以及另一个用于构造 AuthResult 对象的 API，该对象在代码可访问的属性中包含访问令牌。 有关 Office 加载项中的 MSAL.NET 的示例，请参阅: [Office 加载项 Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET) 和 [Outlook 加载项 Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)。 有关在加载项中使用 [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js) 的示例，请参阅 [Office 加载项 Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)。

有关身份验证和授权库的详细信息，请参阅 [Microsoft Graph：推荐的库](authorize-to-microsoft-graph-without-sso.md#recommended-libraries-and-samples)和[其他外部服务：库](auth-external-add-ins.md#libraries)。

## <a name="samples"></a>示例

- [Office 加载项 Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)：一个基于 ASP.NET 的加载项（Excel、Word 或 PowerPoint），它使用 MSAL.NET 库和授权代码流进行登录并获取 Microsoft Graph 数据的访问令牌。
- [Outlook 加载项 Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)：与上面的加载项一样，但 Office 应用程序为 Outlook。
- [Office 加载项 Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)：一个基于 NodeJS 的加载项（Excel、Word 或 PowerPoint），它使用 msal.js 库和隐式流进行登录并获取 Microsoft Graph 数据的访问令牌。

## <a name="see-also"></a>另请参阅

- [在 Office 加载项中授权外部服务](auth-external-add-ins.md)
- [在 Office 加载项中使用对话框 API](dialog-api-in-office-add-ins.md)
