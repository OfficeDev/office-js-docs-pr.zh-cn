---
title: 授权 Microsoft Graph加载项Office Microsoft 外接程序
description: 了解如何通过加载项Graph Microsoft Office Microsoft 外接程序。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8166b7a71767abd0456662dbe8573f59bb2c7e82
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743581"
---
# <a name="authorize-to-microsoft-graph-from-an-office-add-in"></a>授权 Microsoft Graph加载项Office Microsoft 外接程序

加载项可以通过从加载项获取 Microsoft Graph访问令牌，获取对 Microsoft Graph访问Microsoft 标识平台。 像在其他 Web 应用程序中一样使用授权代码流或隐式流，但有一个例外：Microsoft 标识平台 不允许其登录页在 iframe 中打开。 当 Office 加载项在 *Office 网页版* 中运行时，任务窗格是一个 iFrame。 这意味着你需要使用登录对话框 API 在对话框中打开Office页面。 这将影响你使用身份验证和授权帮助程序库的方式。 有关详细信息，请参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。

> [!NOTE]
> 如果要实现 SSO 并计划访问 Microsoft Graph，请参阅使用 [SSO Graph Microsoft 授权](authorize-to-microsoft-graph.md)。

有关使用应用程序对身份验证进行编程Microsoft 标识平台，请参阅Microsoft 标识平台[文档](/azure/active-directory/develop)。 你将在该文档集内找到教程和指南，以及指向相关示例的链接。 同样，您可能需要调整示例代码以在 Office 对话框中运行，以考虑在任务窗格的单独进程中运行的 Office 对话框。

代码获取 Microsoft Graph 的访问令牌后，它会将访问令牌从对话框传递给任务窗格，或者将令牌存储在数据库中并指示任务窗格令牌可用。  (有关详细信息，请参阅使用 Office [对话框 API](auth-with-office-dialog-api.md) 进行身份验证。任务窗格中的 ) Code 从 Microsoft Graph 请求数据，并包括这些请求中的令牌。 有关调用 Microsoft Graph和 Microsoft Graph SDK 的信息，请参阅 [Microsoft Graph文档](/graph/)。

## <a name="recommended-libraries-and-samples"></a>推荐的库和示例

我们建议您在访问 Microsoft Graph 时使用以下库。

- 对于使用服务器端并采用基于网络的框架（如 .NET Core 或 ASP.NET）的加载项，请使用 [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation)。
- 对于使用基于 NodeJS 的服务器端的加载项, 请使用[Passport Azure AD](https://github.com/AzureAD/passport-azure-ad)。
- 对于使用隐式流的加载项，请使用[msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki)。

有关使用 Microsoft 标识平台 (以前称为 "AAD v. 2.0") 的推荐库的详细信息，请参阅[Microsoft 标识平台身份验证库](/azure/active-directory/develop/reference-v2-libraries)。

以下示例获取 Microsoft Graph加载项Office数据。

- [Office 加载项 Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Outlook 加载项 Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Office 加载项 Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
