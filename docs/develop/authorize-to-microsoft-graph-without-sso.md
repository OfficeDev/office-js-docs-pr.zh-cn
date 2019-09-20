---
title: 在不使用 SSO 的情况下对 Microsoft Graph 授权
description: ''
ms.date: 08/07/2019
localization_priority: Priority
ms.openlocfilehash: 1d696783003fc475f98d5d1d37f60348aacec5eb
ms.sourcegitcommit: f781d7cfd980cd866d6d1d00c5b9d16c8a4b7f9b
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/20/2019
ms.locfileid: "37053310"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>在不使用 SSO 的情况下对 Microsoft Graph 授权

可通过从 Azure Active Directory (AAD) 获取 Graph 的访问令牌，获得加载项的 Microsoft Graph 数据的授权。 你可同在任何其他 Web 应用程序中一样，使用授权代码流或隐式流执行此操作，但存在一个例外：AAD 禁止其登录页在 iframe 中打开。 当 Office 加载项在 *Office 网页版*中运行时，任务窗格是一个 iframe。 这意味着你将需要在通过 Office 对话框 API 打开的对话框中打开 AAD 登录屏幕。 这将影响你使用身份验证和授权帮助程序库的方式。 有关详细信息，请参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。

要了解如何使用 AAD 对身份验证进行编程，首先请查看 [Microsoft 标识平台 (v2.0) 概述](/azure/active-directory/develop/v2-overview)。 该文档集中有很多教程和指南，还有相关示例的链接。 再次提醒一下：你可能需要调整示例中的代码以在 Office 对话框中运行, 以考虑该对话框在与任务窗格不同的进程中运行的情况。

代码获取 Microsoft Graph 的访问令牌后，可将访问令牌从对话框传递到任务窗格，或将令牌存储在数据库中, 并向任务窗格发出通知，表明该令牌可用。 （有关详细信息，请参阅[ Office 对话框 API 的身份验证](auth-with-office-dialog-api.md)）任务窗格中的代码请求从 Microsoft Graph 中获得数据, 并将令牌包含在这些请求中。 有关调用 Microsoft Graph 和 Microsoft Graph 的 SDK 的详细信息，请参阅[Microsoft Graph 文档](/graph/)。

## <a name="recommended-libraries-and-samples"></a>推荐的库和示例

建议在不使用 SSO 访问 Microsoft Graph 时使用下列库：

- 对于使用服务器端并采用基于网络的框架（如 .NET Core 或 ASP.NET）的加载项，请使用 [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation)。
- 对于使用基于 NodeJS 的服务器端的加载项, 请使用[Passport Azure AD](https://github.com/AzureAD/passport-azure-ad)。
- 对于使用隐式流的加载项，请使用[msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki)。

有关使用 Microsoft 标识平台 (以前称为 "AAD v. 2.0") 的推荐库的详细信息，请参阅[Microsoft 标识平台身份验证库](/azure/active-directory/develop/reference-v2-libraries)。

以下示例从 Office 加载项获取 Microsoft Graph 数据：

- [Office 加载项 Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Outlook 加载项 Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)

