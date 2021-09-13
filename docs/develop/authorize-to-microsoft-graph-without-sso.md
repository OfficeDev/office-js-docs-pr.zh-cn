---
title: 在不使用 SSO 的情况下对 Microsoft Graph 授权
description: 了解如何在不使用 SSO 的情况下对 Microsoft Graph 授权
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 4f96c65fcc3c90a616f43189e1facebdbf8e9a8c
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149141"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>在不使用 SSO 的情况下对 Microsoft Graph 授权

加载项可以通过从 Azure AD Graph 获取对 Microsoft Graph 的访问令牌Azure Active Directory (Microsoft) 。 像在其他 Web 应用程序中一样使用授权代码流或隐式流，但有一个例外：Azure AD 不允许其登录页在 iframe 中打开。 当 Office 加载项在 *Office 网页版* 中运行时，任务窗格是一个 iFrame。 这意味着你需要在通过对话框 API 打开的对话框中打开 Azure AD Office屏幕。 这将影响你使用身份验证和授权帮助程序库的方式。 有关详细信息，请参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。

有关使用 Azure AD 编程身份验证的信息，请从[Microsoft 标识平台 (v2.0) 概述](/azure/active-directory/develop/v2-overview)开始，你将在该文档集内找到教程和指南，以及指向相关示例的链接。 另外可能需要调整示例中的代码以在 Office 对话框中运行, 以考虑该 Office 对话框在与任务窗格不同的进程中运行的情况。

代码获取 Microsoft Graph 的访问令牌后，它会将访问令牌从对话框传递给任务窗格，或者将令牌存储在数据库中并指示任务窗格令牌可用。  (请参阅使用 Office 对话框[API](auth-with-office-dialog-api.md)进行身份验证了解详细信息。任务窗格中的 ) Code 从 Microsoft Graph 请求数据，并包括这些请求中的令牌。 有关调用 Microsoft Graph 和 Microsoft Graph SDK 的信息，请参阅 Microsoft Graph[文档](/graph/)。

## <a name="recommended-libraries-and-samples"></a>推荐的库和示例

建议您在不使用 SSO 的情况下访问 Microsoft Graph以下库。

- 对于使用服务器端并采用基于网络的框架（如 .NET Core 或 ASP.NET）的加载项，请使用 [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation)。
- 对于使用基于 NodeJS 的服务器端的加载项, 请使用[Passport Azure AD](https://github.com/AzureAD/passport-azure-ad)。
- 对于使用隐式流的加载项，请使用[msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki)。

有关使用 Microsoft 标识平台 (以前称为 "AAD v. 2.0") 的推荐库的详细信息，请参阅[Microsoft 标识平台身份验证库](/azure/active-directory/develop/reference-v2-libraries)。

以下示例获取 Microsoft Graph加载项Office数据。

- [Office 加载项 Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Outlook 加载项 Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Office 加载项 Microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React)
