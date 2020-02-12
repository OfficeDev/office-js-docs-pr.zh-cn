---
title: 在不使用 SSO 的情况下对 Microsoft Graph 授权
description: 了解如何在不使用 SSO 的情况下对 Microsoft Graph 授权
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 828779a766c41088435ff5fdfa693e1d9939c710
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41949658"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>在不使用 SSO 的情况下对 Microsoft Graph 授权

你的加载项可通过从 Azure Active Directory (AAD) 获取 Graph 的访问令牌，获得 Microsoft Graph 数据的授权。 可同在任何其他 Web 应用程序中一样，使用授权代码流或隐式流执行此操作，但存在一个例外：AAD 禁止其登录页在 iframe 中打开。 当 Office 加载项在 *Office 网页版*中运行时，任务窗格是一个 iFrame。 这意味着将需要在通过 Office 对话框 API 打开的对话框中打开 AAD 登录屏幕。 这将影响你使用身份验证和授权帮助程序库的方式。 有关详细信息，请参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。

有关使用 AAD 设置身份验证的信息，请先参阅[Microsoft 标识平台 (v2.0) 概览](/azure/active-directory/develop/v2-overview)，其中可找到此文档集中的教程和指南以及相关示例的链接。 另外可能需要调整示例中的代码以在 Office 对话框中运行, 以考虑该 Office 对话框在与任务窗格不同的进程中运行的情况。

代码获取 Microsoft Graph 的访问令牌后，可将访问令牌从对话框传递到任务窗格，或将令牌存储在数据库中, 并向任务窗格发出通知，表明该令牌可用。 （有关详细信息，请参阅 [Office 对话框 API 的身份验证](auth-with-office-dialog-api.md)）任务窗格中的代码请求从 Microsoft Graph 中获得数据, 并将令牌包含在这些请求中。 有关调用 Microsoft Graph 和 Microsoft Graph SDK 的详细信息，请参阅[Microsoft Graph 文档](/graph/)。

## <a name="recommended-libraries-and-samples"></a>推荐的库和示例

建议在不使用 SSO 访问 Microsoft Graph 时使用下列库：

- 对于使用服务器端并采用基于网络的框架（如 .NET Core 或 ASP.NET）的加载项，请使用 [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation)。
- 对于使用基于 NodeJS 的服务器端的加载项, 请使用[Passport Azure AD](https://github.com/AzureAD/passport-azure-ad)。
- 对于使用隐式流的加载项，请使用[msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki)。

有关使用 Microsoft 标识平台 (以前称为 "AAD v. 2.0") 的推荐库的详细信息，请参阅[Microsoft 标识平台身份验证库](/azure/active-directory/develop/reference-v2-libraries)。

以下示例从 Office 加载项获取 Microsoft Graph 数据：

- [Office 加载项 Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Outlook 加载项 Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Office 加载项 Microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React)
