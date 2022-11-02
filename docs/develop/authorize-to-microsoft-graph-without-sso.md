---
title: 从 Office 加载项授权 Microsoft Graph
description: 了解如何从 Office 加载项授权 Microsoft Graph。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 37dd4be3acb92dc7884972de923d94936fa870f4
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810167"
---
# <a name="authorize-to-microsoft-graph-from-an-office-add-in"></a>从 Office 加载项授权 Microsoft Graph

加载项可以通过从Microsoft 标识平台获取 Microsoft Graph 访问令牌来获取对 Microsoft Graph 数据的授权。 使用授权代码流或隐式流，就像在其他 Web 应用程序中一样，但有一个例外：Microsoft 标识平台不允许其登录页在 iframe 中打开。 当 Office 加载项在 *Office web 版* 中运行时，任务窗格是 iframe。 这意味着需要使用 Office 对话框 API 在对话框中打开登录页。 这将影响你使用身份验证和授权帮助程序库的方式。 有关详细信息，请参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。

> [!NOTE]
> 如果要实现 SSO 并计划访问 Microsoft Graph，请参阅 [使用 SSO 授权 Microsoft Graph](authorize-to-microsoft-graph.md)。

有关使用Microsoft 标识平台对身份验证进行编程的信息，请参阅[Microsoft 标识平台文档](/azure/active-directory/develop)。 你将在该文档集中找到教程和指南，以及相关示例的链接。 再次，可能需要调整示例中的代码以在 Office 对话框中运行，以考虑在任务窗格的单独进程中运行的 Office 对话框。

代码获取 Microsoft Graph 的访问令牌后，要么将访问令牌从对话框传递到任务窗格，要么将令牌存储在数据库中，并向任务窗格发出令牌可用信号。  (有关详细信息，请参阅 [使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md) 。任务窗格中) 代码从 Microsoft Graph 请求数据，并在这些请求中包含令牌。 有关调用 Microsoft Graph 和 Microsoft Graph SDK 的详细信息，请参阅 [Microsoft Graph 文档](/graph/)。

## <a name="recommended-libraries-and-samples"></a>推荐的库和示例

建议在访问 Microsoft Graph 时使用以下库。

- 对于使用服务器端并采用基于网络的框架（如 .NET Core 或 ASP.NET）的加载项，请使用 [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation)。
- 对于使用基于 NodeJS 的服务器端的加载项, 请使用[Passport Azure AD](https://github.com/AzureAD/passport-azure-ad)。
- 对于使用隐式流的加载项，请使用[msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki)。

有关使用 Microsoft 标识平台 (以前称为 "AAD v. 2.0") 的推荐库的详细信息，请参阅[Microsoft 标识平台身份验证库](/azure/active-directory/develop/reference-v2-libraries)。

以下示例从 Office 外接程序获取 Microsoft Graph 数据。

- [Office 加载项 Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Outlook 加载项 Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Office 加载项 Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
