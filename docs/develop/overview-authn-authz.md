---
title: Office 加载项中的身份验证和授权概述
description: ''
ms.date: 01/07/2020
localization_priority: Priority
ms.openlocfilehash: 5086095c711bbf6df98e457092f825690d43229e
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217273"
---
# <a name="overview-of-authentication-and-authorization-in-office-add-ins"></a>Office 加载项中的身份验证和授权概述

Web 应用程序和 Office 加载项默认允许匿名访问，但你可要求用户通过登录进行身份验证。 例如，可以要求用户使用 Microsoft 帐户、Office 365 工作或学校帐户或者其他常用帐户登录。 此任务被称为“用户身份验证”，因为它让加载项能够知道用户的身份。

你的加载项还能够获得用户的同意来访问其 Microsoft Graph 数据（例如其 Office 365 个人资料、OneDrive 文件和 SharePoint 数据），或者访问 Google、Facebook、领英、SalesForce 和 GitHub 等其他外部源中的数据。 此任务被称为“加载项（或应用）授权”，因为要获得授权的是*加载项*，而不是用户。

有两种方式可用来完成这些身份验证。

- **Office 单一登录 (SSO)**：此系统*当前为预览版*，它让用户能在登录到 Office 的同时登录到加载项。 此外，此加载项还可使用用户的 Office 凭据向加载项授予对 Microsoft Graph 的权限。 （不可通过此系统访问非 Microsoft 源。）
- **通过 Azure Active Directory 进行 Web 身份验证和授权**：这是老生常谈，没有特别之处。 它只是在出现 Office SSO 系统之前 Office 加载项（及其他 Web 应用）对用户进行身份验证和授权应用的方式，现在仍在 Office SSO 不可用的场景中使用。

下列流程图展示了需要如同加载项开发人员一样作出的决策。 详细信息请参见本文稍后部分。

![一张图像，它显示了在 Office 加载项中实现身份验证和授权的决策流程图](../images/authflowchart.png)

## <a name="user-authentication-without-sso"></a>在不使用 SSO 的情况下进行用户身份验证

你可如同在任何其他 Web 应用程序中操作一样使用 Azure Active Directory (AAD) 在 Office 加载项中对用户进行身份验证，但存在一个例外：AAD 禁止其登录页在 iFrame 中打开。 当 Office 加载项在 *Office 网页版*中运行时，任务窗格是一个 iFrame。 这意味着你将需要在通过 Office 对话框 API 打开的对话框中打开 AAD 登录屏幕。 这会影响你使用身份验证帮助程序库的方式。 有关详细信息，请参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。

要了解如何使用 AAD 对身份验证进行编程，首先请查看 [Microsoft 标识平台 (v2.0) 概述](/azure/active-directory/develop/v2-overview)。 该文档集中有很多教程和指南，还有相关示例和库的链接。 正如[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)中所述，你可能需要调整示例中的代码以在 Office 对话框中运行。

## <a name="access-to-microsoft-graph-without-sso"></a>在不使用 SSO 的情况下访问 Microsoft Graph

可通过从 Azure Active Directory (AAD) 获取到 Microsoft Graph 的访问令牌，为加载项获得到 Graph 数据的授权。 可在不依赖 Office SSO 的情况下执行此操作。 要详细了解操作方式，请参阅[在不使用 SSO 的情况下访问 Microsoft Graph](authorize-to-microsoft-graph-without-sso.md)（此文中有更多详细信息和示例链接）。

## <a name="user-authentication-with-sso"></a>在使用 SSO 的情况下进行用户身份验证

要使用 SSO 来验证用户的身份，任务窗格或函数文件中的代码会调用 [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) 方法。 如果用户未登录 Office，则 Office 将打开一个对话框，并将其导航到 Azure Active Directory 登录页面。 用户登录后或者在用户已登录时，该方法会返回一个访问令牌。 此令牌是**代理**流中的启动令牌。 （详见[使用 SSO 访问 Microsoft Graph](#access-to-microsoft-graph-with-sso)。）但是，它也可用作 ID 令牌，因为它包含多个对当前用户而言唯一的声明，例如 `preferred_username`、`name`、`sub` 和 `oid`。 要查看指南了解将哪个属性用作最终用户 ID，请参阅 [Microsoft 标识平台访问令牌](https://docs.microsoft.com/azure/active-directory/develop/access-tokens#payload-claims)。 有关上述某一令牌的示例，请参阅[访问令牌示例](sso-in-office-add-ins.md#example-access-token)。

代码从令牌中提取所需的声明后，它将使用该值在你保留的用户表或用户数据库中查找用户。 使用数据库来用户用户首选项或用户帐户状态等用户相关信息。 由于你在使用 SSO，因此你的用户不单独登录到你的加载项，你无需存储用户的密码。

在开始使用 SSO 实现用户身份验证之前，请确保你完全熟悉[为 Office 加载项启用单一登录](sso-in-office-add-ins.md)一文。另请注意下述示例：

- [Office 加载项 NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)，特别是文件 [ssoAuthES6.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/src/auth.ts)。 
- [Office 加载项 ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)。 

但是，这些示例不将令牌用作 ID 令牌。 它们使用此令牌通过**代理**流获得对 Microsoft Graph 的访问权限。

## <a name="access-to-microsoft-graph-with-sso"></a>在使用 SSO 的情况下访问 Microsoft Graph

要使用 SSO 来获取访问 Microsoft Graph 的权限，任务窗格或函数文件中的加载项会调用 [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) 方法。 如果用户未登录 Office，则 Office 将打开一个对话框，并将其导航到 Azure Active Directory 登录页面。 用户登录后或者在用户已登录时，该方法会返回一个访问令牌。 此令牌是**代理**流中的启动令牌。 具体而言，它有一个带 `access_as_user` 值的 `scope` 声明。 要在指南中了解令牌中的声明，请参阅 [Microsoft 标识平台访问令牌](https://docs.microsoft.com/azure/active-directory/develop/access-tokens#payload-claims)。 有关上述某一令牌的示例，请参阅[访问令牌示例](sso-in-office-add-ins.md#example-access-token)。

在代码获取令牌后，它会在**代理**流中使用该令牌来获取第二个令牌，即到 Microsoft Graph 的访问令牌。

在开始实现 Office SSO 之前，请确保你完全熟悉下面两篇文章：

- [为 Office 加载项启用单一登录](sso-in-office-add-ins.md)
- [使用 SSO 对 Microsoft Graph 授权](authorize-to-microsoft-graph.md)

你还应至少阅读此处所列的其中一篇演示文章。 即使你不执行这些步骤，也可在其中了解有关如何实现 Office SSO 和**代理**流的宝贵信息。 

- [创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)
- [创建使用单一登录的 Node.js Office 加载项](create-sso-office-add-ins-nodejs.md)

另请注意下述示例：

- [Office 加载项 NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)
- [Office 加载项 ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)

## <a name="access-to-non-microsoft-data-sources"></a>访问非 Microsoft 数据源

借助 Google、Facebook、领英、SalesForce 和 GitHub 等热门在线服务，开发人员可授权用户访问自己在其他应用中的帐户。 这样，便可在 Office 加载项中添加这些服务。 要概述了解加载项可执行此操作的方法，请参阅[在 Office 加载项中授权外部服务](auth-external-add-ins.md)。

> [!IMPORTANT]
> 开始编码之前，请了解数据源是否允许在 iFrame 中打开其登录屏幕。 当 Office 加载项在 *Office 网页版*中运行时，任务窗格是一个 iFrame。 如果数据源禁止在 iFrame 中打开其登录屏幕，则你需要在通过 Office 对话框 API 打开的对话框中打开登录屏幕。 有关详细信息，请参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。
