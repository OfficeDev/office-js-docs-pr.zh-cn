---
title: 同意管理员访问加载项
description: 了解如何向外接程序授予管理员同意。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 85a230ffd3769b0013081067c88d65d38d43b760
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743779"
---
# <a name="grant-administrator-consent-to-the-add-in"></a>同意管理员访问加载项

> [!NOTE]
> 仅在开发加载项时，才需要执行此过程。 将生产加载项部署到 AppSource 或 Microsoft 365 管理中心 时，用户单独信任它，或者管理员将在安装时同意组织。

在注册 *加载项* 后 [执行此过程](../develop/register-sso-add-in-aad-v2.md)。

1. 浏览到 [Azure 门户 - 应用注册](https://go.microsoft.com/fwlink/?linkid=2083908) 页面以查看应用注册。

1. 使用管理员 ***凭据*** 登录到您的Microsoft 365租户。 例如，MyName@contoso.onmicrosoft.com。

1. Select the app with 显示名称 **$ADD-IN-NAME$**.

1. On the **$ADD-IN-NAME$** page， select **API permissions** then， under the **Configured permissions** section， choose **Grant admin consent for [tenant name]**. 对于 **出现的** 确认，选择"是"。

> [!NOTE]
> 如果你使用的是开发人员帐户，建议采用Microsoft 365[过程](https://developer.microsoft.com/microsoft-365/dev-program)。 但是，如果您愿意，可以旁加载开发中的 SSO 外接程序，并提示用户提供同意表单。 有关详细信息，请参阅旁[加载和Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)[旁加载Office web 版](../testing/sideload-office-add-ins-for-testing.md)。
