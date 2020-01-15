---
title: 使用 Yeoman 生成器创建使用 SSO 的 Office 加载项（预览版）
description: 使用 Yeoman 生成器生成使用单一登录的 Node.js Office 加载项（预览版）。
ms.date: 01/13/2020
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: 3c67fdb2b8582546c13624dcb8a6f139bb638df0
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111104"
---
# <a name="use-the-yeoman-generator-to-create-an-office-add-in-that-uses-single-sign-on-preview"></a>使用 Yeoman 生成器创建使用单一登录的 Node.js Office 加载项（预览版）。

本文将介绍如何使用 Yeoman 生成器创建适用于 Excel、Word 或 PowerPoint ，尽可能使用单一登录（SSO）的 Office 加载项，并在不支持 SSO 时使用替代的用户身份验证方法。

> [!TIP]
> 尝试完成此快速入门前，请查看“[为 Office 加载项启用单一登录](../develop/sso-in-office-add-ins.md)”了解有关 Office 加载项中 SSO 的基本概念。 
 
Yeoman 生成器简化了 SSO 加载项的创建流程，能够自动执行在 Azure 内配置所需的步骤，并生成加载项使用 SSO 所需的代码。 有关介绍如何手动完成 Yeoman 生成器自动运行步骤的详细演练，请参阅“[创建使用单一登录的 Node.js Office 加载项](../develop/create-sso-office-add-ins-nodejs.md)”教程。

## <a name="prerequisites"></a>先决条件

- [Node.js](https://nodejs.org)（版本 10.15.0 或更高版本）

- 最新版本的 [Yeoman](https://github.com/yeoman/yo) 和[适用于 Office 外接程序的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。若要全局安装这些工具，请从命令提示符处运行以下命令：

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

- 一个 Office 365（Office 的订阅版本）账户。 如果还没有 Office 365 账户，可以通过加入 [Office 365 开发人员计划](https://aka.ms/devprogramsignup)获得 90 天免费的可续订 Office 365 订阅。 

- 一个 Office 365 预览体验成员内部版本。 应使用最新的每月版本并从预览体验成员频道构建，但你必须[是 Office 预览体验成员](https://products.office.com/office-insider?tab=tab-1)才能获取此版本。 

    > [!NOTE]
    > 当内部版本进入生产半年频道时，将禁用对该内部版本的预览功能（包括 SSO）的支持。

## <a name="create-the-add-in-project"></a>创建加载项项目

> [!TIP]
> Yeoman 生成器可创建适用于 Excel、Word 或 PowerPoint 的启用 SSO 的 Office 加载项，能够使用 JavaScript 或 TypeScript 类型的脚本创建。 下列说明指定 `JavaScript` 和 `Excel`，但应选择最适合方案的脚本类型和 Office 客户端应用程序。

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **选择项目类型:** `Office Add-in Task Pane project supporting single sign-on`
- **选择脚本类型:** `Javascript`
- **要如何命名加载项?** `My SSO Office Add-in`
- **要支持哪一个 Office 客户端应用程序?** `Excel`

![有关 Yeoman 生成器提示和回答的屏幕截图](../images/yo-office-sso-excel.png)

完成此向导后，生成器会创建项目，并安装支持的 Node 组件。

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>浏览项目

使用 Yeoman 生成器创建的加载项项目包含适用于启用了 SSO 的任务窗格加载项代码。

- 项目根目录中的 **./manifest.xml** 文件定义加载项的设置和功能。

- **./src/taskpane/taskpane.html** 文件包含组成任务窗格的 HTML。
- **./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。
- **./src/taskpane/taskpane.js** 文件包含用于加快任务窗格与 Office 托管应用程序之间的交互的 Office JavaScript API 代码。

- **./src/helpers/documentHelper.js** 文件使用 Office JavaScript 库将 Microsoft Graph 库中的数据添加至 Office 文档。
- **./src/helpers/fallbackauthdialog.html** 文件是加载回退身份验证方法 JavaScript 的无界面页面。
- **./src/helpers/fallbackauthdialog.js** 文件包含用户使用 msal.js 登录的回退身份验证方法 JavaScript。
- **./src/helpers/fallbackauthhelper.js** 文件包含任务窗格 JavaScript，当不支持 SSO 身份验证时，在方案中调用回退身份验证方法。
- **./src/helpers/ssoauthhelper.js** 文件包含调用 SSO API、`getAccessToken` 的 JavaScript ，接收引导令牌，针对 Microsoft Graph 访问令牌启动引导令牌交换，同时调用 Microsoft Graph 以获得数据。

- 项目根目录中的 **/ENV** 文件定义了加载项项目所使用的常量。
    > [!NOTE]
    > 此文件中定义的部分常量用于简化 SSO 流程。 可能需要更新此文件中的数值以匹配特定的方案。 例如，加载项需要 `User.Read`之外的其他内容时，则可以更新该文件来指定不同的范围。

## <a name="configure-sso"></a>配置 SSO

此时，加载项项目已创建并含有简化 SSO 流程所需的代码。 接下来，完成以下步骤，为你的加载项配置 SSO。

1. 导航到项目的根文件夹。

    ```command&nbsp;line
    cd "My SSO Office Add-in"
    ```

2. 运行下列命令，为加载项配置 SSO。

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > 如果租户已配置为需要双因素验证，则此命令将失败。 在此情况中，需要按照“[创建使用单一登录的 Node.js Office 加载项](../develop/create-sso-office-add-ins-nodejs.md)”教程所述，手动完成 Azure 应用程序注册和 SSO 配置步骤。

3. Web 浏览器窗口将打开，并提示登录 Azure。 使用现有的 Office 365 管理员凭据登录到 Azure。 这些凭据将用于在 Azure 中注册新的应用程序并配置 SSO 所需的设置。

    > [!NOTE]
    > 在此步骤中，如果使用非管理员凭据登录 Azure，`configure-sso` 脚本将无法向组织中的用户提供该加载项的管理员许可。 因此，该加载项的用户无法使用 SSO，系统将提示用户登录。

4. 输入凭据后，关闭浏览器窗口并返回命令提示符。 随着 SSO 配置流程的继续，将看到写入控制台的状态消息。 正如控制台消息所述，加载项项目中的文件会自动更新 SSO 流程所需的数据。

## <a name="try-it-out"></a>试用

1. SSO 配置流程完成后，运行以下命令生成项目、启动本地 Web 服务器，并旁加载之前在 Office 客户端应用程序中选定的加载项。

    > [!NOTE]
    > Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。 如果系统在运行以下命令后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。

    ```command&nbsp;line
    npm start
    ```

2. 在运行上一个命令（即 Excel、 Word 或 PowerPoin）时打开的 Office 客户端应用程序中，确保登录的用户与在[上一节](#configure-sso)第 3 步中配置 SSO 时用于连接至 Azure 所使用的 Office 365 管理员账户是同一 Office 365 组织的成员。 执行此操作，将为成功进行 SSO 建立了相应的条件。 

3. 在 Office 客户端应用程序中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。 下图显示 Excel 中的该按钮。

    ![Excel 加载项按钮](../images/excel-quickstart-addin-3b.png)

4. 在任务窗格底部，选择 “**获取我的用户配置文件信息**”按钮以开始 SSO 流程。 

    > [!NOTE] 
    > 如果此时尚未登录到 Office，系统将提示登录。 如前所述，如果希望成功完成 SSO，登录的用户与在[上一节](#configure-sso)第 3 步中配置 SSO 时用于连接至 Azure 所使用的 Office 365 管理员账户是同一 Office 365 组织的成员。

5. 如果对话框窗口显示代表加载项请求权限，则表示 你的方案不支持 SSO，并且加载项已退回至替代的用户身份验证方法。 租户管理员未向用户授予访问 Microsoft Graph 的许可，或未使用有效的 Microsoft 帐户或 Office 365 （“工作或学校”）帐户登录 Office 时，可能会出现这种情况。 选择对话框窗口中的“**接受**”按钮以继续。

    ![权限请求对话框](../images/sso-permissions-request.png)

    > [!NOTE]
    > 用户接受此权限请求后，以后将不会再收到提示。

6. 加载项检索已登录用户的配置文件信息并写入至文档中。 下图显示了写入至 Excel 工作表的配置文件信息的实例。

    ![Excel 工作表中的用户配置文件信息](../images/sso-user-profile-info-excel.png)

## <a name="next-steps"></a>后续步骤

祝贺你成功创建了可能使用 SSO 的任务窗格加载项，并在不支持 SSO 时，使用替代用户身份验证方法。 若要详细了解有关 Yeoman 生成器自动完成的 SSO 配置步骤，以及有助于 SSO 流程的代码，参见“[创建使用单一登录的 Node.js Office 加载项](../develop/create-sso-office-add-ins-nodejs.md)”教程。

## <a name="see-also"></a>另请参阅

- [为 Office 加载项启用单一登录](../develop/sso-in-office-add-ins.md)
- [创建使用单一登录的 Node.js Office 加载项](../develop/create-sso-office-add-ins-nodejs.md)
- [排查单一登录 (SSO) 错误消息](../develop/troubleshoot-sso-in-office-add-ins.md)