---
title: 单一登录 (SSO) 快速入门
description: 使用 Yeoman 生成器生成使用单一登录的 Node.js Office 加载项。
ms.date: 09/07/2022
ms.prod: non-product-specific
ms.localizationpriority: high
ms.openlocfilehash: ecbecfd7e475c224451735c7a864f6de2c230d07
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/30/2022
ms.locfileid: "68239369"
---
# <a name="single-sign-on-sso-quick-start"></a>单一登录 (SSO) 快速入门

本文将使用 Office 外接程序的 Yeoman 生成器创建用于 Excel、Outlook、Word 或 PowerPoint 的 Office 加载项，该外接程序使用单一登录 (SSO) 。

> [!NOTE]
> 由适用于 Office 外接程序的 Yeoman 生成器提供的 SSO 模板仅在 localhost 上运行，无法部署。 如果要为生产目的使用 SSO 构建新的 Office 外接程序，请按照 [创建使用单一登录的 Node.js Office 加](../develop/create-sso-office-add-ins-nodejs.md)载项中的说明进行操作。

## <a name="prerequisites"></a>先决条件

- [Node.js](https://nodejs.org)（最新[LTS](https://nodejs.org/about/releases) 版本）。

- 最新版本的 [Yeoman](https://github.com/yeoman/yo) 和[适用于 Office 加载项的 Yeoman 生成器](../develop/yeoman-generator-overview.md)。若要全局安装这些工具，请从命令提示符处运行以下命令。

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

- 如果你使用的是 Mac，并且计算机上未安装 Azure CLI，则必须安装 [Homebrew](https://brew.sh/)。 在此快速入门过程中运行的 SSO 配置脚本将使用 Homebrew 来安装 Azure CLI，然后将使用 Azure CLI 在 Azure 中配置 SSO。

## <a name="create-the-add-in-project"></a>创建加载项项目

> [!TIP]
> Yeoman 生成器可以使用脚本类型 JavaScript 或 TypeScript 为 Excel、Outlook、Word 或 PowerPoint 创建启用 SSO 的 Office 外接程序。 下列说明指定 `JavaScript` 和 `Excel`，但应选择最适合方案的脚本类型和 Office 客户端应用程序。

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **选择项目类型:** `Office Add-in Task Pane project supporting single sign-on (localhost)`
- **选择脚本类型:** `JavaScript`
- **要如何命名加载项?** `My Office Add-in`
- **要支持哪一个 Office 客户端应用程序?** 选择`Excel`、`Outlook`或 `Word``Powerpoint`。

:::image type="content" source="../images/yo-office-sso-excel.png" alt-text="命令行接口中 Yeoman 生成器的提示和答案。":::

完成此向导后，生成器会创建项目，并安装支持的 Node 组件。

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>浏览项目

使用 Yeoman 生成器创建的加载项项目包含适用于启用了 SSO 的任务窗格加载项代码。

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="configure-sso"></a>配置 SSO

现已创建外接程序项目并包含为简化 SSO 过程而必需的代码，请完成以下步骤，为外接程序配置 SSO。

1. 转到项目的根文件夹。

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. 运行下列命令，为加载项配置 SSO。

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > 如果租户已配置为需要双因素验证，则此命令将失败。 在此方案中，需要按照使用 [单一登录教程的“创建 Node.js Office 外](../develop/create-sso-office-add-ins-nodejs.md) 接程序”中的所有步骤手动完成 Azure 应用注册和 SSO 配置步骤。

3. Web 浏览器窗口将打开，并提示登录 Azure。 使用现有的 Microsoft 365 管理员凭据登录到 Azure。 这些凭据将用于在 Azure 中注册新的应用程序并配置 SSO 所需的设置。

    > [!NOTE]
    > 在此步骤中，如果使用非管理员凭据登录 Azure，`configure-sso` 脚本将无法向组织中的用户提供该加载项的管理员许可。 因此，该加载项的用户无法使用 SSO，系统将提示用户登录。

4. 输入凭据后，关闭浏览器窗口并返回命令提示符。 随着 SSO 配置流程的继续，将看到写入控制台的状态消息。 正如控制台消息所述，加载项项目中的文件会自动更新 SSO 流程所需的数据。

## <a name="test-your-add-in"></a>测试加载项

如果已创建 Excel、Word 或 PowerPoint 加载项，请完成以下部分中的步骤来尝试。 如果已创建 Outlook 加载项，请改为完成 [Outlook](#outlook) 部分中的步骤。

### <a name="excel-word-and-powerpoint"></a>Excel、Word 和 PowerPoint

完成以下步骤以测试 Excel、Word 或 PowerPoint 加载项。

1. SSO 配置流程完成后，运行以下命令生成项目、启动本地 Web 服务器，并旁加载之前在 Office 客户端应用程序中选定的加载项。

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. 当运行上一命令时，Excel、Word 或 PowerPoint 打开时，请确保使用用户帐户登录，该用户帐户是 Microsoft 365 组织的成员，与在 [上一部分](#configure-sso)的步骤 3 中配置 SSO 时用来连接到 Azure 的 Microsoft 365 管理员帐户相同。 执行此操作，将为成功进行 SSO 建立了相应的条件。

3. 在 Office 客户端应用程序中，选择 **“开始** ”选项卡，然后选择 **“显示任务窗格** ”打开加载项任务窗格。

    :::image type="content" source="../images/excel-quickstart-addin-3b.png" alt-text="Excel 加载项按钮。":::

4. 在任务窗格底部，选择“获取我的用户配置文件信息”按钮以开始 SSO 流程。

5. 如果对话框窗口显示代表加载项请求权限，则表示 你的方案不支持 SSO，并且加载项已退回至替代的用户身份验证方法。 当租户管理员未授予外接程序访问 Microsoft Graph 的许可，或者用户未使用有效的 Microsoft 帐户或 Microsoft 365 教育版 或 Work 帐户登录到 Office 时，可能会发生这种情况。 选择 **“接受** ”以继续。

    ![显示突出显示“接受”按钮的权限请求对话框屏幕截图。](../images/sso-permissions-request.png)

    > [!NOTE]
    > 用户接受此权限请求后，以后将不会再收到提示。

6. 加载项检索已登录用户的配置文件信息并写入至文档中。 下图显示了写入至 Excel 工作表的配置文件信息的实例。

    ![显示 Excel 工作表中用户配置文件信息的屏幕截图。](../images/sso-user-profile-info-excel.png)

### <a name="outlook"></a>Outlook

完成以下步骤以试用 Outlook 加载项。

1. SSO 配置过程完成后，运行以下命令生成项目并启动本地 Web 服务器。

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. 按照[旁加载 Outlook 加载项以供测试](../outlook/sideload-outlook-add-ins-for-testing.md)中的说明操作，旁加载加载项。 确保用于登录 Outlook 的用户与在[上一节](#configure-sso)第 3 步中配置 SSO 时用于连接至 Azure 的 Microsoft 365 管理员帐户是同一 Microsoft 365 组织的成员。 执行此操作，将为成功进行 SSO 建立了相应的条件。

3. 在 Outlook 中，撰写一封新邮件。

4. 在消息撰写窗口中，选择 **“显示任务窗格** ”按钮以打开加载项任务窗格。

    ![。显示 Outlook 撰写邮件窗口中突出显示的加载项功能区按钮屏幕截图。](../images/outlook-sso-ribbon-button.png)

5. 在任务窗格底部，选择“获取我的用户配置文件信息”按钮以开始 SSO 流程。

6. 如果对话框窗口显示代表加载项请求权限，则表示 你的方案不支持 SSO，并且加载项已退回至替代的用户身份验证方法。 当租户管理员未授予外接程序访问 Microsoft Graph 的许可，或者用户未使用有效的 Microsoft 帐户或 Microsoft 365 教育版 或 Work 帐户登录到 Office 时，可能会发生这种情况。 选择 **“接受** ”以继续。

    ![突出显示“接受”按钮的权限请求对话框屏幕截图。](../images/sso-permissions-request.png)

    > [!NOTE]
    > 用户接受此权限请求后，以后将不会再收到提示。

7. 加载项检索已登录用户的配置文件信息并写入至电子邮件的正文中。

    ![显示 Outlook 撰写邮件窗口中的用户配置文件信息的屏幕截图。](../images/sso-user-profile-info-outlook.png)

## <a name="next-steps"></a>后续步骤

祝贺你成功创建了可能使用 SSO 的任务窗格加载项，并在不支持 SSO 时，使用替代用户身份验证方法。 若要了解如何自定义加载项以添加需要不同权限的新功能，请参阅 “[自定义启用了 Node.js SSO 的加载项](sso-quickstart-customize.md)”。

## <a name="see-also"></a>另请参阅

- [为 Office 加载项启用单一登录](../develop/sso-in-office-add-ins.md)
- [自定义启用了 Node.js SSO 的加载项](sso-quickstart-customize.md)
- [创建使用单一登录的 Node.js Office 加载项](../develop/create-sso-office-add-ins-nodejs.md)
- [排查单一登录 (SSO) 错误消息](../develop/troubleshoot-sso-in-office-add-ins.md)
- [使用Visual Studio Code发布](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)