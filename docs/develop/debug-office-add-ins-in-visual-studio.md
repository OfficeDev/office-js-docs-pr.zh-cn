---
title: 在 Visual Studio 中调试 Office 加载项
description: 使用 Visual Studio 调试Office桌面客户端中的Office加载项Windows。
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 49d52bd9b34b6f03dcf8b333cff816632c47c1c9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743528"
---
# <a name="debug-office-add-ins-in-visual-studio"></a>在 Visual Studio 中调试 Office 加载项

本文介绍如何调试使用 Office 2022 中的 Office 外接程序项目模板之一创建的 Visual Studio 外接程序中的客户端代码。  有关在外接程序中调试服务器端代码Office，请参阅调试 [Office 外接程序 - 服务器端还是客户端？](../testing/debug-add-ins-overview.md#server-side-or-client-side)。

> [!NOTE]
> 你无法通过 Visual Studio 在 Mac 上的 Office 中调试外接程序。 有关在 Mac 上调试的信息，请参阅Office [Mac 上的调试加载项](../testing/debug-office-add-ins-on-ipad-and-mac.md)。

## <a name="review-the-build-and-debug-properties"></a>查看生成和调试属性

在开始调试之前，请查看每个项目的属性以确认 Visual Studio将打开所需的 Office 应用程序，并正确设置其他生成和调试属性。

### <a name="add-in-project-properties"></a>外接程序项目属性

打开 **加载项** 项目的"属性"窗口，查看项目属性。

1. 在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。

2. 在菜单栏中，依次选择“**视图**” > “**属性窗口**”。

下表介绍了外接程序项目的属性。

|属性|说明|
|:-----|:-----|
|**启动操作**|指定外接程序的调试模式。 这应设置为 **Microsoft Edge** 加载项Outlook的加载项。 对于所有其他Office应用程序，应设置为Office **客户端**。|
|**启动文档**<br/>（仅限 Excel、PowerPoint 和 Word 外接程序）|指定要在启动项目时打开的文档。 在新项目中，此名称设置为 **[New Excel Workbook]**、**[New Word Document]** 或 **[New PowerPoint Presentation]**。 若要指定特定文档，请按照使用现有文档调试 [外接程序中的步骤操作](#use-an-existing-document-to-debug-the-add-in)。|
|**Web 项目**|指定与外接程序关联的 Web 项目的名称。|
|**电子邮件地址**<br/>（仅限 Outlook 外接程序）|指定你想在 Exchange Server 或 Exchange Online 中用来测试 Outlook 外接程序的用户帐户的电子邮件地址。 如果留空，则开始调试时将提示你输入电子邮件地址。|
|**EWS Url**<br/>（仅限 Outlook 外接程序）|指定 web Exchange URL (例如： `https://www.contoso.com/ews/exchange.aspx`) 。 此属性可留空。|
|**OWA Url**<br/>（仅限 Outlook 外接程序）|指定Outlook 网页版 URL (例如： `https://www.contoso.com/owa`) 。 此属性可留空。|
|**使用多重身份验证**<br/>（仅限 Outlook 加载项）|指定指示是否应该使用多重身份验证的布尔值。 默认值为 **false**，但该属性没有实际效果。 如果您通常必须提供第二个因素才能登录到电子邮件帐户，则当您开始调试时，系统会提示您。 |
|**用户名**<br/>（仅限 Outlook 外接程序）|指定你想在 Exchange Server 或 Exchange Online 中用来测试 Outlook 外接程序的用户帐户的名称。 此属性可留空。|
|**项目文件**|指定包含生成、配置和有关项目的其他信息的文件名称。|
|**项目文件夹**|指定项目文件的位置。|

> [!NOTE]
> 对于 Outlook 外接程序，你可以选择在“**属性**”窗口中为一个或多个 *Outlook 外接程序* 属性指定值，但这样做并不是必须的。

### <a name="web-application-project-properties"></a>Web 应用程序项目属性

打开 **Web 应用程序** 项目的"属性"窗口，查看项目属性。

1. 在 **"解决方案资源管理器**"中，选择 Web 应用程序项目。

2. 在菜单栏中，依次选择“**视图**” > “**属性窗口**”。

下表介绍了与 Office 外接程序项目最相关的 Web 应用程序项目的属性。

|属性|说明|
|:-----|:-----|
|**SSL 已启用**|指定是否在站点上启用 SSL。 对于 Office 外接程序项目，此属性应设置为 **True**。|
|**SSL URL**|指定站点的安全 HTTPS URL。 只读。|
|**URL**|指定站点的 HTTP URL。 只读。|
|**项目文件**|指定包含生成、配置和有关项目的其他信息的文件名称。|
|**项目文件夹**|指定项目文件的位置。 只读。 Visual Studio 在运行时生成的清单文件将写入到此位置的 `bin\Debug\OfficeAppManifests` 文件夹中。|

## <a name="debug-an-excel-powerpoint-or-word-add-in-project"></a>调试Excel、PowerPoint或 Word 加载项项目

本节介绍如何启动和调试 Excel、PowerPoint 或 Word 外接程序。

### <a name="start-the-excel-powerpoint-or-word-add-in-project"></a>启动 Excel、PowerPoint 或 Word 加载项项目

通过从菜单栏中选择 **"调试** > **""** 开始调试"或按 F5 按钮启动项目。 Visual Studio将自动生成解决方案，并启动Office应用程序。

在Visual Studio项目时，它将执行以下任务：

1. 创建 XML 清单文件的副本并添加到  `_ProjectName_\bin\Debug\OfficeAppManifests` 目录中。 托管Office的应用程序在启动加载项并调试加载项Visual Studio会使用该副本。

2. 在 Windows 计算机上创建一组注册表项，这些注册表项使加载项能够显示在Office应用程序中。

3. 生成 Web 应用程序项目，然后将它部署到本地 IIS Web 服务器 `https://localhost` () 。

4. 如果这是已部署到本地 IIS Web 服务器的第一个加载项项目，系统可能会提示你向当前用户的受信任根证书存储安装 Self-Signed 证书。 若要使 IIS Express 正确显示加载项内容，这是必需的操作。

> [!NOTE]
> Office使用 Edge 旧版 Webview 控件 (EdgeHTML) 在 Windows 计算机上运行外接程序，Visual Studio 可能会提示你添加本地网络环回豁免。 Webview 控件需要此权限才能访问部署到本地 IIS Web 服务器的网站。 还可以在 Visual Studio 中的“工具” > “选项” > “Office 工具(Web)” > “Web 加载项调试”下随时更改此设置。 若要了解在 Windows 计算机上使用哪些浏览器控件，请参阅 Office [外接程序使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

接下来，Visual Studio 会执行以下操作：

1. 将令牌替换为起始页的完全限定地址 ( (`https://localhost:44302/Home.html` 例如，) ，修改 XML 清单文件)  (`_ProjectName_\bin\Debug\OfficeAppManifests` `~remoteAppUrl` 的 [SourceLocation](../reference/manifest/sourcelocation.md) 元素。

2. 在 IIS Express 中启动 Web 应用程序项目。

3. 验证清单。 要查看项目中 XML 清单文件的验证规则，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。 

   > [!IMPORTANT]
   > Office安装Visual Studio清单 XSD 文件已过期。 如果收到清单的验证错误，则第一个疑难解答步骤应为将其中一个或多个文件替换为最新版本。 有关详细说明，请参阅[清单架构验证错误Visual Studio项目中](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects)。

4. 打开Office应用程序并旁加载外接程序。

### <a name="debug-the-excel-powerpoint-or-word-add-in"></a>调试Excel、PowerPoint或 Word 加载项

1. 在加载项应用程序中启动Office加载项。 例如，如果是任务窗格加载项，它将向"主页"功能区添加一个按钮 (例如，"显示 **任务** 窗格"按钮) 。 选择功能区中的按钮。 

   > [!NOTE]
   > 如果加载项不是由加载项旁加载Visual Studio，可以手动旁加载。 在Excel、PowerPoint或 Word 中，选择"插入"选项卡，然后选择"我的外接程序"右边 **的向下箭头**。
   >
   > ![Screenshot showing Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted.](../images/excel-cf-register-add-in-1b.png)
   >
   > 在可用外接程序列表中，找到“**开发人员外接程序**”部分并选择你的外接程序进行注册。

   > [!TIP]
   > 首次打开任务窗格时，它可能显示为空白。 如果是这样，它应在你稍后步骤中启动调试工具时正确呈现。

3. 打开 ["个性"菜单](../design/task-pane-add-ins.md#personality-menu) ，然后选择" **附加调试器"**。 这将打开 Webview 控件的调试工具，Office在加载项计算机上运行Windows工具。 您可以设置断点并逐步执行代码，如以下文章之一所述：

    - [使用适用于 Internet Explorer 的开发人员工具调试加载项](../testing/debug-add-ins-using-f12-tools-ie.md)
    - [使用旧版 Edge 开发人员工具调试加载项](../testing/debug-add-ins-using-devtools-edge-legacy.md)
    - [使用 Microsoft Edge（基于 Chromium）中的开发人员工具调试加载项](../testing/debug-add-ins-using-devtools-edge-chromium.md)

4. 若要更改代码，请首先停止调试Visual Studio并关闭Office应用程序。 进行更改，并开始新的调试会话。

## <a name="debug-an-outlook-add-in-project"></a>调试Outlook加载项项目

本节介绍如何启动和调试Outlook加载项。

### <a name="start-the-outlook-add-in-project"></a>启动Outlook加载项项目

通过从菜单栏中选择 **"调试** > **""** 开始调试"或按 F5 按钮启动项目。 Visual Studio将自动生成解决方案，并启动Outlook租户的 Microsoft 365 页面。

当Visual Studio项目时，它将执行以下任务。

1. 提示您输入登录凭据。 如果系统要求你重复登录，或者收到未经授权错误，则对于你的租户上的帐户，可能会Microsoft 365基本身份验证。 在这种情况下，请尝试使用 Microsoft 帐户。 您还可以尝试在"Web 外接程序项目属性"窗格中Outlook"使用多重身份验证"属性设置为 **True**。 请参阅 [加载项项目属性](#add-in-project-properties)。

1. 创建 XML 清单文件的副本并添加到 `_ProjectName_\bin\Debug\OfficeAppManifests` 目录中。 Outlook开始运行并调试外接程序Visual Studio此副本。

2. 生成 Web 应用程序项目，然后将它部署到本地 IIS Web 服务器 `https://localhost` () 。

3. 如果这是已部署到本地 IIS Web 服务器的第一个加载项项目，系统可能会提示你向当前用户的受信任根证书存储安装 Self-Signed 证书。 若要使 IIS Express 正确显示加载项内容，这是必需的操作。

> [!NOTE]
> Office使用 Edge 旧版 Webview 控件 (EdgeHTML) 在 Windows 计算机上运行外接程序，Visual Studio 可能会提示你添加本地网络环回豁免。 Webview 控件需要此权限才能访问部署到本地 IIS Web 服务器的网站。 还可以在 Visual Studio 中的“工具” > “选项” > “Office 工具(Web)” > “Web 加载项调试”下随时更改此设置。 若要了解在 Windows 计算机上使用哪些浏览器控件，请参阅 Office [外接程序使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

接下来，Visual Studio 会执行以下操作：

1. 将令牌替换为起始页的完全限定地址 ( (`https://localhost:44302/Home.html` 例如，) ，修改 XML 清单文件)  (`_ProjectName_\bin\Debug\OfficeAppManifests` `~remoteAppUrl` 的 [SourceLocation](../reference/manifest/sourcelocation.md) 元素。

2. 在 IIS Express 中启动 Web 应用程序项目。

3. 验证清单。 要查看项目中 XML 清单文件的验证规则，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。 

   > [!IMPORTANT]
   > Office安装Visual Studio清单 XSD 文件已过期。 如果收到清单的验证错误，则第一个疑难解答步骤应为将其中一个或多个文件替换为最新版本。 有关详细说明，请参阅[清单架构验证错误Visual Studio项目中](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects)。

4. 在 Outlook 中打开Microsoft 365租户的 Microsoft Edge。

### <a name="debug-the-outlook-add-in"></a>调试Outlook加载项

1. 在"Outlook"页中，选择电子邮件或约会项目以在其自己的窗口中打开它。 

2. 按 F12 打开 Edge 调试工具。

3. 工具打开后，启动加载项。 例如，在消息顶部的工具栏中，选择"更多应用"按钮，然后从打开的标注中选择加载项。

   ![Screenshot showing the More apps button and the callout that it opens with the add-in's name and icon visible with other app icons.](../images/outlook-more-apps-button.png)

4. 使用以下文章之一中的说明设置断点并逐步执行代码。 它们各自都有一个指向更详细指南的链接。

   - [使用旧版 Edge 开发人员工具调试加载项](../testing/debug-add-ins-using-devtools-edge-legacy.md)
   - [使用 Microsoft Edge（基于 Chromium）中的开发人员工具调试加载项](../testing/debug-add-ins-using-devtools-edge-chromium.md)

   > [!TIP]
   > 若要调试在方法`Office.initialize``Office.onReady`或外接程序打开时运行的方法中运行的代码，请设置断点，然后关闭并重新打开外接程序。 有关这些方法详细信息，请参阅初始化Office[加载项](../develop/initialize-add-in.md)。

5. 若要更改代码，请首先停止调试会话，Visual Studio并关闭Outlook页面。 进行更改，并开始新的调试会话。

## <a name="use-an-existing-document-to-debug-the-add-in"></a>使用现有文档调试外接程序

如果你有一个文档包含要在调试 Excel、PowerPoint 或 Word 外接程序时使用的测试数据，则可以将 Visual Studio 配置为在启动项目时打开该文档。 若要指定在调试外接程序时要使用的现有文档，请完成以下步骤。

1. 在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。

2. 从菜单栏中，选择“**项目**” > “**添加现有项**”。

3. 在“**添加现有项**”对话框中，找到并选择要添加的文档。

4. 选择“**添加**”按钮以将文档添加到项目中。

5. 在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。

6. 在菜单栏中，依次选择“**视图**” > “**属性窗口**”。

7. 在“**属性**”窗口中，选择“**启动文档**”列表，然后选择添加到项目中的文档。 该项目现在配置为在该文档中启动外接程序。

## <a name="next-steps"></a>后续步骤

在外接程序正常工作后，请参阅[部署和发布 Office 外接程序](../publish/publish.md)，以了解可用于将外接程序分发给用户的方法。
