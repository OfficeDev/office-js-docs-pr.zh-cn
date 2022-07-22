---
title: 在 Visual Studio 中调试 Office 加载项
description: 使用 Visual Studio 在 Windows 上的 Office 桌面客户端中调试 Office 加载项。
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 09693f81c069aba97740265fa88bf117a937c742
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958711"
---
# <a name="debug-office-add-ins-in-visual-studio"></a>在 Visual Studio 中调试 Office 加载项

本文介绍如何在使用 Visual Studio 2022 中的某个 Office 外接程序项目模板创建的 Office 外接程序中调试客户端代码。  有关在 Office 外接程序中调试服务器端代码的信息，请参阅 Office 外接程序调[试概述 - 服务器端或客户端？](../testing/debug-add-ins-overview.md#server-side-or-client-side)

> [!NOTE]
> 不能使用 Visual Studio 在 Office on Mac 中调试加载项。 有关在 Mac 上调试的信息，请参阅 [Mac 上的调试 Office 加载项](../testing/debug-office-add-ins-on-ipad-and-mac.md)。

## <a name="review-the-build-and-debug-properties"></a>查看生成和调试属性

在开始调试之前，请查看每个项目的属性，确认 Visual Studio 将打开所需的 Office 应用程序，并适当设置其他生成和调试属性。

### <a name="add-in-project-properties"></a>外接程序项目属性

打开加载项项目的 **“属性** ”窗口以查看项目属性。

1. 在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。

2. 在菜单栏中，依次选择“**视图**” > “**属性窗口**”。

下表介绍了外接程序项目的属性。

|属性|说明|
|:-----|:-----|
|**启动操作**|指定外接程序的调试模式。 这应设置为 Outlook 加载项的 **Microsoft Edge** 。 对于所有其他 Office 应用程序，应将其设置为 **Office 桌面客户端**。|
|**启动文档**<br/>（仅限 Excel、PowerPoint 和 Word 外接程序）|指定要在启动项目时打开的文档。 在新项目中，此设置为 **[新建 Excel 工作簿]**、 **[新建 Word 文档]** 或 **[新建 PowerPoint 演示文稿]**。 若要指定特定文档，请按照 [使用现有文档中的步骤调试加载项](#use-an-existing-document-to-debug-the-add-in)。|
|**Web 项目**|指定与外接程序关联的 Web 项目的名称。|
|**电子邮件地址**<br/>（仅限 Outlook 外接程序）|指定你想在 Exchange Server 或 Exchange Online 中用来测试 Outlook 外接程序的用户帐户的电子邮件地址。 如果留空，则在开始调试时，系统会提示你输入电子邮件地址。|
|**EWS Url**<br/>（仅限 Outlook 外接程序）|指定 Exchange Web 服务 URL (例如： `https://www.contoso.com/ews/exchange.aspx`) 。 此属性可以留空。|
|**OWA Url**<br/>（仅限 Outlook 外接程序）|指定Outlook 网页版 URL (例如： `https://www.contoso.com/owa`) 。 此属性可以留空。|
|**使用多重身份验证**<br/>（仅限 Outlook 加载项）|指定指示是否应使用多重身份验证的布尔值。 默认值为 **false**，但该属性没有实际效果。 如果通常必须提供第二个因素才能登录到电子邮件帐户，则在开始调试时会提示你。 |
|**用户名**<br/>（仅限 Outlook 外接程序）|指定你想在 Exchange Server 或 Exchange Online 中用来测试 Outlook 外接程序的用户帐户的名称。 此属性可以留空。|
|**项目文件**|指定包含生成、配置和有关项目的其他信息的文件名称。|
|**项目文件夹**|指定项目文件的位置。|

> [!NOTE]
> 对于 Outlook 外接程序，你可以选择在“**属性**”窗口中为一个或多个 *Outlook 外接程序* 属性指定值，但这样做并不是必须的。

### <a name="web-application-project-properties"></a>Web 应用程序项目属性

打开 Web 应用程序项目的 **“属性”** 窗口以查看项目属性。

1. 在 **解决方案资源管理器** 中，选择 Web 应用程序项目。

2. 在菜单栏中，依次选择“**视图**” > “**属性窗口**”。

下表介绍了与 Office 外接程序项目最相关的 Web 应用程序项目的属性。

|属性|说明|
|:-----|:-----|
|**SSL 已启用**|指定是否在站点上启用 SSL。 对于 Office 外接程序项目，此属性应设置为 **True**。|
|**SSL URL**|指定站点的安全 HTTPS URL。 只读。|
|**URL**|指定站点的 HTTP URL。 只读。|
|**项目文件**|指定包含生成、配置和有关项目的其他信息的文件名称。|
|**项目文件夹**|指定项目文件的位置。 只读。 Visual Studio 在运行时生成的清单文件将写入到此位置的 `bin\Debug\OfficeAppManifests` 文件夹中。|

## <a name="debug-an-excel-powerpoint-or-word-add-in-project"></a>调试 Excel、PowerPoint 或 Word 外接程序项目

本部分介绍如何启动和调试 Excel、PowerPoint 或 Word 加载项。

### <a name="start-the-excel-powerpoint-or-word-add-in-project"></a>启动 Excel、PowerPoint 或 Word 加载项项目

通过从菜单栏中选择 **“调试** > **开始调试** ”或按 F5 按钮启动项目。 Visual Studio 将自动生成解决方案并启动 Office 主机应用程序。

Visual Studio 生成项目时，它将执行以下任务：

1. 创建 XML 清单文件的副本并将其添加到  `_ProjectName_\bin\Debug\OfficeAppManifests` 目录。 在启动 Visual Studio 并调试外接程序时，托管外接程序的 Office 应用程序会使用此副本。

2. 在 Windows 计算机上创建一组注册表条目，使加载项能够显示在 Office 应用程序中。

3. 生成 Web 应用程序项目，然后将其部署到本地 IIS Web 服务器 (`https://localhost`) 。

4. 如果这是部署到本地 IIS Web 服务器的第一个外接程序项目，系统可能会提示你将Self-Signed证书安装到当前用户的受信任根证书存储。 若要使 IIS Express 正确显示加载项内容，这是必需的操作。

> [!NOTE]
> 如果 Office 使用 Edge 旧版 Webview 控件 (EdgeHTML) 在 Windows 计算机上运行加载项，则 Visual Studio 可能会提示你添加本地网络环回豁免。 Webview 控件必须能够访问部署到本地 IIS Web 服务器的网站。 还可以在 Visual Studio 中的“工具” > “选项” > “Office 工具(Web)” > “Web 加载项调试”下随时更改此设置。 若要了解 Windows 计算机上使用的浏览器控件，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

接下来，Visual Studio 会执行以下操作：

1. 修改复制到`_ProjectName_\bin\Debug\OfficeAppManifests`目录) 的 XML 清单文件 (的 [SourceLocation](/javascript/api/manifest/sourcelocation) 元素，将令牌替换`~remoteAppUrl`为起始页 (的完全限定地址，例如 `https://localhost:44302/Home.html`) 。

2. 在 IIS Express 中启动 Web 应用程序项目。

3. 验证清单。 要查看项目中 XML 清单文件的验证规则，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。 

   > [!IMPORTANT]
   > Visual Studio 安装的 Office 清单 XSD 文件已过时。 如果收到清单的验证错误，第一个故障排除步骤应该是将其中一个或多个文件替换为最新版本。 有关详细说明，请参阅 [Visual Studio 项目中的清单架构验证错误](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects)。

4. 打开 Office 应用程序并旁加载加载项。

### <a name="debug-the-excel-powerpoint-or-word-add-in"></a>调试 Excel、PowerPoint 或 Word 加载项

1. 在 Office 应用程序中启动加载项。 例如，如果它是任务窗格外接程序，它将向 **主页** 功能区添加一个按钮 (例如，“ **显示任务窗格** ”按钮) 。 选择功能区中的按钮。 

   > [!NOTE]
   > 如果外接程序未由 Visual Studio 旁加载，则可以手动旁加载它。 在 Excel、PowerPoint 或 Word 中，选择 **“插入** ”选项卡，然后选择“ **我的外接程序**”右侧的向下箭头。
   >
   > ![显示 Windows 上 Excel 中的“插入”功能区的屏幕截图，其中突出显示了“我的外接程序”箭头。](../images/excel-cf-register-add-in-1b.png)
   >
   > 在可用外接程序列表中，找到“**开发人员外接程序**”部分并选择你的外接程序进行注册。

   > [!TIP]
   > 首次打开任务窗格时，该窗格可能显示为空白。 如果是这样，则在稍后的步骤中启动调试工具时，它应正确呈现。

3. 打开 [个性菜单](../design/task-pane-add-ins.md#personality-menu) ，然后选择 **“附加调试器**”。 这将打开 Office 用于在 Windows 计算机上运行加载项的 Web 视图控件的调试工具。 可以按照以下文章之一中所述设置断点并逐步执行代码：

    - [使用适用于 Internet Explorer 的开发人员工具调试加载项](../testing/debug-add-ins-using-f12-tools-ie.md)
    - [使用旧版 Edge 开发人员工具调试加载项](../testing/debug-add-ins-using-devtools-edge-legacy.md)
    - [使用 Microsoft Edge（基于 Chromium）中的开发人员工具调试加载项](../testing/debug-add-ins-using-devtools-edge-chromium.md)

4. 若要对代码进行更改，请先在 Visual Studio 中停止调试会话并关闭 Office 应用程序。 进行更改，并启动新的调试会话。

## <a name="debug-an-outlook-add-in-project"></a>调试 Outlook 加载项项目

本部分介绍如何启动和调试 Outlook 加载项。

### <a name="start-the-outlook-add-in-project"></a>启动 Outlook 加载项项目

通过从菜单栏中选择 **“调试** > **开始调试** ”或按 F5 按钮启动项目。 Visual Studio 将自动生成解决方案并启动 Microsoft 365 租户的 Outlook 页面。

当 Visual Studio 生成项目时，它将执行以下任务。

1. 提示输入登录凭据。 如果被要求重复登录或收到未经授权的错误，则 Microsoft 365 租户上的帐户可能会禁用基本身份验证。 在这种情况下，请尝试使用 Microsoft 帐户。 还可以在 Outlook Web 外接程序项目属性窗格中尝试将属性 **“使用多重身份验证** ”设置为 **True** 。 请参阅 [加载项项目属性](#add-in-project-properties)。

1. 创建 XML 清单文件的副本并将其添加到 `_ProjectName_\bin\Debug\OfficeAppManifests` 目录。 启动 Visual Studio 并调试外接程序时，Outlook 会使用此副本。

2. 生成 Web 应用程序项目，然后将其部署到本地 IIS Web 服务器 (`https://localhost`) 。

3. 如果这是部署到本地 IIS Web 服务器的第一个外接程序项目，系统可能会提示你将Self-Signed证书安装到当前用户的受信任根证书存储。 若要使 IIS Express 正确显示加载项内容，这是必需的操作。

> [!NOTE]
> 如果 Office 使用 Edge 旧版 Webview 控件 (EdgeHTML) 在 Windows 计算机上运行加载项，则 Visual Studio 可能会提示你添加本地网络环回豁免。 Webview 控件必须能够访问部署到本地 IIS Web 服务器的网站。 还可以在 Visual Studio 中的“工具” > “选项” > “Office 工具(Web)” > “Web 加载项调试”下随时更改此设置。 若要了解 Windows 计算机上使用的浏览器控件，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

接下来，Visual Studio 会执行以下操作：

1. 修改复制到`_ProjectName_\bin\Debug\OfficeAppManifests`目录) 的 XML 清单文件 (的 [SourceLocation](/javascript/api/manifest/sourcelocation) 元素，将令牌替换`~remoteAppUrl`为起始页 (的完全限定地址，例如 `https://localhost:44302/Home.html`) 。

2. 在 IIS Express 中启动 Web 应用程序项目。

3. 验证清单。 要查看项目中 XML 清单文件的验证规则，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。 

   > [!IMPORTANT]
   > Visual Studio 安装的 Office 清单 XSD 文件已过时。 如果收到清单的验证错误，第一个故障排除步骤应该是将其中一个或多个文件替换为最新版本。 有关详细说明，请参阅 [Visual Studio 项目中的清单架构验证错误](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects)。

4. 打开 Microsoft Edge 中 Microsoft 365 租户的 Outlook 页面。

### <a name="debug-the-outlook-add-in"></a>调试 Outlook 加载项

1. 在 Outlook 页面中，选择电子邮件或约会项目以在自己的窗口中打开它。 

2. 按 F12 打开 Edge 调试工具。

3. 打开该工具后，启动加载项。 例如，在邮件顶部的工具栏中，选择 **“更多应用** ”按钮，然后从打开的标注中选择加载项。

   ![显示“更多应用”按钮及其打开的标注的屏幕截图，其中加载项的名称和图标与其他应用图标一起可见。](../images/outlook-more-apps-button.png)

4. 使用以下文章之一中的说明设置断点并逐步执行代码。 它们都有一个指向更详细指南的链接。

   - [使用旧版 Edge 开发人员工具调试加载项](../testing/debug-add-ins-using-devtools-edge-legacy.md)
   - [使用 Microsoft Edge（基于 Chromium）中的开发人员工具调试加载项](../testing/debug-add-ins-using-devtools-edge-chromium.md)

   > [!TIP]
   > 若要调试在函数中 `Office.initialize` 运行的代码或 `Office.onReady` 在加载项打开时运行的函数，请设置断点，然后关闭并重新打开加载项。 有关这些函数的详细信息，请参阅 [“初始化 Office 加载项](../develop/initialize-add-in.md)”。

5. 若要更改代码，请先在 Visual Studio 中停止调试会话并关闭 Outlook 页面。 进行更改，并启动新的调试会话。

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
