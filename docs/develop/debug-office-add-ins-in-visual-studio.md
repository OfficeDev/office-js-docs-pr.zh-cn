---
title: 在 Visual Studio 中调试 Office 加载项
description: 使用 Visual Studio 在 Windows 上的 Office 桌面客户端中调试 Office 加载项
ms.date: 12/31/2019
localization_priority: Normal
ms.openlocfilehash: 018bfa24424514598d323c29d165e3e8ec066a8e
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093656"
---
# <a name="debug-office-add-ins-in-visual-studio"></a>在 Visual Studio 中调试 Office 加载项

本文介绍如何使用 Visual Studio 2019 在 Windows 上的 Office 桌面客户端中调试 Office 加载项。 如果使用的是 Visual Studio 的其他版本，操作步骤可能略有不同。 

> [!NOTE]
> 无法使用 Visual Studio 在 Office 网页版或 Mac 版 Office 中调试加载项。 若要了解如何在这些平台上进行调试，请参阅[在 Office 网页版中调试 Office 加载项](../testing/debug-add-ins-in-office-online.md)或[在 Mac 上调试 Office 加载项](../testing/debug-office-add-ins-on-ipad-and-mac.md)。

## <a name="enable-debugging-for-add-in-commands-and-ui-less-code"></a>对加载项命令和无 UI 的代码启用调试

当 Visual Studio 调试 Windows 上的 Office 时，加载项托管在 Microsoft Internet Explorer 或 Microsoft Edge 浏览器实例中。 若要确定开发计算机上使用的浏览器，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。
> [!NOTE]
> 以下过程不再需要 JS_Debug 环境变量。 有关详细信息，请参阅 Microsoft 开发人员社区支持论坛中的 [Office Web 加载项中的调试行为](https://developercommunity.visualstudio.com/content/problem/740413/office-development-inconsistent-script-debugging-b.html)。

[!include[Enable debugging on Microsoft Edge DevTools](../includes/enable-debugging-on-edge-devtools.md)]

## <a name="review-the-build-and-debug-properties"></a>查看生成和调试属性

在开始调试之前，请检查每个项目的属性以确认 Visual Studio 将打开所需的主机应用程序，并已正确设置其他生成和调试属性。

### <a name="add-in-project-properties"></a>外接程序项目属性

打开外接程序项目的“**属性**”窗口以查看项目属性：

1. 在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。

2. 在菜单栏中，依次选择“**视图**” > “**属性窗口**”。

下表介绍了外接程序项目的属性。

|**属性**|**说明**|
|:-----|:-----|
|**启动操作**|指定外接程序的调试模式。 目前，Office 外接程序项目仅支持 **Office 桌面客户端**模式。|
|**启动文档**<br/>（仅限 Excel、PowerPoint 和 Word 外接程序）|指定要在启动项目时打开的文档。|
|**Web 项目**|指定与外接程序关联的 Web 项目的名称。|
|**电子邮件地址**<br/>（仅限 Outlook 外接程序）|指定你想在 Exchange Server 或 Exchange Online 中用来测试 Outlook 外接程序的用户帐户的电子邮件地址。|
|**EWS Url**<br/>（仅限 Outlook 外接程序）|Exchange Web 服务 URL（例如：`https://www.contoso.com/ews/exchange.aspx`）。 |
|**OWA Url**<br/>（仅限 Outlook 外接程序）|Outlook 网页版 URL（例如：`https://www.contoso.com/owa`）。|
|**使用多重身份验证**<br/>（仅限 Outlook 加载项）|布尔值，指示是否应使用多重身份验证。|
|**用户名**<br/>（仅限 Outlook 外接程序）|指定你想在 Exchange Server 或 Exchange Online 中用来测试 Outlook 外接程序的用户帐户的名称。|
|**项目文件**|指定包含生成、配置和有关项目的其他信息的文件名称。|
|**项目文件夹**|项目文件的位置。|

> [!NOTE]
> 对于 Outlook 外接程序，你可以选择在“**属性**”窗口中为一个或多个 *Outlook 外接程序*属性指定值，但这样做并不是必须的。

### <a name="web-application-project-properties"></a>Web 应用程序项目属性

打开 Web 应用程序项目的“**属性**”窗口以查看项目属性：

1. 在 "**解决方案资源管理器**" 中，选择 "web 应用程序" 项目。

2. 在菜单栏中，依次选择“**视图**” > “**属性窗口**”。

下表介绍了与 Office 外接程序项目最相关的 Web 应用程序项目的属性。

|**属性**|**说明**|
|:-----|:-----|
|**SSL 已启用**|指定是否在站点上启用 SSL。 对于 Office 外接程序项目，此属性应设置为 **True**。|
|**SSL URL**|指定站点的安全 HTTPS URL。 只读。|
|**URL**|指定站点的 HTTP URL。 只读。|
|**项目文件**|指定包含生成、配置和有关项目的其他信息的文件名称。|
|**项目文件夹**|指定项目文件的位置。 只读。 Visual Studio 在运行时生成的清单文件将写入到此位置的 `bin\Debug\OfficeAppManifests` 文件夹中。|

## <a name="use-an-existing-document-to-debug-the-add-in"></a>使用现有文档调试外接程序

如果你有一个文档包含要在调试 Excel、PowerPoint 或 Word 外接程序时使用的测试数据，则可以将 Visual Studio 配置为在启动项目时打开该文档。 若要指定在调试外接程序时要使用的现有文档，请完成以下步骤。

1. 在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。

2. 从菜单栏中，选择“**项目**” > “**添加现有项**”。

3. 在“**添加现有项**”对话框中，找到并选择要添加的文档。

4. 选择“**添加**”按钮以将文档添加到项目中。

5. 在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。

6. 在菜单栏中，依次选择“**视图**” > “**属性窗口**”。

7. 在“**属性**”窗口中，选择“**启动文档**”列表，然后选择添加到项目中的文档。 该项目现在配置为在该文档中启动外接程序。

## <a name="start-the-project"></a>启动项目

从菜单栏中依次选择“**调试**” > “**开始调试**”，可启动项目。 Visual Studio 将自动生成解决方案并启动 Office 以托管外接程序。

> [!NOTE]
> 启动 Outlook 外接程序项目时，系统会提示你输入登录凭据。 如果要求您反复登录，或者如果您收到未经授权的错误，则可能会对 Microsoft 365 租户上的帐户禁用基本身份验证。 在这种情况下，请尝试使用 Microsoft 帐户。 可能还需要在“Outlook Web 加载项”项目属性对话框中将属性“使用多重身份验证”设置为 True。

当 Visual Studio 生成项目时，它执行以下任务：

1. 创建 XML 清单文件的副本并将其添加到 `_ProjectName_\bin\Debug\OfficeAppManifests` 目录。 启动 Visual Studio 并调试外接程序时，主机应用程序将使用此副本。

2. 在计算机上创建一组允许外接程序在主机应用程序中显示的注册表项。

3. 生成 Web 应用程序项目，然后将其部署到本地 IIS Web 服务器 (https://localhost))。

4. 如果这是你已部署到本地 IIS Web 服务器的第一个加载项项目，系统可能会提示你将自签名证书安装到当前用户的受信任的根证书存储中。 若要使 IIS Express 正确显示加载项内容，这是必需的操作。

> [!NOTE]
> 在 Windows 10 上运行时，最新版本的 Office 可能会使用较新的 Web 控件来显示加载项内容。 如果是这种情况，Visual Studio 可能会提示你添加本地网络环回豁免。 在 Office 主机应用程序中，需要这样做才能使 Web 控件访问部署到本地 IIS Web 服务器的网站。 还可以在 Visual Studio 中的“工具” > “选项” > “Office 工具(Web)” > “Web 加载项调试”下随时更改此设置****************。

接下来，Visual Studio 会执行以下操作：

1. 通过将 `~remoteAppUrl` 标记替换为起始页的完全限定地址（例如，`https://localhost:44302/Home.html`）来修改 XML 清单文件的 [SourceLocation](../reference/manifest/sourcelocation.md) 元素。

2. 在 IIS Express 中启动 Web 应用程序项目。

3. 打开主机应用程序。

当您构建项目时，Visual Studio 不会在“输出”**** 窗口中显示验证错误。 Visual Studio 报告“**错误列表**”窗口中出现的错误和警告。 通过在代码和文本编辑器中显示不同颜色的波浪下划线（称为波浪线），Visual Studio 还报告验证错误。 通过这些标志，你可以得知 Visual Studio 在你的代码中检测到的问题。 有关如何启用或禁用验证的详细信息，请参阅[选项、文本编辑器、JavaScript、IntelliSense](/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2019)。

要查看项目中 XML 清单文件的验证规则，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。

## <a name="debug-the-code-for-an-excel-powerpoint-or-word-add-in"></a>调试 Excel、PowerPoint 或 Word 外接程序的代码

如果在[启动项目](#start-the-project)后，在主机应用程序（Excel、PowerPoint 或 Word）中显示的文档中看不到外接程序，请在主机应用程序中手动启动外接程序。 例如，通过选择“**主页**”选项卡功能区中的“**显示任务窗格**”按钮来启动任务窗格外接程序。在 Excel、PowerPoint 或 Word 中显示外接程序后，你可以通过执行以下操作来调试代码：

1. 在 Excel、PowerPoint 或 Word 中，选择“**插入**”选项卡，然后选择“**我的外接程序**”右侧的向下箭头。

    ![Windows 版 Excel 的“插入”功能区及突出显示的“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)

2. 在可用外接程序列表中，找到“**开发人员外接程序**”部分并选择你的外接程序进行注册。

3. 在 Visual Studio 中，在代码中设置断点。

4. 在 Excel、PowerPoint 或 Word 中，与外接程序进行交互。

5. 在 Visual Studio 中命中断点时，根据需要逐步执行代码。

你可以更改代码并在外接程序中查看这些更改的效果，而无需关闭主机应用程序并重新启动该项目。 保存对代码的更改后，只需在主机应用程序中重新加载外接程序。 例如，通过选择任务窗格的右上角来激活[个性菜单](../design/task-pane-add-ins.md#personality-menu)，然后选择“**重新加载**”，便可重新加载任务窗格外接程序。

## <a name="debug-the-code-for-an-outlook-add-in"></a>调试 Outlook 外接程序的代码

在你已[启动项目](#start-the-project)，且 Visual Studio 启动 Outlook 来托管外接程序后，打开电子邮件或约会项目。 

Outlook activates the add-in for the item as long as the activation criteria are met. The add-in bar appears at the top of the Inspector window or Reading Pane, and your Outlook add-in appears as a button in the add-in bar. If your add-in has an add-in command, a button will appear in the ribbon, either in the default tab or a specified custom tab, and the add-in will not appear in the add-in bar.

若要查看 Outlook 外接程序，请选择对应 Outlook 外接程序的按钮。 在 Outlook 中显示外接程序后，你可以通过执行以下操作来调试代码：

1. 在 Visual Studio 中，在代码中设置断点。

2. 在 Outlook 中，与外接程序进行交互。

3. 在 Visual Studio 中命中断点时，根据需要逐步执行代码。

你可以更改代码并在外接程序中查看这些更改的效果，而无需关闭 Outlook 并重新启动该项目。 保存对代码的更改后，只需打开外接程序的快捷菜单（在 Outlook 中），然后选择“**重新加载**”。

## <a name="next-steps"></a>后续步骤

在外接程序正常工作后，请参阅[部署和发布 Office 外接程序](../publish/publish.md)，以了解可用于将外接程序分发给用户的方法。
