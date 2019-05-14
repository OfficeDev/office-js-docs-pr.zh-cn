---
title: 在 Visual Studio 中创建和调试 Office 外接程序
description: 使用 Visual Studio 在 Windows 上的 Office 桌面客户端中创建和调试 Office 加载项
ms.date: 05/08/2019
localization_priority: Priority
ms.openlocfilehash: c60599ed63c327d10b157e642e109542c3cefc47
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952094"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a>在 Visual Studio 中创建和调试 Office 外接程序

本文介绍如何使用 Visual Studio 2017 为 Excel、Word、PowerPoint 或 Outlook 创建 Office 外接程序，并在 Windows 上的 Office 桌面客户端中调试外接程序。 如果使用的是 Visual Studio 的其他版本，操作步骤可能略有不同。

> [!NOTE]
> Visual Studio 不支持为 OneNote 或 Project 创建 Office 外接程序，但你可以使用 [Office 外接程序的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)来创建这些类型的外接程序。
> - 若要开始使用 OneNote 的外接程序，请参阅[生成首个 OneNote 外接程序](../quickstarts/onenote-quickstart.md)。
>
> - 若要开始使用 Project 的外接程序，请参阅[生成首个 Project 外接程序](../quickstarts/project-quickstart.md)。

## <a name="prerequisites"></a>先决条件

- 安装了 **Office/SharePoint 开发**工作负载的 [Visual Studio 2017](https://www.visualstudio.com/vs/)

    > [!TIP]
    > 如果之前已安装 Visual Studio 2017，请[使用 Visual Studio 安装程序](/visualstudio/install/modify-visual-studio)，以确保安装 **Office/SharePoint 开发**工作负载。 如果尚未安装此工作负载，请使用 Visual Studio 安装程序进行[安装](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads)。

- Office 2013 或更高版本

    > [!TIP]
    > 如果你还没有 Office，则可加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)以获取 Office 365 订阅，或者你可以[注册免费 1 个月的试用版](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735)。

## <a name="create-the-add-in-project-in-visual-studio"></a>在 Visual Studio 中创建外接程序项目

首先完成以下三个步骤，然后完成后续部分中与你正在创建的外接程序类型相对应的步骤。 

1. 打开 Visual Studio，在 Visual Studio 菜单栏中，依次选择“**文件**” > “**新建**” > “**项目**”。

2. 在 **Visual C#** 或 **Visual Basic** 下的项目类型列表中，展开 **Office/SharePoint**，选择“**外接程序**”，然后选择要创建的外接程序项目的类型。 

3. 命名此项目，然后选择“**确定**”。

### <a name="word-web-add-in-or-outlook-web-add-in"></a>Word Web 外接程序或 Outlook Web 外接程序

如果你已选择创建 **Word Web 外接程序**或 **Outlook Web 外接程序**，Visual Studio 将创建一个解决方案，并在“**解决方案资源管理器**”中显示这两个项目。 接下来，你可以[浏览 Visual Studio 解决方案](#explore-the-visual-studio-solution)。 

### <a name="powerpoint-web-add-in"></a>PowerPoint Web 外接程序

如果你已选择创建 **PowerPoint Web 外接程序**，则会出现“**创建 Office 外接程序**”对话框。 

- 若要创建任务窗格外接程序，请选择“**向 PowerPoint 添加新功能**”，然后选择“**完成**”按钮以创建 Visual Studio 解决方案。

- 若要创建内容外接程序，请选择“**向 PowerPoint 幻灯片插入内容**”，然后选择“**完成**”按钮以创建 Visual Studio 解决方案。

接下来，你可以[浏览 Visual Studio 解决方案](#explore-the-visual-studio-solution)。

### <a name="excel-web-add-in"></a>Excel Web 外接程序

如果你已选择创建 **Excel Web 外接程序**，则会出现“**创建 Office 外接程序**”对话框。 

- 若要创建任务窗格外接程序，请选择“**向 Excel 添加新功能**”，然后选择“**完成**”按钮以创建 Visual Studio 解决方案。

- 若要创建内容外接程序，请选择“**向 Excel 电子表格插入内容**”，选择“**下一步**”按钮，选择以下选项之一，然后选择“**完成**”按钮以创建 Visual Studio 解决方案：

    - **基本外接程序** - 使用最少的入门代码创建内容外接程序项目

    - **文档可视化外接程序** - 使用入门代码创建内容外接程序项目，以实现可视化并绑定到数据  

### <a name="explore-the-visual-studio-solution"></a>浏览 Visual Studio 解决方案

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

## <a name="modify-your-add-in-settings"></a>修改外接程序设置

若要修改外接程序的设置，请编辑外接程序项目中的 XML 清单文件。 在“**解决方案资源管理器**”中，展开外接程序项目节点，展开包含 XML 清单的文件夹并选择 XML 清单。 你可以指向该文件中的任何元素以查看说明该元素用途的工具提示。 有关清单文件的详细信息，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。

## <a name="develop-the-contents-of-your-add-in"></a>开发外接程序的内容

加载项项目允许您修改描述加载项的设置，而 Web 应用程序提供加载项中显示的内容。 

Web 应用程序项目包含可用于实现入门的默认 HTML 文件、JavaScript 文件和 CSS 文件。 其中一些文件包含对其他 JavaScript 库的引用，包括适用于 Office 的 JavaScript API。 你可以通过更新这些文件和/或添加更多 HTML 和 JavaScript 文件来开发外接程序。 下表描述了创建 Visual Studio 解决方案时 Web 应用程序项目包含的默认文件。

|**文件名**|**说明**|
|:-----|:-----|
|**Home.html**<br/>（Excel、PowerPoint、Word）<br/><br/>**MessageRead.html**<br/>(Outlook)|外接程序的默认 HTML 页面。 在文档、电子邮件或约会项目中激活该外接程序时，此页面将显示为外接程序内的第一个页面。 此文件包含入门所需的所有文件引用。 你可以通过将 HTML 代码添加到此文件来开始开发外接程序。|
|**Home.js**<br/>（Excel、PowerPoint、Word）<br/><br/>**MessageRead.js**<br/>(Outlook)|与 **Home.html** 页面（Excel、PowerPoint、Word）或 **MessageRead.html** 页面 (Outlook) 关联的 JavaScript 文件。 此文件应包含特定于 **Home.html** 页面（Excel、PowerPoint、Word）或 **MessageRead.html** 页面 (Outlook) 行为的任何代码。 此文件包含一些可帮你入门的示例代码。|
|**Home.css**<br/>（Excel、PowerPoint、Word）<br/><br/>**MessageRead.css**<br/>(Outlook)|定义要应用于外接程序的默认样式。 我们建议对设计和样式使用 Office UI Fabric。 有关详细信息，请参阅 [Office 外接程序中的 Office UI Fabric](../design/office-ui-fabric.md)。|

> [!NOTE]
> 你无需使用这些文件。 你可以随意将其他文件添加到项目并改为使用这些文件。 如果要将另一个 HTML 文件显示为外接程序的初始页面，请打开清单编辑器，然后将 **SourceLocation** 属性设置为该文件的名称。

## <a name="debug-your-add-in"></a>调试外接程序

你可以使用 Visual Studio 在 Windows 上的 Office 桌面客户端中调试外接程序，如以下部分所述：

- [查看生成和调试属性](#review-the-build-and-debug-properties)
- [使用现有文档调试外接程序](#use-an-existing-document-to-debug-the-add-in)
- [启动项目](#start-the-project)
- [调试 Excel、PowerPoint 或 Word 外接程序的代码](#debug-the-code-for-an-excel-powerpoint-or-word-add-in)
- [调试 Outlook 外接程序的代码](#debug-the-code-for-an-outlook-add-in)

> [!NOTE]
> 你无法使用 Visual Studio 在 Office Online 或 Office for Mac 中调试 Office 外接程序。 有关在这些平台上进行调试的信息，请参阅[在 Office Online 中调试 Office 外接程序](../testing/debug-add-ins-in-office-online.md)或[在 iPad 和 Mac 上调试 Office 外接程序](../testing/debug-office-add-ins-on-ipad-and-mac.md)

### <a name="review-the-build-and-debug-properties"></a>查看生成和调试属性

在开始调试之前，请检查每个项目的属性以确认 Visual Studio 将打开所需的主机应用程序，并已正确设置其他生成和调试属性。

#### <a name="add-in-project-properties"></a>外接程序项目属性

打开外接程序项目的“**属性**”窗口以查看项目属性：

1. 在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。

2. 在菜单栏中，依次选择“**视图**” >  “**属性窗口**”。

下表介绍了外接程序项目的属性。

|**属性**|**说明**|
|:-----|:-----|
|**启动操作**|指定外接程序的调试模式。 目前，Office 外接程序项目仅支持 **Office 桌面客户端**模式。|
|**启动文档**<br/>（仅限 Excel、PowerPoint 和 Word 外接程序）|指定要在启动项目时打开的文档。|
|**Web 项目**|指定与外接程序关联的 Web 项目的名称。|
|**电子邮件地址**<br/>（仅限 Outlook 外接程序）|指定你想在 Exchange Server 或 Exchange Online 中用来测试 Outlook 外接程序的用户帐户的电子邮件地址。|
|**EWS Url**<br/>（仅限 Outlook 外接程序）|Exchange Web 服务 URL（例如：`https://www.contoso.com/ews/exchange.aspx`）。 |
|**OWA Url**<br/>（仅限 Outlook 外接程序）|Outlook Web App URL（例如，`https://www.contoso.com/owa`）。|
|**使用多重身份验证**<br/>（仅限 Outlook 加载项）|布尔值，指示是否应使用多重身份验证。|
|**用户名**<br/>（仅限 Outlook 外接程序）|指定你想在 Exchange Server 或 Exchange Online 中用来测试 Outlook 外接程序的用户帐户的名称。|
|**项目文件**|指定包含生成、配置和有关项目的其他信息的文件名称。|
|**项目文件夹**|项目文件的位置。|

> [!NOTE]
> 对于 Outlook 外接程序，你可以选择在“**属性**”窗口中为一个或多个 *Outlook 外接程序*属性指定值，但这样做并不是必须的。

#### <a name="web-application-project-properties"></a>Web 应用程序项目属性

打开 Web 应用程序项目的“**属性**”窗口以查看项目属性：

1. 在“**解决方案资源管理器**”中，选择 Web 应用程序项目。

2. 在菜单栏中，依次选择“**视图**” >  “**属性窗口**”。

下表介绍了与 Office 外接程序项目最相关的 Web 应用程序项目的属性。

|**属性**|**说明**|
|:-----|:-----|
|**SSL 已启用**|指定是否在站点上启用 SSL。 对于 Office 外接程序项目，此属性应设置为 **True**。|
|**SSL URL**|指定站点的安全 HTTPS URL。 只读。|
|**URL**|指定站点的 HTTP URL。 只读。|
|**项目文件**|指定包含生成、配置和有关项目的其他信息的文件名称。|
|**项目文件夹**|指定项目文件的位置。 只读。 Visual Studio 在运行时生成的清单文件将写入到此位置的 `bin\Debug\OfficeAppManifests` 文件夹中。|

### <a name="use-an-existing-document-to-debug-the-add-in"></a>使用现有文档调试外接程序

如果你有一个文档包含要在调试 Excel、PowerPoint 或 Word 外接程序时使用的测试数据，则可以将 Visual Studio 配置为在启动项目时打开该文档。 若要指定在调试外接程序时要使用的现有文档，请完成以下步骤。

1. 在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。

2. 从菜单栏中，选择“**项目**” > “**添加现有项**”。

3. 在“**添加现有项**”对话框中，找到并选择要添加的文档。

4. 选择“**添加**”按钮以将文档添加到项目中。

5. 在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。

6. 在菜单栏中，依次选择“**视图**” > “**属性窗口**”。

7. 在“**属性**”窗口中，选择“**启动文档**”列表，然后选择添加到项目中的文档。 该项目现在配置为在该文档中启动外接程序。

### <a name="start-the-project"></a>启动项目

从菜单栏中依次选择“**调试**” > “**开始调试**”，可启动项目。 Visual Studio 将自动生成解决方案并启动 Office 以托管外接程序。

> [!NOTE]
> 启动 Outlook 外接程序项目时，系统会提示你输入登录凭据。 如果系统要求你重复登录，或者如果收到指示未经授权的错误，则可能会禁用 Office 365 租户上帐户的基本身份验证。 在这种情况下，请尝试使用 Microsoft 帐户。 可能还需要在“Outlook Web 加载项”项目属性对话框中将属性“使用多重身份验证”设置为 True。

当 Visual Studio 生成项目时，它执行以下任务：

1. 创建 XML 清单文件的副本并将其添加到 `_ProjectName_\bin\Debug\OfficeAppManifests` 目录。 启动 Visual Studio 并调试外接程序时，主机应用程序将使用此副本。

2. 在计算机上创建一组允许外接程序在主机应用程序中显示的注册表项。

3. 生成 Web 应用程序项目，然后将其部署到本地 IIS Web 服务器 (https://localhost))。

4. 如果这是你已部署到本地 IIS Web 服务器的第一个加载项项目，系统可能会提示你将自签名证书安装到当前用户的受信任的根证书存储中。 若要使 IIS Express 正确显示加载项内容，这是必需的操作。


> [!NOTE]
> 在 Windows 10 上运行时，最新版本的 Office 可能会使用较新的 Web 控件来显示加载项内容。 如果是这种情况，Visual Studio 可能会提示你添加本地网络环回豁免。 在 Office 主机应用程序中，需要这样做才能使 Web 控件访问部署到本地 IIS Web 服务器的网站。 还可以在 Visual Studio 中的“工具” > “选项” > “Office 工具(Web)” > “Web 加载项调试”下随时更改此设置****************。


接下来，Visual Studio 会执行以下操作：

1. 通过将 `~remoteAppUrl` 标记替换为起始页的完全限定地址（例如，`https://localhost:44302/Home.html`）来修改 XML 清单文件的 [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) 元素。

2. 在 IIS Express 中启动 Web 应用程序项目。

3. 打开主机应用程序。

生成项目时，Visual Studio 不会显示“**输出**”窗口中的验证错误。 Visual Studio 报告“**错误列表**”窗口中出现的错误和警告。 通过在代码和文本编辑器中显示不同颜色的波浪下划线（称为波浪线），Visual Studio 还报告验证错误。 通过这些标志，你可以得知 Visual Studio 在你的代码中检测到的问题。 有关详细信息，请参阅[代码和文本编辑器](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)。 有关如何启用或禁用验证的详细信息，请参阅[选项、文本编辑器、JavaScript、IntelliSense](/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2017)。

要查看项目中 XML 清单文件的验证规则，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。

### <a name="debug-the-code-for-an-excel-powerpoint-or-word-add-in"></a>调试 Excel、PowerPoint 或 Word 外接程序的代码

如果在[启动项目](#start-the-project)后，在主机应用程序（Excel、PowerPoint 或 Word）中显示的文档中看不到外接程序，请在主机应用程序中手动启动外接程序。 例如，通过选择“**主页**”选项卡功能区中的“**显示任务窗格**”按钮来启动任务窗格外接程序。在 Excel、PowerPoint 或 Word 中显示外接程序后，你可以通过执行以下操作来调试代码：

1. 在 Excel、PowerPoint 或 Word 中，选择“**插入**”选项卡，然后选择“**我的外接程序**”右侧的向下箭头。

    ![Windows 版 Excel 的“插入”功能区及突出显示的“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)

2. 在可用外接程序列表中，找到“**开发人员外接程序**”部分并选择你的外接程序进行注册。

3. 在 Visual Studio 中，在代码中设置断点。

4. 在 Excel、PowerPoint 或 Word 中，与外接程序进行交互。

5. 在 Visual Studio 中命中断点时，根据需要逐步执行代码。

你可以更改代码并在外接程序中查看这些更改的效果，而无需关闭主机应用程序并重新启动该项目。 保存对代码的更改后，只需在主机应用程序中重新加载外接程序。 例如，通过选择任务窗格的右上角来激活[个性菜单](../design/task-pane-add-ins.md#personality-menu)，然后选择“**重新加载**”，便可重新加载任务窗格外接程序。

### <a name="debug-the-code-for-an-outlook-add-in"></a>调试 Outlook 外接程序的代码

在你已[启动项目](#start-the-project)，且 Visual Studio 启动 Outlook 来托管外接程序后，打开电子邮件或约会项目。 

只要满足激活条件，Outlook 便会为项目激活外接程序。外接程序栏显示在"检查器"窗口或阅读窗格的顶部，Outlook 外接程序显示为外接程序栏中的一个按钮。如果您的外接程序有外接程序命令，那么在默认选项卡或指定的自定义选项卡中将有一个按钮显示在功能区中，而该外接程序将不会显示在外接程序栏中。

若要查看 Outlook 外接程序，请选择对应 Outlook 外接程序的按钮。 在 Outlook 中显示外接程序后，你可以通过执行以下操作来调试代码：

1. 在 Visual Studio 中，在代码中设置断点。

2. 在 Outlook 中，与外接程序进行交互。

3. 在 Visual Studio 中命中断点时，根据需要逐步执行代码。

你可以更改代码并在外接程序中查看这些更改的效果，而无需关闭 Outlook 并重新启动该项目。 保存对代码的更改后，只需打开外接程序的快捷菜单（在 Outlook 中），然后选择“**重新加载**”。

## <a name="next-steps"></a>后续步骤

在外接程序正常工作后，请参阅[部署和发布 Office 外接程序](../publish/publish.md)，以了解可用于将外接程序分发给用户的方法。
