---
title: 在 Visual Studio 中创建和调试 Office 加载项
description: ''
ms.date: 10/01/2018
ms.openlocfilehash: 224a4781b894e9bf165d279c30ca16d18bea956d
ms.sourcegitcommit: c400a220783b03a739449e2d3ff00bbffe5ec7c1
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2018
ms.locfileid: "25681838"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a>在 Visual Studio 中创建和调试 Office 加载项

本文介绍如何使用 Visual Studio 创建第一个 Office 加载项。本文中的步骤基于 Visual Studio 2017。如果使用的是 Visual Studio 的其他版本，操作步骤可能略有不同。

> [!NOTE]
> 若要开始创建 OneNote 加载项，请参阅[生成首个 OneNote 加载项](../onenote/onenote-add-ins-getting-started.md)。

## <a name="create-an-office-add-in-project-in-visual-studio"></a>在 Visual Studio 中创建 Office 加载项项目


首先，请确保已安装 [Office 开发人员工具](https://www.visualstudio.com/features/office-tools-vs.aspx)和一版 Microsoft Office。可以加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)，也可以按照下面的说明操作，以获取[最新版](../develop/install-latest-office-version.md)。

1. 在 Visual Studio 菜单栏中，依次选择**文件** > **新建** > **项目**。
2. 在 **Visual C#** 或 **Visual Basic** 下的项目类型列表中，展开 **Office/SharePoint**，选择 **Web 加载项**，然后选择加载项项目之一。
3. 命名此项目，再选择**确定**以创建项目。

在 Visual Studio 2017 中，选择**确定**后，以下加载项项目模板会有额外选择：

**PowerPoint**
- 你可以选择**将新功能添加到 PowerPoint**，这会创建任务窗格加载项。
- 或者，可以选择**将内容插入 PowerPoint 幻灯片**，这会创建内容加载项。

**Excel** 
- 你可以选择**将新功能添加到 Excel**，这会创建任务窗格加载项。
- 或者，可以选择**将内容插入 Excel 电子表格**，这会创建内容加载项。
    - 如果创建内容加载项，你可以有**基本加载项**的额外选择，这会议最少起始代码创建内容加载项项目。
    - 或者，可以选择**文档可视化加载项**，这包括可视化并绑定到数据的起始代码。

完成该向导后，Visual Studio 会为你创建包含两个项目的解决方案。 你将看到默认 Home.html 页面打开。

|**项目**|**描述**|
|:-----|:-----|
|加载项项目|仅包含一个 XML 清单文件，该文件包含描述你加载项的所有设置。这些设置可帮助 Office 主机确定应何时激活加载项，以及在何处显示加载项。Visual Studio 会为你生成此文件的内容，以便你能够立即运行项目并使用加载项。你可以通过使用清单编辑器来随时更改这些设置。|
|Web 应用程序项目|包含加载项的内容页面，其中包括开发 Office 感知 HTML 和 JavaScript 页面所需的全部文件和文件引用。 在用户开发加载项期间，Visual Studio 在本地 IIS 服务器上托管 Web 应用。 准备好发布时，你必须找到承载此项目的服务器。 如果要了解有关 ASP.NET Web 应用程序项目的更多信息，请参阅 ASP.NET Web 项目。[ ](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx)|

## <a name="modify-your-add-in-settings"></a>修改你的加载项设置


若要修改加载项设置，请编辑项目的 XML 清单文件。 在**解决方案资源管理器**中，展开加载项项目节点、展开包含 XML 清单的文件夹并选择 XML 清单。 你可以指向该文件中的任何元素以查看说明该元素的用途的工具提示。 有关清单文件的详细信息，请参阅 [Office 加载项 XML 清单](../develop/add-in-manifests.md)。


## <a name="develop-the-contents-of-your-add-in"></a>开发加载项的内容

加载项项目允许你修改描述加载项的设置，而 Web 应用程序提供加载项中显示的内容。 

Web 应用程序项目包含可以开始使用的默认 HTML 页面和 JavaScript 文件。 这些文件包含对其他 JavaScript 库的引用，包括适用于 Office 的 JavaScript API。 更新这些文件，并添加更多的 HTML 和 JavaScript 文件可以开发加载项。 下表介绍了默认 HTML 和 JavaScript 文件。

> [!NOTE]
> 根据所用项目模板的类型，下表中的文件可能位于 web 项目的根文件夹或 **Home** 文件夹。

|**文件**|**描述**|
|:-----|:-----|
|**Home.html**|加载项的默认 HTML 页面。 在文档、电子邮件或约会项目中激活此页面时，它会显示为加载项内的第一个页面。 此文件包含你开始使用需要的所有文件引用。 你可以将 HTML 代码添加到此文件，开始开发加载项。|
|**Home.js**|与 Home.html 页面关联的 JavaScript 文件。 你可以将特定于 Home.html 页面的行为的任何代码置于 Home.js 文件中。 Home.js 文件包含一些可帮你入门的示例代码。|
|**Home.css**|定义要应用到加载项的默认样式。 我们建议为设计和样式使用 Office UI Fabric。 有关详细信息，请参阅 [Office 加载项中的 Office UI Fabric](../design/office-ui-fabric.md)。|

> [!NOTE]
> 你无需使用这些文件。 你可以随意将其他文件添加到项目并改为使用这些。 如果你想让其他 HTML 文件显示为加载项的初始页，请打开清单编辑器，然后将 **SourceLocation** 属性指向文件名。

## <a name="debug-your-add-in"></a>调试加载项

Visual Studio 提供生成和调试属性，以帮助调试加载项。

### <a name="review-the-build-and-debug-properties"></a>查看生成和调试属性

在启动解决方案之前，请确认 Visual Studio 将打开你需要的主机应用程序。该信息连同与构建和调试加载项有关的其他几个属性一起显示在项目的属性页中。

### <a name="to-open-the-property-pages-of-a-project"></a>打开项目的属性页

1. 在**解决方案资源管理器**中，选择基本加载项项目（非 Web 项目）。    
2. 在菜单栏上，依次选择**视图** >   **属性窗口**。
    
下表介绍了项目的属性。



|**属性**|**描述**|
|:-----|:-----|
|**启动操作**|指定是否在 Office 桌面客户端或在指定浏览器的 Office Online 客户端调试加载项。|
|**启动文档**（仅限内容和任务窗格加载项）|指定要在启动项目时打开的文档。|
|**Web 项目**|指定与加载项关联的 Web 项目的名称。|
|**电子邮件地址**（仅限 Outlook 加载项）|指定 Exchange Server 或 Exchange Online 中你想用来测试你的 Outlook 加载项的用户帐户的电子邮件地址。|
|**EWS URL**（仅限 Outlook 加载项）|Exchange Web 服务 URL（例如：https://www.contoso.com/ews/exchange.aspx)。 |
|**OWA URL**（仅限 Outlook 加载项）|Outlook Web App URL（例如，https://www.contoso.com/owa)。|
|**用户名**（仅限 Outlook 加载项）|指定 Exchange Server 或 Exchange Online 中的用户帐户名称。|
|**项目文件**|指定包含生成、配置和有关项目的其他信息的文件名称。|
|**项目文件夹**|项目文件的位置。|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a>使用现有文档调试加载项（仅限内容和任务窗格加载项）

你可以将文档添加到加载项项目。如果你有包含要用于加载项的测试数据的文档，Visual Studio 将在你启动项目时为你打开该文档。

### <a name="to-use-an-existing-document-to-debug-the-add-in"></a>使用现有文档调试加载项

1. 在**解决方案资源管理器**中，选择加载项项目文件夹。
    
    > [!NOTE]
    > 选择加载项项目，而不是 Web 应用程序项目。

2. 在**项目**菜单中，选择**添加现有项**。
    
3. 在**添加现有项**对话框中，找到并选择要添加的文档。
    
4. 选择**添加**按钮以向你的项目添加文档。
    
5. 在**解决方案资源管理器**中，选择加载项项目文件夹。
6. 在菜单栏上，依次选择**视图** >  **属性窗口**。
7. 在属性窗口中，选择**启动文档**列表，然后选择添加到项目的文档。 现在，项目将配置为在现有的文档中启动加载项。

### <a name="start-the-solution"></a>启动解决方案

选择**调试** > **启动调试**，从菜单栏启动解决方案。 Visual Studio 将自动生成解决方案，并启动 Office 来承载你的加载项。

当 Visual Studio 生成项目时，它将执行以下任务：

1. 创建 XML 清单文件的副本并将其添加到  _ProjectName_\Output 目录。主机应用程序将在你启动 Visual Studio 并调试加载项时使用此副本。
    
2. 在计算机上创建一组允许加载项在主机应用程序中显示的注册表项。
    
3. 生成网络应用程序项目，然后将其部署到本地 IIS Web 服务器（http://localhost) 
    
接下来，Visual Studio 会执行以下操作：

1. 修改 XML 显示文件的 [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation?view=office-js)元素，通过将 ～remoteAppUrl 标记替换为起始页的完全限定地址（例如，http://localhost/MyAgave.html)）。
    
2. 在 IIS Express 中启动 Web 应用程序项目。
    
3. 打开主机应用程序。 
    
生成项目时，Visual Studio 不会显示“**输出**”窗口中的验证错误。Visual Studio 报告“**错误列表**”窗口中出现的错误和警告。通过在代码和文本编辑器中显示不同颜色的波浪下划线（称为波浪线），Visual Studio 还报告验证错误。通过这些标志，你可以得知 Visual Studio 在代码中检测到的问题。有关详细信息，请参阅 [代码和文本编辑器](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)。有关如何启用或禁用验证的详细信息，请参阅： 

- [选项、文本编辑器、JavaScript 和 IntelliSense](https://docs.microsoft.com/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2015)
    
- [操作方法：为 Visual Web Developer 中的 HTML 编辑设置验证选项](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)
    
- [有关 CSS，请参阅验证、CSS、文本编辑器和“选项”对话框](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)
    
若要查看项目中 XML 清单文件的验证规则，请参阅 [Office 加载项 XML 清单](../develop/add-in-manifests.md)。

### <a name="show-an-add-in-in-excel-or-word-and-step-through-your-code"></a>在 Excel 或 Word 中显示加载项并单步调试代码

如果你将加载项项目的“**启动文档**”属性设置为 Excel 或 Word，Visual Studio 会创建一个新文档，加载项会出现。 如果你将加载项项目的**启动文档**属性设置为使用现有文档，Visual Studio 会打开该文档，但是你必须手动插入加载项。

1. 在 Excel 或 Word 中的**插入**选项卡上个，选择**我的加载项**下拉列表。 从下拉箭头而不是按钮本身选择列表，这将打开 **Office 加载项**对话框。
2. 在**开发人员加载项**下，选择加载项。

然后，在 Visual Studio 中，可以设置中断点并与你的加载项进行交互，逐行执行 HTML 或 JavaScript 文件中的文件。

### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a>在 Outlook 中显示 Outlook 加载项并单步调试代码

若要在 Outlook 中查看加载项，请打开一个电子邮件或约会项目。

只要满足激活条件，Outlook 便会为项目激活加载项。加载项栏显示在"检查器"窗口或阅读窗格的顶部，Outlook 加载项显示为加载项栏中的一个按钮。如果你的加载项有加载项命令，那么在默认选项卡或指定的自定义选项卡中将有一个按钮显示在功能区中，而该加载项将不会显示在加载项栏中。

若要查看 Outlook 加载项，请选择 Outlook 加载项的按钮。

然后，在 Visual Studio 中，可以设置中断点并与你的加载项进行交互，逐行执行 HTML 或 JavaScript 文件中的文件。

你还可以更改代码并在 Outlook 加载项中查看这些更改的效果，而不必关闭 Office 加载项并再次启动项目。在 Outlook 中，只需打开 Outlook 加载项的快捷菜单，然后选择**重新加载**即可。


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a>修改代码并继续调试加载项，而不必再次启动项目

你可以更改代码并在加载项中查看这些更改的效果，无需关闭主机应用程序并重新启动该项目。 更改并保存代码后，打开加载项的快捷菜单，然后选择**重新加载**。
    

## <a name="next-steps"></a>后续步骤

- [部署和发布 Office 加载项](../publish/publish.md)
    
