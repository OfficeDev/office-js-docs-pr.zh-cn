---
title: 在 Visual Studio 中创建和调试 Office 加载项
description: ''
ms.date: 12/04/2017
---


# <a name="create-and-debug-office-add-ins-in-visual-studio"></a>在 Visual Studio 中创建和调试 Office 加载项

本文介绍如何使用 Visual Studio 创建第一个 Office 外接程序。本文中的步骤基于 Visual Studio 2015。如果使用的是 Visual Studio 的其他版本，操作步骤可能略有不同。

> [!NOTE]
> 若要开始创建 OneNote 加载项，请参阅[生成首个 OneNote 加载项](../onenote/onenote-add-ins-getting-started.md)。

## <a name="create-an-office-add-in-project-in-visual-studio"></a>在 Visual Studio 中创建 Office 加载项项目


开始使用前，确保已安装 [Office 开发人员工具](https://www.visualstudio.com/features/office-tools-vs.aspx)和 Microsoft Office 的某个版本。可以加入 [Office 365 开发人员计划](https://dev.office.com/devprogram)，或按照以下说明获取[最新版本](../develop/install-latest-office-version.md)。


1. 在 Visual Studio 菜单栏中，依次选择“文件”**** > “新建”**** > “项目”****。
    
2. 在“**Visual C#**”或“**Visual Basic**”下的项目类型列表中，展开“**Office/SharePoint**”，选择“**Web 外接程序**”，然后选择外接程序项目之一。  
    
3. 命名此项目，再选择“确定”****以创建项目。
    
4. 此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”****中。默认的 Home.html 页面在 Visual Studio 中打开。
    
在 Visual Studio 2015 中，部分加载项项目模板已更新为反映其他功能：


- 内容加载项除了可以显示在 Excel 电子表格中，还可以显示在 Access 和 PowerPoint 文档的正文中。您也可以选择"基本项目"选项，从而可通过最少的起始代码创建基本内容加载项项目，或者选择"文档可视化项目"选项（仅适用于 Access 和 Excel）来创建更多功能全面的内容加载项，其中包含可视化和绑定到数据的起始代码。
    
- Outlook 外接程序包含的选项不仅可用于将您的外接程序包含在电子邮件或约会中，还可用于指定撰写及阅读电子邮件或约会时外接程序是否可用。
    

> [!NOTE]
> 在 Visual Studio 中，大多数选项的含义都可以根据说明进行理解，但“电子邮件”****复选框除外。若要创建 Outlook 加载项，不仅会与邮件项一起出现，还会与会议请求、响应和取消一起出现，请选中此复选框。

完成向导后，Visual Studio 便会创建解决方案，其中包含两个项目。



|**项目**|**说明**|
|:-----|:-----|
|外接程序项目|仅包含一个 XML 清单文件，该文件包含描述您加载项的所有设置。这些设置可帮助 Office 主机确定应何时激活加载项，以及在何处显示加载项。Visual Studio 会为您生成此文件的内容，以便您能够立即运行项目并使用加载项。您可以通过使用清单编辑器来随时更改这些设置。|
|Web 应用程序项目|包含加载项的内容页面，包括开发可识别 Office 的 HTML 和 JavaScript 页面所需的所有文件和文件引用。在您开发加载项时，Visual Studio 会在本地 IIS 服务器上承载 Web 应用程序。准备好进行发布后，必须找出一个服务器来承载此项目。如果要了解有关 ASP.NET Web 应用程序项目的更多信息，请参阅 [ASP.NET Web 项目](http://msdn.microsoft.com/zh-cn/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx)。|

## <a name="modify-your-add-in-settings"></a>修改您的外接程序设置


若要修改外接程序设置，请编辑项目的 XML 清单文件。在“**解决方案资源管理器**”中，展开外接程序项目节点、展开包含 XML 清单的文件夹并选择 XML 清单。你可以指向该文件中的任何元素以查看说明该元素用途的工具提示。有关清单文件的详细信息，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。


## <a name="develop-the-contents-of-your-add-in"></a>开发外接程序的内容


加载项项目允许您修改描述加载项的设置，而 Web 应用程序提供加载项中显示的内容。 

Web 应用程序项目包含一个可用于入门的默认 HTML 页和 Javascriptshort 文件。该项目也包含您向项目添加的所有页面所共有的一个 JavaScript 文件。这些文件包含对其他 JavaScript 库（包括适用于 Office 的 JavaScript API）的引用，因此很方便。 

随着您的加载项变得更复杂，可以添加更多 HTML 和 JavaScript 文件。您可以将默认 HTML 和 JavaScript 文件的内容用作引用类型的示例，您可能希望将该类型添加到项目中的其他页以使其与您的加载项一起工作。下表介绍了默认 HTML 和 JavaScript 文件。



|**文件**|**说明**|
|:-----|:-----|
|**Home.html**|位于项目的**主**文件夹中，此为外接程序的默认 HTML 页面。在文档、电子邮件或约会项目中激活此页面时，它会显示为外接程序内的第一个页面。此文件很方便，因为它包含你入门所需的所有文件引用。准备好创建第一个外接程序时，只需向此文件添加 HTML 代码即可。|
|**Home.js**|位于项目的**主**文件夹中，此为与 Home.js 页面相关联的 JavaScript 文件。你可以将特定于 Home.html 页面的行为的任何代码置于 Home.js 文件中。Home.js 文件包含一些可帮你入门的示例代码。|
|**App.js**|位于项目的**外接程序**文件夹中，此为整个外接程序的默认 JavaScript 文件。你可以将对你外接程序的多个页面的行为通用的代码置于 App.js 文件中。App.js 文件包含一些可帮你入门的示例代码。|

> [!NOTE]
> 不一定要使用这些文件。可以随意向项目中添加其他文件，并改用这些文件。若要让其他 HTML 文件显示为加载项的初始网页，请打开清单编辑器，再将“SourceLocation”****属性指向相应的文件名称。


## <a name="debug-your-add-in"></a>调试加载项


当您准备启动加载项时，请查看与构建和调试相关的属性，然后启动解决方案。


### <a name="review-the-build-and-debug-properties"></a>查看生成和调试属性

在启动解决方案之前，请确认 Visual Studio 将打开您需要的主机应用程序。该信息连同与构建和调试加载项有关的其他几个属性一起显示在项目的属性页中。


### <a name="to-open-the-property-pages-of-a-project"></a>打开项目的属性页


1. 在“**解决方案资源管理器**”中，选择项目名称。
    
2. 在菜单栏上，依次选择“**视图**”和“**属性窗口**”。
    
下表介绍了项目的属性。



|**属性**|**说明**|
|:-----|:-----|
|**启动操作**|指定是否在 Office 桌面客户端或在指定浏览器的 Office Online 客户端调试外接程序。|
|**启动文档**（仅限内容和任务窗格加载项）|指定要在启动项目时打开的文档。|
|**Web 项目**|指定与外接程序关联的 Web 项目的名称。|
|**电子邮件地址**（仅限 Outlook 外接程序）|指定 Exchange Server 或 Exchange Online 中您想用来测试您的 Outlook 外接程序的用户帐户的电子邮件地址。|
|**EWS Url**（仅限 Outlook 外接程序）|Exchange Web 服务 URL（例如：https://www.contoso.com/ews/exchange.aspx）。 |
|**OWA Url**（仅限 Outlook 外接程序）|Outlook Web App URL（例如：https://www.contoso.com/owa）。|
|**用户名**（仅限 Outlook 外接程序）|指定 Exchange Server 或 Exchange Online 中的用户帐户名称。|
|**项目文件**|指定包含生成、配置和有关项目的其他信息的文件名称。|
|**项目文件夹**|项目文件的位置。|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a>使用现有文档调试加载项（仅限内容和任务窗格加载项）


您可以将文档添加到加载项项目。如果您有包含要用于加载项的测试数据的文档，Visual Studio 将在您启动项目时为您打开该文档。


### <a name="to-use-an-existing-document-to-debug-the-add-in"></a>使用现有文档调试加载项


1. 在“解决方案资源管理器”****中，选择加载项项目文件夹。
    
    > [!NOTE]
    > 选择加载项项目，而不是 Web 应用项目。

2. 在“项目”****菜单中，选择“添加现有项”****。
    
3. 在“**添加现有项**”对话框中，找到并选择要添加的文档。
    
4. 选择“**添加**”按钮以向你的项目添加文档。
    
5. 在“**解决方案资源管理器**”中，打开项目的快捷菜单，然后选择“**属性**”。
    
    显示项目的属性页。
    
6. 在“**启动文档**”列表中，选择要添加到项目的文档，然后选择“**确定**”按钮关闭属性页。
    

### <a name="start-the-solution"></a>启动解决方案


启动 Visual Studio 时将自动生成解决方案。你可以通过依次选择“**调试**、“**启动**”，从“**菜单**”栏中启动解决方案。 


> [!NOTE]
> 如果 Internet Explorer 中未启用脚本调试，将无法在 Visual Studio 中启动调试器。若要启用脚本调试，可以打开“Internet 选项”****对话框，选择“高级”****选项卡，再清除“禁用脚本调试(Internet Explorer)”****和“禁用脚本调试(其他)”****复选框。

此时，Visual Studio 生成项目，并执行以下操作：


1. 创建 XML 清单文件的副本并将其添加到  _ProjectName_\Output 目录。主机应用程序将在您启动 Visual Studio 并调试加载项时使用此副本。
    
2. 在计算机上创建一组允许加载项在主机应用程序中显示的注册表项。
    
3. 生成 Web 应用程序项目，然后将其部署到本地 IIS Web 服务器 (http://localhost)。 
    
接下来，Visual Studio 会执行以下操作：


1. 通过将 ~remoteAppUrl 令牌替换为起始页的完全限定地址（例如，http://localhost/MyAgave.html）修改 XML 清单文件的 [SourceLocation](http://msdn.microsoft.com/zh-cn/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx) 元素。
    
2. 在 IIS Express 中启动 Web 应用程序项目。
    
3. 打开主机应用程序。 
    
生成项目时，Visual Studio 不会显示“**输出**”窗口中的验证错误。Visual Studio 报告“**错误列表**”窗口中出现的错误和警告。通过在代码和文本编辑器中显示不同颜色的波浪下划线（称为波浪线），Visual Studio 还报告验证错误。通过这些标志，你可以得知 Visual Studio 在代码中检测到的问题。有关详细信息，请参阅 [代码和文本编辑器](https://msdn.microsoft.com/zh-cn/library/se2f663y(v=vs.140).aspx)。有关如何启用或禁用验证的详细信息，请参阅： 

- 
  [选项、文本编辑器、JavaScript 和 IntelliSense](https://msdn.microsoft.com/zh-cn/library/hh362485(v=vs.140).aspx)
    
- 
  [操作方法：为 Visual Web Developer 中的 HTML 编辑设置验证选项](https://msdn.microsoft.com/zh-cn/library/0byxkfet(v=vs.100).aspx)
    
- 
  [有关 CSS，请参阅验证、CSS、文本编辑器和“选项”对话框](https://msdn.microsoft.com/zh-cn/library/se2f663y(v=vs.140).aspx)
    
若要查看项目中 XML 清单文件的验证规则，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。


### <a name="show-an-add-in-in-excel-word-or-project-and-step-through-your-code"></a>在 Excel、Word 或 Project 中显示加载项并单步调试代码


如果将外接程序项目的“**启动文档**”属性设置为 Excel 或 Word，Visual Studio 会创建一个新文档，外接程序会出现。如果将外接程序项目的“**启动文档**”属性设置为使用现有文档，Visual Studio 会打开该文档，但是你必须手动插入外接程序。如果将“**启动文档**”设置为“**Microsoft Project**”，则还需要手动插入外接程序。


### <a name="to-show-an-office-add-in-in-excel-or-word"></a>在 Excel 或 Word 中显示 Office 外接程序


1. 在 Excel 或 Word 中的“**插入**”选项卡上，选择“**Office 外接程序**”。
    
2. 在出现的列表中选择您的加载项。
    

### <a name="to-show-an-office-add-in-in-project"></a>在 Project 中显示 Office 外接程序


1. 在 Project 中的“**项目**”选项卡上，选择“**Office 外接程序**”。
    
2. 在出现的列表中选择您的加载项。
    
在 Visual Studio 中，您随后可以设置断点。然后，当您与加载项交互时，可对 HTML、JavaScript 和 C# 或 VB 代码文件中的代码进行单步调试。


### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a>在 Outlook 中显示 Outlook 外接程序并单步调试代码


若要在 Outlook 中查看加载项，请打开一个电子邮件或约会项目。

只要满足激活条件，Outlook 便会为项目激活外接程序。外接程序栏显示在"检查器"窗口或阅读窗格的顶部，Outlook 外接程序显示为外接程序栏中的一个按钮。如果您的外接程序有外接程序命令，那么在默认选项卡或指定的自定义选项卡中将有一个按钮显示在功能区中，而该外接程序将不会显示在外接程序栏中。

若要查看 Outlook 外接程序，请选择 Outlook 外接程序的按钮。

在 Visual Studio 中，可以设置断点。然后，与 Outlook 外接程序交互并逐句调试 HTML、JavaScript 和 C# 或 VB 代码文件中的代码。 

你还可以更改代码并在 Outlook 外接程序中查看这些更改的效果，而不必关闭 Office 外接程序并再次启动项目。在 Outlook 中，只需打开 Outlook 外接程序的快捷菜单，然后选择“**重新加载**”即可。


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a>修改代码并继续调试加载项，而不必再次启动项目


你可以更改代码并在外接程序中查看这些更改的效果，无需关闭主机应用程序并重新启动该项目。更改代码后，打开外接程序的快捷菜单，然后选择“**重新加载**”。当重新加载外接程序时，它会与 Visual Studio 调试器断开连接。因此，你可以查看所做更改的效果，但是在将 Visual Studio 调试器附加到所有可用的 Iexplore.exe 进程之前，将无法再次单步执行代码。


### <a name="to-attach-the-visual-studio-debugger-to-all-of-the-available-iexploreexe-processes"></a>将 Visual Studio 调试器附加到所有可用的 Iexplore.exe 进程


1. 在 Visual Studio 中，依次选择“**调试**”、“**附加到进程**”。
    
2. 在“**附加到进程**”对话框中，选择所有可用的“**Iexplore.exe**”进程，然后选择“**附加**”按钮。
    

## <a name="next-steps"></a>后续步骤

- [部署和发布 Office 外接程序](../publish/publish.md)
    
