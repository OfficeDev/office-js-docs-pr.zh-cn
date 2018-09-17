---
title: 在 Visual Studio 中创建和调试 Office 加载项
description: ''
ms.date: 03/14/2018
ms.openlocfilehash: 2e5c08a72ec97e26000d6ea7e53dd1d0f2c9e6dc
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945353"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a><span data-ttu-id="9e084-102">在 Visual Studio 中创建和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="9e084-102">Create and debug Office Add-ins in Visual Studio</span></span>

<span data-ttu-id="9e084-p101">本文介绍如何使用 Visual Studio 创建第一个 Office 外接程序。本文中的步骤基于 Visual Studio 2015。如果使用的是 Visual Studio 的其他版本，操作步骤可能略有不同。</span><span class="sxs-lookup"><span data-stu-id="9e084-p101">This article describes how to use Visual Studio to create your first Office Add-in. The steps in this article based on Visual Studio 2015. If you're using another version of Visual Studio, the procedures might vary slightly.</span></span>

> [!NOTE]
> <span data-ttu-id="9e084-106">若要开始创建 OneNote 加载项，请参阅[生成首个 OneNote 加载项](../onenote/onenote-add-ins-getting-started.md)。</span><span class="sxs-lookup"><span data-stu-id="9e084-106">To get started with an add-in for OneNote, see [Build your first OneNote add-in](../onenote/onenote-add-ins-getting-started.md).</span></span>

## <a name="create-an-office-add-in-project-in-visual-studio"></a><span data-ttu-id="9e084-107">在 Visual Studio 中创建 Office 加载项项目</span><span class="sxs-lookup"><span data-stu-id="9e084-107">Create an Office Add-in project in Visual Studio</span></span>


<span data-ttu-id="9e084-p102">首先，请确保已安装 [Office 开发人员工具](https://www.visualstudio.com/features/office-tools-vs.aspx)和一版 Microsoft Office。可以加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)，也可以按照下面的说明操作，以获取[最新版](../develop/install-latest-office-version.md)。</span><span class="sxs-lookup"><span data-stu-id="9e084-p102">To get started, make sure you have the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) installed, and a version of Microsoft Office. You can join the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program), or follow these instructions to get the [latest version](../develop/install-latest-office-version.md).</span></span>


1. <span data-ttu-id="9e084-110">在 Visual Studio 菜单栏中，依次选择“文件”\*\*\*\* > “新建”\*\*\*\* > “项目”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="9e084-110">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="9e084-111">在**Visual C#** 或**Visual Basic**下的项目类型列表中，展开**Office/SharePoint**，选择**Web 外接程序**，然后选择外接程序项目之一。</span><span class="sxs-lookup"><span data-stu-id="9e084-111">In the list of project types under  **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose  **Web Add-ins**, and then select one of the Add-in projects.</span></span>  
    
3. <span data-ttu-id="9e084-112">命名此项目，再选择“确定”\*\*\*\* 以创建项目。</span><span class="sxs-lookup"><span data-stu-id="9e084-112">Name the project, and then choose  **OK** to create the project.</span></span>
    
4. <span data-ttu-id="9e084-p103">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”\*\*\*\* 中。默认的 Home.html 页面在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="9e084-p103">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The default Home.html page opens in Visual Studio.</span></span>
    
<span data-ttu-id="9e084-115">在 Visual Studio 2015 中，部分加载项项目模板已更新为反映其他功能：</span><span class="sxs-lookup"><span data-stu-id="9e084-115">In Visual Studio 2015, some of the add-in project templates have been updated to reflect additional functionality:</span></span>


- <span data-ttu-id="9e084-p104">内容加载项除了可以显示在 Excel 电子表格中，还可以显示在 Access 和 PowerPoint 文档的正文中。您也可以选择"基本项目"选项，从而可通过最少的起始代码创建基本内容加载项项目，或者选择"文档可视化项目"选项（仅适用于 Access 和 Excel）来创建更多功能全面的内容加载项，其中包含可视化和绑定到数据的起始代码。</span><span class="sxs-lookup"><span data-stu-id="9e084-p104">Content add-ins can appear in the body of Access and PowerPoint documents, in addition to Excel spreadsheets. You can also choose the Basic Project option to create a basic content add-in project with minimal starter code, or the Document Visualization Project option (for Access and Excel only) to create a more full-featured content add-in that includes starter code to visualize and bind to data.</span></span>
    
- <span data-ttu-id="9e084-118">Outlook 外接程序包含的选项不仅可用于将您的外接程序包含在电子邮件或约会中，还可用于指定撰写及阅读电子邮件或约会时外接程序是否可用。</span><span class="sxs-lookup"><span data-stu-id="9e084-118">Outlook add-ins include options not just for including your add-in in email messages or appointments, but also for specifying whether the add-in is available when an email message or appointment is being composed as well as read.</span></span>
    

> [!NOTE]
> <span data-ttu-id="9e084-p105">在 Visual Studio 中，大多数选项的含义都可以根据说明进行理解，但“电子邮件”\*\*\*\* 复选框除外。若要创建 Outlook 加载项，不仅会与邮件项一起出现，还会与会议请求、响应和取消一起出现，请选中此复选框。</span><span class="sxs-lookup"><span data-stu-id="9e084-p105">In Visual Studio most options are understandable from their descriptions except for the  **Email Message** checkbox. Use that checkbox if you want to create an Outlook add-in that appears not just with mail items, but also with meeting requests, responses, and cancellations.</span></span>

<span data-ttu-id="9e084-121">完成向导后，Visual Studio 便会创建解决方案，其中包含两个项目。</span><span class="sxs-lookup"><span data-stu-id="9e084-121">When you've completed the wizard, Visual Studio creates a solution for you that contains two projects.</span></span>



|<span data-ttu-id="9e084-122">**项目**</span><span class="sxs-lookup"><span data-stu-id="9e084-122">**Project**</span></span>|<span data-ttu-id="9e084-123">**说明**</span><span class="sxs-lookup"><span data-stu-id="9e084-123">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="9e084-124">加载项项目</span><span class="sxs-lookup"><span data-stu-id="9e084-124">Add-in project</span></span>|<span data-ttu-id="9e084-p106">仅包含一个 XML 清单文件，该文件包含描述您加载项的所有设置。这些设置可帮助 Office 主机确定应何时激活加载项，以及在何处显示加载项。Visual Studio 会为您生成此文件的内容，以便您能够立即运行项目并使用加载项。您可以通过使用清单编辑器来随时更改这些设置。</span><span class="sxs-lookup"><span data-stu-id="9e084-p106">Contains only an XML manifest file, which contains all the settings that describe your add-in. These settings help the Office host determine when your add-in should be activated and where the add-in should appear. Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately. You change these settings any time by using the Manifest editor.</span></span>|
|<span data-ttu-id="9e084-129">Web 应用程序项目</span><span class="sxs-lookup"><span data-stu-id="9e084-129">Web application project</span></span>|<span data-ttu-id="9e084-p107">包含加载项的内容页面，包括开发可识别 Office 的 HTML 和 JavaScript 页面所需的所有文件和文件引用。在您开发加载项时，Visual Studio 会在本地 IIS 服务器上承载 Web 应用程序。准备好进行发布后，必须找出一个服务器来承载此项目。如果要了解有关 ASP.NET Web 应用程序项目的更多信息，请参阅 [ASP.NET Web 项目](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx)。</span><span class="sxs-lookup"><span data-stu-id="9e084-p107">Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages. While you develop your add-in, Visual Studio hosts the web application on your local IIS server. When you're ready to publish, you'll have to find a server to host this project.To learn more about ASP.NET web application projects, see [ASP.NET Web Projects](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).</span></span>|

## <a name="modify-your-add-in-settings"></a><span data-ttu-id="9e084-133">修改您的外接程序设置</span><span class="sxs-lookup"><span data-stu-id="9e084-133">Modify your add-in settings</span></span>


<span data-ttu-id="9e084-p108">若要修改外接程序设置，请编辑项目的 XML 清单文件。在“**解决方案资源管理器**”中，展开外接程序项目节点、展开包含 XML 清单的文件夹并选择 XML 清单。你可以指向该文件中的任何元素以查看说明该元素用途的工具提示。有关清单文件的详细信息，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="9e084-p108">To modify the settings of your add-in, edit the XML manifest file of the project. In  **Solution Explorer**, expand the add-in project node, expand the folder that contains the XML manifest, and choose the XML manifest. You can point to any element in the file to view a tooltip that describes the purpose of the element. For more information about the manfiest file, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>


## <a name="develop-the-contents-of-your-add-in"></a><span data-ttu-id="9e084-138">开发外接程序的内容</span><span class="sxs-lookup"><span data-stu-id="9e084-138">Develop the contents of your add-in</span></span>


<span data-ttu-id="9e084-139">加载项项目允许您修改描述加载项的设置，而 Web 应用程序提供加载项中显示的内容。</span><span class="sxs-lookup"><span data-stu-id="9e084-139">While the add-in project lets you modify the settings that describe your add-in, the web application provides the content that appears in the add-in.</span></span> 

<span data-ttu-id="9e084-p109">Web 应用程序项目包含一个可用于入门的默认 HTML 页和 Javascriptshort 文件。该项目也包含您向项目添加的所有页面所共有的一个 JavaScript 文件。这些文件包含对其他 JavaScript 库（包括适用于 Office 的 JavaScript API）的引用，因此很方便。</span><span class="sxs-lookup"><span data-stu-id="9e084-p109">The web application project contains a default HTML page and JavaScript file that you can use to get started. The project also contains a JavaScript file that is common to all pages that you add to your project. These files are convenient because they contain references to other JavaScript libraries including the JavaScript API for Office.</span></span> 

<span data-ttu-id="9e084-p110">随着您的加载项变得更复杂，可以添加更多 HTML 和 JavaScript 文件。您可以将默认 HTML 和 JavaScript 文件的内容用作引用类型的示例，您可能希望将该类型添加到项目中的其他页以使其与您的加载项一起工作。下表介绍了默认 HTML 和 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="9e084-p110">As your add-in becomes more sophisticated, you can add more HTML and JavaScript files. You can use the contents of the default HTML and JavaScript files as examples of the types of references you might want to add to other pages in your project to make them work with your add-in. The following table describes default HTML and JavaScript files.</span></span>



|<span data-ttu-id="9e084-146">**文件**</span><span class="sxs-lookup"><span data-stu-id="9e084-146">**File**</span></span>|<span data-ttu-id="9e084-147">**说明**</span><span class="sxs-lookup"><span data-stu-id="9e084-147">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="9e084-148">**Home.html**</span><span class="sxs-lookup"><span data-stu-id="9e084-148">**Home.html**</span></span>|<span data-ttu-id="9e084-p111">位于项目的**主**文件夹中，此为外接程序的默认 HTML 页面。在文档、电子邮件或约会项目中激活此页面时，它会显示为外接程序内的第一个页面。此文件很方便，因为它包含你入门所需的所有文件引用。准备好创建第一个外接程序时，只需向此文件添加 HTML 代码即可。</span><span class="sxs-lookup"><span data-stu-id="9e084-p111">Located in the  **Home** folder of the project, this is default HTML page of the add-in. This page appears as the first page inside of the add-in when it is activated in a document, email message or appointment item. This file is convenient because it contains all of the file references that you need to get started. When you are ready to create your first add-in, just add your HTML code to this file.</span></span>|
|<span data-ttu-id="9e084-153">**Home.js**</span><span class="sxs-lookup"><span data-stu-id="9e084-153">**Home.js**</span></span>|<span data-ttu-id="9e084-p112">位于项目的**主**文件夹中，此为与 Home.js 页面相关联的 JavaScript 文件。你可以将特定于 Home.html 页面的行为的任何代码置于 Home.js 文件中。Home.js 文件包含一些可帮你入门的示例代码。</span><span class="sxs-lookup"><span data-stu-id="9e084-p112">Located in the  **Home** folder of the project, this is the JavaScript file associated with the Home.js page. You can place any code that is specific to the behavior of the Home.html page in the Home.js file. The Home.js file contains some example code to get you started.</span></span>|
|<span data-ttu-id="9e084-157">**App.js**</span><span class="sxs-lookup"><span data-stu-id="9e084-157">**App.js**</span></span>|<span data-ttu-id="9e084-p113">位于项目的**外接程序**文件夹中，此为整个外接程序的默认 JavaScript 文件。你可以将对你外接程序的多个页面的行为通用的代码置于 App.js 文件中。App.js 文件包含一些可帮你入门的示例代码。</span><span class="sxs-lookup"><span data-stu-id="9e084-p113">Located in the  **Add-in** folder of the project, this is the default JavaScript file of the entire add-in. You can place code that is common to the behavior of multiple pages of your add-in in the App.js file. The App.js file contains some example code to get you started.</span></span>|

> [!NOTE]
> <span data-ttu-id="9e084-p114">不一定要使用这些文件。可以随意向项目中添加其他文件，并改用这些文件。若要让其他 HTML 文件显示为加载项的初始网页，请打开清单编辑器，再将“SourceLocation”\*\*\*\* 属性指向相应的文件名称。</span><span class="sxs-lookup"><span data-stu-id="9e084-p114">You don't have to use these files. Feel free to add other files to the project and use those instead. If you want another HTML file to appear as the initial page of the add-in, open the manifest editor, and then point the  **SourceLocation** property to the name of the file.</span></span>


## <a name="debug-your-add-in"></a><span data-ttu-id="9e084-164">调试加载项</span><span class="sxs-lookup"><span data-stu-id="9e084-164">Debug your add-in</span></span>


<span data-ttu-id="9e084-165">当您准备启动加载项时，请查看与构建和调试相关的属性，然后启动解决方案。</span><span class="sxs-lookup"><span data-stu-id="9e084-165">When you are ready to start your add-in, review build and debug related properties, and then start the solution.</span></span>


### <a name="review-the-build-and-debug-properties"></a><span data-ttu-id="9e084-166">查看生成和调试属性</span><span class="sxs-lookup"><span data-stu-id="9e084-166">Review the build and debug properties</span></span>

<span data-ttu-id="9e084-p115">在启动解决方案之前，请确认 Visual Studio 将打开您需要的主机应用程序。该信息连同与构建和调试加载项有关的其他几个属性一起显示在项目的属性页中。</span><span class="sxs-lookup"><span data-stu-id="9e084-p115">Before you start the solution, verify that Visual Studio will open the host application that you want. That information appears in the property pages of the project along with several other properties that relate to building and debugging the add-in.</span></span>


### <a name="to-open-the-property-pages-of-a-project"></a><span data-ttu-id="9e084-169">打开项目的属性页</span><span class="sxs-lookup"><span data-stu-id="9e084-169">To open the property pages of a project</span></span>


1. <span data-ttu-id="9e084-170">在 **解决方案资源管理器**中，选择基本加载项项目（非 Web 项目）。</span><span class="sxs-lookup"><span data-stu-id="9e084-170">In  **Solution Explorer**, choose the basic add-in project (not the Web project).</span></span>
    
2. <span data-ttu-id="9e084-171">在菜单栏上，依次选择“**视图**”和“**属性窗口**”。</span><span class="sxs-lookup"><span data-stu-id="9e084-171">On the menu bar, choose  **View**,  **Properties Window**.</span></span>
    
<span data-ttu-id="9e084-172">下表介绍了项目的属性。</span><span class="sxs-lookup"><span data-stu-id="9e084-172">The following table describes the properties of the project.</span></span>



|<span data-ttu-id="9e084-173">**属性**</span><span class="sxs-lookup"><span data-stu-id="9e084-173">**Property**</span></span>|<span data-ttu-id="9e084-174">**说明**</span><span class="sxs-lookup"><span data-stu-id="9e084-174">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="9e084-175">**启动操作**</span><span class="sxs-lookup"><span data-stu-id="9e084-175">**Start Action**</span></span>|<span data-ttu-id="9e084-176">指定是否在 Office 桌面客户端或在指定浏览器的 Office Online 客户端调试外接程序。</span><span class="sxs-lookup"><span data-stu-id="9e084-176">Specifies whether to debug your add-in in an Office desktop client or in an Office Online client in the specified browser.</span></span>|
|<span data-ttu-id="9e084-177">**启动文档**（仅限内容和任务窗格加载项）</span><span class="sxs-lookup"><span data-stu-id="9e084-177">**Start Document** (Content and task pane add-ins only)</span></span>|<span data-ttu-id="9e084-178">指定要在启动项目时打开的文档。</span><span class="sxs-lookup"><span data-stu-id="9e084-178">Specifies what document to open when you start the project.</span></span>|
|<span data-ttu-id="9e084-179">**Web 项目**</span><span class="sxs-lookup"><span data-stu-id="9e084-179">**Web Project**</span></span>|<span data-ttu-id="9e084-180">指定与外接程序关联的 Web 项目的名称。</span><span class="sxs-lookup"><span data-stu-id="9e084-180">Specifies the name of the web project associated with the add-in.</span></span>|
|<span data-ttu-id="9e084-181">**电子邮件地址**（仅限 Outlook 外接程序）</span><span class="sxs-lookup"><span data-stu-id="9e084-181">**Email Address** (Outlook add-ins only)</span></span>|<span data-ttu-id="9e084-182">指定 Exchange Server 或 Exchange Online 中您想用来测试您的 Outlook 外接程序的用户帐户的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="9e084-182">Specifies the email address of the user account in Exchange Server or Exchange Online that you want to test your Outlook add-in with.</span></span>|
|<span data-ttu-id="9e084-183">**EWS URL**（仅限 Outlook 加载项）</span><span class="sxs-lookup"><span data-stu-id="9e084-183">**EWS Url** (Outlook add-ins only)</span></span>|<span data-ttu-id="9e084-184">Exchange Web 服务 URL（例如：https://www.contoso.com/ews/exchange.aspx)。</span><span class="sxs-lookup"><span data-stu-id="9e084-184">Exchange Web service URL (For example: https://www.contoso.com/ews/exchange.aspx).</span></span> |
|<span data-ttu-id="9e084-185">**OWA URL**（仅限 Outlook 加载项）</span><span class="sxs-lookup"><span data-stu-id="9e084-185">**OWA Url** (Outlook add-ins only)</span></span>|<span data-ttu-id="9e084-186">Outlook Web App URL（例如，https://www.contoso.com/owa)。</span><span class="sxs-lookup"><span data-stu-id="9e084-186">Outlook Web App URL (For example: https://www.contoso.com/owa).</span></span>|
|<span data-ttu-id="9e084-187">**用户名**（仅限 Outlook 加载项）</span><span class="sxs-lookup"><span data-stu-id="9e084-187">**User name** (Outlook add-ins only)</span></span>|<span data-ttu-id="9e084-188">指定 Exchange Server 或 Exchange Online 中的用户帐户名称。</span><span class="sxs-lookup"><span data-stu-id="9e084-188">Specifies the name of your user account in Exchange Server or Exchange Online.</span></span>|
|<span data-ttu-id="9e084-189">**项目文件**</span><span class="sxs-lookup"><span data-stu-id="9e084-189">**Project File**</span></span>|<span data-ttu-id="9e084-190">指定包含生成、配置和有关项目的其他信息的文件名称。</span><span class="sxs-lookup"><span data-stu-id="9e084-190">Specifies the name of the file containing build, configuration, and other information about the project.</span></span>|
|<span data-ttu-id="9e084-191">**项目文件夹**</span><span class="sxs-lookup"><span data-stu-id="9e084-191">**Project Folder**</span></span>|<span data-ttu-id="9e084-192">项目文件的位置。</span><span class="sxs-lookup"><span data-stu-id="9e084-192">The location of the project file.</span></span>|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a><span data-ttu-id="9e084-193">使用现有文档调试加载项（仅限内容和任务窗格加载项）</span><span class="sxs-lookup"><span data-stu-id="9e084-193">Use an existing document to debug the add-in (content and task pane add-ins only)</span></span>


<span data-ttu-id="9e084-p116">您可以将文档添加到加载项项目。如果您有包含要用于加载项的测试数据的文档，Visual Studio 将在您启动项目时为您打开该文档。</span><span class="sxs-lookup"><span data-stu-id="9e084-p116">You can add documents to the add-in project. If you have a document that contains test data that you want to use with your add-in, Visual Studio opens that document for you when you start the project.</span></span>


### <a name="to-use-an-existing-document-to-debug-the-add-in"></a><span data-ttu-id="9e084-196">使用现有文档调试加载项</span><span class="sxs-lookup"><span data-stu-id="9e084-196">To use an existing document to debug the add-in</span></span>


1. <span data-ttu-id="9e084-197">在“解决方案资源管理器”\*\*\*\* 中，选择加载项项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="9e084-197">In  **Solution Explorer**, choose the add-in project folder.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="9e084-198">选择加载项项目，而不是 Web 应用项目。</span><span class="sxs-lookup"><span data-stu-id="9e084-198">Choose the add-in project and not the web application project.</span></span>

2. <span data-ttu-id="9e084-199">在“项目”\*\*\*\* 菜单中，选择“添加现有项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="9e084-199">On the  **Project** menu, choose **Add Existing Item**.</span></span>
    
3. <span data-ttu-id="9e084-200">在“**添加现有项**”对话框中，找到并选择要添加的文档。</span><span class="sxs-lookup"><span data-stu-id="9e084-200">In the  **Add Existing Item** dialog box, locate and select the document that you want to add.</span></span>
    
4. <span data-ttu-id="9e084-201">选择“**添加**”按钮以向你的项目添加文档。</span><span class="sxs-lookup"><span data-stu-id="9e084-201">Choose the  **Add** button to add the document to your project.</span></span>
    
5. <span data-ttu-id="9e084-202">在“**解决方案资源管理器**”中，打开项目的快捷菜单，然后选择“**属性**”。</span><span class="sxs-lookup"><span data-stu-id="9e084-202">In  **Solution Explorer**, open the shortcut menu for the project, and then choose  **Properties**.</span></span>
    
    <span data-ttu-id="9e084-203">显示项目的属性页。</span><span class="sxs-lookup"><span data-stu-id="9e084-203">The property pages for the project appear.</span></span>
    
6. <span data-ttu-id="9e084-204">在“**启动文档**”列表中，选择要添加到项目的文档，然后选择“**确定**”按钮关闭属性页。</span><span class="sxs-lookup"><span data-stu-id="9e084-204">In the  **Start Document** list, choose the document that you added to the project, and then choose the **OK** button to close the property pages.</span></span>
    

### <a name="start-the-solution"></a><span data-ttu-id="9e084-205">启动解决方案</span><span class="sxs-lookup"><span data-stu-id="9e084-205">Start the solution</span></span>


<span data-ttu-id="9e084-p117">启动 Visual Studio 时将自动生成解决方案。你可以通过依次选择“**调试**、“**启动**”，从“**菜单**”栏中启动解决方案。</span><span class="sxs-lookup"><span data-stu-id="9e084-p117">Visual Studio will automatically build the solution when you start it. You can start the solution from the  **Menu** bar by choosing **Debug**,  **Start**.</span></span> 


> [!NOTE]
> <span data-ttu-id="9e084-p118">如果 Internet Explorer 中未启用脚本调试，将无法在 Visual Studio 中启动调试器。若要启用脚本调试，可以打开“Internet 选项”\*\*\*\* 对话框，选择“高级”\*\*\*\* 选项卡，再清除“禁用脚本调试(Internet Explorer)”\*\*\*\* 和“禁用脚本调试(其他)”\*\*\*\* 复选框。</span><span class="sxs-lookup"><span data-stu-id="9e084-p118">If script debugging isn't enabled in Internet Explorer, you won't be able to start the debugger in Visual Studio. You can enable script debugging by opening the  **Internet Options** dialog box, choosing the **Advanced** tab, and then clearing the **Disable Script Debugging (Internet Explorer)** and **Disable Script Debugging (Other)** check boxes.</span></span>

<span data-ttu-id="9e084-210">此时，Visual Studio 生成项目，并执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="9e084-210">Visual Studio builds the project and does the following:</span></span>


1. <span data-ttu-id="9e084-p119">创建 XML 清单文件的副本并将其添加到  _ProjectName_\Output 目录。主机应用程序将在您启动 Visual Studio 并调试加载项时使用此副本。</span><span class="sxs-lookup"><span data-stu-id="9e084-p119">Creates a copy of the XML manifest file and adds it to  _ProjectName_\Output directory. The host application consumes this copy when you start Visual Studio and debug the add-in.</span></span>
    
2. <span data-ttu-id="9e084-213">在计算机上创建一组允许加载项在主机应用程序中显示的注册表项。</span><span class="sxs-lookup"><span data-stu-id="9e084-213">Creates a set of registry entries on your computer that enable the add-in to appear in the host application.</span></span>
    
3. <span data-ttu-id="9e084-214">生成网络应用程序项目，然后将其部署到本地 IIS Web 服务器（http://localhost)</span><span class="sxs-lookup"><span data-stu-id="9e084-214">Builds the web application project, and then deploys it to the local IIS web server (http://localhost).</span></span> 
    
<span data-ttu-id="9e084-215">接下来，Visual Studio 会执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="9e084-215">Next, Visual Studio does the following:</span></span>


1. <span data-ttu-id="9e084-216">修改 XML 显示文件的 [SourceLocation](https://docs.microsoft.com/javascript/office/manifest/sourcelocation?view=office-js)元素，通过将 ～remoteAppUrl 标记替换为起始页的完全限定地址（例如，http://localhost/MyAgave.html)）。</span><span class="sxs-lookup"><span data-stu-id="9e084-216">Modifies the SourceLocation element of the XML manifest file by replacing the ~remoteAppUrl token with the fully qualified address of the start page (for example, http://localhost/MyAgave.html).</span></span>
    
2. <span data-ttu-id="9e084-217">在 IIS Express 中启动 Web 应用程序项目。</span><span class="sxs-lookup"><span data-stu-id="9e084-217">Starts the web application project in IIS Express.</span></span>
    
3. <span data-ttu-id="9e084-218">打开主机应用程序。</span><span class="sxs-lookup"><span data-stu-id="9e084-218">Opens the host application.</span></span> 
    
<span data-ttu-id="9e084-p120">生成项目时，Visual Studio 不会显示“**输出**”窗口中的验证错误。Visual Studio 报告“**错误列表**”窗口中出现的错误和警告。通过在代码和文本编辑器中显示不同颜色的波浪下划线（称为波浪线），Visual Studio 还报告验证错误。通过这些标志，你可以得知 Visual Studio 在代码中检测到的问题。有关详细信息，请参阅 [代码和文本编辑器](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)。有关如何启用或禁用验证的详细信息，请参阅：</span><span class="sxs-lookup"><span data-stu-id="9e084-p120">Visual Studio doesn't show validation errors in the  **OUTPUT** window when you build the project. Visual Studio reports errors and warnings in the **ERRORLIST** window as they occur. Visual Studio also reports validation errors by showing wavy underlines (known as squiggles) of different colors in the code and text editor. These marks notify you of problems that Visual Studio detected in your code. For more information, see [Code and Text Editor](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx). For more information about how to enable or disable validation, see:</span></span> 

- [<span data-ttu-id="9e084-225">选项、文本编辑器、JavaScript 和 IntelliSense</span><span class="sxs-lookup"><span data-stu-id="9e084-225">Options, Text Editor, JavaScript, IntelliSense</span></span>](https://docs.microsoft.com/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2015)
    
- <span data-ttu-id="9e084-226">[操作方法：为 Visual Web Developer 中的 HTML 编辑设置验证选项](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)</span><span class="sxs-lookup"><span data-stu-id="9e084-226">[How to: Set Validation Options for HTML Editing in Visual Web Developer](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)</span></span>
    
- <span data-ttu-id="9e084-227">[有关 CSS，请参阅验证、CSS、文本编辑器和“选项”对话框](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)</span><span class="sxs-lookup"><span data-stu-id="9e084-227">[CSS, see Validation, CSS, Text Editor, Options Dialog Box](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)</span></span>
    
<span data-ttu-id="9e084-228">若要查看项目中 XML 清单文件的验证规则，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="9e084-228">To review the validation rules of the XML manifest file in your project, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>


### <a name="show-an-add-in-in-excel-word-or-project-and-step-through-your-code"></a><span data-ttu-id="9e084-229">在 Excel、Word 或 Project 中显示加载项并单步调试代码</span><span class="sxs-lookup"><span data-stu-id="9e084-229">Show an add-in in Excel, Word, or Project and step through your code</span></span>


<span data-ttu-id="9e084-p121">如果将外接程序项目的“**启动文档**”属性设置为 Excel 或 Word，Visual Studio 会创建一个新文档，外接程序会出现。如果将外接程序项目的“**启动文档**”属性设置为使用现有文档，Visual Studio 会打开该文档，但是你必须手动插入外接程序。如果将“**启动文档**”设置为“**Microsoft Project**”，则还需要手动插入外接程序。</span><span class="sxs-lookup"><span data-stu-id="9e084-p121">If you set the  **Start Document** property of the add-in project to Excel or Word, Visual Studio creates a new document and the add-in appears. If you set the **Start Document** property of the add-in project to use an existing document, Visual Studio opens the document, but you have to insert the add-in manually. If you set the **Start Document** to **Microsoft Project**, you also have to insert the add-in manually.</span></span>


### <a name="to-show-an-office-add-in-in-excel-or-word"></a><span data-ttu-id="9e084-233">在 Excel 或 Word 中显示 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="9e084-233">To show an Office Add-in in Excel or Word</span></span>


1. <span data-ttu-id="9e084-234">在 Excel 或 Word 中的“**插入**”选项卡上，选择“**Office 外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="9e084-234">In Excel or Word, on the  **Insert** tab, choose **Office Add-ins**.</span></span>
    
2. <span data-ttu-id="9e084-235">在出现的列表中选择您的加载项。</span><span class="sxs-lookup"><span data-stu-id="9e084-235">In the list that appears, choose your add-in.</span></span>
    

### <a name="to-show-an-office-add-in-in-project"></a><span data-ttu-id="9e084-236">在 Project 中显示 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="9e084-236">To show an Office Add-in in Project</span></span>


1. <span data-ttu-id="9e084-237">在 Project 中的“**项目**”选项卡上，选择“**Office 外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="9e084-237">In Project, on the  **Project** tab, choose **Office Add-ins**.</span></span>
    
2. <span data-ttu-id="9e084-238">在出现的列表中选择您的加载项。</span><span class="sxs-lookup"><span data-stu-id="9e084-238">In the list that appears, choose your add-in.</span></span>
    
<span data-ttu-id="9e084-p122">在 Visual Studio 中，您随后可以设置断点。然后，当您与加载项交互时，可对 HTML、JavaScript 和 C# 或 VB 代码文件中的代码进行单步调试。</span><span class="sxs-lookup"><span data-stu-id="9e084-p122">In Visual Studio, you can then set break-points. Then, as you interact with your add-in and step through the code in your HTML, JavaScript, and C# or VB code files.</span></span>


### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a><span data-ttu-id="9e084-241">在 Outlook 中显示 Outlook 外接程序并单步调试代码</span><span class="sxs-lookup"><span data-stu-id="9e084-241">Show the Outlook add-in in Outlook and step through your code</span></span>


<span data-ttu-id="9e084-242">若要在 Outlook 中查看加载项，请打开一个电子邮件或约会项目。</span><span class="sxs-lookup"><span data-stu-id="9e084-242">To view the add-in in Outlook, open an email message or appointment item.</span></span>

<span data-ttu-id="9e084-p123">只要满足激活条件，Outlook 便会为项目激活外接程序。外接程序栏显示在"检查器"窗口或阅读窗格的顶部，Outlook 外接程序显示为外接程序栏中的一个按钮。如果您的外接程序有外接程序命令，那么在默认选项卡或指定的自定义选项卡中将有一个按钮显示在功能区中，而该外接程序将不会显示在外接程序栏中。</span><span class="sxs-lookup"><span data-stu-id="9e084-p123">Outlook activates the add-in for the item as long as the activation criteria are met. The add-in bar appears at the top of the Inspector window or Reading Pane, and your Outlook add-in appears as a button in the add-in bar. If your add-in has an add-in command, a button will appear in the ribbon, either in the default tab or a specified custom tab, and the add-in will not appear in the add-in bar.</span></span>

<span data-ttu-id="9e084-246">若要查看 Outlook 外接程序，请选择 Outlook 外接程序的按钮。</span><span class="sxs-lookup"><span data-stu-id="9e084-246">To view your Outlook add-in, choose the button for your Outlook add-in.</span></span>

<span data-ttu-id="9e084-p124">在 Visual Studio 中，可以设置断点。然后，与 Outlook 外接程序交互并逐句调试 HTML、JavaScript 和 C# 或 VB 代码文件中的代码。</span><span class="sxs-lookup"><span data-stu-id="9e084-p124">In Visual Studio, you can set break-points. Then, as you interact with your Outlook add-in and step through the code in your HTML, JavaScript, and C# or VB code files.</span></span> 

<span data-ttu-id="9e084-p125">你还可以更改代码并在 Outlook 外接程序中查看这些更改的效果，而不必关闭 Office 外接程序并再次启动项目。在 Outlook 中，只需打开 Outlook 外接程序的快捷菜单，然后选择“**重新加载**”即可。</span><span class="sxs-lookup"><span data-stu-id="9e084-p125">You can also change your code and review the effects of those changes in your Outlook add-in without having to close the Office Add-in and start the project again. In Outlook, just open the shortcut menu for the Outlook add-in, and then choose  **Reload**.</span></span>


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a><span data-ttu-id="9e084-251">修改代码并继续调试加载项，而不必再次启动项目</span><span class="sxs-lookup"><span data-stu-id="9e084-251">Modify code and continue to debug the add-in without having to start the project again</span></span>


<span data-ttu-id="9e084-p126">你可以更改代码并在外接程序中查看这些更改的效果，无需关闭主机应用程序并重新启动该项目。更改代码后，打开外接程序的快捷菜单，然后选择“**重新加载**”。当重新加载外接程序时，它会与 Visual Studio 调试器断开连接。因此，你可以查看所做更改的效果，但是在将 Visual Studio 调试器附加到所有可用的 Iexplore.exe 进程之前，将无法再次单步执行代码。</span><span class="sxs-lookup"><span data-stu-id="9e084-p126">You can change your code and review the effects of those changes in your add-in without having to close the host application and start the project again. After you change your code, open the shortcut menu for the add-in, and then choose  **Reload**. When you reload the add-in it becomes disconnected with the Visual Studio debugger. Therefore, you can view the effects of your change, but you cannot step through your code again until you attach the Visual Studio debugger to all of the available Iexplore.exe processes.</span></span>


### <a name="to-attach-the-visual-studio-debugger-to-all-of-the-available-iexploreexe-processes"></a><span data-ttu-id="9e084-256">将 Visual Studio 调试器附加到所有可用的 Iexplore.exe 进程</span><span class="sxs-lookup"><span data-stu-id="9e084-256">To attach the Visual Studio debugger to all of the available Iexplore.exe processes</span></span>


1. <span data-ttu-id="9e084-257">在 Visual Studio 中，依次选择“**调试**”、“**附加到进程**”。</span><span class="sxs-lookup"><span data-stu-id="9e084-257">In Visual Studio, choose  **DEBUG**,  **Attach to Process**.</span></span>
    
2. <span data-ttu-id="9e084-258">在“**附加到进程**”对话框中，选择所有可用的“**Iexplore.exe**”进程，然后选择“**附加**”按钮。</span><span class="sxs-lookup"><span data-stu-id="9e084-258">In the  **Attach to Process** dialog box, choose all of the available **Iexplore.exe** processes, and then choose the **Attach** button.</span></span>
    

## <a name="next-steps"></a><span data-ttu-id="9e084-259">后续步骤</span><span class="sxs-lookup"><span data-stu-id="9e084-259">Next steps</span></span>

- [<span data-ttu-id="9e084-260">部署和发布 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="9e084-260">Deploy and publish your Office Add-in</span></span>](../publish/publish.md)
    
