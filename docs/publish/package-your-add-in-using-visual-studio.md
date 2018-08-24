---
title: 使用 Visual Studio 打包加载项以准备发布
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: 89f59d06ff305e0d0fd062a36f7e9f756415df45
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925246"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="9814c-102">使用 Visual Studio 打包加载项以准备发布</span><span class="sxs-lookup"><span data-stu-id="9814c-102">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="9814c-103">Office 加载项包包含 XML [清单文件](../develop/add-in-manifests.md)，它可用于发布加载项。</span><span class="sxs-lookup"><span data-stu-id="9814c-103">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="9814c-104">必须单独发布项目的 Web 应用程序文件。</span><span class="sxs-lookup"><span data-stu-id="9814c-104">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="9814c-105">本文介绍如何使用 Visual Studio 2015 部署 Web 项目并打包加载项。</span><span class="sxs-lookup"><span data-stu-id="9814c-105">This article describes how to deploy your web project and package your add-in by using Visual Studio 2015.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a><span data-ttu-id="9814c-106">使用 Visual Studio 2015 部署 Web 项目</span><span class="sxs-lookup"><span data-stu-id="9814c-106">To deploy your web project using Visual Studio 2015</span></span>

<span data-ttu-id="9814c-107">完成以下步骤以使用 Visual Studio 2015 部署 Web 项目。</span><span class="sxs-lookup"><span data-stu-id="9814c-107">Complete the following steps to deploy your web project using Visual Studio 2015.</span></span>

1. <span data-ttu-id="9814c-108">在“解决方案资源管理器”**** 中，打开加载项项目的快捷菜单，然后选择“发布”****。</span><span class="sxs-lookup"><span data-stu-id="9814c-108">In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.</span></span>
    
    <span data-ttu-id="9814c-109">将显示“**发布外接程序**”页。</span><span class="sxs-lookup"><span data-stu-id="9814c-109">The  **Publish your add-in** page appears.</span></span>
    
2. <span data-ttu-id="9814c-110">选择“当前配置文件”**** 下拉列表中的配置文件，或选择“新建…”**** 新建配置文件。</span><span class="sxs-lookup"><span data-stu-id="9814c-110">In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="9814c-111">发布配置文件指定要部署到的服务器、登录服务器所需的凭据、要部署的数据库和其他部署选项。</span><span class="sxs-lookup"><span data-stu-id="9814c-111">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

    <span data-ttu-id="9814c-p102">如果你选择“**新建...**”，将会显示“**创建发布配置文件**”向导。可以使用此向导从托管提供程序（如 Microsoft Azure）的网站导入发布配置文件，或创建新配置文件并添加你的服务器、凭据以及下一过程中的其他设置。</span><span class="sxs-lookup"><span data-stu-id="9814c-p102">If you choose  **New ...**, the  **Create publishing profile** wizard appears. You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.</span></span>
    
    <span data-ttu-id="9814c-114">有关导入发布配置文件或创建新发布配置文件的详细信息，请参阅 [创建发布配置文件](http://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile)。</span><span class="sxs-lookup"><span data-stu-id="9814c-114">For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](http://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span></span>
    
3. <span data-ttu-id="9814c-115">在“**发布外接程序**”页中，选择“**部署 Web 项目**”链接。</span><span class="sxs-lookup"><span data-stu-id="9814c-115">In the  **Publish your add-in** page, choose the **Deploy your web project** link.</span></span>
    
    <span data-ttu-id="9814c-p103">出现 **“发布 Web”** 对话框。有关使用此向导的详细信息，请参阅[如何：在 Visual Studio 中使用“一键式发布”部署 Web 项目](http://msdn.microsoft.com/library/dd465337.aspx)。</span><span class="sxs-lookup"><span data-stu-id="9814c-p103">The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](http://msdn.microsoft.com/library/dd465337.aspx).</span></span>
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a><span data-ttu-id="9814c-118">使用 Visual Studio 2015 打包加载项的具体步骤</span><span class="sxs-lookup"><span data-stu-id="9814c-118">To package your add-in using Visual Studio 2015</span></span>

<span data-ttu-id="9814c-119">完成以下步骤以使用 Visual Studio 2015 打包加载项。</span><span class="sxs-lookup"><span data-stu-id="9814c-119">Complete the following steps to package your add-in using Visual Studio 2015.</span></span>

1. <span data-ttu-id="9814c-120">在“发布加载项”**** 页中，选择“打包加载项”**** 链接。</span><span class="sxs-lookup"><span data-stu-id="9814c-120">In the **Publish your add-in** page, choose the **Package the add-in** link.</span></span>
    
    <span data-ttu-id="9814c-121">此时，“发布 Office 和 SharePoint 加载项”**** 向导显示。</span><span class="sxs-lookup"><span data-stu-id="9814c-121">The **Publish Office and SharePoint Add-ins** wizard appears.</span></span>
    
2. <span data-ttu-id="9814c-122">在“网站托管在哪里?”**** 下拉列表中，选择或输入托管加载项内容文件的网站的 HTTPS URL，再选择“完成”****。</span><span class="sxs-lookup"><span data-stu-id="9814c-122">In the **Where is your website hosted?** dropdown list, select or enter the HTTPS URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span> 
    
    <span data-ttu-id="9814c-p104">必须指定以 HTTPS 前缀开头的 URL，才能完成此向导。若要使用网站的 HTTP 终结点，可以在创建包后使用文本编辑器打开 XML 清单文件，并将网站的 HTTPS 前缀替换为 HTTP 前缀。</span><span class="sxs-lookup"><span data-stu-id="9814c-p104">You must specify a URL that begins with the HTTPS prefix to complete this wizard. If you want to use an HTTP endpoint for your website, you can open the XML manifest file in a text editor after the package has been created and replace the HTTPS prefix of your website with an HTTP prefix.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="9814c-125"> Azure 网站自动提供 HTTPS 端点。</span><span class="sxs-lookup"><span data-stu-id="9814c-125">Azure websites automatically provide an HTTPS endpoint.</span></span>

    <span data-ttu-id="9814c-126">Visual Studio 生成发布加载项所需的文件，并打开发布输出文件夹。</span><span class="sxs-lookup"><span data-stu-id="9814c-126">Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder.</span></span> 
    
<span data-ttu-id="9814c-p105">如果计划将加载项提交到 AppSource，可以选择“执行验证检查”**** 链接，以发现将会导致加载项被拒绝的任何问题。应先解决所有问题，再将加载项提交到应用商店。</span><span class="sxs-lookup"><span data-stu-id="9814c-p105">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** link to identify any issues that will prevent your add-in from being accepted. You should address all issues before you submit your add-in to the store.</span></span>

<span data-ttu-id="9814c-p106">现在，可以将 XML 清单上传到适当位置，以[发布加载项](../publish/publish.md)。XML 清单位于 `app.publish` 文件夹的 `OfficeAppManifests` 中。例如：</span><span class="sxs-lookup"><span data-stu-id="9814c-p106">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a><span data-ttu-id="9814c-132">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9814c-132">See also</span></span>

- [<span data-ttu-id="9814c-133">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="9814c-133">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="9814c-134">将解决方案提交到 AppSource 和 Office 应用商店</span><span class="sxs-lookup"><span data-stu-id="9814c-134">Make your solutions available in AppSource and within Office</span></span>](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
