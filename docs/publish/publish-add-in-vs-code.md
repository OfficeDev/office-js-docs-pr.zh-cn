---
title: 使用 Azure 和 Visual Studio Code加载项
description: 如何使用加载项和加载项Visual Studio Code Azure Active Directory
ms.date: 08/12/2020
localization_priority: Normal
ms.openlocfilehash: ab8daf3dfb87c809cd812da45246ce2d5ca9e743
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076936"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a><span data-ttu-id="74ab0-103">发布使用 Visual Studio Code 开发的加载项</span><span class="sxs-lookup"><span data-stu-id="74ab0-103">Publish an add-in developed with Visual Studio Code</span></span>

<span data-ttu-id="74ab0-104">本文介绍如何发布使用 Yeoman 生成器创建并使用 [Visual Studio Code (VS Code)](https://code.visualstudio.com) 或任何其他编辑器开发的 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="74ab0-104">This article describes how to publish an Office Add-in that you created using the Yeoman generator and developed with [Visual Studio Code (VS Code)](https://code.visualstudio.com) or any other editor.</span></span>

> [!NOTE]
> <span data-ttu-id="74ab0-105">要了解如何发布使用 Visual Studio 创建的 Office 加载项，请参阅[使用 Visual Studio 发布加载项](package-your-add-in-using-visual-studio.md)。</span><span class="sxs-lookup"><span data-stu-id="74ab0-105">For information about publishing an Office Add-in that you created using Visual Studio, see [Publish your add-in using Visual Studio](package-your-add-in-using-visual-studio.md).</span></span>

## <a name="publishing-an-add-in-for-other-users-to-access"></a><span data-ttu-id="74ab0-106">发布加载项供其他人用户访问</span><span class="sxs-lookup"><span data-stu-id="74ab0-106">Publishing an add-in for other users to access</span></span>

<span data-ttu-id="74ab0-107">Office 加载项由一个 Web 应用程序和一个清单文件构成。</span><span class="sxs-lookup"><span data-stu-id="74ab0-107">An Office Add-in consists of a web application and a manifest file.</span></span> <span data-ttu-id="74ab0-108">Web 应用程序定义加载项的用户界面和功能，清单指定 Web 应用程序的位置并定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="74ab0-108">The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.</span></span>

<span data-ttu-id="74ab0-109">开发时，可以在本地 Web 服务器上运行加载项 `localhost` ， () 。</span><span class="sxs-lookup"><span data-stu-id="74ab0-109">While you're developing, you can run the add-in on your local web server (`localhost`).</span></span> <span data-ttu-id="74ab0-110">准备好发布它供其他用户访问时，需要部署 Web 应用程序并更新清单以指定已部署应用程序的 URL。</span><span class="sxs-lookup"><span data-stu-id="74ab0-110">When you're ready to publish it for other users to access, you'll need to deploy the web application and update the manifest to specify the URL of the deployed application.</span></span>

<span data-ttu-id="74ab0-111">当加载项根据需要运行时，可以使用 Visual Studio Code 扩展直接Azure 存储它。</span><span class="sxs-lookup"><span data-stu-id="74ab0-111">When your add-in is working as desired, you can publish it directly through Visual Studio Code using the Azure Storage extension.</span></span>

## <a name="using-visual-studio-code-to-publish"></a><span data-ttu-id="74ab0-112">使用Visual Studio Code发布</span><span class="sxs-lookup"><span data-stu-id="74ab0-112">Using Visual Studio Code to publish</span></span>

>[!NOTE]
> <span data-ttu-id="74ab0-113">这些步骤仅适用于使用 Yeoman 生成器创建的项目。</span><span class="sxs-lookup"><span data-stu-id="74ab0-113">These steps only work for projects created with the Yeoman generator.</span></span>

1. <span data-ttu-id="74ab0-114">从项目根文件夹中打开项目，Visual Studio Code (VS Code) 。</span><span class="sxs-lookup"><span data-stu-id="74ab0-114">Open your project from its root folder in Visual Studio Code (VS Code).</span></span>
2. <span data-ttu-id="74ab0-115">从"扩展"视图中VS Code，搜索Azure 存储扩展并安装它。</span><span class="sxs-lookup"><span data-stu-id="74ab0-115">From the Extensions view in VS Code, search for the Azure Storage extension and install it.</span></span>
3. <span data-ttu-id="74ab0-116">安装后，Azure 图标将添加到活动栏。</span><span class="sxs-lookup"><span data-stu-id="74ab0-116">Once installed, an Azure icon is added to the Activity Bar.</span></span> <span data-ttu-id="74ab0-117">选择它以访问扩展。</span><span class="sxs-lookup"><span data-stu-id="74ab0-117">Select it to access the extension.</span></span> <span data-ttu-id="74ab0-118">如果活动栏处于隐藏状态，你将无法访问扩展。</span><span class="sxs-lookup"><span data-stu-id="74ab0-118">If your Activity Bar is hidden, you won't be able to access the extension.</span></span> <span data-ttu-id="74ab0-119">通过选择"显示活动栏 **">">"显示活动栏"。**</span><span class="sxs-lookup"><span data-stu-id="74ab0-119">Show the Activity Bar by selecting **View > Appearance > Show Activity Bar**.</span></span>
4. <span data-ttu-id="74ab0-120">在扩展中时，通过选择"登录到 Azure" **登录到 Azure 帐户**。</span><span class="sxs-lookup"><span data-stu-id="74ab0-120">When in the extension, sign in to your Azure account by selecting **Sign in to Azure**.</span></span> <span data-ttu-id="74ab0-121">如果还没有 Azure 帐户，也可以选择"创建免费的 Azure 帐户"来创建 **Azure 帐户**。</span><span class="sxs-lookup"><span data-stu-id="74ab0-121">You can also create an Azure account if you don't already have one by selecting **Create a free Azure account**.</span></span> <span data-ttu-id="74ab0-122">按照提供的步骤设置帐户。</span><span class="sxs-lookup"><span data-stu-id="74ab0-122">Follow the provided steps to set up your account.</span></span>
5. <span data-ttu-id="74ab0-123">登录 Azure 帐户后，你将看到 Azure 存储帐户显示在扩展中。</span><span class="sxs-lookup"><span data-stu-id="74ab0-123">Once you have signed in to your Azure account, you'll see your Azure storage accounts appear in the extension.</span></span> <span data-ttu-id="74ab0-124">如果还没有存储帐户，则需要使用"新建存储帐户"选项 **创建一** 个存储帐户。</span><span class="sxs-lookup"><span data-stu-id="74ab0-124">If you don't already have a storage account, you'll need to create one using the **Create new storage account** option.</span></span> <span data-ttu-id="74ab0-125">将存储帐户命名为全局唯一名称，仅使用"a-z"和"0-9"。</span><span class="sxs-lookup"><span data-stu-id="74ab0-125">Name your storage account a globally unique name, using only 'a-z' and '0-9'.</span></span> <span data-ttu-id="74ab0-126">请注意，默认情况下，这将创建一个存储帐户和一个同名的资源组。</span><span class="sxs-lookup"><span data-stu-id="74ab0-126">Note that by default, this creates a storage account and a resource group with the same name.</span></span> <span data-ttu-id="74ab0-127">它会自动将存储帐户置于美国西部。</span><span class="sxs-lookup"><span data-stu-id="74ab0-127">It automatically puts the storage account in West US.</span></span> <span data-ttu-id="74ab0-128">这可以通过 Azure 帐户 [在线调整](https://portal.azure.com/)。</span><span class="sxs-lookup"><span data-stu-id="74ab0-128">This can be adjusted online through [your Azure account](https://portal.azure.com/).</span></span>
6. <span data-ttu-id="74ab0-129">选择并按住 (右键单击) 存储帐户"，选择"**配置静态网站"。**</span><span class="sxs-lookup"><span data-stu-id="74ab0-129">Select and hold (right-click) your storage account, choosing **Configure static website**.</span></span> <span data-ttu-id="74ab0-130">将要求您输入索引文档名称和 404 文档名称。</span><span class="sxs-lookup"><span data-stu-id="74ab0-130">You'll be asked to enter the index document name and the 404 document name.</span></span> <span data-ttu-id="74ab0-131">将索引文档名称从默认更改为 `index.html` **`taskpane.html`** 。</span><span class="sxs-lookup"><span data-stu-id="74ab0-131">Change the index document name from the default `index.html` to **`taskpane.html`**.</span></span> <span data-ttu-id="74ab0-132">您也可以决定更改 404 文档名称，但不要求更改。</span><span class="sxs-lookup"><span data-stu-id="74ab0-132">You may decide to also change the 404 document name but are not required to.</span></span>
7. <span data-ttu-id="74ab0-133">选择并按住 (再次右键) 存储"，这次选择"浏览静态 **网站"。**</span><span class="sxs-lookup"><span data-stu-id="74ab0-133">Select and hold (right-click) your storage again, this time choosing **Browse static website**.</span></span> <span data-ttu-id="74ab0-134">从打开的浏览器窗口中，复制网站 URL。</span><span class="sxs-lookup"><span data-stu-id="74ab0-134">From the browser window that opens, copy the website URL.</span></span>
8. <span data-ttu-id="74ab0-135">在 VS Code 中，打开项目的清单文件 () ，将本地主机 URL (（如) ）的任何引用更改为已复制的 `manifest.xml` `https://localhost:3000` URL。</span><span class="sxs-lookup"><span data-stu-id="74ab0-135">In VS Code, open your project's manifest file (`manifest.xml`) and change any reference to your localhost URL (such as `https://localhost:3000`) to the URL you've copied.</span></span> <span data-ttu-id="74ab0-136">此终结点是新创建的存储帐户的静态网站 URL。</span><span class="sxs-lookup"><span data-stu-id="74ab0-136">This endpoint is the static website URL for your newly created storage account.</span></span> <span data-ttu-id="74ab0-137">保存对清单文件所做的更改。</span><span class="sxs-lookup"><span data-stu-id="74ab0-137">Save the changes to your manifest file.</span></span>
9. <span data-ttu-id="74ab0-138">打开命令行提示符并导航到加载项项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="74ab0-138">Open a command line prompt and navigate to the root directory of your add-in project.</span></span> <span data-ttu-id="74ab0-139">然后运行以下命令，准备用于生产部署的所有文件。</span><span class="sxs-lookup"><span data-stu-id="74ab0-139">Then run the following command to prepare all files for production deployment.</span></span>

    ```command&nbsp;line
    npm run build
    ```

    <span data-ttu-id="74ab0-140">生成完成后，加载项项目的根目录中的 **dist** 文件夹将包含要在后续步骤中部署的文件。</span><span class="sxs-lookup"><span data-stu-id="74ab0-140">When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy in subsequent steps.</span></span>

10. <span data-ttu-id="74ab0-141">若要部署，请选择文件资源管理器，选择并按住 (右键单击 **") "，** 然后选择"部署到静态网站 **"。**</span><span class="sxs-lookup"><span data-stu-id="74ab0-141">To deploy, select the Files explorer, select and hold (right-click) your **dist** folder, and choose **Deploy to Static Website**.</span></span> <span data-ttu-id="74ab0-142">当系统提示时，选择之前创建的存储帐户。</span><span class="sxs-lookup"><span data-stu-id="74ab0-142">When prompted, select the storage account you created previously.</span></span>

![部署到静态网站。](../images/deploy-to-static-website.png)

11. <span data-ttu-id="74ab0-144">部署完成后 **，将显示"** 浏览到网站"消息，您可以选择该消息打开已部署应用代码的主终结点。</span><span class="sxs-lookup"><span data-stu-id="74ab0-144">When deployment is complete, a **Browse to website** message appears which you can select to open the primary endpoint of the deployed app code.</span></span>

## <a name="see-also"></a><span data-ttu-id="74ab0-145">另请参阅</span><span class="sxs-lookup"><span data-stu-id="74ab0-145">See also</span></span>

- [<span data-ttu-id="74ab0-146">使用 Visual Studio Code 开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="74ab0-146">Develop Office Add-ins with Visual Studio Code</span></span>](../develop/develop-add-ins-vscode.md)
- [<span data-ttu-id="74ab0-147">部署和发布 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="74ab0-147">Deploy and publish your Office Add-in</span></span>](../publish/publish.md)
