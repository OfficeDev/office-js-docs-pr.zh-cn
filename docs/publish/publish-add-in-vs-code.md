---
title: 使用代码和 Azure Visual Studio外接程序
description: 如何使用 Code 和 Azure Active Directory Visual Studio加载项
ms.date: 08/12/2020
localization_priority: Normal
ms.openlocfilehash: 3552e4eebacc84fc2b8e37782c97b4e03e96e508
ms.sourcegitcommit: 7faa0932b953a4983a80af70f49d116c3236d81a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/21/2020
ms.locfileid: "46845507"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a><span data-ttu-id="e24af-103">发布使用 Visual Studio Code 开发的加载项</span><span class="sxs-lookup"><span data-stu-id="e24af-103">Publish an add-in developed with Visual Studio Code</span></span>

<span data-ttu-id="e24af-104">本文介绍如何发布使用 Yeoman 生成器创建并使用 [Visual Studio Code (VS Code)](https://code.visualstudio.com) 或任何其他编辑器开发的 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="e24af-104">This article describes how to publish an Office Add-in that you created using the Yeoman generator and developed with [Visual Studio Code (VS Code)](https://code.visualstudio.com) or any other editor.</span></span>

> [!NOTE]
> <span data-ttu-id="e24af-105">要了解如何发布使用 Visual Studio 创建的 Office 加载项，请参阅[使用 Visual Studio 发布加载项](package-your-add-in-using-visual-studio.md)。</span><span class="sxs-lookup"><span data-stu-id="e24af-105">For information about publishing an Office Add-in that you created using Visual Studio, see [Publish your add-in using Visual Studio](package-your-add-in-using-visual-studio.md).</span></span>

## <a name="publishing-an-add-in-for-other-users-to-access"></a><span data-ttu-id="e24af-106">发布加载项供其他人用户访问</span><span class="sxs-lookup"><span data-stu-id="e24af-106">Publishing an add-in for other users to access</span></span>

<span data-ttu-id="e24af-107">Office 加载项由一个 Web 应用程序和一个清单文件构成。</span><span class="sxs-lookup"><span data-stu-id="e24af-107">An Office Add-in consists of a web application and a manifest file.</span></span> <span data-ttu-id="e24af-108">Web 应用程序定义加载项的用户界面和功能，清单指定 Web 应用程序的位置并定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="e24af-108">The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.</span></span>

<span data-ttu-id="e24af-109">在开发过程中，你可以在本地 Web 服务器上运行该加载项， (具体 `localhost`) 。</span><span class="sxs-lookup"><span data-stu-id="e24af-109">While you're developing, you can run the add-in on your local web server (`localhost`).</span></span> <span data-ttu-id="e24af-110">当您准备好发布它供其他用户访问时，你将需要部署 Web 应用程序并更新清单以指定已部署应用程序的 URL。</span><span class="sxs-lookup"><span data-stu-id="e24af-110">When you're ready to publish it for other users to access, you'll need to deploy the web application and update the manifest to specify the URL of the deployed application.</span></span>

<span data-ttu-id="e24af-111">如果外接程序可按需工作，则可以使用 Azure 存储扩展直接通过 Visual Studio Code 发布它。</span><span class="sxs-lookup"><span data-stu-id="e24af-111">When your add-in is working as desired, you can publish it directly through Visual Studio Code using the Azure Storage extension.</span></span>

## <a name="using-visual-studio-code-to-publish"></a><span data-ttu-id="e24af-112">使用Visual Studio发布</span><span class="sxs-lookup"><span data-stu-id="e24af-112">Using Visual Studio Code to publish</span></span>

>[!NOTE]
> <span data-ttu-id="e24af-113">这些步骤仅适用于使用 Yeoman 生成器创建的项目。</span><span class="sxs-lookup"><span data-stu-id="e24af-113">These steps only work for projects created with the Yeoman generator.</span></span>

1. <span data-ttu-id="e24af-114">在 VS Code 代码编辑器的代码Visual Studio从其 (根文件夹中) 。</span><span class="sxs-lookup"><span data-stu-id="e24af-114">Open your project from its root folder in Visual Studio Code (VS Code).</span></span>
2. <span data-ttu-id="e24af-115">从 VS Code 中的扩展视图，搜索 Azure 存储扩展并将其安装。</span><span class="sxs-lookup"><span data-stu-id="e24af-115">From the Extensions view in VS Code, search for the Azure Storage extension and install it.</span></span>
3. <span data-ttu-id="e24af-116">安装完成后，会向活动栏中添加一个 Azure 图标。</span><span class="sxs-lookup"><span data-stu-id="e24af-116">Once installed, an Azure icon is added to the Activity Bar.</span></span> <span data-ttu-id="e24af-117">选择它可访问扩展。</span><span class="sxs-lookup"><span data-stu-id="e24af-117">Select it to access the extension.</span></span> <span data-ttu-id="e24af-118">如果你的活动栏处于隐藏状态，你将无法访问扩展。</span><span class="sxs-lookup"><span data-stu-id="e24af-118">If your Activity Bar is hidden, you won't be able to access the extension.</span></span> <span data-ttu-id="e24af-119">通过选择视图和显示活动 **栏>显示>栏**。</span><span class="sxs-lookup"><span data-stu-id="e24af-119">Show the Activity Bar by selecting **View > Appearance > Show Activity Bar**.</span></span>
4. <span data-ttu-id="e24af-120">在扩展中时，选择"登录 Azure"，登录 **Azure 帐户**。</span><span class="sxs-lookup"><span data-stu-id="e24af-120">When in the extension, sign in to your Azure account by selecting **Sign in to Azure**.</span></span> <span data-ttu-id="e24af-121">如果你还没有 Azure 帐户，也可以选择"创建免费的 Azure 帐户 **"帐户。**</span><span class="sxs-lookup"><span data-stu-id="e24af-121">You can also create an Azure account if you don't already have one by selecting **Create a free Azure account**.</span></span> <span data-ttu-id="e24af-122">请按照提供的步骤设置帐户。</span><span class="sxs-lookup"><span data-stu-id="e24af-122">Follow the provided steps to set up your account.</span></span>
5. <span data-ttu-id="e24af-123">登录到 Azure 帐户后，你将会看到你的 Azure 存储帐户的显示有扩展中。</span><span class="sxs-lookup"><span data-stu-id="e24af-123">Once you have signed in to your Azure account, you'll see your Azure storage accounts appear in the extension.</span></span> <span data-ttu-id="e24af-124">如果还没有存储帐户，则需要使用"创建新存储帐户 **"选项创建一个存储帐户** 。</span><span class="sxs-lookup"><span data-stu-id="e24af-124">If you don't already have a storage account, you'll need to create one using the **Create new storage account** option.</span></span> <span data-ttu-id="e24af-125">对存储帐户命名一个全局唯一名称，仅使用"a-z"和"0-9"。</span><span class="sxs-lookup"><span data-stu-id="e24af-125">Name your storage account a globally unique name, using only 'a-z' and '0-9'.</span></span> <span data-ttu-id="e24af-126">请注意，默认情况下，这将创建一个同名的存储帐户和资源组。</span><span class="sxs-lookup"><span data-stu-id="e24af-126">Note that by default, this creates a storage account and a resource group with the same name.</span></span> <span data-ttu-id="e24af-127">它将自动将存储帐户放置于美国西部。</span><span class="sxs-lookup"><span data-stu-id="e24af-127">It automatically puts the storage account in West US.</span></span> <span data-ttu-id="e24af-128">可以通过 Azure 帐户联机 [调整该操作](https://portal.azure.com/)。</span><span class="sxs-lookup"><span data-stu-id="e24af-128">This can be adjusted online through [your Azure account](https://portal.azure.com/).</span></span>
6. <span data-ttu-id="e24af-129">选择并 (右键单击) ，选择"配置**静态网站"。**</span><span class="sxs-lookup"><span data-stu-id="e24af-129">Select and hold (right-click) your storage account, choosing **Configure static website**.</span></span> <span data-ttu-id="e24af-130">系统将要求您输入索引文档的名称和 404 的文档名称。</span><span class="sxs-lookup"><span data-stu-id="e24af-130">You'll be asked to enter the index document name and the 404 document name.</span></span> <span data-ttu-id="e24af-131">将索引文档的名称从默认更改为 `index.html` **`taskpane.html`** .</span><span class="sxs-lookup"><span data-stu-id="e24af-131">Change the index document name from the default `index.html` to **`taskpane.html`**.</span></span> <span data-ttu-id="e24af-132">您可能还要更改 404 个文档名称，但不是必需的。</span><span class="sxs-lookup"><span data-stu-id="e24af-132">You may decide to also change the 404 document name but are not required to.</span></span>
7. <span data-ttu-id="e24af-133">现在选择 (静态网站) 右键单击，然后右键单击 **邮箱**。</span><span class="sxs-lookup"><span data-stu-id="e24af-133">Select and hold (right-click) your storage again, this time choosing **Browse static website**.</span></span> <span data-ttu-id="e24af-134">从打开的浏览器窗口中复制网站 URL。</span><span class="sxs-lookup"><span data-stu-id="e24af-134">From the browser window that opens, copy the website URL.</span></span>
8. <span data-ttu-id="e24af-135">在 VS Code 中，打开项目的清单文件 (`manifest.xml`) 并将对 localhost URL (如 `https://localhost:3000`) 已复制的 URL 的任何引用。</span><span class="sxs-lookup"><span data-stu-id="e24af-135">In VS Code, open your project's manifest file (`manifest.xml`) and change any reference to your localhost URL (such as `https://localhost:3000`) to the URL you've copied.</span></span> <span data-ttu-id="e24af-136">此终结点是您新建的存储帐户的静态网站 URL。</span><span class="sxs-lookup"><span data-stu-id="e24af-136">This endpoint is the static website URL for your newly created storage account.</span></span> <span data-ttu-id="e24af-137">保存对清单文件所做的更改。</span><span class="sxs-lookup"><span data-stu-id="e24af-137">Save the changes to your manifest file.</span></span>
9. <span data-ttu-id="e24af-138">打开命令行提示符并导航到加载项项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="e24af-138">Open a command line prompt and navigate to the root directory of your add-in project.</span></span> <span data-ttu-id="e24af-139">然后运行以下命令，为生产部署准备所有文件。</span><span class="sxs-lookup"><span data-stu-id="e24af-139">Then run the following command to prepare all files for production deployment.</span></span>

    ```command&nbsp;line
    npm run build
    ```

    <span data-ttu-id="e24af-140">生成完成后，加载项项目的根目录中的 **dist** 文件夹将包含要在后续步骤中部署的文件。</span><span class="sxs-lookup"><span data-stu-id="e24af-140">When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy in subsequent steps.</span></span>

10. <span data-ttu-id="e24af-141">若要部署，请选择文件资源管理器，选择并 (右键单击) **不同的文件夹**，然后选择"**部署到静态网站"。**</span><span class="sxs-lookup"><span data-stu-id="e24af-141">To deploy, select the Files explorer, select and hold (right-click) your **dist** folder, and choose **Deploy to Static Website**.</span></span> <span data-ttu-id="e24af-142">出现提示时，选择之前创建的存储帐户。</span><span class="sxs-lookup"><span data-stu-id="e24af-142">When prompted, select the storage account you created previously.</span></span>

![部署到静态网站](../images/deploy-to-static-website.png)

11. <span data-ttu-id="e24af-144">部署完成后，将显示 **"浏览至网站** "消息，你可以选择该消息来打开已部署应用代码的主要终结点。</span><span class="sxs-lookup"><span data-stu-id="e24af-144">When deployment is complete, a **Browse to website** message appears which you can select to open the primary endpoint of the deployed app code.</span></span>

## <a name="see-also"></a><span data-ttu-id="e24af-145">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e24af-145">See also</span></span>

- [<span data-ttu-id="e24af-146">使用 Visual Studio Code 开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="e24af-146">Develop Office Add-ins with Visual Studio Code</span></span>](../develop/develop-add-ins-vscode.md)
- [<span data-ttu-id="e24af-147">部署和发布 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="e24af-147">Deploy and publish your Office Add-in</span></span>](../publish/publish.md)
