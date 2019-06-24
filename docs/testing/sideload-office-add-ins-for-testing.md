---
title: 在 Office 网页版中旁加载 Office 加载项进行测试
description: 通过旁加载在 Office 网页版中测试 Office 加载项
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: e382e4de62df28584da2c827433af3102486e057
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35126944"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a><span data-ttu-id="99994-103">在 Office 网页版中旁加载 Office 加载项进行测试</span><span class="sxs-lookup"><span data-stu-id="99994-103">Sideload Office Add-ins in Office on the web for testing</span></span>

<span data-ttu-id="99994-104">可以通过使用旁加载来安装 Office 加载项以供测试，而无需先将它添加到加载项目录中。</span><span class="sxs-lookup"><span data-stu-id="99994-104">You can install an Office Add-in for testing without having to first put it in an add-in catalog by using sideloading.</span></span> <span data-ttu-id="99994-105">在 Office 365 或 Office 网页版中都可以进行旁加载。</span><span class="sxs-lookup"><span data-stu-id="99994-105">Sideloading can be done in either Office 365 or Office Online.</span></span> <span data-ttu-id="99994-106">该过程使用的两个平台略有不同。</span><span class="sxs-lookup"><span data-stu-id="99994-106">The procedure is slightly different for the two platforms.</span></span> 

<span data-ttu-id="99994-107">当旁加载外接程序时，外接程序清单存储在浏览器的本地存储区中，因此如果清除浏览器的缓存，或切换到另一个浏览器，就必须再次旁加载该外接程序。</span><span class="sxs-lookup"><span data-stu-id="99994-107">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>


> [!NOTE]
> <span data-ttu-id="99994-p102">如本文所述，Word、Excel 和 PowerPoint 支持旁加载。若要旁加载 Outlook 外接程序，请参阅[旁加载 Outlook 外接程序进行测试](/outlook/add-ins/sideload-outlook-add-ins-for-testing)。</span><span class="sxs-lookup"><span data-stu-id="99994-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

<span data-ttu-id="99994-110">下面的视频逐步展示了如何在 Office 网页版或桌面上旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="99994-110">The following video walks you through the process of sideloading your add-in in Office desktop or Office Online.</span></span>


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a><span data-ttu-id="99994-111">在 Office 网页版中旁加载 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="99994-111">Sideload an Office Add-in in Office on the web</span></span>

1. <span data-ttu-id="99994-112">打开 [Microsoft Office 网页版](https://office.live.com/)。</span><span class="sxs-lookup"><span data-stu-id="99994-112">Open [Microsoft Office on the web](https://office.live.com/).</span></span>
    
2. <span data-ttu-id="99994-113">在“**立即开始使用在线应用**”中，选择 **Excel**、**Word** 或 **PowerPoint**；然后打开一个新文档。</span><span class="sxs-lookup"><span data-stu-id="99994-113">In  **Get started with the online apps now**, choose  **Excel**,  **Word**, or  **PowerPoint**; and then open a new document.</span></span>
    
3. <span data-ttu-id="99994-114">打开功能区上的“**插入**”选项卡，然后在“**外接程序**”部分中，选择“**Office 外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="99994-114">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="99994-115">在“Office 加载项”\*\*\*\* 对话框中，依次选择“我的加载项”\*\*\*\* 选项卡、“管理我的加载项”\*\*\*\* 和“上传我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="99994-115">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![“Office 加载项”对话框，右上方有“管理我的加载项”下拉列表，其中有下拉选项“上传我的加载项”](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="99994-117">**转到**加载项清单文件，再选择“上传”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="99994-117">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

6. <span data-ttu-id="99994-p103">验证是否已安装外接程序。例如，如果它是一个外接程序命令，它应显示在功能区或上下文菜单上。如果它是一个任务窗格外接程序，则应显示窗格。</span><span class="sxs-lookup"><span data-stu-id="99994-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
><span data-ttu-id="99994-122">若要使用 Microsoft Edge 测试 Office 加载项，需要执行两个配置步骤：</span><span class="sxs-lookup"><span data-stu-id="99994-122">To test your Office Add-in with Edge, two configuration steps are required:</span></span> 
>
> - <span data-ttu-id="99994-123">在 Windows 命令提示符下，运行以下行：`CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span><span class="sxs-lookup"><span data-stu-id="99994-123">In a Windows Command Prompt, run the following line: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span></span>
>
> - <span data-ttu-id="99994-124">在 Microsoft Edge 搜索栏中输入“**about:flags**”以调出“开发人员设置”选项。</span><span class="sxs-lookup"><span data-stu-id="99994-124">Enter “**about:flags**” in the Edge search bar to bring up the Developer Settings options.</span></span>  <span data-ttu-id="99994-125">选中“**允许本地主机环回**”选项，然后重新启动 Microsoft Edge。</span><span class="sxs-lookup"><span data-stu-id="99994-125">Check the “**Allow localhost loopback**” option and restart Edge.</span></span>

>    ![选中此框后，Edge 会允许本地主机环回选项。](../images/allow-localhost-loopback.png)


## <a name="sideload-an-office-add-in-in-office-365"></a><span data-ttu-id="99994-127">在 Office 365 上旁加载 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="99994-127">Sideload an Office Add-in in Office 365</span></span>

1. <span data-ttu-id="99994-128">登录 Office 365 帐户。</span><span class="sxs-lookup"><span data-stu-id="99994-128">Sign in to your Office 365 account.</span></span>
    
2. <span data-ttu-id="99994-129">打开工具栏最左端的应用启动器，选择“Excel”\*\*\*\*、“Word”\*\*\*\* 或“PowerPoint”\*\*\*\*，再新建文档。</span><span class="sxs-lookup"><span data-stu-id="99994-129">Open the App Launcher on the left end of the toolbar and select  **Excel**,  **Word**, or  **PowerPoint**, and then create a new document.</span></span>
    
3. <span data-ttu-id="99994-130">步骤 3 - 6 与上一部分**在 Office 网页版中旁加载 Office 加载项**相同。</span><span class="sxs-lookup"><span data-stu-id="99994-130">Steps 3 - 6 are the same as in the preceding section **Sideload an Office Add-in in Office Online**.</span></span>


## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="99994-131">使用 Visual Studio 时旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="99994-131">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="99994-132">如果使用 Visual Studio 来开发加载项，则旁加载的过程类似。</span><span class="sxs-lookup"><span data-stu-id="99994-132">If you're using Visual Studio to develop your add-in, the process to sideload is similar.</span></span> <span data-ttu-id="99994-133">唯一的区别是，必须更新清单中 **SourceURL** 元素的值以包含部署加载项位置的完整 URL。</span><span class="sxs-lookup"><span data-stu-id="99994-133">The only difference is that you must update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span>

> [!NOTE]
> <span data-ttu-id="99994-134">虽然可以将加载项从 Visual Studio 旁加载到 Office 网页版，但无法从 Visual Studio 调试它们。</span><span class="sxs-lookup"><span data-stu-id="99994-134">Although you can sideload add-ins from Visual Studio to Office Online, you cannot debug them from Visual Studio.</span></span> <span data-ttu-id="99994-135">若要进行调试，需要使用浏览器调试工具。</span><span class="sxs-lookup"><span data-stu-id="99994-135">To debug you will need to use the browser debugging tools.</span></span> <span data-ttu-id="99994-136">有关详细信息，请参阅[在 Office 网页版中调试加载项](debug-add-ins-in-office-online.md)。</span><span class="sxs-lookup"><span data-stu-id="99994-136">For more information, see [Debug add-ins in Office Online](debug-add-ins-in-office-online.md).</span></span>

1. <span data-ttu-id="99994-137">在 Visual Studio 中，通过选择**视图** -> **属性窗口**来显示**属性**窗口。</span><span class="sxs-lookup"><span data-stu-id="99994-137">In Visual Studio, show the **Properties** window by choosing **View** -> **Properties Window**.</span></span>
2. <span data-ttu-id="99994-138">在**解决方案资源管理器**中，选择 Web 项目。</span><span class="sxs-lookup"><span data-stu-id="99994-138">In the **Solution Explorer**, select the web project.</span></span> <span data-ttu-id="99994-139">这将在**属性**窗口中显示项目的属性。</span><span class="sxs-lookup"><span data-stu-id="99994-139">This will display properties for the project in the **Properties** window.</span></span>
3. <span data-ttu-id="99994-140">在“属性”窗口中复制 **SSL URL**。</span><span class="sxs-lookup"><span data-stu-id="99994-140">In the Properties window, copy the **SSL URL**.</span></span>
4. <span data-ttu-id="99994-141">在加载项项目中，打开清单 XML 文件。</span><span class="sxs-lookup"><span data-stu-id="99994-141">In the add-in project, open the manifest XML file.</span></span> <span data-ttu-id="99994-142">请确保正在编辑源 XML。</span><span class="sxs-lookup"><span data-stu-id="99994-142">Be sure you are editing the source XML.</span></span> <span data-ttu-id="99994-143">对于某些项目类型，Visual Studio 将打开 XML 的可视视图，它不适用于下一步骤。</span><span class="sxs-lookup"><span data-stu-id="99994-143">For some project types Visual Studio will open a visual view of the XML which will not work for the next step.</span></span>
5. <span data-ttu-id="99994-144">使用刚复制的 SSL URL 来搜索和替换 **~remoteAppUrl/** 的所有实例。</span><span class="sxs-lookup"><span data-stu-id="99994-144">Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied.</span></span> <span data-ttu-id="99994-145">将看到多个替换，具体取决于项目类型。将显示新 URL，类似于 `https://localhost:44300/Home.html`。</span><span class="sxs-lookup"><span data-stu-id="99994-145">You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.</span></span>
6. <span data-ttu-id="99994-146">保存 XML 文件。</span><span class="sxs-lookup"><span data-stu-id="99994-146">Save the XML file.</span></span>
7. <span data-ttu-id="99994-147">右键单击 Web 项目，然后选择**调试** -> **启动新实例**。</span><span class="sxs-lookup"><span data-stu-id="99994-147">Right click the web project and choose **Debug** -> **Start new instance**.</span></span> <span data-ttu-id="99994-148">这将在不启动 Office 的情况下运行 Web 项目。</span><span class="sxs-lookup"><span data-stu-id="99994-148">This will run the web project without launching Office.</span></span>
8. <span data-ttu-id="99994-149">从 Office 网页版，使用之前[在 Office 网页版中加载 Office 加载项](#sideload-an-office-add-in-in-office-on-the-web)中所述的步骤旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="99994-149">From Office Online, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office Online](#sideload-an-office-add-in-in-office-on-the-web).</span></span>
