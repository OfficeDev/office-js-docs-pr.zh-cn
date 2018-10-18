---
title: 在 Office Online 中旁加载 Office 加载项以供测试
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 10e236366012bb402b968d0f61ea64326bb9172d
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925302"
---
# <a name="sideload-office-add-ins-in-office-online-for-testing"></a><span data-ttu-id="5d154-102">在 Office Online 中旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="5d154-102">Sideload Office Add-ins in Office Online for testing</span></span>

<span data-ttu-id="5d154-p101">您可以安装 Office 外接程序进行测试，而无需首先使用旁加载将其放在外接程序目录中。在 Office 365 或 Office Online 中都可以进行旁加载。该过程对两个平台略有不同。</span><span class="sxs-lookup"><span data-stu-id="5d154-p101">You can install an Office Add-in for testing without having to first put it in an add-in catalog by using sideloading. Sideloading can be done on either Office 365 or Office Online. The procedure is slightly different for the two platforms.</span></span> 

<span data-ttu-id="5d154-106">当旁加载外接程序时，外接程序清单存储在浏览器的本地存储区中，因此如果清除浏览器的缓存，或切换到另一个浏览器，就必须再次旁加载该外接程序。</span><span class="sxs-lookup"><span data-stu-id="5d154-106">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>


> [!NOTE]
> <span data-ttu-id="5d154-p102">如本文所述，Word、Excel 和 PowerPoint 支持旁加载。若要旁加载 Outlook 外接程序，请参阅[旁加载 Outlook 外接程序进行测试](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)。</span><span class="sxs-lookup"><span data-stu-id="5d154-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

<span data-ttu-id="5d154-109">下面的视频逐步展示了如何在 Office 桌面或 Office Online 上旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="5d154-109">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-on-office-365"></a><span data-ttu-id="5d154-110">在 Office 365 上旁加载 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="5d154-110">Sideload an Office Add-in on Office 365</span></span>


1. <span data-ttu-id="5d154-111">登录 Office 365 帐户。</span><span class="sxs-lookup"><span data-stu-id="5d154-111">Sign in to your Office 365 account.</span></span>
    
2. <span data-ttu-id="5d154-112">打开工具栏最左端的应用启动器，选择“Excel”\*\*\*\*、“Word”\*\*\*\* 或“PowerPoint”\*\*\*\*，再新建文档。</span><span class="sxs-lookup"><span data-stu-id="5d154-112">Open the App Launcher on the left end of the toolbar and select  **Excel**,  **Word**, or  **PowerPoint**, and then create a new document.</span></span>
    
3. <span data-ttu-id="5d154-113">打开功能区上的“**插入**”选项卡，然后在“**外接程序**”部分中，选择“**Office 外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="5d154-113">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="5d154-114">在“Office 加载项”\*\*\*\* 对话框中，依次选择“我的组织”\*\*\*\* 选项卡和“上传我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5d154-114">On the  **Office Add-ins** dialog, select the **MY ORGANIZATION** tab, and then **Upload My Add-in**.</span></span>
    
    ![标题为“Office 加载项”的对话框，左上角附近有链接“上传我的加载项”](../images/office-add-ins.png)

5.  <span data-ttu-id="5d154-116">**转到**加载项清单文件，再选择“上传”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5d154-116">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![包含“浏览”、“上传”和“取消”按钮的“上传加载项”对话框](../images/upload-add-in.png)

6. <span data-ttu-id="5d154-p103">验证是否已安装外接程序。例如，如果它是一个外接程序命令，它应显示在功能区或上下文菜单上。如果它是一个任务窗格外接程序，则应显示窗格。</span><span class="sxs-lookup"><span data-stu-id="5d154-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in the pane should appear.</span></span>
    

## <a name="sideload-an-office-add-in-on-office-online"></a><span data-ttu-id="5d154-121">将 Office 外接程序旁加载在 Office Online 上</span><span class="sxs-lookup"><span data-stu-id="5d154-121">Sideload an Office Add-in on Office Online</span></span>


1. <span data-ttu-id="5d154-122">打开 [Microsoft Office Online](https://office.live.com/)。</span><span class="sxs-lookup"><span data-stu-id="5d154-122">Open [Microsoft Office Online](https://office.live.com/).</span></span>
    
2. <span data-ttu-id="5d154-123">在“**立即开始使用在线应用**”中，选择 **Excel**、**Word** 或 **PowerPoint**；然后打开一个新文档。</span><span class="sxs-lookup"><span data-stu-id="5d154-123">In  **Get started with the online apps now**, choose  **Excel**,  **Word**, or  **PowerPoint**; and then open a new document.</span></span>
    
3. <span data-ttu-id="5d154-124">打开功能区上的“**插入**”选项卡，然后在“**外接程序**”部分中，选择“**Office 外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="5d154-124">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="5d154-125">在“Office 加载项”\*\*\*\* 对话框中，依次选择“我的加载项”\*\*\*\* 选项卡、“管理我的加载项”\*\*\*\* 和“上传我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5d154-125">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![“Office 加载项”对话框，右上方有“管理我的加载项”下拉列表，其中有下拉选项“上传我的加载项”](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="5d154-127">**转到**加载项清单文件，再选择“上传”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5d154-127">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

6. <span data-ttu-id="5d154-p104">验证是否已安装外接程序。例如，如果它是一个外接程序命令，它应显示在功能区或上下文菜单上。如果它是一个任务窗格外接程序，则应显示窗格。</span><span class="sxs-lookup"><span data-stu-id="5d154-p104">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
><span data-ttu-id="5d154-132">要使用 Edge 测试您的 Office 加载项，请在 Edge 搜索栏中输入 "**about:flags**" 以调出“开发者设置”选项。</span><span class="sxs-lookup"><span data-stu-id="5d154-132">To test your Office Add-in with Edge, enter “**about:flags**” in the Edge search bar to bring up the Developer Settings options.</span></span>  <span data-ttu-id="5d154-133">选中“\*\* 允许本地主机环回\*\*”选项并重启 Edge。</span><span class="sxs-lookup"><span data-stu-id="5d154-133">Check the “**Allow localhost loopback**” option and restart Edge.</span></span>

>    ![选中此框后，Edge 将允许本地主机环回。](../images/allow-localhost-loopback.png)

## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="5d154-135">使用 Visual Studio 时旁加载外接程序</span><span class="sxs-lookup"><span data-stu-id="5d154-135">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="5d154-p106">如果使用 Visual Studio 来开发外接程序，则旁加载的过程类似。唯一的区别是，必须更新清单中 **SourceURL** 元素的值以包含部署外接程序位置的完整 URL。</span><span class="sxs-lookup"><span data-stu-id="5d154-p106">If you're using Visual Studio to develop your add-in, the process to sideload is similar. The only difference is that you will have to update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span> 

<span data-ttu-id="5d154-p107">如果当前正在开发外接程序，则找到外接程序 manifest.xml 文件，并更新 **SourceLocation** 元素值以包含绝对 URI。Visual Studio 将放置一个令牌以供 localhost 部署。</span><span class="sxs-lookup"><span data-stu-id="5d154-p107">If you're currently developing your add-in, locate your add-in manifest.xml file, and update the **SourceLocation** element value to include an absolute URI. Visual Studio will put in a token for your localhost deployment.</span></span>

<span data-ttu-id="5d154-140">例如：</span><span class="sxs-lookup"><span data-stu-id="5d154-140">For example:</span></span> 

```xml
<SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
```
