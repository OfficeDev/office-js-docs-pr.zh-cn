---
title: 旁加载 Office 加载项以供测试
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: 42af5d0665fc6cb1135103789adcb4414c4763ff
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703803"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="33f8d-102">旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="33f8d-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="33f8d-103">您可以通过以下方法之一安装 Office 加载项，以便在 Windows 上运行的 Office 客户端中进行测试：</span><span class="sxs-lookup"><span data-stu-id="33f8d-103">You can install an Office Add-in for testing in an Office client running on Windows by one of the following methods:</span></span>

- <span data-ttu-id="33f8d-104">使用共享文件夹目录将清单发布到网络文件共享位置（说明见下文）</span><span class="sxs-lookup"><span data-stu-id="33f8d-104">Using a shared folder catalog to publish the manifest to a network file share (instructions below)</span></span>
- [<span data-ttu-id="33f8d-105">从加载项项目文件夹根目录运行 "**npm run sideload**" 命令。</span><span class="sxs-lookup"><span data-stu-id="33f8d-105">Running the "**npm run sideload**" command from the root of the add-in project folder.</span></span>](sideload-office-addin-using-sideload-command.md)

    > [!NOTE]
    > <span data-ttu-id="33f8d-106">"npm run sideload" 方法仅适用于 Excel、Word 和 PowerPoint 加载项。</span><span class="sxs-lookup"><span data-stu-id="33f8d-106">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins).</span></span>

<span data-ttu-id="33f8d-107">如果不在 Windows 上测试 Word、Excel 或 PowerPoint 加载项，则请参阅以下主题之一来旁加载加载项：</span><span class="sxs-lookup"><span data-stu-id="33f8d-107">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="33f8d-108">在 Office Online 中旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="33f8d-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="33f8d-109">在 iPad 和 Mac 上旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="33f8d-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

<span data-ttu-id="33f8d-110">下面的视频逐步展示了如何在 Office 桌面或 Office Online 上使用共享文件夹目录旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="33f8d-110">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="33f8d-111">共享文件夹</span><span class="sxs-lookup"><span data-stu-id="33f8d-111">Share a folder</span></span>

1. <span data-ttu-id="33f8d-112">在想要托管外接程序的 Windows 计算机上，转到你想用作共享文件夹目录的文件夹的父文件夹或驱动器号。</span><span class="sxs-lookup"><span data-stu-id="33f8d-112">On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="33f8d-113">打开（右键单击）文件夹的上下文菜单并选择“**属性**”。</span><span class="sxs-lookup"><span data-stu-id="33f8d-113">Open the context menu for the folder (right-click) and choose **Properties**.</span></span>

3. <span data-ttu-id="33f8d-114">打开“**共享**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="33f8d-114">Open the **Sharing** tab.</span></span>

4. <span data-ttu-id="33f8d-p101">在“**选择人员...**”页上，添加你自己以及想要与其共享外接程序的其他任何人。如果他们都是安全组的成员，那么可以添加该组。将至少需要该文件夹的**读/写**权限。</span><span class="sxs-lookup"><span data-stu-id="33f8d-p101">On the **Choose people ...** page, add yourself and and anyone else with whom you want to share your add-in. If they are all members of a security group, you can add the group. You will need at least **Read/Write** permission to the folder.</span></span> 

5. <span data-ttu-id="33f8d-118">依次选择“**共享**”、“ > **完成**”和“ > **关闭**”。</span><span class="sxs-lookup"><span data-stu-id="33f8d-118">Choose **Share** > **Done** > **Close**.</span></span>


## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="33f8d-119">将共享文件夹指定为受信任的目录</span><span class="sxs-lookup"><span data-stu-id="33f8d-119">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="33f8d-120">在 Excel、Word 或 PowerPoint 中打开一个新文档。</span><span class="sxs-lookup"><span data-stu-id="33f8d-120">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="33f8d-121">选择“**文件**”选项卡，然后选择“**选项**”。</span><span class="sxs-lookup"><span data-stu-id="33f8d-121">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="33f8d-122">选择“**信任中心**”，然后选择“**信任中心设置**”按钮。</span><span class="sxs-lookup"><span data-stu-id="33f8d-122">Choose **Trust Center**, and then choose the  **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="33f8d-123">选择“**受信任的外接程序目录**”。</span><span class="sxs-lookup"><span data-stu-id="33f8d-123">Choose  **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="33f8d-124">在“**目录 URL**”框中，输入共享文件夹目录的完整网络路径，然后选择“**添加目录**”。</span><span class="sxs-lookup"><span data-stu-id="33f8d-124">In the  **Catalog Url** box, enter the full network path to the shared folder catalog, and then choose **Add Catalog**.</span></span>
    
6. <span data-ttu-id="33f8d-125">选中“**显示在菜单中**”复选框，然后选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="33f8d-125">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

7. <span data-ttu-id="33f8d-126">关闭 Office 应用程序，你的更改将生效。</span><span class="sxs-lookup"><span data-stu-id="33f8d-126">Close the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="33f8d-127">旁加载外接程序</span><span class="sxs-lookup"><span data-stu-id="33f8d-127">Sideload your add-in</span></span>

1. <span data-ttu-id="33f8d-p102">放入在共享文件夹目录中进行测试的所有外接程序的清单文件。请务必将 Web 应用程序本身部署到 Web 服务器。务必在清单文件的 **SourceLocation** 元素中指定 URL。</span><span class="sxs-lookup"><span data-stu-id="33f8d-p102">Put the manifest file of any add-in that you are testing in the shared folder catalog. Note that you deploy the web application itself to a web server. Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="33f8d-131">在 Excel、Word 或 PowerPoint 中，选择功能区上“插入”**** 选项卡中的“我的加载项”****。</span><span class="sxs-lookup"><span data-stu-id="33f8d-131">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="33f8d-132">在“**Office 外接程序**”对话框的顶部，选择“**共享文件夹**”。</span><span class="sxs-lookup"><span data-stu-id="33f8d-132">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="33f8d-133">依次选择加载项名称和“确定”****，以插入加载项。</span><span class="sxs-lookup"><span data-stu-id="33f8d-133">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="33f8d-134">另请参阅</span><span class="sxs-lookup"><span data-stu-id="33f8d-134">See also</span></span>

- [<span data-ttu-id="33f8d-135">验证并排查清单问题</span><span class="sxs-lookup"><span data-stu-id="33f8d-135">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="33f8d-136">发布 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="33f8d-136">Publish your Office Add-in</span></span>](../publish/publish.md)
    
