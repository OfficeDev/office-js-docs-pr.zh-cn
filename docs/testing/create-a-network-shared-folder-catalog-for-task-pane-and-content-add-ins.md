---
title: 旁加载 Office 加载项以供测试
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: b143999422866dba9b43432359c12f3607261c60
ms.sourcegitcommit: e094aaa06d9aff3d13f8ffd3429d4a31f0b65b81
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/03/2018
ms.locfileid: "21782810"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="8c787-102">旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="8c787-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="8c787-103">您可以通过将清单发布到网络文件共享来安装 Office 加载项，以便在 Windows 上运行的 Office 客户端中进行测试（说明如下）。</span><span class="sxs-lookup"><span data-stu-id="8c787-103">You can install an Office Add-in for testing in an Office client running on Windows by using a shared folder catalog to publish the manifest to a network file share.</span></span>

> [!NOTE]
> <span data-ttu-id="8c787-104">如果您的加载项项目是使用 [**yo office** 工具](https://github.com/OfficeDev/generator-office)创建的，那么有一种替代方法可以为您提供旁加载功能。</span><span class="sxs-lookup"><span data-stu-id="8c787-104">If your add-in project was created with the [**yo office** tool](https://github.com/OfficeDev/generator-office), there is an alternative way of sideloading it that might work for you.</span></span> <span data-ttu-id="8c787-105">有关详情，请参阅 [使用 sideload 命令旁加载 Office 加载项](sideload-office-addin-using-sideload-command.md)。</span><span class="sxs-lookup"><span data-stu-id="8c787-105">Sideload Office Add-ins using the sideload command</span></span>

<span data-ttu-id="8c787-106">本文仅适用于在 Windows 上测试 Word、Excel 或 PowerPoint 加载项。</span><span class="sxs-lookup"><span data-stu-id="8c787-106">This article applies only to testing a Word, Excel, or PowerPoint add-ins on Windows.</span></span> <span data-ttu-id="8c787-107">如果要在其他平台上进行测试或想要测试 Outlook 加载项，请参阅以下主题之一来旁加载加载项：</span><span class="sxs-lookup"><span data-stu-id="8c787-107">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="8c787-108">在 Office Online 中旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="8c787-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="8c787-109">在 iPad 和 Mac 上旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="8c787-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="8c787-110">旁加载 Outlook 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="8c787-110">Sideload Outlook add-ins for testing</span></span>](../../../../outlook/add-ins/sideload-outlook-add-ins-for-testing)


<span data-ttu-id="8c787-111">下面的视频逐步展示了如何在 Office 桌面或 Office Online 上使用共享文件夹目录旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="8c787-111">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="8c787-112">共享文件夹</span><span class="sxs-lookup"><span data-stu-id="8c787-112">Share a folder</span></span>

1. <span data-ttu-id="8c787-113">在想要托管外接程序的 Windows 计算机上，转到你想用作共享文件夹目录的文件夹的父文件夹或驱动器号。</span><span class="sxs-lookup"><span data-stu-id="8c787-113">On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="8c787-114">打开（右键单击）文件夹的上下文菜单并选择“**属性**”。</span><span class="sxs-lookup"><span data-stu-id="8c787-114">Open the context menu for the folder (right-click) and choose **Properties**.</span></span>

3. <span data-ttu-id="8c787-115">打开“**共享**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="8c787-115">Open the **Sharing** tab.</span></span>

4. <span data-ttu-id="8c787-p103">在“**选择人员...**”页上，添加你自己以及想要与其共享外接程序的其他任何人。如果他们都是安全组的成员，那么可以添加该组。将至少需要该文件夹的**读/写**权限。</span><span class="sxs-lookup"><span data-stu-id="8c787-p103">On the **Choose people ...** page, add yourself and and anyone else with whom you want to share your add-in. If they are all members of a security group, you can add the group. You will need at least **Read/Write** permission to the folder.</span></span> 

5. <span data-ttu-id="8c787-119">依次选择“**共享**”、“ > **完成**”和“ > **关闭**”。</span><span class="sxs-lookup"><span data-stu-id="8c787-119">Choose **Share** > **Done** > **Close**.</span></span>


## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="8c787-120">将共享文件夹指定为受信任的目录</span><span class="sxs-lookup"><span data-stu-id="8c787-120">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="8c787-121">在 Excel、Word 或 PowerPoint 中打开一个新文档。</span><span class="sxs-lookup"><span data-stu-id="8c787-121">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="8c787-122">选择“**文件**”选项卡，然后选择“**选项**”。</span><span class="sxs-lookup"><span data-stu-id="8c787-122">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="8c787-123">选择“**信任中心**”，然后选择“**信任中心设置**”按钮。</span><span class="sxs-lookup"><span data-stu-id="8c787-123">Choose **Trust Center**, and then choose the  **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="8c787-124">选择“**受信任的外接程序目录**”。</span><span class="sxs-lookup"><span data-stu-id="8c787-124">Choose  **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="8c787-125">在“**目录 URL**”框中，输入共享文件夹目录的完整网络路径，然后选择“**添加目录**”。</span><span class="sxs-lookup"><span data-stu-id="8c787-125">In the  **Catalog Url** box, enter the full network path to the shared folder catalog, and then choose **Add Catalog**.</span></span>
    
6. <span data-ttu-id="8c787-126">选中“**显示在菜单中**”复选框，然后选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="8c787-126">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

7. <span data-ttu-id="8c787-127">关闭 Office 应用程序，你的更改将生效。</span><span class="sxs-lookup"><span data-stu-id="8c787-127">Close the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="8c787-128">旁加载外接程序</span><span class="sxs-lookup"><span data-stu-id="8c787-128">Sideload your add-in</span></span>

1. <span data-ttu-id="8c787-p104">放入在共享文件夹目录中进行测试的所有外接程序的清单文件。请务必将 Web 应用程序本身部署到 Web 服务器。务必在清单文件的 **SourceLocation** 元素中指定 URL。</span><span class="sxs-lookup"><span data-stu-id="8c787-p104">Put the manifest file of any add-in that you are testing in the shared folder catalog. Note that you deploy the web application itself to a web server. Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="8c787-132">在 Excel、Word 或 PowerPoint 中，选择功能区上“插入”**** 选项卡中的“我的加载项”****。</span><span class="sxs-lookup"><span data-stu-id="8c787-132">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="8c787-133">在“**Office 外接程序**”对话框的顶部，选择“**共享文件夹**”。</span><span class="sxs-lookup"><span data-stu-id="8c787-133">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="8c787-134">依次选择加载项名称和“确定”****，以插入加载项。</span><span class="sxs-lookup"><span data-stu-id="8c787-134">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="8c787-135">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8c787-135">See also</span></span>

- [<span data-ttu-id="8c787-136">验证并排查清单问题</span><span class="sxs-lookup"><span data-stu-id="8c787-136">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="8c787-137">发布 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="8c787-137">Publish your Office Add-in</span></span>](../publish/publish.md)
    
