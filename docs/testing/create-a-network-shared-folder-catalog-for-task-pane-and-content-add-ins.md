---
title: 旁加载 Office 加载项以供测试
description: ''
ms.date: 10/17/2018
localization_priority: Priority
ms.openlocfilehash: 7d7aa136e7268183c5a064f16315d433ee53a38b
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388841"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="de46d-102">旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="de46d-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="de46d-103">你可以安装 Office 外接程序以在 Windows 上运行的 Office 客户端中进行测试（通过使用共享文件夹，以将清单发布到网络文件共享）。</span><span class="sxs-lookup"><span data-stu-id="de46d-103">You can install an Office Add-in for testing in an Office client running on Windows by publishing the manifest to a network file share (instructions below).</span></span>

> [!NOTE]
> <span data-ttu-id="de46d-104">如果你的加载项项目是使用 [**yo office** 工具](https://github.com/OfficeDev/generator-office)创建的，则可以使用另一种方法来旁加载它，这可能对你有用。</span><span class="sxs-lookup"><span data-stu-id="de46d-104">If your add-in project was created with the [**yo office** tool](https://github.com/OfficeDev/generator-office), there is an alternative way of sideloading it that might work for you.</span></span> <span data-ttu-id="de46d-105">有关详细信息，请参阅[使用旁加载命令旁加载 Office 加载项](sideload-office-addin-using-sideload-command.md)。</span><span class="sxs-lookup"><span data-stu-id="de46d-105">For details, see [Sideload Office Add-ins using the sideload command](sideload-office-addin-using-sideload-command.md).</span></span>

<span data-ttu-id="de46d-106">本文仅适用于在 Windows 上测试 Word、Excel 或 PowerPoint 加载项。</span><span class="sxs-lookup"><span data-stu-id="de46d-106">This article applies only to testing a Word, Excel, or PowerPoint add-ins on Windows.</span></span> <span data-ttu-id="de46d-107">如果要在其他平台上进行测试或要测试 Outlook 加载项，请参阅以下主题之一以旁加载你的加载项：</span><span class="sxs-lookup"><span data-stu-id="de46d-107">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="de46d-108">在 Office Online 中旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="de46d-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="de46d-109">在 iPad 和 Mac 上旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="de46d-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="de46d-110">旁加载 Outlook 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="de46d-110">Sideload Outlook add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)


<span data-ttu-id="de46d-111">下面的视频逐步展示了如何使用共享文件夹目录在 Office 桌面或 Office Online 上旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="de46d-111">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="de46d-112">共享文件夹</span><span class="sxs-lookup"><span data-stu-id="de46d-112">Share a folder</span></span>

1. <span data-ttu-id="de46d-113">在想要托管外接程序的 Windows 计算机上，转到你想用作共享文件夹目录的文件夹的父文件夹或驱动器号。</span><span class="sxs-lookup"><span data-stu-id="de46d-113">In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="de46d-114">打开要用作共享文件夹目录的文件夹的上下文菜单（右键单击该文件夹），然后选择“**属性**”。</span><span class="sxs-lookup"><span data-stu-id="de46d-114">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="de46d-115">在“**属性**”对话框窗口中，打开“**共享**”选项卡，然后选择“**共享**”按钮。</span><span class="sxs-lookup"><span data-stu-id="de46d-115">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![已突出显示“共享”选项卡和“共享”按钮的文件夹“属性”对话框](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="de46d-117">在**网络访问**对话框窗口中，添加你自己以及要与其共享加载项的任何其他用户和/或组。</span><span class="sxs-lookup"><span data-stu-id="de46d-117">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="de46d-118">你至少需要对该文件夹的**读/写**权限。</span><span class="sxs-lookup"><span data-stu-id="de46d-118">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="de46d-119">选择要与其共享的人员后，请选择“**共享**”按钮。</span><span class="sxs-lookup"><span data-stu-id="de46d-119">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="de46d-120">当你看到确认**你的文件夹已共享**的消息时，请记下紧跟文件夹名称显示的完整网络路径。</span><span class="sxs-lookup"><span data-stu-id="de46d-120">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="de46d-121">（当你[将共享文件夹指定为受信任的目录](#specify-the-shared-folder-as-a-trusted-catalog)时，你需要将此值输入为**目录UR **，如本文下一节所述。）选择“**完成**”按钮以关闭“**网络访问**”对话框窗口。</span><span class="sxs-lookup"><span data-stu-id="de46d-121">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![已突出显示共享路径的“网络访问”对话框](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="de46d-123">选择“**关闭**”按钮以关闭“**属性**”对话框窗口。</span><span class="sxs-lookup"><span data-stu-id="de46d-123">Choose the **Close** button to close the **Properties** dialog window.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="de46d-124">将共享文件夹指定为受信任的目录</span><span class="sxs-lookup"><span data-stu-id="de46d-124">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="de46d-125">在 Excel、Word 或 PowerPoint 中打开一个新文档。</span><span class="sxs-lookup"><span data-stu-id="de46d-125">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="de46d-126">选择“**文件**”选项卡，然后选择“**选项**”。</span><span class="sxs-lookup"><span data-stu-id="de46d-126">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="de46d-127">选择“**信任中心**”，然后选择“**信任中心设置**”按钮。</span><span class="sxs-lookup"><span data-stu-id="de46d-127">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="de46d-128">选择“**受信任的加载项目录**”。</span><span class="sxs-lookup"><span data-stu-id="de46d-128">Choose **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="de46d-129">在“**目录 Url**”框中，输入你之前[共享](#share-a-folder)的文件夹的完整网络路径。</span><span class="sxs-lookup"><span data-stu-id="de46d-129">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="de46d-130">如果在共享文件夹时未能记下文件夹的完整网络路径，则可以从文件夹的“**属性**”对话框窗口中获取它，如以下屏幕截图所示。</span><span class="sxs-lookup"><span data-stu-id="de46d-130">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span> 

    ![已突出显示“共享”选项卡和网络路径的文件夹“属性”对话框](../images/sideload-windows-properties-dialog-2.png)
    
6. <span data-ttu-id="de46d-132">在“**目录 Url**”框中输入文件夹的完整网络路径后，选择“**添加目录**”按钮。</span><span class="sxs-lookup"><span data-stu-id="de46d-132">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="de46d-133">选中新添加项目的“**在菜单中显示**”复选框，然后选择“**确定**”按钮以关闭“**信任中心**”对话框窗口。</span><span class="sxs-lookup"><span data-stu-id="de46d-133">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![已选择目录的“信任中心”对话框](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="de46d-135">选择“**确定**”按钮以关闭“**Word 选项**”对话框窗口。</span><span class="sxs-lookup"><span data-stu-id="de46d-135">Choose the **OK** button to close the **Word Options** dialog window.</span></span>

9. <span data-ttu-id="de46d-136">关闭并重新打开 Office 应用程序，以使更改生效。</span><span class="sxs-lookup"><span data-stu-id="de46d-136">Close and reopen the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="de46d-137">旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="de46d-137">Sideload your add-in</span></span>


1. <span data-ttu-id="de46d-138">放入在共享文件夹目录中进行测试的所有加载项的清单 XML 文件。</span><span class="sxs-lookup"><span data-stu-id="de46d-138">Put the manifest file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="de46d-139">请务必将 Web 应用程序本身部署到 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="de46d-139">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="de46d-140">务必在清单文件的 **SourceLocation** 元素中指定 URL。</span><span class="sxs-lookup"><span data-stu-id="de46d-140">Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="de46d-141">在 Excel、Word 或 PowerPoint 中，选择功能区上“**插入**”选项卡中的“**我的加载项**”。</span><span class="sxs-lookup"><span data-stu-id="de46d-141">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="de46d-142">在“**Office 外接程序**”对话框的顶部，选择“**共享文件夹**”。</span><span class="sxs-lookup"><span data-stu-id="de46d-142">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="de46d-143">依次选择加载项名称和“确定”\*\*\*\*，以插入加载项。</span><span class="sxs-lookup"><span data-stu-id="de46d-143">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="de46d-144">另请参阅</span><span class="sxs-lookup"><span data-stu-id="de46d-144">See also</span></span>

- [<span data-ttu-id="de46d-145">验证并排查清单问题</span><span class="sxs-lookup"><span data-stu-id="de46d-145">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="de46d-146">发布 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="de46d-146">Publish your Office Add-in</span></span>](../publish/publish.md)
    
