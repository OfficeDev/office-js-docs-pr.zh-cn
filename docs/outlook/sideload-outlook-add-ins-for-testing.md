---
title: 旁加载 Outlook 外接程序进行测试
description: 利用旁加载来安装 Outlook 外接程序以供测试，无需先将其置于外接程序目录中。
ms.date: 12/01/2020
localization_priority: Normal
ms.openlocfilehash: dea2125ccd64eba2e3f1695c8ca1111a710321a4
ms.sourcegitcommit: c2fd7f982f3da748ef6be5c3a7434d859f8b46b9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/02/2020
ms.locfileid: "49530925"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="b720f-103">旁加载 Outlook 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="b720f-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="b720f-104">可以使用旁加载安装 Outlook 外接程序进行测试，而无需首先将其置于外接程序目录中。</span><span class="sxs-lookup"><span data-stu-id="b720f-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-web"></a><span data-ttu-id="b720f-105">在 Outlook 网页版中旁加载外接程序</span><span class="sxs-lookup"><span data-stu-id="b720f-105">Sideload an add-in in Outlook on the web</span></span>

<span data-ttu-id="b720f-106">在 web 上的 Outlook 中旁加载外接程序的过程取决于您使用的是新版本还是经典版本。</span><span class="sxs-lookup"><span data-stu-id="b720f-106">The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.</span></span>

- <span data-ttu-id="b720f-107">如果邮箱工具栏类似于下图，请参阅[在全新 Outlook 网页版中旁加载外接程序](#sideload-an-add-in-in-the-new-outlook-on-the-web)。</span><span class="sxs-lookup"><span data-stu-id="b720f-107">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span></span>

    ![全新 Outlook 网页版工具栏的部分屏幕截图](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="b720f-109">如果邮箱工具栏类似于下图，请参阅[在经典 Outlook 网页版中旁加载外接程序](#sideload-an-add-in-in-classic-outlook-on-the-web)。</span><span class="sxs-lookup"><span data-stu-id="b720f-109">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#sideload-an-add-in-in-classic-outlook-on-the-web).</span></span>

    ![经典 Outlook 网页版工具栏的部分屏幕截图](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="b720f-111">如果你的组织在邮箱工具栏中添加了自己的徽标，则你看到的界面可能会与前面的图像略有不同。</span><span class="sxs-lookup"><span data-stu-id="b720f-111">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="sideload-an-add-in-in-the-new-outlook-on-the-web"></a><span data-ttu-id="b720f-112">在全新 Outlook 网页版中旁加载外接程序</span><span class="sxs-lookup"><span data-stu-id="b720f-112">Sideload an add-in in the new Outlook on the web</span></span>

1. <span data-ttu-id="b720f-113">转到 [Office 365 中的 Outlook](https://outlook.office.com)。</span><span class="sxs-lookup"><span data-stu-id="b720f-113">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="b720f-114">在 Outlook 网页版中，创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="b720f-114">In Outlook on the web, create a new message.</span></span>

1. <span data-ttu-id="b720f-115">从新邮件的底部选择 **...**，然后从出现的菜单中选择“**获取外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="b720f-115">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![全新 Outlook 网页版中的邮件撰写窗口（突出显示了“获取外接程序”选项）](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="b720f-117">在“**Outlook 外接程序**”对话框中，选择“**我的外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="b720f-117">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![全新 Outlook 网页版中的“Outlook 外接程序”对话框（已选中“我的外接程序”）](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="b720f-119">在对话框底部找到“**自定义外接程序**”部分。</span><span class="sxs-lookup"><span data-stu-id="b720f-119">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="b720f-120">选择“**添加自定义外接程序**”链接，然后选择“**从文件添加**”。</span><span class="sxs-lookup"><span data-stu-id="b720f-120">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![指向“从文件添加”选项的“管理外接程序”屏幕截图](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="b720f-p102">找到自定义外接程序的清单文件并进行安装。在安装过程中接受所有提示。</span><span class="sxs-lookup"><span data-stu-id="b720f-p102">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="sideload-an-add-in-in-classic-outlook-on-the-web"></a><span data-ttu-id="b720f-124">在经典 Outlook 网页版中旁加载外接程序</span><span class="sxs-lookup"><span data-stu-id="b720f-124">Sideload an add-in in classic Outlook on the web</span></span>

1. <span data-ttu-id="b720f-125">转到 [Office 365 中的 Outlook](https://outlook.office.com)。</span><span class="sxs-lookup"><span data-stu-id="b720f-125">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="b720f-126">选择右上部分的齿轮图标，然后选择“**管理外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="b720f-126">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Outlook 网页版屏幕截图（指向“管理外接程序”选项）](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="b720f-128">在“管理加载项”页中，选择“加载项”，然后选择“我的加载项”。</span><span class="sxs-lookup"><span data-stu-id="b720f-128">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Outlook 网页版应用商店对话框（已选中“我的外接程序”）](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="b720f-130">在对话框底部找到“**自定义外接程序**”部分。</span><span class="sxs-lookup"><span data-stu-id="b720f-130">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="b720f-131">选择“**添加自定义外接程序**”链接，然后选择“**从文件添加**”。</span><span class="sxs-lookup"><span data-stu-id="b720f-131">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![指向“从文件添加”选项的“管理外接程序”屏幕截图](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="b720f-p104">找到您的自定义外接程序清单文件并进行安装。在安装过程中接受所有提示。</span><span class="sxs-lookup"><span data-stu-id="b720f-p104">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-desktop"></a><span data-ttu-id="b720f-135">在 Outlook 桌面版中旁加载外接程序</span><span class="sxs-lookup"><span data-stu-id="b720f-135">Sideload an add-in in Outlook on the desktop</span></span>

### <a name="outlook-2016-or-later"></a><span data-ttu-id="b720f-136">Outlook 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="b720f-136">Outlook 2016 or later</span></span>

1. <span data-ttu-id="b720f-137">在 Windows 或 Mac 上打开 Outlook 2016 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="b720f-137">Open Outlook 2016 or later on Windows or Mac.</span></span>

1. <span data-ttu-id="b720f-138">选择功能区上的“**获取外接程序**”按钮。</span><span class="sxs-lookup"><span data-stu-id="b720f-138">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![Outlook 2016 功能区指向 "获取外接程序" 按钮](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > <span data-ttu-id="b720f-140">如果您在您的 Outlook 版本中看不到 " **获取外接程序** " 按钮，请选择：</span><span class="sxs-lookup"><span data-stu-id="b720f-140">If you don't see the **Get Add-ins** button in your version of Outlook, select:</span></span>
    >
    > - <span data-ttu-id="b720f-141">功能区上的 "**存储**" 按钮（如果有）。</span><span class="sxs-lookup"><span data-stu-id="b720f-141">**Store** button on the ribbon, if available.</span></span>
    >
    >   <span data-ttu-id="b720f-142">OR</span><span class="sxs-lookup"><span data-stu-id="b720f-142">OR</span></span>
    >
    > - <span data-ttu-id="b720f-143">**文件**"菜单上，选择"**信息**"选项卡上的"**管理外接程序**"按钮，以在 Outlook 网页中打开"**外** 接程序 "对话框。</span><span class="sxs-lookup"><span data-stu-id="b720f-143">**File** menu, then select the **Manage Add-ins** button on the **Info** tab to open the **Add-ins** dialog in Outlook on the web.</span></span><br><span data-ttu-id="b720f-144">您可以在 [Outlook 网页](#sideload-an-add-in-in-outlook-on-the-web)版本的旁加载中查看有关上一节中的 web 体验的详细信息。</span><span class="sxs-lookup"><span data-stu-id="b720f-144">You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#sideload-an-add-in-in-outlook-on-the-web).</span></span>

1. <span data-ttu-id="b720f-145">如果对话框顶部附近有选项卡，请确保选择了 " **外接程序** " 选项卡。</span><span class="sxs-lookup"><span data-stu-id="b720f-145">If there are tabs near the top of the dialog, ensure that the **Add-ins** tab is selected.</span></span> <span data-ttu-id="b720f-146">选择 **"我的外接程序"**。</span><span class="sxs-lookup"><span data-stu-id="b720f-146">Choose **My add-ins**.</span></span>

    ![Outlook 2016 应用商店对话框（已选中“我的外接程序”）](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="b720f-148">在对话框底部找到“自定义加载项”部分。</span><span class="sxs-lookup"><span data-stu-id="b720f-148">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="b720f-149">选择“添加自定义加载项”链接，然后选择“从文件添加”。</span><span class="sxs-lookup"><span data-stu-id="b720f-149">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![“应用商店”屏幕截图（指向“从文件添加”选项）](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="b720f-p107">找到您的自定义外接程序清单文件并进行安装。在安装过程中接受所有提示。</span><span class="sxs-lookup"><span data-stu-id="b720f-p107">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="outlook-2013"></a><span data-ttu-id="b720f-153">Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="b720f-153">Outlook 2013</span></span>

1. <span data-ttu-id="b720f-154">在 Windows 上打开 Outlook 2013。</span><span class="sxs-lookup"><span data-stu-id="b720f-154">Open Outlook 2013 on Windows.</span></span>

1. <span data-ttu-id="b720f-155">选择 "**文件**" 菜单，然后选择 "**信息**" 选项卡上的 "**管理外接程序**" 按钮。Outlook 将在浏览器中打开 web 版本。</span><span class="sxs-lookup"><span data-stu-id="b720f-155">Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open the web version in a browser.</span></span>

1. <span data-ttu-id="b720f-156">按照您的 Outlook 网页版上的 [旁加载中的加载](#sideload-an-add-in-in-outlook-on-the-web) 项中的步骤，在 web 上的 outlook 的 "web" 部分中执行。</span><span class="sxs-lookup"><span data-stu-id="b720f-156">Follow the steps in the [Sideload an add-in in Outlook on the web](#sideload-an-add-in-in-outlook-on-the-web) section according to your version of Outlook on the web.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="b720f-157">删除旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="b720f-157">Remove a sideloaded add-in</span></span>

<span data-ttu-id="b720f-158">若要从 Outlook 中删除旁加载外接程序，请使用本文中前面所述的步骤，在列出已安装加载项的对话框的 " **自定义外接程序** " 部分中查找该外接程序。选择外接程序的省略号 (`...`) ，然后选择 " **删除** " 以删除该特定外接程序。</span><span class="sxs-lookup"><span data-stu-id="b720f-158">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the the add-in and then choose **Remove** to remove that specific add-in.</span></span>