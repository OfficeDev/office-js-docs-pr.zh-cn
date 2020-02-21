---
title: 旁加载 Outlook 外接程序进行测试
description: 利用旁加载来安装 Outlook 外接程序以供测试，无需先将其置于外接程序目录中。
ms.date: 06/24/2019
localization_priority: Normal
ms.openlocfilehash: b177e6adbc4ac702b7bd9dcec38f2fe2d2f29cf1
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165887"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="3e401-103">旁加载 Outlook 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="3e401-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="3e401-104">可以使用旁加载安装 Outlook 外接程序进行测试，而无需首先将其置于外接程序目录中。</span><span class="sxs-lookup"><span data-stu-id="3e401-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>


## <a name="sideload-an-add-in-in-outlook-in-office-365"></a><span data-ttu-id="3e401-105">在 Office 365 的 Outlook 中旁加载外接程序</span><span class="sxs-lookup"><span data-stu-id="3e401-105">Sideload an add-in in Outlook in Office 365</span></span>

<span data-ttu-id="3e401-106">在 Office 365 的 Outlook 中旁加载外接程序的过程取决于使用的是全新 Outlook 网页版还是经典 Outlook 网页版。</span><span class="sxs-lookup"><span data-stu-id="3e401-106">The process for sideloading an add-in in Outlook in Office 365 depends upon whether you are using the new Outlook on the web or classic Outlook on the web.</span></span>

- <span data-ttu-id="3e401-107">如果邮箱工具栏类似于下图，请参阅[在全新 Outlook 网页版中旁加载外接程序](#sideload-an-add-in-in-the-new-outlook-on-the-web)。</span><span class="sxs-lookup"><span data-stu-id="3e401-107">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span></span>

    ![全新 Outlook 网页版工具栏的部分屏幕截图](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="3e401-109">如果邮箱工具栏类似于下图，请参阅[在经典 Outlook 网页版中旁加载外接程序](#sideload-an-add-in-in-classic-outlook-on-the-web)。</span><span class="sxs-lookup"><span data-stu-id="3e401-109">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#sideload-an-add-in-in-classic-outlook-on-the-web).</span></span>

    ![经典 Outlook 网页版工具栏的部分屏幕截图](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="3e401-111">如果你的组织在邮箱工具栏中添加了自己的徽标，则你看到的界面可能会与前面的图像略有不同。</span><span class="sxs-lookup"><span data-stu-id="3e401-111">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="sideload-an-add-in-in-the-new-outlook-on-the-web"></a><span data-ttu-id="3e401-112">在全新 Outlook 网页版中旁加载外接程序</span><span class="sxs-lookup"><span data-stu-id="3e401-112">Sideload an add-in in the new Outlook on the web</span></span>

1. <span data-ttu-id="3e401-113">转到 [Office 365 中的 Outlook](https://outlook.office.com)。</span><span class="sxs-lookup"><span data-stu-id="3e401-113">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="3e401-114">在 Outlook 网页版中，创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="3e401-114">In Outlook on the web, create a new message.</span></span>   

1. <span data-ttu-id="3e401-115">从新邮件的底部选择 **...**，然后从出现的菜单中选择“**获取外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="3e401-115">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![全新 Outlook 网页版中的邮件撰写窗口（突出显示了“获取外接程序”选项）](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="3e401-117">在“**Outlook 外接程序**”对话框中，选择“**我的外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="3e401-117">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![全新 Outlook 网页版中的“Outlook 外接程序”对话框（已选中“我的外接程序”）](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="3e401-119">在对话框底部找到“**自定义外接程序**”部分。</span><span class="sxs-lookup"><span data-stu-id="3e401-119">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="3e401-120">选择“**添加自定义外接程序**”链接，然后选择“**从文件添加**”。</span><span class="sxs-lookup"><span data-stu-id="3e401-120">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![指向“从文件添加”选项的“管理外接程序”屏幕截图](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="3e401-p102">找到自定义外接程序的清单文件并进行安装。在安装过程中接受所有提示。</span><span class="sxs-lookup"><span data-stu-id="3e401-p102">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="sideload-an-add-in-in-classic-outlook-on-the-web"></a><span data-ttu-id="3e401-124">在经典 Outlook 网页版中旁加载外接程序</span><span class="sxs-lookup"><span data-stu-id="3e401-124">Sideload an add-in in classic Outlook on the web</span></span>

1. <span data-ttu-id="3e401-125">转到 [Office 365 中的 Outlook](https://outlook.office.com)。</span><span class="sxs-lookup"><span data-stu-id="3e401-125">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="3e401-126">选择右上部分的齿轮图标，然后选择“**管理外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="3e401-126">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Outlook 网页版屏幕截图（指向“管理外接程序”选项）](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="3e401-128">在“管理加载项”\*\*\*\* 页中，选择“加载项”\*\*\*\*，然后选择“我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3e401-128">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Outlook 网页版应用商店对话框（已选中“我的外接程序”）](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="3e401-130">在对话框底部找到“**自定义外接程序**”部分。</span><span class="sxs-lookup"><span data-stu-id="3e401-130">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="3e401-131">选择“**添加自定义外接程序**”链接，然后选择“**从文件添加**”。</span><span class="sxs-lookup"><span data-stu-id="3e401-131">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![指向“从文件添加”选项的“管理外接程序”屏幕截图](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="3e401-p104">找到您的自定义外接程序清单文件并进行安装。在安装过程中接受所有提示。</span><span class="sxs-lookup"><span data-stu-id="3e401-p104">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-desktop"></a><span data-ttu-id="3e401-135">在 Outlook 桌面版中旁加载外接程序</span><span class="sxs-lookup"><span data-stu-id="3e401-135">Sideload an add-in in Outlook on the desktop</span></span>

1. <span data-ttu-id="3e401-136">打开 Windows 版 Outlook 2013 或更高版本，或 Mac 版 Outlook 2016 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="3e401-136">Open Outlook 2013 or later on Windows, or Outlook 2016 or later on Mac.</span></span>

1. <span data-ttu-id="3e401-137">选择功能区上的“**获取外接程序**”按钮。</span><span class="sxs-lookup"><span data-stu-id="3e401-137">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![Outlook 2016 功能区（指向“应用商店”按钮）](../images/outlook-sideload-desktop-store.png)

    > [!NOTE]
    > <span data-ttu-id="3e401-139">如果没有在使用的 Outlook 版本中看到“**获取外接程序**”按钮，请改为选择功能区上的“**应用商店**”按钮。</span><span class="sxs-lookup"><span data-stu-id="3e401-139">If you don't see the **Get Add-ins** button in your version of Outlook, select the **Store** button on the ribbon instead.</span></span>

1. <span data-ttu-id="3e401-140">选择“**外接程序**”，然后选择“**我的外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="3e401-140">Select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Outlook 2016 应用商店对话框（已选中“我的外接程序”）](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="3e401-142">在对话框底部找到“自定义加载项”\*\*\*\* 部分。</span><span class="sxs-lookup"><span data-stu-id="3e401-142">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="3e401-143">选择“添加自定义加载项”\*\*\*\* 链接，然后选择“从文件添加”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3e401-143">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![“应用商店”屏幕截图（指向“从文件添加”选项）](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="3e401-p106">找到您的自定义外接程序清单文件并进行安装。在安装过程中接受所有提示。</span><span class="sxs-lookup"><span data-stu-id="3e401-p106">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="3e401-147">删除旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="3e401-147">Remove a sideloaded add-in</span></span>

<span data-ttu-id="3e401-148">若要从 Outlook 中删除旁加载加载项，请使用本文前面所述的步骤，在列出已安装加载项的对话框的 "**自定义外接程序**" 部分中查找加载项。选择外接程序的省略号`...`（），然后选择 "**删除**" 以删除该特定外接程序。</span><span class="sxs-lookup"><span data-stu-id="3e401-148">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the the add-in and then choose **Remove** to remove that specific add-in.</span></span>