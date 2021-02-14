---
title: 旁加载 Outlook 外接程序进行测试
description: 利用旁加载来安装 Outlook 外接程序以供测试，无需先将其置于外接程序目录中。
ms.date: 02/10/2021
localization_priority: Normal
ms.openlocfilehash: b783b815af84a7fd8b4abd52cdd8e0925bfb9ecf
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234245"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="b663b-103">旁加载 Outlook 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="b663b-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="b663b-104">可以使用旁加载安装 Outlook 外接程序进行测试，而无需首先将其置于外接程序目录中。</span><span class="sxs-lookup"><span data-stu-id="b663b-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>

## <a name="sideload-automatically"></a><span data-ttu-id="b663b-105">自动旁加载</span><span class="sxs-lookup"><span data-stu-id="b663b-105">Sideload automatically</span></span>

<span data-ttu-id="b663b-106">如果使用适用于 Office 加载项的 [Yeoman](https://github.com/OfficeDev/generator-office)生成器创建了 Outlook 外接程序，则最好通过命令行完成旁加载。</span><span class="sxs-lookup"><span data-stu-id="b663b-106">If you created your Outlook add-in using [the Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), sideloading is best done through the command line.</span></span> <span data-ttu-id="b663b-107">这将利用我们的工具和在一个命令中跨所有受支持的设备旁加载。</span><span class="sxs-lookup"><span data-stu-id="b663b-107">This will take advantage of our tooling and sideload across all of your supported devices in one command.</span></span>

1. <span data-ttu-id="b663b-108">使用命令行导航到 Yeoman 生成的加载项项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="b663b-108">Using the command line, navigate to the root directory of your Yeoman generated add-in project.</span></span> <span data-ttu-id="b663b-109">运行命令 `npm start`。</span><span class="sxs-lookup"><span data-stu-id="b663b-109">Run the command `npm start`.</span></span>

2. <span data-ttu-id="b663b-110">Outlook 外接程序将自动旁加载到台式计算机上 Outlook。</span><span class="sxs-lookup"><span data-stu-id="b663b-110">Your Outlook add-in will automatically sideload to Outlook on your desktop computer.</span></span> <span data-ttu-id="b663b-111">将显示一个对话框，说明尝试旁加载外接程序，并列出清单文件的名称和位置。</span><span class="sxs-lookup"><span data-stu-id="b663b-111">You'll see a dialog appear, stating there is an attempt to sideload the add-in, listing the name and the location of the manifest file.</span></span> <span data-ttu-id="b663b-112">选择 **"** 确定"，这将注册清单。</span><span class="sxs-lookup"><span data-stu-id="b663b-112">Select **OK**, which will register the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b663b-113">如果清单包含错误或清单路径无效，您将收到一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="b663b-113">If the manifest contains an error or the path to the manifest is invalid, you'll receive an error message.</span></span>

3. <span data-ttu-id="b663b-114">如果清单不包含任何错误且路径有效，则现在您的外接程序将旁加载，并且可在桌面和 Web 上的 Outlook 中使用。</span><span class="sxs-lookup"><span data-stu-id="b663b-114">If your manifest contains no errors and the path is valid, your add-in will now be sideloaded and available on both your desktop and in Outlook on the web.</span></span> <span data-ttu-id="b663b-115">它还将安装在所有受支持的设备上。</span><span class="sxs-lookup"><span data-stu-id="b663b-115">It will also be installed across all your supported devices.</span></span>

## <a name="sideload-manually"></a><span data-ttu-id="b663b-116">手动旁加载</span><span class="sxs-lookup"><span data-stu-id="b663b-116">Sideload manually</span></span>

<span data-ttu-id="b663b-117">尽管我们强烈建议通过命令行自动旁加载，如上一节所述，但您也可以基于 Outlook 客户端手动旁加载 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="b663b-117">Though we strongly recommend sideloading automatically through the command line as covered in the previous section, you can also manually sideload an Outlook add-in based on the Outlook client.</span></span>

### <a name="outlook-on-the-web"></a><span data-ttu-id="b663b-118">Outlook 网页版</span><span class="sxs-lookup"><span data-stu-id="b663b-118">Outlook on the web</span></span>

<span data-ttu-id="b663b-119">在 Outlook 网页版中旁加载外接程序的过程取决于您使用的是新版本还是经典版本。</span><span class="sxs-lookup"><span data-stu-id="b663b-119">The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.</span></span>

- <span data-ttu-id="b663b-120">如果邮箱工具栏类似于下图，请参阅[在全新 Outlook 网页版中旁加载外接程序](#new-outlook-on-the-web)。</span><span class="sxs-lookup"><span data-stu-id="b663b-120">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#new-outlook-on-the-web).</span></span>

    ![全新 Outlook 网页版工具栏的部分屏幕截图](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="b663b-122">如果邮箱工具栏类似于下图，请参阅[在经典 Outlook 网页版中旁加载外接程序](#classic-outlook-on-the-web)。</span><span class="sxs-lookup"><span data-stu-id="b663b-122">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#classic-outlook-on-the-web).</span></span>

    ![经典 Outlook 网页版工具栏的部分屏幕截图](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="b663b-124">如果你的组织在邮箱工具栏中添加了自己的徽标，则你看到的界面可能会与前面的图像略有不同。</span><span class="sxs-lookup"><span data-stu-id="b663b-124">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="new-outlook-on-the-web"></a><span data-ttu-id="b663b-125">新 Outlook 网页</span><span class="sxs-lookup"><span data-stu-id="b663b-125">New Outlook on the web</span></span>

1. <span data-ttu-id="b663b-126">转到 [Outlook 网页版](https://outlook.office.com)。</span><span class="sxs-lookup"><span data-stu-id="b663b-126">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="b663b-127">创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="b663b-127">Create a new message.</span></span>

1. <span data-ttu-id="b663b-128">从新邮件的底部选择 **...**，然后从出现的菜单中选择“**获取外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="b663b-128">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![全新 Outlook 网页版中的邮件撰写窗口（突出显示了“获取外接程序”选项）](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="b663b-130">在“**Outlook 外接程序**”对话框中，选择“**我的外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="b663b-130">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![全新 Outlook 网页版中的“Outlook 外接程序”对话框（已选中“我的外接程序”）](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="b663b-132">在对话框底部找到“**自定义外接程序**”部分。</span><span class="sxs-lookup"><span data-stu-id="b663b-132">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="b663b-133">选择“**添加自定义外接程序**”链接，然后选择“**从文件添加**”。</span><span class="sxs-lookup"><span data-stu-id="b663b-133">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![指向“从文件添加”选项的“管理外接程序”屏幕截图](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="b663b-p106">找到您的自定义外接程序清单文件并进行安装。在安装过程中接受所有提示。</span><span class="sxs-lookup"><span data-stu-id="b663b-p106">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="classic-outlook-on-the-web"></a><span data-ttu-id="b663b-137">经典 Outlook 网页</span><span class="sxs-lookup"><span data-stu-id="b663b-137">Classic Outlook on the web</span></span>

1. <span data-ttu-id="b663b-138">转到 [Outlook 网页版](https://outlook.office.com)。</span><span class="sxs-lookup"><span data-stu-id="b663b-138">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="b663b-139">选择右上部分的齿轮图标，然后选择“**管理外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="b663b-139">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Outlook 网页版屏幕截图（指向“管理外接程序”选项）](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="b663b-141">在“管理加载项”页中，选择“加载项”，然后选择“我的加载项”。</span><span class="sxs-lookup"><span data-stu-id="b663b-141">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Outlook 网页版应用商店对话框（已选中“我的外接程序”）](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="b663b-143">在对话框底部找到“**自定义外接程序**”部分。</span><span class="sxs-lookup"><span data-stu-id="b663b-143">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="b663b-144">选择“**添加自定义外接程序**”链接，然后选择“**从文件添加**”。</span><span class="sxs-lookup"><span data-stu-id="b663b-144">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![指向“从文件添加”选项的“管理外接程序”屏幕截图](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="b663b-p108">找到您的自定义外接程序清单文件并进行安装。在安装过程中接受所有提示。</span><span class="sxs-lookup"><span data-stu-id="b663b-p108">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="outlook-on-the-desktop"></a><span data-ttu-id="b663b-148">桌面上的 Outlook</span><span class="sxs-lookup"><span data-stu-id="b663b-148">Outlook on the desktop</span></span>

#### <a name="outlook-2016-or-later"></a><span data-ttu-id="b663b-149">Outlook 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="b663b-149">Outlook 2016 or later</span></span>

1. <span data-ttu-id="b663b-150">在 Windows 或 Mac 上打开 Outlook 2016 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="b663b-150">Open Outlook 2016 or later on Windows or Mac.</span></span>

1. <span data-ttu-id="b663b-151">选择功能区上的“**获取外接程序**”按钮。</span><span class="sxs-lookup"><span data-stu-id="b663b-151">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![指向"获取外接程序"按钮的 Outlook 2016 功能区](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > <span data-ttu-id="b663b-153">如果在 Outlook 版本中看不到"获取 **外接程序** "按钮，请选择：</span><span class="sxs-lookup"><span data-stu-id="b663b-153">If you don't see the **Get Add-ins** button in your version of Outlook, select:</span></span>
    >
    > - <span data-ttu-id="b663b-154">**功能** 区上的"存储"按钮（如果可用）。</span><span class="sxs-lookup"><span data-stu-id="b663b-154">**Store** button on the ribbon, if available.</span></span>
    >
    >   <span data-ttu-id="b663b-155">或</span><span class="sxs-lookup"><span data-stu-id="b663b-155">OR</span></span>
    >
    > - <span data-ttu-id="b663b-156">**文件** 菜单，然后选择"信息 **"** 选项卡上的"管理外接程序"按钮，以在 Outlook网页中打开"加载项"对话框。</span><span class="sxs-lookup"><span data-stu-id="b663b-156">**File** menu, then select the **Manage Add-ins** button on the **Info** tab to open the **Add-ins** dialog in Outlook on the web.</span></span><br><span data-ttu-id="b663b-157">可以在上一节中查看有关 Web 体验的更多内容，在 Outlook 网页版中旁 [加载外接程序](#outlook-on-the-web)。</span><span class="sxs-lookup"><span data-stu-id="b663b-157">You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#outlook-on-the-web).</span></span>

1. <span data-ttu-id="b663b-158">如果对话框顶部附近有选项卡，请确保选择了" **加载项** "选项卡。</span><span class="sxs-lookup"><span data-stu-id="b663b-158">If there are tabs near the top of the dialog, ensure that the **Add-ins** tab is selected.</span></span> <span data-ttu-id="b663b-159">选择 **"我的外接程序"。**</span><span class="sxs-lookup"><span data-stu-id="b663b-159">Choose **My add-ins**.</span></span>

    ![Outlook 2016 应用商店对话框（已选中“我的外接程序”）](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="b663b-161">在对话框底部找到“自定义加载项”部分。</span><span class="sxs-lookup"><span data-stu-id="b663b-161">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="b663b-162">选择“添加自定义加载项”链接，然后选择“从文件添加”。</span><span class="sxs-lookup"><span data-stu-id="b663b-162">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![“应用商店”屏幕截图（指向“从文件添加”选项）](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="b663b-p111">找到您的自定义外接程序清单文件并进行安装。在安装过程中接受所有提示。</span><span class="sxs-lookup"><span data-stu-id="b663b-p111">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

#### <a name="outlook-2013"></a><span data-ttu-id="b663b-166">Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="b663b-166">Outlook 2013</span></span>

1. <span data-ttu-id="b663b-167">在 Windows 上打开 Outlook 2013。</span><span class="sxs-lookup"><span data-stu-id="b663b-167">Open Outlook 2013 on Windows.</span></span>

1. <span data-ttu-id="b663b-168">选择 **"** 文件"菜单，然后选择" **信息"** 选项卡上的"管理外接程序 **"** 按钮。Outlook 将在浏览器中打开 Web 版本。</span><span class="sxs-lookup"><span data-stu-id="b663b-168">Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open the web version in a browser.</span></span>

1. <span data-ttu-id="b663b-169">按照 Outlook 网页版中的 [旁](#outlook-on-the-web) 加载外接程序部分中的步骤操作。</span><span class="sxs-lookup"><span data-stu-id="b663b-169">Follow the steps in the [Sideload an add-in in Outlook on the web](#outlook-on-the-web) section according to your version of Outlook on the web.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="b663b-170">删除旁加载的加载项</span><span class="sxs-lookup"><span data-stu-id="b663b-170">Remove a sideloaded add-in</span></span>

<span data-ttu-id="b663b-171">在所有版本的 Outlook 中，删除旁加载的外接程序的关键是列出已安装的外接程序的"我的外接程序"对话框。选择外接程序 () 省略 `...` 号，然后选择"删除 **"。**</span><span class="sxs-lookup"><span data-stu-id="b663b-171">On all versions of Outlook, the key to removing a sideloaded add-in is the **My Add-ins** dialog which lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then select **Remove**.</span></span>

<span data-ttu-id="b663b-172">若要 **导航到** Outlook 客户端的"我的外接程序"对话框，请使用本文前面部分中列出的用于手动旁 [](#sideload-manually)加载的最后步骤。</span><span class="sxs-lookup"><span data-stu-id="b663b-172">To navigate to the **My Add-ins** dialog box for your Outlook client, use the last steps listed for [manual sideloading](#sideload-manually) in the previous sections of this article.</span></span>

<span data-ttu-id="b663b-173">若要从 Outlook 中删除旁加载的外接程序，请使用本文前面介绍的步骤在列出已安装加载项的对话框的"自定义加载项"部分查找外接程序。选择外接程序 () 省略号，然后选择"删除"以删除该特定 `...` 加载项。 </span><span class="sxs-lookup"><span data-stu-id="b663b-173">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the the add-in and then choose **Remove** to remove that specific add-in.</span></span>

