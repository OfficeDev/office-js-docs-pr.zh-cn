---
title: 在 iPad 和 Mac 上旁加载 Office 加载项以供测试
description: 通过旁加载在 iPad 和 Mac 上测试 Office 外接程序
ms.date: 02/18/2020
localization_priority: Normal
ms.openlocfilehash: 1a1cb804a72aa182480d06009cf30b41a37276d2
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292200"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="1afd5-103">在 iPad 和 Mac 上旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="1afd5-103">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="1afd5-p101">若要查看加载项在 iOS 版 Office 中如何运行，可以使用 iTunes 将加载项的清单旁加载到 iPad，或直接将加载项的清单旁加载到 Mac 版 Office 中。此操作并不能使你在运行时对其设置断点和调试代码，但你可以查看其行为方式，并验证 UI 可用且正确呈现。</span><span class="sxs-lookup"><span data-stu-id="1afd5-p101">To see how your add-in will run in Office on iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office on Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span>

## <a name="prerequisites-for-office-on-ios"></a><span data-ttu-id="1afd5-106">iOS 版 Office 的先决条件</span><span class="sxs-lookup"><span data-stu-id="1afd5-106">Prerequisites for Office on iOS</span></span>

- <span data-ttu-id="1afd5-107">安装了 [iTunes](https://www.apple.com/itunes/download/) 的 Windows 或 Mac 计算机。</span><span class="sxs-lookup"><span data-stu-id="1afd5-107">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>

- <span data-ttu-id="1afd5-108">安装了 [iPad 版 Excel](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) 的 iPad（运行 iOS 8.2 或更高版本）以及同步电缆。</span><span class="sxs-lookup"><span data-stu-id="1afd5-108">An iPad running iOS 8.2 or later with [Excel on iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) installed, and a sync cable.</span></span>

- <span data-ttu-id="1afd5-109">你想要测试的外接程序的清单 .xml 文件。</span><span class="sxs-lookup"><span data-stu-id="1afd5-109">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="prerequisites-for-office-on-mac"></a><span data-ttu-id="1afd5-110">Mac 版 Office 的先决条件</span><span class="sxs-lookup"><span data-stu-id="1afd5-110">Prerequisites for Office on Mac</span></span>

- <span data-ttu-id="1afd5-111">在已安装 [Mac 版 Office](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) 的情况下可运行 OS X v10.10 "Yosemite" 或更高版本的 Mac。</span><span class="sxs-lookup"><span data-stu-id="1afd5-111">A Mac running OS X v10.10 "Yosemite" or later with [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>

- <span data-ttu-id="1afd5-112">Mac 版本 15.18 (160109) 上的 Word。</span><span class="sxs-lookup"><span data-stu-id="1afd5-112">Word on Mac version 15.18 (160109).</span></span>

- <span data-ttu-id="1afd5-113">Mac 版本 15.19 (160206) 上的 Excel。</span><span class="sxs-lookup"><span data-stu-id="1afd5-113">Excel on Mac version 15.19 (160206).</span></span>

- <span data-ttu-id="1afd5-114">Mac 版本 15.24 (160614) 上的 PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1afd5-114">PowerPoint on Mac version 15.24 (160614)</span></span>

- <span data-ttu-id="1afd5-115">你想要测试的外接程序的清单 .xml 文件。</span><span class="sxs-lookup"><span data-stu-id="1afd5-115">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad"></a><span data-ttu-id="1afd5-116">在 iPad 版 Excel 或 iPad 版 Word 上旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="1afd5-116">Sideload an add-in on Excel or Word on iPad</span></span>

1. <span data-ttu-id="1afd5-117">使用同步电缆将 iPad 连接到你的计算机。</span><span class="sxs-lookup"><span data-stu-id="1afd5-117">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="1afd5-118">如果你是首次将 iPad 连接到你的计算机，系统会提示你 **信任此计算机？**。</span><span class="sxs-lookup"><span data-stu-id="1afd5-118">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="1afd5-119">选择“**信任**”继续执行操作。</span><span class="sxs-lookup"><span data-stu-id="1afd5-119">Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="1afd5-120">在 iTunes 中，选择菜单栏下的“iPad”\*\*\*\* 图标。</span><span class="sxs-lookup"><span data-stu-id="1afd5-120">In iTunes, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="1afd5-121">在 iTunes 左侧的“设置”\*\*\*\* 下，选择“应用程序”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="1afd5-121">Under **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="1afd5-122">在 iTunes 右侧，向下滚动到“文件共享”\*\*\*\*，然后在“外接程序”\*\*\*\* 列下选择“Excel”\*\*\*\* 或“Word”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="1afd5-122">On the right side of iTunes, scroll down to **File Sharing**, and then choose **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="1afd5-123">在 " **Excel** " 或 " **Word 文档** " 列底部，选择 " **添加文件**"，然后选择要旁加载的加载项的清单 .xml 文件。</span><span class="sxs-lookup"><span data-stu-id="1afd5-123">At the bottom of the **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span>

6. <span data-ttu-id="1afd5-124">在你的 iPad 上打开 Excel 或 Word 应用。</span><span class="sxs-lookup"><span data-stu-id="1afd5-124">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="1afd5-125">如果 Excel 或 Word 应用程序已在运行，请选择 " **主页** " 按钮，然后关闭并重新启动该应用。</span><span class="sxs-lookup"><span data-stu-id="1afd5-125">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

7. <span data-ttu-id="1afd5-126">打开一个文档。</span><span class="sxs-lookup"><span data-stu-id="1afd5-126">Open a document.</span></span>

8. <span data-ttu-id="1afd5-127">在 "**插入**" 选项卡上选择 "**外接程序**"。您的旁加载外接程序可在 "**外接程序**" UI 中的 "**开发人员**" 标题下插入。</span><span class="sxs-lookup"><span data-stu-id="1afd5-127">Choose **Add-ins** on the **Insert** tab. Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![在 Excel 应用程序中插入的加载项](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a><span data-ttu-id="1afd5-129">在 Mac 版 Office 中旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="1afd5-129">Sideload an add-in in Office on Mac</span></span>

> [!NOTE]
> <span data-ttu-id="1afd5-130">若要旁加载 Mac 版 Outlook 加载项，请参阅[旁加载 Outlook 加载项进行测试](../outlook/sideload-outlook-add-ins-for-testing.md)。</span><span class="sxs-lookup"><span data-stu-id="1afd5-130">To sideload an Outlook add-in on Mac, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="1afd5-131">打开 " **终端** " 并转到以下文件夹之一，您将在其中保存外接程序的清单文件。</span><span class="sxs-lookup"><span data-stu-id="1afd5-131">Open **Terminal** and go to one of the following folders where you'll save your add-in's manifest file.</span></span> <span data-ttu-id="1afd5-132">如果 `wef` 文件夹在你的计算机上不存在，请创建它。</span><span class="sxs-lookup"><span data-stu-id="1afd5-132">If the `wef` folder doesn't exist on your computer, create it.</span></span>

    - <span data-ttu-id="1afd5-133">对于 Word：`/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="1afd5-133">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span></span>    
    - <span data-ttu-id="1afd5-134">对于 Excel：`/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="1afd5-134">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span></span>
    - <span data-ttu-id="1afd5-135">对于 PowerPoint：`/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="1afd5-135">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span></span>

2. <span data-ttu-id="1afd5-136">使用命令**Finder** `open .` (包含句点或点) 的命令在查找器中打开文件夹。</span><span class="sxs-lookup"><span data-stu-id="1afd5-136">Open the folder in **Finder** using the command `open .` (including the period or dot).</span></span> <span data-ttu-id="1afd5-137">将你的外接程序的清单文件复制到该文件夹中。</span><span class="sxs-lookup"><span data-stu-id="1afd5-137">Copy your add-in's manifest file to this folder.</span></span>

    ![Mac 版 Office 中的 Wef 文件夹](../images/all-my-files.png)

3. <span data-ttu-id="1afd5-p106">打开 Word，然后打开一个文档。如果 Word 已运行，则重新启动它。</span><span class="sxs-lookup"><span data-stu-id="1afd5-p106">Open Word, and then open a document. Restart Word if it's already running.</span></span>

4. <span data-ttu-id="1afd5-141">在 Word 中，选择 "在外接程序中**插入**  >  **外**接程序  >  **"** (下拉菜单) ，然后选择您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="1afd5-141">In Word, choose **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>

    ![Mac 版 Office 中的“我的加载项”](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="1afd5-p107">旁加载的加载项不会显示在“我的加载项”对话框中。它们仅显示在下拉菜单中（单击“插入”\*\*\*\* 选项卡上“我的加载项”右侧的向下小箭头）。旁加载的加载项在此菜单中的“开发人员加载项”\*\*\*\* 标题下列出。</span><span class="sxs-lookup"><span data-stu-id="1afd5-p107">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span>

5. <span data-ttu-id="1afd5-146">验证加载项是否在 Word 中显示。</span><span class="sxs-lookup"><span data-stu-id="1afd5-146">Verify that your add-in is displayed in Word.</span></span>

    ![Mac 版 Office 中显示的 Office 加载项](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="1afd5-148">删除旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="1afd5-148">Remove a sideloaded add-in</span></span>

<span data-ttu-id="1afd5-149">您可以通过清除计算机上的 Office 缓存来删除以前的旁加载外接程序。</span><span class="sxs-lookup"><span data-stu-id="1afd5-149">You can remove a previously sideloaded add-in by clearing the Office cache on your computer.</span></span> <span data-ttu-id="1afd5-150">有关如何清除每个平台和应用程序的缓存的详细信息，请参阅 [清除 Office 缓存](clear-cache.md)中的一文。</span><span class="sxs-lookup"><span data-stu-id="1afd5-150">Details on how to clear the cache for each platform and application can be found in the article [Clear the Office cache](clear-cache.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="1afd5-151">另请参阅</span><span class="sxs-lookup"><span data-stu-id="1afd5-151">See also</span></span>

- [<span data-ttu-id="1afd5-152">在 iPad 和 Mac 上调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="1afd5-152">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
