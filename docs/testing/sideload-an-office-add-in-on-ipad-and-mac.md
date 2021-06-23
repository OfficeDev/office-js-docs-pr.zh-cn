---
title: 在 iPad 和 Mac 上旁加载 Office 加载项以供测试
description: 通过旁Office在 iPad 和 Mac 上测试加载项。
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: b3d7d7fa3ee69e849c112c888b66fa9deed23d88
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076201"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="fe888-103">在 iPad 和 Mac 上旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="fe888-103">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="fe888-p101">若要查看加载项在 iOS 版 Office 中如何运行，可以使用 iTunes 将加载项的清单旁加载到 iPad，或直接将加载项的清单旁加载到 Mac 版 Office 中。此操作并不能使你在运行时对其设置断点和调试代码，但你可以查看其行为方式，并验证 UI 可用且正确呈现。</span><span class="sxs-lookup"><span data-stu-id="fe888-p101">To see how your add-in will run in Office on iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office on Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span>

## <a name="prerequisites-for-office-on-ios"></a><span data-ttu-id="fe888-106">iOS 版 Office 的先决条件</span><span class="sxs-lookup"><span data-stu-id="fe888-106">Prerequisites for Office on iOS</span></span>

- <span data-ttu-id="fe888-107">安装了 [iTunes](https://www.apple.com/itunes/download/) 的 Windows 或 Mac 计算机。</span><span class="sxs-lookup"><span data-stu-id="fe888-107">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>
  > [!IMPORTANT]
  > <span data-ttu-id="fe888-108">如果您运行的是 macOS Catalina，[则 iTunes](https://support.apple.com/HT210200)不再可用，因此您应该按照本文稍后使用[macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina)在 Excel 或 iPad 上的 Word 旁加载外接程序部分中的说明进行操作。</span><span class="sxs-lookup"><span data-stu-id="fe888-108">If you're running macOS Catalina, [iTunes is no longer available](https://support.apple.com/HT210200) so you should follow the instructions in the section [Sideload an add-in on Excel or Word on iPad using macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) later in this article.</span></span>

- <span data-ttu-id="fe888-109">运行iPad iOS 8.2 或更高版本，Excel [](https://apps.apple.com/app/microsoft-excel/id586683407) [Word](https://apps.apple.com/app/microsoft-word/id586447913)和同步电缆。</span><span class="sxs-lookup"><span data-stu-id="fe888-109">An iPad running iOS 8.2 or later with [Excel](https://apps.apple.com/app/microsoft-excel/id586683407) or [Word](https://apps.apple.com/app/microsoft-word/id586447913) installed, and a sync cable.</span></span>

- <span data-ttu-id="fe888-110">你想要测试的外接程序的清单 .xml 文件。</span><span class="sxs-lookup"><span data-stu-id="fe888-110">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="prerequisites-for-office-on-mac"></a><span data-ttu-id="fe888-111">Mac 版 Office 的先决条件</span><span class="sxs-lookup"><span data-stu-id="fe888-111">Prerequisites for Office on Mac</span></span>

- <span data-ttu-id="fe888-112">在已安装 [Mac 版 Office](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) 的情况下可运行 OS X v10.10 "Yosemite" 或更高版本的 Mac。</span><span class="sxs-lookup"><span data-stu-id="fe888-112">A Mac running OS X v10.10 "Yosemite" or later with [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>

- <span data-ttu-id="fe888-113">Mac 版本 15.18 (160109) 上的 Word。</span><span class="sxs-lookup"><span data-stu-id="fe888-113">Word on Mac version 15.18 (160109).</span></span>

- <span data-ttu-id="fe888-114">Mac 版本 15.19 (160206) 上的 Excel。</span><span class="sxs-lookup"><span data-stu-id="fe888-114">Excel on Mac version 15.19 (160206).</span></span>

- <span data-ttu-id="fe888-115">Mac 版本 15.24 (160614) 上的 PowerPoint</span><span class="sxs-lookup"><span data-stu-id="fe888-115">PowerPoint on Mac version 15.24 (160614)</span></span>

- <span data-ttu-id="fe888-116">你想要测试的外接程序的清单 .xml 文件。</span><span class="sxs-lookup"><span data-stu-id="fe888-116">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a><span data-ttu-id="fe888-117">使用 iTunes 在 Excel 或 Word iPad旁加载外接程序</span><span class="sxs-lookup"><span data-stu-id="fe888-117">Sideload an add-in on Excel or Word on iPad using iTunes</span></span>

1. <span data-ttu-id="fe888-118">使用同步电缆将 iPad 连接到你的计算机。</span><span class="sxs-lookup"><span data-stu-id="fe888-118">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="fe888-119">如果是首次将 iPad连接到计算机，系统将提示"信任此计算机 **？"。**</span><span class="sxs-lookup"><span data-stu-id="fe888-119">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="fe888-120">选择“**信任**”继续执行操作。</span><span class="sxs-lookup"><span data-stu-id="fe888-120">Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="fe888-121">在 iTunes 中，选择菜单栏下的“iPad”图标。</span><span class="sxs-lookup"><span data-stu-id="fe888-121">In iTunes, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="fe888-122">在 iTunes 左侧的“设置”下，选择“应用程序”。</span><span class="sxs-lookup"><span data-stu-id="fe888-122">Under **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="fe888-123">在 iTunes 右侧，向下滚动到“文件共享”，然后在“外接程序”列下选择“Excel”或“Word”。</span><span class="sxs-lookup"><span data-stu-id="fe888-123">On the right side of iTunes, scroll down to **File Sharing**, and then choose **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="fe888-124">At the bottom of the **Excel** or **Word Documents** column， choose **Add File**， and then select the manifest .xml file of the add-in you want to sideload.</span><span class="sxs-lookup"><span data-stu-id="fe888-124">At the bottom of the **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span>

6. <span data-ttu-id="fe888-125">在你的 iPad 上打开 Excel 或 Word 应用。</span><span class="sxs-lookup"><span data-stu-id="fe888-125">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="fe888-126">如果Excel Word 应用已在运行，请选择"开始"按钮，然后关闭并重新启动应用。</span><span class="sxs-lookup"><span data-stu-id="fe888-126">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

7. <span data-ttu-id="fe888-127">打开一个文档。</span><span class="sxs-lookup"><span data-stu-id="fe888-127">Open a document.</span></span>

8. <span data-ttu-id="fe888-128">选择"插入"选项卡上的"外接程序"。 (在"插入"选项卡上，您可能需要水平滚动，直到您看到"外接程序"按钮。) 您的旁加载的外接程序可在外接程序 **UI** 中的"开发人员"标题下插入。  </span><span class="sxs-lookup"><span data-stu-id="fe888-128">Choose **Add-ins** on the **Insert** tab. (On the **Insert** tab, you may need to scroll horizontally until you see the **Add-ins** button.) Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![在加载项应用中插入Excel加载项。](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a><span data-ttu-id="fe888-130">使用 macOS 加泰罗尼亚语在 Excel 或 Word iPad旁加载外接程序</span><span class="sxs-lookup"><span data-stu-id="fe888-130">Sideload an add-in on Excel or Word on iPad using macOS Catalina</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fe888-131">随着 macOS Catalina 的引入，Apple 不再使用 Mac 上的 [iTunes，](https://support.apple.com/HT210200) 以及将应用旁加载到 **Finder 所需的集成功能**。</span><span class="sxs-lookup"><span data-stu-id="fe888-131">With the introduction of macOS Catalina, [Apple discontinued iTunes on Mac](https://support.apple.com/HT210200) and integrated functionality required to sideload apps into **Finder**.</span></span>

1. <span data-ttu-id="fe888-132">使用同步电缆将 iPad 连接到你的计算机。</span><span class="sxs-lookup"><span data-stu-id="fe888-132">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="fe888-133">如果是首次将 iPad连接到计算机，系统将提示"信任此计算机 **？"。**</span><span class="sxs-lookup"><span data-stu-id="fe888-133">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="fe888-134">选择“**信任**”继续执行操作。</span><span class="sxs-lookup"><span data-stu-id="fe888-134">Choose **Trust** to continue.</span></span> <span data-ttu-id="fe888-135">还可能会询问您是否是新的iPad或是否正在还原一个。</span><span class="sxs-lookup"><span data-stu-id="fe888-135">You may also be asked if this is a new iPad or if you're restoring one.</span></span>

2. <span data-ttu-id="fe888-136">在"查找器"中的"位置 **"** 下 **，选择iPad** 栏下方的"搜索"图标。</span><span class="sxs-lookup"><span data-stu-id="fe888-136">In Finder, under **Locations**, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="fe888-137">在"查找器"窗口顶部，单击"文件"，然后找到 **"Excel** **Word"。**</span><span class="sxs-lookup"><span data-stu-id="fe888-137">On the top of the Finder window, click on **Files**, and then locate **Excel** or **Word**.</span></span>

4. <span data-ttu-id="fe888-138">从不同的"查找器"窗口中，拖放要旁加载到第一个"查找器"窗口中的 Excel 或 Word文件的外接程序的 **manifest.xml** 文件。</span><span class="sxs-lookup"><span data-stu-id="fe888-138">From a different Finder window, drag and drop the manifest.xml file of the add-in you want to side load onto the **Excel** or **Word** file in the first Finder window.</span></span>

5. <span data-ttu-id="fe888-139">在你的 iPad 上打开 Excel 或 Word 应用。</span><span class="sxs-lookup"><span data-stu-id="fe888-139">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="fe888-140">如果Excel Word 应用已在运行，请选择"开始"按钮，然后关闭并重新启动应用。</span><span class="sxs-lookup"><span data-stu-id="fe888-140">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

6. <span data-ttu-id="fe888-141">打开一个文档。</span><span class="sxs-lookup"><span data-stu-id="fe888-141">Open a document.</span></span>

7. <span data-ttu-id="fe888-142">选择"插入"选项卡上的"外接程序"。 (在"插入"选项卡上，您可能需要水平滚动，直到您看到"外接程序"按钮。) 您的旁加载的外接程序可在外接程序 **UI** 中的"开发人员"标题下插入。  </span><span class="sxs-lookup"><span data-stu-id="fe888-142">Choose **Add-ins** on the **Insert** tab. (On the **Insert** tab, you may need to scroll horizontally until you see the **Add-ins** button.) Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![在加载项应用中插入Excel加载项。](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a><span data-ttu-id="fe888-144">在 Mac 版 Office 中旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="fe888-144">Sideload an add-in in Office on Mac</span></span>

> [!NOTE]
> <span data-ttu-id="fe888-145">若要旁加载 Mac 版 Outlook 加载项，请参阅[旁加载 Outlook 加载项进行测试](../outlook/sideload-outlook-add-ins-for-testing.md)。</span><span class="sxs-lookup"><span data-stu-id="fe888-145">To sideload an Outlook add-in on Mac, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="fe888-146">打开 **终端** ，然后转到以下文件夹之一，你将保存加载项的清单文件。</span><span class="sxs-lookup"><span data-stu-id="fe888-146">Open **Terminal** and go to one of the following folders where you'll save your add-in's manifest file.</span></span> <span data-ttu-id="fe888-147">如果 `wef` 文件夹在你的计算机上不存在，请创建它。</span><span class="sxs-lookup"><span data-stu-id="fe888-147">If the `wef` folder doesn't exist on your computer, create it.</span></span>

    - <span data-ttu-id="fe888-148">对于 Word：`/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="fe888-148">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span></span>
    - <span data-ttu-id="fe888-149">对于 Excel：`/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="fe888-149">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span></span>
    - <span data-ttu-id="fe888-150">对于 PowerPoint：`/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="fe888-150">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span></span>

2. <span data-ttu-id="fe888-151">使用命令命令 **打开 Finder** 中的文件夹 `open .` (包括句点或点) 。</span><span class="sxs-lookup"><span data-stu-id="fe888-151">Open the folder in **Finder** using the command `open .` (including the period or dot).</span></span> <span data-ttu-id="fe888-152">将你的外接程序的清单文件复制到该文件夹中。</span><span class="sxs-lookup"><span data-stu-id="fe888-152">Copy your add-in's manifest file to this folder.</span></span>

    ![Mac 上 Office Wef 文件夹。](../images/all-my-files.png)

3. <span data-ttu-id="fe888-p108">打开 Word，然后打开一个文档。如果 Word 已运行，则重新启动它。</span><span class="sxs-lookup"><span data-stu-id="fe888-p108">Open Word, and then open a document. Restart Word if it's already running.</span></span>

4. <span data-ttu-id="fe888-156">在 Word中，选择"插入外接程序我的  >    >   ("下拉菜单) "，然后选择您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="fe888-156">In Word, choose **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>

    ![在 Mac 上Office我的外接程序。](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="fe888-p109">旁加载的加载项不会显示在“我的加载项”对话框中。它们仅显示在下拉菜单中（单击“插入”选项卡上“我的加载项”右侧的向下小箭头）。旁加载的加载项在此菜单中的“开发人员加载项”标题下列出。</span><span class="sxs-lookup"><span data-stu-id="fe888-p109">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span>

5. <span data-ttu-id="fe888-161">验证加载项是否在 Word 中显示。</span><span class="sxs-lookup"><span data-stu-id="fe888-161">Verify that your add-in is displayed in Word.</span></span>

    ![Office在 Mac 上的 Office 中显示的外接程序。](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="fe888-163">删除旁加载的外接程序</span><span class="sxs-lookup"><span data-stu-id="fe888-163">Remove a sideloaded add-in</span></span>

<span data-ttu-id="fe888-164">可以通过清除计算机上以前旁加载的Office缓存来删除。</span><span class="sxs-lookup"><span data-stu-id="fe888-164">You can remove a previously sideloaded add-in by clearing the Office cache on your computer.</span></span> <span data-ttu-id="fe888-165">有关如何清除每个平台和应用程序缓存的详细信息，请参阅清除每个Office[缓存。](clear-cache.md)</span><span class="sxs-lookup"><span data-stu-id="fe888-165">Details on how to clear the cache for each platform and application can be found in the article [Clear the Office cache](clear-cache.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="fe888-166">另请参阅</span><span class="sxs-lookup"><span data-stu-id="fe888-166">See also</span></span>

- [<span data-ttu-id="fe888-167">在 iPad 和 Mac 上调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="fe888-167">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
