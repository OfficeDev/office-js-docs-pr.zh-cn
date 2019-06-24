---
title: 在 iPad 和 Mac 上旁加载 Office 加载项以供测试
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: b97da347064fdceda847403785e0bf792846dfaa
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35126965"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="b29c3-102">在 iPad 和 Mac 上旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="b29c3-102">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="b29c3-p101">若要查看加载项在 iOS 版 Office 中如何运行，可以使用 iTunes 将加载项的清单旁加载到 iPad，或直接将加载项的清单旁加载到 Mac 版 Office 中。此操作并不能使你在运行时对其设置断点和调试代码，但你可以查看其行为方式，并验证 UI 可用且正确呈现。</span><span class="sxs-lookup"><span data-stu-id="b29c3-p101">To see how your add-in will run in Office for iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office for Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span> 

## <a name="prerequisites-for-office-on-ios"></a><span data-ttu-id="b29c3-105">iOS 版 Office 的先决条件</span><span class="sxs-lookup"><span data-stu-id="b29c3-105">Prerequisites for Office for iOS</span></span>

- <span data-ttu-id="b29c3-106">安装了 [iTunes](https://www.apple.com/itunes/download/) 的 Windows 或 Mac 计算机。</span><span class="sxs-lookup"><span data-stu-id="b29c3-106">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>
    
- <span data-ttu-id="b29c3-107">安装了 [iPad 版 Excel](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) 的 iPad（运行 iOS 8.2 或更高版本）以及同步电缆。</span><span class="sxs-lookup"><span data-stu-id="b29c3-107">An iPad running iOS 8.2 or later with [Excel for iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) installed, and a sync cable.</span></span>
    
- <span data-ttu-id="b29c3-108">你想要测试的外接程序的清单 .xml 文件。</span><span class="sxs-lookup"><span data-stu-id="b29c3-108">The manifest .xml file for the add-in you want to test.</span></span>
    

## <a name="prerequisites-for-office-on-mac"></a><span data-ttu-id="b29c3-109">Mac 版 Office 的先决条件</span><span class="sxs-lookup"><span data-stu-id="b29c3-109">Prerequisites for Office for Mac</span></span>

- <span data-ttu-id="b29c3-110">在已安装 [Mac 版 Office](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) 的情况下可运行 OS X v10.10 "Yosemite" 或更高版本的 Mac。</span><span class="sxs-lookup"><span data-stu-id="b29c3-110">A Mac running OS X v10.10 "Yosemite" or later with [Office for Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>
    
- <span data-ttu-id="b29c3-111">Mac 版本 15.18 (160109) 上的 Word。</span><span class="sxs-lookup"><span data-stu-id="b29c3-111">Word for Mac version 15.18 (160109)</span></span>
   
- <span data-ttu-id="b29c3-112">Mac 版本 15.19 (160206) 上的 Excel。</span><span class="sxs-lookup"><span data-stu-id="b29c3-112">Excel for Mac version 15.19 (160206)</span></span>

- <span data-ttu-id="b29c3-113">Mac 版本 15.24 (160614) 上的 PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b29c3-113">PowerPoint for Mac version 15.24 (160614)</span></span>
    
- <span data-ttu-id="b29c3-114">你想要测试的外接程序的清单 .xml 文件。</span><span class="sxs-lookup"><span data-stu-id="b29c3-114">The manifest .xml file for the add-in you want to test.</span></span>
    

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad"></a><span data-ttu-id="b29c3-115">在 iPad 版 Excel 或 iPad 版 Word 上旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="b29c3-115">Sideload an add-in on Excel or Word for iPad</span></span>

1. <span data-ttu-id="b29c3-p102">使用同步电缆将 iPad 连接到你的计算机。如果是第一次将 iPad 连接到计算机，系统将提示“**信任此计算机？**”。选择“**信任**”继续执行操作。</span><span class="sxs-lookup"><span data-stu-id="b29c3-p102">Use a sync cable to connect your iPad to your computer. If you're connecting the iPad to your computer for the first time, you'll be prompted with  **Trust This Computer?**. Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="b29c3-119">在 iTunes 中，选择菜单栏下的“**iPad**”图标。</span><span class="sxs-lookup"><span data-stu-id="b29c3-119">In iTunes, choose the  **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="b29c3-120">在 iTunes 左侧的"设置"下，选择"应用程序"。</span><span class="sxs-lookup"><span data-stu-id="b29c3-120">Under  **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="b29c3-121">在 iTunes 右侧，向下滚动到"文件共享"，然后在"外接程序"列下选择"Excel"或"Word"。</span><span class="sxs-lookup"><span data-stu-id="b29c3-121">On the right side of iTunes, scroll down to  **File Sharing**, and then choose  **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="b29c3-122">在"Excel"或"Word 文档"列底部，选择"添加文件"，然后选择您要旁加载的外接程序的清单 .xml 文件。</span><span class="sxs-lookup"><span data-stu-id="b29c3-122">At the bottom of the  **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span> 
    
6. <span data-ttu-id="b29c3-p103">在你的 iPad 上打开 Excel 或 Word 应用。如果 Excel 或 Word 应用已运行，则选择“**首页**”按钮，然后关闭并重新启动该应用。</span><span class="sxs-lookup"><span data-stu-id="b29c3-p103">Open the Excel or Word app on your iPad. If the Excel or Word app is already running, choose the  **Home** button, and then close and restart the app.</span></span>
    
7. <span data-ttu-id="b29c3-125">打开一个文档。</span><span class="sxs-lookup"><span data-stu-id="b29c3-125">Open a document.</span></span>
    
8. <span data-ttu-id="b29c3-126">选择“**插入**”选项卡上的“**外接程序**”。旁加载的外接程序可在“**外接程序**”UI 中的“**开发人员**”标题下插入。</span><span class="sxs-lookup"><span data-stu-id="b29c3-126">Choose  **Add-ins** on the **Insert** tab. Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>
    
    ![在 Excel 应用程序中插入的加载项](../images/excel-insert-add-in.png)


## <a name="sideload-an-add-in-in-office-on-mac"></a><span data-ttu-id="b29c3-128">在 Mac 版 Office 中旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="b29c3-128">Sideload an add-in on Office for Mac</span></span>

> [!NOTE]
> <span data-ttu-id="b29c3-129">若要旁加载 Mac 版 Outlook 加载项，请参阅[旁加载 Outlook 加载项进行测试](/outlook/add-ins/sideload-outlook-add-ins-for-testing)。</span><span class="sxs-lookup"><span data-stu-id="b29c3-129">To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

1. <span data-ttu-id="b29c3-p104">打开“**终端**”并转到以下文件夹之一，你将在其中保存外接程序的清单文件。如果 `wef` 文件夹在你的计算机上不存在，请创建它。</span><span class="sxs-lookup"><span data-stu-id="b29c3-p104">Open  **Terminal** and go to one of the following folders where you'll save your add-in's manifest file. If the `wef` folder doesn't exist on your computer, create it.</span></span>
    
    - <span data-ttu-id="b29c3-132">对于 Word：`/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="b29c3-132">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`</span></span>    
    - <span data-ttu-id="b29c3-133">对于 Excel：`/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="b29c3-133">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`</span></span>
    - <span data-ttu-id="b29c3-134">对于 PowerPoint：`/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="b29c3-134">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`</span></span>
    
2. <span data-ttu-id="b29c3-p105">在“**查找程序**”中使用命令 `open .`（包括句点或点）打开该文件夹。将你的外接程序的清单文件复制到该文件夹中。</span><span class="sxs-lookup"><span data-stu-id="b29c3-p105">Open the folder in  **Finder** using the command `open .` (including the period or dot). Copy your add-in's manifest file to this folder.</span></span>
    
    ![Mac 版 Office 中的 Wef 文件夹](../images/all-my-files.png)

3. <span data-ttu-id="b29c3-p106">打开 Word，然后打开一个文档。如果 Word 已运行，则重新启动它。</span><span class="sxs-lookup"><span data-stu-id="b29c3-p106">Open Word, and then open a document. Restart Word if it's already running.</span></span>
    
4. <span data-ttu-id="b29c3-140">在 Word 中，选择“**插入**” > “**外接程序**” > “**我的外接程序**”（下拉菜单），然后选择外接程序。</span><span class="sxs-lookup"><span data-stu-id="b29c3-140">In Word, choose  **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>
    
    ![Mac 版 Office 中的“我的加载项”](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="b29c3-p107">旁加载的加载项不会显示在“我的加载项”对话框中。它们仅显示在下拉菜单中（单击“插入”\*\*\*\* 选项卡上“我的加载项”右侧的向下小箭头）。旁加载的加载项在此菜单中的“开发人员加载项”\*\*\*\* 标题下列出。</span><span class="sxs-lookup"><span data-stu-id="b29c3-p107">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span> 
    
5. <span data-ttu-id="b29c3-145">验证加载项是否在 Word 中显示。</span><span class="sxs-lookup"><span data-stu-id="b29c3-145">Verify that your add-in is displayed in Word.</span></span>
    
    ![Mac 版 Office 中显示的 Office 加载项](../images/lorem-ipsum-wikipedia.png)
    
### <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="b29c3-147">在 Mac 上清除 Office 应用程序缓存</span><span class="sxs-lookup"><span data-stu-id="b29c3-147">Clearing the Office application's cache on a Mac or iPad</span></span>

<span data-ttu-id="b29c3-148">出于性能方面的考虑，加载项通常在 Mac 版 Office 中缓存。</span><span class="sxs-lookup"><span data-stu-id="b29c3-148">Add-ins are cached often in Office for Mac, for performance reasons.</span></span> <span data-ttu-id="b29c3-149">通常情况下，将通过重载外接程序清除缓存。</span><span class="sxs-lookup"><span data-stu-id="b29c3-149">Normally, the cache is cleared by reloading the add-in.</span></span> <span data-ttu-id="b29c3-150">如果同一文档中存在多个外接程序，则重载后自动清除缓存的过程可能不可靠。</span><span class="sxs-lookup"><span data-stu-id="b29c3-150">If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span>

<span data-ttu-id="b29c3-151">在 Mac 上，通过删除 `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` 文件夹中的内容可以手动清除缓存。</span><span class="sxs-lookup"><span data-stu-id="b29c3-151">On a Mac, you can clear the cache manually by deleting everything in the `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span> 

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="see-also"></a><span data-ttu-id="b29c3-152">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b29c3-152">See also</span></span>

- [<span data-ttu-id="b29c3-153">在 iPad 和 Mac 上调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="b29c3-153">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
