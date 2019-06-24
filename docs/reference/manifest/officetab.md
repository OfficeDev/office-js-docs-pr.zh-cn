---
title: 清单文件中的 OfficeTab 元素
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: d073d712cec2fd58e957ffe8f344d7443d1e896e
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127560"
---
# <a name="officetab-element"></a><span data-ttu-id="5d4f5-102">OfficeTab 元素</span><span class="sxs-lookup"><span data-stu-id="5d4f5-102">OfficeTab element</span></span>

<span data-ttu-id="5d4f5-p101">定义在其上显示外接程序命令的功能区选项卡。这可以是默认的选项卡（“**主页**”、“**消息**”或“**会议**”），或是由外接程序定义的自定义选项卡。此元素是必需的。</span><span class="sxs-lookup"><span data-stu-id="5d4f5-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="5d4f5-106">子元素</span><span class="sxs-lookup"><span data-stu-id="5d4f5-106">Child elements</span></span>

|  <span data-ttu-id="5d4f5-107">元素</span><span class="sxs-lookup"><span data-stu-id="5d4f5-107">Element</span></span> |  <span data-ttu-id="5d4f5-108">必需</span><span class="sxs-lookup"><span data-stu-id="5d4f5-108">Required</span></span>  |  <span data-ttu-id="5d4f5-109">说明</span><span class="sxs-lookup"><span data-stu-id="5d4f5-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="5d4f5-110">组</span><span class="sxs-lookup"><span data-stu-id="5d4f5-110">Group</span></span>      | <span data-ttu-id="5d4f5-111">是</span><span class="sxs-lookup"><span data-stu-id="5d4f5-111">Yes</span></span> |  <span data-ttu-id="5d4f5-p102">定义一组命令。对于每个外接程序，只能将一个组添加到默认选项卡。</span><span class="sxs-lookup"><span data-stu-id="5d4f5-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="5d4f5-114">下面是主机的有效选项卡 `id` 值。</span><span class="sxs-lookup"><span data-stu-id="5d4f5-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="5d4f5-115">以**粗体显示**的值在桌面和联机状态中均受支持 (例如, Windows 和 web 上的 word 中的 word 2016 或更高版本)。</span><span class="sxs-lookup"><span data-stu-id="5d4f5-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="5d4f5-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="5d4f5-116">Outlook</span></span>

- <span data-ttu-id="5d4f5-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="5d4f5-118">Word</span><span class="sxs-lookup"><span data-stu-id="5d4f5-118">Word</span></span>

- <span data-ttu-id="5d4f5-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-119">**TabHome**</span></span>
- <span data-ttu-id="5d4f5-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-120">**TabInsert**</span></span>
- <span data-ttu-id="5d4f5-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="5d4f5-121">TabWordDesign</span></span>
- <span data-ttu-id="5d4f5-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="5d4f5-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="5d4f5-123">TabReferences</span></span>
- <span data-ttu-id="5d4f5-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="5d4f5-124">TabMailings</span></span>
- <span data-ttu-id="5d4f5-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="5d4f5-125">TabReviewWord</span></span>
- <span data-ttu-id="5d4f5-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-126">**TabView**</span></span>
- <span data-ttu-id="5d4f5-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="5d4f5-127">TabDeveloper</span></span>
- <span data-ttu-id="5d4f5-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="5d4f5-128">TabAddIns</span></span>
- <span data-ttu-id="5d4f5-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="5d4f5-129">TabBlogPost</span></span>
- <span data-ttu-id="5d4f5-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="5d4f5-130">TabBlogInsert</span></span>
- <span data-ttu-id="5d4f5-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="5d4f5-131">TabPrintPreview</span></span>
- <span data-ttu-id="5d4f5-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="5d4f5-132">TabOutlining</span></span>
- <span data-ttu-id="5d4f5-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="5d4f5-133">TabConflicts</span></span>
- <span data-ttu-id="5d4f5-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="5d4f5-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="5d4f5-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="5d4f5-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="5d4f5-136">Excel</span><span class="sxs-lookup"><span data-stu-id="5d4f5-136">Excel</span></span>

- <span data-ttu-id="5d4f5-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-137">**TabHome**</span></span>
- <span data-ttu-id="5d4f5-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-138">**TabInsert**</span></span>
- <span data-ttu-id="5d4f5-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="5d4f5-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="5d4f5-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="5d4f5-140">TabFormulas</span></span>
- <span data-ttu-id="5d4f5-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-141">**TabData**</span></span>
- <span data-ttu-id="5d4f5-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-142">**TabReview**</span></span>
- <span data-ttu-id="5d4f5-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-143">**TabView**</span></span>
- <span data-ttu-id="5d4f5-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="5d4f5-144">TabDeveloper</span></span>
- <span data-ttu-id="5d4f5-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="5d4f5-145">TabAddIns</span></span>
- <span data-ttu-id="5d4f5-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="5d4f5-146">TabPrintPreview</span></span>
- <span data-ttu-id="5d4f5-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="5d4f5-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="5d4f5-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5d4f5-148">PowerPoint</span></span>

- <span data-ttu-id="5d4f5-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-149">**TabHome**</span></span>
- <span data-ttu-id="5d4f5-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-150">**TabInsert**</span></span>
- <span data-ttu-id="5d4f5-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-151">**TabDesign**</span></span>
- <span data-ttu-id="5d4f5-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-152">**TabTransitions**</span></span>
- <span data-ttu-id="5d4f5-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-153">**TabAnimations**</span></span>
- <span data-ttu-id="5d4f5-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="5d4f5-154">TabSlideShow</span></span>
- <span data-ttu-id="5d4f5-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="5d4f5-155">TabReview</span></span>
- <span data-ttu-id="5d4f5-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-156">**TabView**</span></span>
- <span data-ttu-id="5d4f5-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="5d4f5-157">TabDeveloper</span></span>
- <span data-ttu-id="5d4f5-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="5d4f5-158">TabAddIns</span></span>
- <span data-ttu-id="5d4f5-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="5d4f5-159">TabPrintPreview</span></span>
- <span data-ttu-id="5d4f5-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="5d4f5-160">TabMerge</span></span>
- <span data-ttu-id="5d4f5-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="5d4f5-161">TabGrayscale</span></span>
- <span data-ttu-id="5d4f5-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="5d4f5-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="5d4f5-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="5d4f5-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="5d4f5-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="5d4f5-164">TabSlideMaster</span></span>
- <span data-ttu-id="5d4f5-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="5d4f5-165">TabHandoutMaster</span></span>
- <span data-ttu-id="5d4f5-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="5d4f5-166">TabNotesMaster</span></span>
- <span data-ttu-id="5d4f5-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="5d4f5-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="5d4f5-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="5d4f5-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="5d4f5-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="5d4f5-169">OneNote</span></span>

- <span data-ttu-id="5d4f5-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-170">**TabHome**</span></span>
- <span data-ttu-id="5d4f5-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-171">**TabInsert**</span></span>
- <span data-ttu-id="5d4f5-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="5d4f5-172">**TabView**</span></span>
- <span data-ttu-id="5d4f5-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="5d4f5-173">TabDeveloper</span></span>
- <span data-ttu-id="5d4f5-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="5d4f5-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="5d4f5-175">组</span><span class="sxs-lookup"><span data-stu-id="5d4f5-175">Group</span></span>

<span data-ttu-id="5d4f5-p104">选项卡中的一组 UI 扩展点。一组可以有多达六个控件。需要 **id** 属性且每个 **id** 在清单内必须是唯一的。**id** 是一个最大长度为 125 个字符的字符串。查看 [Group 元素](group.md)。</span><span class="sxs-lookup"><span data-stu-id="5d4f5-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="5d4f5-180">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="5d4f5-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
