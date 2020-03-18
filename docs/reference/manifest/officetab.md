---
title: 清单文件中的 OfficeTab 元素
description: OfficeTab 元素定义在其中显示外接程序命令的功能区选项卡。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 1d1810f3d3a206f72bf9544814a3fdaaa556476e
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720489"
---
# <a name="officetab-element"></a><span data-ttu-id="b71ec-103">OfficeTab 元素</span><span class="sxs-lookup"><span data-stu-id="b71ec-103">OfficeTab element</span></span>

<span data-ttu-id="b71ec-104">定义在其上显示外接程序命令的功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="b71ec-104">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="b71ec-105">这可以是默认的选项卡（"**主页**"、"**消息**" 或 "**会议**"），也可以是由外接程序定义的自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="b71ec-105">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="b71ec-106">此元素是必需的。</span><span class="sxs-lookup"><span data-stu-id="b71ec-106">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b71ec-107">子元素</span><span class="sxs-lookup"><span data-stu-id="b71ec-107">Child elements</span></span>

|  <span data-ttu-id="b71ec-108">元素</span><span class="sxs-lookup"><span data-stu-id="b71ec-108">Element</span></span> |  <span data-ttu-id="b71ec-109">必需</span><span class="sxs-lookup"><span data-stu-id="b71ec-109">Required</span></span>  |  <span data-ttu-id="b71ec-110">说明</span><span class="sxs-lookup"><span data-stu-id="b71ec-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b71ec-111">组</span><span class="sxs-lookup"><span data-stu-id="b71ec-111">Group</span></span>      | <span data-ttu-id="b71ec-112">是</span><span class="sxs-lookup"><span data-stu-id="b71ec-112">Yes</span></span> |  <span data-ttu-id="b71ec-p102">定义一组命令。对于每个外接程序，只能将一个组添加到默认选项卡。</span><span class="sxs-lookup"><span data-stu-id="b71ec-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="b71ec-115">下面是主机的有效选项卡 `id` 值。</span><span class="sxs-lookup"><span data-stu-id="b71ec-115">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="b71ec-116">以**粗体显示**的值在桌面和联机状态中均受支持（例如，Windows 和 web 上的 word 中的 word 2016 或更高版本）。</span><span class="sxs-lookup"><span data-stu-id="b71ec-116">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="b71ec-117">Outlook</span><span class="sxs-lookup"><span data-stu-id="b71ec-117">Outlook</span></span>

- <span data-ttu-id="b71ec-118">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="b71ec-118">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="b71ec-119">Word</span><span class="sxs-lookup"><span data-stu-id="b71ec-119">Word</span></span>

- <span data-ttu-id="b71ec-120">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="b71ec-120">**TabHome**</span></span>
- <span data-ttu-id="b71ec-121">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="b71ec-121">**TabInsert**</span></span>
- <span data-ttu-id="b71ec-122">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="b71ec-122">TabWordDesign</span></span>
- <span data-ttu-id="b71ec-123">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="b71ec-123">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="b71ec-124">TabReferences</span><span class="sxs-lookup"><span data-stu-id="b71ec-124">TabReferences</span></span>
- <span data-ttu-id="b71ec-125">TabMailings</span><span class="sxs-lookup"><span data-stu-id="b71ec-125">TabMailings</span></span>
- <span data-ttu-id="b71ec-126">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="b71ec-126">TabReviewWord</span></span>
- <span data-ttu-id="b71ec-127">**TabView**</span><span class="sxs-lookup"><span data-stu-id="b71ec-127">**TabView**</span></span>
- <span data-ttu-id="b71ec-128">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="b71ec-128">TabDeveloper</span></span>
- <span data-ttu-id="b71ec-129">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="b71ec-129">TabAddIns</span></span>
- <span data-ttu-id="b71ec-130">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="b71ec-130">TabBlogPost</span></span>
- <span data-ttu-id="b71ec-131">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="b71ec-131">TabBlogInsert</span></span>
- <span data-ttu-id="b71ec-132">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="b71ec-132">TabPrintPreview</span></span>
- <span data-ttu-id="b71ec-133">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="b71ec-133">TabOutlining</span></span>
- <span data-ttu-id="b71ec-134">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="b71ec-134">TabConflicts</span></span>
- <span data-ttu-id="b71ec-135">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="b71ec-135">TabBackgroundRemoval</span></span>
- <span data-ttu-id="b71ec-136">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="b71ec-136">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="b71ec-137">Excel</span><span class="sxs-lookup"><span data-stu-id="b71ec-137">Excel</span></span>

- <span data-ttu-id="b71ec-138">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="b71ec-138">**TabHome**</span></span>
- <span data-ttu-id="b71ec-139">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="b71ec-139">**TabInsert**</span></span>
- <span data-ttu-id="b71ec-140">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="b71ec-140">TabPageLayoutExcel</span></span>
- <span data-ttu-id="b71ec-141">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="b71ec-141">TabFormulas</span></span>
- <span data-ttu-id="b71ec-142">**TabData**</span><span class="sxs-lookup"><span data-stu-id="b71ec-142">**TabData**</span></span>
- <span data-ttu-id="b71ec-143">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="b71ec-143">**TabReview**</span></span>
- <span data-ttu-id="b71ec-144">**TabView**</span><span class="sxs-lookup"><span data-stu-id="b71ec-144">**TabView**</span></span>
- <span data-ttu-id="b71ec-145">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="b71ec-145">TabDeveloper</span></span>
- <span data-ttu-id="b71ec-146">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="b71ec-146">TabAddIns</span></span>
- <span data-ttu-id="b71ec-147">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="b71ec-147">TabPrintPreview</span></span>
- <span data-ttu-id="b71ec-148">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="b71ec-148">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="b71ec-149">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b71ec-149">PowerPoint</span></span>

- <span data-ttu-id="b71ec-150">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="b71ec-150">**TabHome**</span></span>
- <span data-ttu-id="b71ec-151">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="b71ec-151">**TabInsert**</span></span>
- <span data-ttu-id="b71ec-152">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="b71ec-152">**TabDesign**</span></span>
- <span data-ttu-id="b71ec-153">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="b71ec-153">**TabTransitions**</span></span>
- <span data-ttu-id="b71ec-154">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="b71ec-154">**TabAnimations**</span></span>
- <span data-ttu-id="b71ec-155">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="b71ec-155">TabSlideShow</span></span>
- <span data-ttu-id="b71ec-156">TabReview</span><span class="sxs-lookup"><span data-stu-id="b71ec-156">TabReview</span></span>
- <span data-ttu-id="b71ec-157">**TabView**</span><span class="sxs-lookup"><span data-stu-id="b71ec-157">**TabView**</span></span>
- <span data-ttu-id="b71ec-158">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="b71ec-158">TabDeveloper</span></span>
- <span data-ttu-id="b71ec-159">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="b71ec-159">TabAddIns</span></span>
- <span data-ttu-id="b71ec-160">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="b71ec-160">TabPrintPreview</span></span>
- <span data-ttu-id="b71ec-161">TabMerge</span><span class="sxs-lookup"><span data-stu-id="b71ec-161">TabMerge</span></span>
- <span data-ttu-id="b71ec-162">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="b71ec-162">TabGrayscale</span></span>
- <span data-ttu-id="b71ec-163">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="b71ec-163">TabBlackAndWhite</span></span>
- <span data-ttu-id="b71ec-164">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="b71ec-164">TabBroadcastPresentation</span></span>
- <span data-ttu-id="b71ec-165">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="b71ec-165">TabSlideMaster</span></span>
- <span data-ttu-id="b71ec-166">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="b71ec-166">TabHandoutMaster</span></span>
- <span data-ttu-id="b71ec-167">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="b71ec-167">TabNotesMaster</span></span>
- <span data-ttu-id="b71ec-168">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="b71ec-168">TabBackgroundRemoval</span></span>
- <span data-ttu-id="b71ec-169">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="b71ec-169">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="b71ec-170">OneNote</span><span class="sxs-lookup"><span data-stu-id="b71ec-170">OneNote</span></span>

- <span data-ttu-id="b71ec-171">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="b71ec-171">**TabHome**</span></span>
- <span data-ttu-id="b71ec-172">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="b71ec-172">**TabInsert**</span></span>
- <span data-ttu-id="b71ec-173">**TabView**</span><span class="sxs-lookup"><span data-stu-id="b71ec-173">**TabView**</span></span>
- <span data-ttu-id="b71ec-174">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="b71ec-174">TabDeveloper</span></span>
- <span data-ttu-id="b71ec-175">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="b71ec-175">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="b71ec-176">组</span><span class="sxs-lookup"><span data-stu-id="b71ec-176">Group</span></span>

<span data-ttu-id="b71ec-177">选项卡中的一组 UI 扩展点。</span><span class="sxs-lookup"><span data-stu-id="b71ec-177">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="b71ec-178">**Id**属性是必需的，并且每个**id**在清单中必须是唯一的。</span><span class="sxs-lookup"><span data-stu-id="b71ec-178">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="b71ec-179">**Id**是最多为125个字符的字符串。</span><span class="sxs-lookup"><span data-stu-id="b71ec-179">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="b71ec-180">查看 [Group 元素](group.md)。</span><span class="sxs-lookup"><span data-stu-id="b71ec-180">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="b71ec-181">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="b71ec-181">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
