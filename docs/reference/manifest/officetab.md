---
title: 清单文件中的 OfficeTab 元素
description: OfficeTab 元素定义在其中显示外接程序命令的功能区选项卡。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 25e8044d8b3264bf9ee64c54487566bf11f0065e
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292298"
---
# <a name="officetab-element"></a><span data-ttu-id="9b351-103">OfficeTab 元素</span><span class="sxs-lookup"><span data-stu-id="9b351-103">OfficeTab element</span></span>

<span data-ttu-id="9b351-104">定义在其上显示外接程序命令的功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="9b351-104">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="9b351-105">这可以是默认选项卡 (" **主页**"、" **邮件**" 或 " **会议**) "，也可以是由加载项定义的自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="9b351-105">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="9b351-106">此元素是必需的。</span><span class="sxs-lookup"><span data-stu-id="9b351-106">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="9b351-107">子元素</span><span class="sxs-lookup"><span data-stu-id="9b351-107">Child elements</span></span>

|  <span data-ttu-id="9b351-108">元素</span><span class="sxs-lookup"><span data-stu-id="9b351-108">Element</span></span> |  <span data-ttu-id="9b351-109">必需</span><span class="sxs-lookup"><span data-stu-id="9b351-109">Required</span></span>  |  <span data-ttu-id="9b351-110">说明</span><span class="sxs-lookup"><span data-stu-id="9b351-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="9b351-111">组</span><span class="sxs-lookup"><span data-stu-id="9b351-111">Group</span></span>      | <span data-ttu-id="9b351-112">是</span><span class="sxs-lookup"><span data-stu-id="9b351-112">Yes</span></span> |  <span data-ttu-id="9b351-p102">定义一组命令。对于每个外接程序，只能将一个组添加到默认选项卡。</span><span class="sxs-lookup"><span data-stu-id="9b351-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="9b351-115">以下是按应用程序的有效选项卡 `id` 值。</span><span class="sxs-lookup"><span data-stu-id="9b351-115">The following are valid tab `id` values by application.</span></span> <span data-ttu-id="9b351-116">以 **粗体显示** 的值在桌面和联机 (中均受支持（例如，word 2016 或更高版本位于 web 上的 Windows 和 word) 。</span><span class="sxs-lookup"><span data-stu-id="9b351-116">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="9b351-117">Outlook</span><span class="sxs-lookup"><span data-stu-id="9b351-117">Outlook</span></span>

- <span data-ttu-id="9b351-118">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="9b351-118">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="9b351-119">Word</span><span class="sxs-lookup"><span data-stu-id="9b351-119">Word</span></span>

- <span data-ttu-id="9b351-120">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="9b351-120">**TabHome**</span></span>
- <span data-ttu-id="9b351-121">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="9b351-121">**TabInsert**</span></span>
- <span data-ttu-id="9b351-122">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="9b351-122">TabWordDesign</span></span>
- <span data-ttu-id="9b351-123">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="9b351-123">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="9b351-124">TabReferences</span><span class="sxs-lookup"><span data-stu-id="9b351-124">TabReferences</span></span>
- <span data-ttu-id="9b351-125">TabMailings</span><span class="sxs-lookup"><span data-stu-id="9b351-125">TabMailings</span></span>
- <span data-ttu-id="9b351-126">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="9b351-126">TabReviewWord</span></span>
- <span data-ttu-id="9b351-127">**TabView**</span><span class="sxs-lookup"><span data-stu-id="9b351-127">**TabView**</span></span>
- <span data-ttu-id="9b351-128">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="9b351-128">TabDeveloper</span></span>
- <span data-ttu-id="9b351-129">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="9b351-129">TabAddIns</span></span>
- <span data-ttu-id="9b351-130">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="9b351-130">TabBlogPost</span></span>
- <span data-ttu-id="9b351-131">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="9b351-131">TabBlogInsert</span></span>
- <span data-ttu-id="9b351-132">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="9b351-132">TabPrintPreview</span></span>
- <span data-ttu-id="9b351-133">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="9b351-133">TabOutlining</span></span>
- <span data-ttu-id="9b351-134">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="9b351-134">TabConflicts</span></span>
- <span data-ttu-id="9b351-135">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="9b351-135">TabBackgroundRemoval</span></span>
- <span data-ttu-id="9b351-136">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="9b351-136">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="9b351-137">Excel</span><span class="sxs-lookup"><span data-stu-id="9b351-137">Excel</span></span>

- <span data-ttu-id="9b351-138">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="9b351-138">**TabHome**</span></span>
- <span data-ttu-id="9b351-139">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="9b351-139">**TabInsert**</span></span>
- <span data-ttu-id="9b351-140">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="9b351-140">TabPageLayoutExcel</span></span>
- <span data-ttu-id="9b351-141">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="9b351-141">TabFormulas</span></span>
- <span data-ttu-id="9b351-142">**TabData**</span><span class="sxs-lookup"><span data-stu-id="9b351-142">**TabData**</span></span>
- <span data-ttu-id="9b351-143">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="9b351-143">**TabReview**</span></span>
- <span data-ttu-id="9b351-144">**TabView**</span><span class="sxs-lookup"><span data-stu-id="9b351-144">**TabView**</span></span>
- <span data-ttu-id="9b351-145">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="9b351-145">TabDeveloper</span></span>
- <span data-ttu-id="9b351-146">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="9b351-146">TabAddIns</span></span>
- <span data-ttu-id="9b351-147">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="9b351-147">TabPrintPreview</span></span>
- <span data-ttu-id="9b351-148">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="9b351-148">TabBackgroundRemoval</span></span>

### <a name="powerpoint"></a><span data-ttu-id="9b351-149">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="9b351-149">PowerPoint</span></span>

- <span data-ttu-id="9b351-150">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="9b351-150">**TabHome**</span></span>
- <span data-ttu-id="9b351-151">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="9b351-151">**TabInsert**</span></span>
- <span data-ttu-id="9b351-152">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="9b351-152">**TabDesign**</span></span>
- <span data-ttu-id="9b351-153">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="9b351-153">**TabTransitions**</span></span>
- <span data-ttu-id="9b351-154">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="9b351-154">**TabAnimations**</span></span>
- <span data-ttu-id="9b351-155">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="9b351-155">TabSlideShow</span></span>
- <span data-ttu-id="9b351-156">TabReview</span><span class="sxs-lookup"><span data-stu-id="9b351-156">TabReview</span></span>
- <span data-ttu-id="9b351-157">**TabView**</span><span class="sxs-lookup"><span data-stu-id="9b351-157">**TabView**</span></span>
- <span data-ttu-id="9b351-158">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="9b351-158">TabDeveloper</span></span>
- <span data-ttu-id="9b351-159">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="9b351-159">TabAddIns</span></span>
- <span data-ttu-id="9b351-160">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="9b351-160">TabPrintPreview</span></span>
- <span data-ttu-id="9b351-161">TabMerge</span><span class="sxs-lookup"><span data-stu-id="9b351-161">TabMerge</span></span>
- <span data-ttu-id="9b351-162">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="9b351-162">TabGrayscale</span></span>
- <span data-ttu-id="9b351-163">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="9b351-163">TabBlackAndWhite</span></span>
- <span data-ttu-id="9b351-164">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="9b351-164">TabBroadcastPresentation</span></span>
- <span data-ttu-id="9b351-165">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="9b351-165">TabSlideMaster</span></span>
- <span data-ttu-id="9b351-166">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="9b351-166">TabHandoutMaster</span></span>
- <span data-ttu-id="9b351-167">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="9b351-167">TabNotesMaster</span></span>
- <span data-ttu-id="9b351-168">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="9b351-168">TabBackgroundRemoval</span></span>
- <span data-ttu-id="9b351-169">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="9b351-169">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="9b351-170">OneNote</span><span class="sxs-lookup"><span data-stu-id="9b351-170">OneNote</span></span>

- <span data-ttu-id="9b351-171">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="9b351-171">**TabHome**</span></span>
- <span data-ttu-id="9b351-172">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="9b351-172">**TabInsert**</span></span>
- <span data-ttu-id="9b351-173">**TabView**</span><span class="sxs-lookup"><span data-stu-id="9b351-173">**TabView**</span></span>
- <span data-ttu-id="9b351-174">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="9b351-174">TabDeveloper</span></span>
- <span data-ttu-id="9b351-175">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="9b351-175">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="9b351-176">组</span><span class="sxs-lookup"><span data-stu-id="9b351-176">Group</span></span>

<span data-ttu-id="9b351-177">选项卡中的一组 UI 扩展点。</span><span class="sxs-lookup"><span data-stu-id="9b351-177">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="9b351-178">**Id**属性是必需的，并且每个**id**在清单中必须是唯一的。</span><span class="sxs-lookup"><span data-stu-id="9b351-178">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="9b351-179">**Id**是最多为125个字符的字符串。</span><span class="sxs-lookup"><span data-stu-id="9b351-179">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="9b351-180">查看 [Group 元素](group.md)。</span><span class="sxs-lookup"><span data-stu-id="9b351-180">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="9b351-181">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="9b351-181">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
