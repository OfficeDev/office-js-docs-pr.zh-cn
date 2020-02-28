---
title: 清单文件中的 OfficeTab 元素
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: b8458233ba93e98fe0bd8d51f5734b1fece65864
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324832"
---
# <a name="officetab-element"></a><span data-ttu-id="6fbf4-102">OfficeTab 元素</span><span class="sxs-lookup"><span data-stu-id="6fbf4-102">OfficeTab element</span></span>

<span data-ttu-id="6fbf4-103">定义在其上显示外接程序命令的功能区选项卡。</span><span class="sxs-lookup"><span data-stu-id="6fbf4-103">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="6fbf4-104">这可以是默认的选项卡（"**主页**"、"**消息**" 或 "**会议**"），也可以是由外接程序定义的自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="6fbf4-104">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="6fbf4-105">此元素是必需的。</span><span class="sxs-lookup"><span data-stu-id="6fbf4-105">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="6fbf4-106">子元素</span><span class="sxs-lookup"><span data-stu-id="6fbf4-106">Child elements</span></span>

|  <span data-ttu-id="6fbf4-107">元素</span><span class="sxs-lookup"><span data-stu-id="6fbf4-107">Element</span></span> |  <span data-ttu-id="6fbf4-108">必需</span><span class="sxs-lookup"><span data-stu-id="6fbf4-108">Required</span></span>  |  <span data-ttu-id="6fbf4-109">说明</span><span class="sxs-lookup"><span data-stu-id="6fbf4-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="6fbf4-110">组</span><span class="sxs-lookup"><span data-stu-id="6fbf4-110">Group</span></span>      | <span data-ttu-id="6fbf4-111">是</span><span class="sxs-lookup"><span data-stu-id="6fbf4-111">Yes</span></span> |  <span data-ttu-id="6fbf4-p102">定义一组命令。对于每个外接程序，只能将一个组添加到默认选项卡。</span><span class="sxs-lookup"><span data-stu-id="6fbf4-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="6fbf4-114">下面是主机的有效选项卡 `id` 值。</span><span class="sxs-lookup"><span data-stu-id="6fbf4-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="6fbf4-115">以**粗体显示**的值在桌面和联机状态中均受支持（例如，Windows 和 web 上的 word 中的 word 2016 或更高版本）。</span><span class="sxs-lookup"><span data-stu-id="6fbf4-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="6fbf4-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="6fbf4-116">Outlook</span></span>

- <span data-ttu-id="6fbf4-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="6fbf4-118">Word</span><span class="sxs-lookup"><span data-stu-id="6fbf4-118">Word</span></span>

- <span data-ttu-id="6fbf4-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-119">**TabHome**</span></span>
- <span data-ttu-id="6fbf4-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-120">**TabInsert**</span></span>
- <span data-ttu-id="6fbf4-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="6fbf4-121">TabWordDesign</span></span>
- <span data-ttu-id="6fbf4-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="6fbf4-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="6fbf4-123">TabReferences</span></span>
- <span data-ttu-id="6fbf4-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="6fbf4-124">TabMailings</span></span>
- <span data-ttu-id="6fbf4-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="6fbf4-125">TabReviewWord</span></span>
- <span data-ttu-id="6fbf4-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-126">**TabView**</span></span>
- <span data-ttu-id="6fbf4-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="6fbf4-127">TabDeveloper</span></span>
- <span data-ttu-id="6fbf4-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="6fbf4-128">TabAddIns</span></span>
- <span data-ttu-id="6fbf4-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="6fbf4-129">TabBlogPost</span></span>
- <span data-ttu-id="6fbf4-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="6fbf4-130">TabBlogInsert</span></span>
- <span data-ttu-id="6fbf4-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="6fbf4-131">TabPrintPreview</span></span>
- <span data-ttu-id="6fbf4-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="6fbf4-132">TabOutlining</span></span>
- <span data-ttu-id="6fbf4-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="6fbf4-133">TabConflicts</span></span>
- <span data-ttu-id="6fbf4-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="6fbf4-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="6fbf4-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="6fbf4-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="6fbf4-136">Excel</span><span class="sxs-lookup"><span data-stu-id="6fbf4-136">Excel</span></span>

- <span data-ttu-id="6fbf4-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-137">**TabHome**</span></span>
- <span data-ttu-id="6fbf4-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-138">**TabInsert**</span></span>
- <span data-ttu-id="6fbf4-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="6fbf4-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="6fbf4-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="6fbf4-140">TabFormulas</span></span>
- <span data-ttu-id="6fbf4-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-141">**TabData**</span></span>
- <span data-ttu-id="6fbf4-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-142">**TabReview**</span></span>
- <span data-ttu-id="6fbf4-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-143">**TabView**</span></span>
- <span data-ttu-id="6fbf4-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="6fbf4-144">TabDeveloper</span></span>
- <span data-ttu-id="6fbf4-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="6fbf4-145">TabAddIns</span></span>
- <span data-ttu-id="6fbf4-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="6fbf4-146">TabPrintPreview</span></span>
- <span data-ttu-id="6fbf4-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="6fbf4-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="6fbf4-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="6fbf4-148">PowerPoint</span></span>

- <span data-ttu-id="6fbf4-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-149">**TabHome**</span></span>
- <span data-ttu-id="6fbf4-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-150">**TabInsert**</span></span>
- <span data-ttu-id="6fbf4-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-151">**TabDesign**</span></span>
- <span data-ttu-id="6fbf4-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-152">**TabTransitions**</span></span>
- <span data-ttu-id="6fbf4-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-153">**TabAnimations**</span></span>
- <span data-ttu-id="6fbf4-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="6fbf4-154">TabSlideShow</span></span>
- <span data-ttu-id="6fbf4-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="6fbf4-155">TabReview</span></span>
- <span data-ttu-id="6fbf4-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-156">**TabView**</span></span>
- <span data-ttu-id="6fbf4-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="6fbf4-157">TabDeveloper</span></span>
- <span data-ttu-id="6fbf4-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="6fbf4-158">TabAddIns</span></span>
- <span data-ttu-id="6fbf4-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="6fbf4-159">TabPrintPreview</span></span>
- <span data-ttu-id="6fbf4-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="6fbf4-160">TabMerge</span></span>
- <span data-ttu-id="6fbf4-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="6fbf4-161">TabGrayscale</span></span>
- <span data-ttu-id="6fbf4-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="6fbf4-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="6fbf4-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="6fbf4-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="6fbf4-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="6fbf4-164">TabSlideMaster</span></span>
- <span data-ttu-id="6fbf4-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="6fbf4-165">TabHandoutMaster</span></span>
- <span data-ttu-id="6fbf4-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="6fbf4-166">TabNotesMaster</span></span>
- <span data-ttu-id="6fbf4-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="6fbf4-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="6fbf4-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="6fbf4-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="6fbf4-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="6fbf4-169">OneNote</span></span>

- <span data-ttu-id="6fbf4-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-170">**TabHome**</span></span>
- <span data-ttu-id="6fbf4-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-171">**TabInsert**</span></span>
- <span data-ttu-id="6fbf4-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="6fbf4-172">**TabView**</span></span>
- <span data-ttu-id="6fbf4-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="6fbf4-173">TabDeveloper</span></span>
- <span data-ttu-id="6fbf4-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="6fbf4-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="6fbf4-175">组</span><span class="sxs-lookup"><span data-stu-id="6fbf4-175">Group</span></span>

<span data-ttu-id="6fbf4-176">选项卡中的一组 UI 扩展点。</span><span class="sxs-lookup"><span data-stu-id="6fbf4-176">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="6fbf4-177">**Id**属性是必需的，并且每个**id**在清单中必须是唯一的。</span><span class="sxs-lookup"><span data-stu-id="6fbf4-177">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="6fbf4-178">**Id**是最多为125个字符的字符串。</span><span class="sxs-lookup"><span data-stu-id="6fbf4-178">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="6fbf4-179">查看 [Group 元素](group.md)。</span><span class="sxs-lookup"><span data-stu-id="6fbf4-179">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="6fbf4-180">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="6fbf4-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
