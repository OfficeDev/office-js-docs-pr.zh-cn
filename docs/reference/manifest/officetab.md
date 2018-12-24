---
title: 清单文件中的 OfficeTab 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 721064687c3c892b565a94e418815726cc0817f5
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432870"
---
# <a name="officetab-element"></a><span data-ttu-id="cc1cd-102">OfficeTab 元素</span><span class="sxs-lookup"><span data-stu-id="cc1cd-102">OfficeTab element</span></span>

<span data-ttu-id="cc1cd-p101">定义在其上显示外接程序命令的功能区选项卡。这可以是默认的选项卡（“**主页**”、“**消息**”或“**会议**”），或是由外接程序定义的自定义选项卡。此元素是必需的。</span><span class="sxs-lookup"><span data-stu-id="cc1cd-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="cc1cd-106">子元素</span><span class="sxs-lookup"><span data-stu-id="cc1cd-106">Child elements</span></span>

|  <span data-ttu-id="cc1cd-107">元素</span><span class="sxs-lookup"><span data-stu-id="cc1cd-107">Element</span></span> |  <span data-ttu-id="cc1cd-108">必需</span><span class="sxs-lookup"><span data-stu-id="cc1cd-108">Required</span></span>  |  <span data-ttu-id="cc1cd-109">说明</span><span class="sxs-lookup"><span data-stu-id="cc1cd-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="cc1cd-110">组</span><span class="sxs-lookup"><span data-stu-id="cc1cd-110">Group</span></span>      | <span data-ttu-id="cc1cd-111">必需</span><span class="sxs-lookup"><span data-stu-id="cc1cd-111">Yes</span></span> |  <span data-ttu-id="cc1cd-p102">定义一组命令。对于每个外接程序，只能将一个组添加到默认选项卡。</span><span class="sxs-lookup"><span data-stu-id="cc1cd-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="cc1cd-114">下面是主机的有效选项卡 `id` 值。</span><span class="sxs-lookup"><span data-stu-id="cc1cd-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="cc1cd-115">以**粗体** 显示的值在桌面和联机状态中均受支持（例如，适用于 Windows 的 Word 2016 或更高版本和 Word Online）。</span><span class="sxs-lookup"><span data-stu-id="cc1cd-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="cc1cd-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="cc1cd-116">Outlook</span></span>

- <span data-ttu-id="cc1cd-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="cc1cd-118">Word</span><span class="sxs-lookup"><span data-stu-id="cc1cd-118">Word</span></span>

- <span data-ttu-id="cc1cd-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-119">**TabHome**</span></span>
- <span data-ttu-id="cc1cd-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-120">**TabInsert**</span></span>
- <span data-ttu-id="cc1cd-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="cc1cd-121">TabWordDesign</span></span>
- <span data-ttu-id="cc1cd-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="cc1cd-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="cc1cd-123">TabReferences</span></span>
- <span data-ttu-id="cc1cd-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="cc1cd-124">TabMailings</span></span>
- <span data-ttu-id="cc1cd-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="cc1cd-125">TabReviewWord</span></span>
- <span data-ttu-id="cc1cd-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-126">**TabView**</span></span>
- <span data-ttu-id="cc1cd-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="cc1cd-127">TabDeveloper</span></span>
- <span data-ttu-id="cc1cd-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="cc1cd-128">TabAddIns</span></span>
- <span data-ttu-id="cc1cd-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="cc1cd-129">TabBlogPost</span></span>
- <span data-ttu-id="cc1cd-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="cc1cd-130">TabBlogInsert</span></span>
- <span data-ttu-id="cc1cd-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="cc1cd-131">TabPrintPreview</span></span>
- <span data-ttu-id="cc1cd-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="cc1cd-132">TabOutlining</span></span>
- <span data-ttu-id="cc1cd-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="cc1cd-133">TabConflicts</span></span>
- <span data-ttu-id="cc1cd-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="cc1cd-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="cc1cd-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="cc1cd-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="cc1cd-136">Excel</span><span class="sxs-lookup"><span data-stu-id="cc1cd-136">Excel</span></span>

- <span data-ttu-id="cc1cd-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-137">**TabHome**</span></span>
- <span data-ttu-id="cc1cd-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-138">**TabInsert**</span></span>
- <span data-ttu-id="cc1cd-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="cc1cd-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="cc1cd-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="cc1cd-140">TabFormulas</span></span>
- <span data-ttu-id="cc1cd-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-141">**TabData**</span></span>
- <span data-ttu-id="cc1cd-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-142">**TabReview**</span></span>
- <span data-ttu-id="cc1cd-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-143">**TabView**</span></span>
- <span data-ttu-id="cc1cd-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="cc1cd-144">TabDeveloper</span></span>
- <span data-ttu-id="cc1cd-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="cc1cd-145">TabAddIns</span></span>
- <span data-ttu-id="cc1cd-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="cc1cd-146">TabPrintPreview</span></span>
- <span data-ttu-id="cc1cd-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="cc1cd-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="cc1cd-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="cc1cd-148">PowerPoint</span></span>

- <span data-ttu-id="cc1cd-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-149">**TabHome**</span></span>
- <span data-ttu-id="cc1cd-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-150">**TabInsert**</span></span>
- <span data-ttu-id="cc1cd-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-151">**TabDesign**</span></span>
- <span data-ttu-id="cc1cd-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-152">**TabTransitions**</span></span>
- <span data-ttu-id="cc1cd-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-153">**TabAnimations**</span></span>
- <span data-ttu-id="cc1cd-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="cc1cd-154">TabSlideShow</span></span>
- <span data-ttu-id="cc1cd-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="cc1cd-155">TabReview</span></span>
- <span data-ttu-id="cc1cd-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-156">**TabView**</span></span>
- <span data-ttu-id="cc1cd-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="cc1cd-157">TabDeveloper</span></span>
- <span data-ttu-id="cc1cd-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="cc1cd-158">TabAddIns</span></span>
- <span data-ttu-id="cc1cd-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="cc1cd-159">TabPrintPreview</span></span>
- <span data-ttu-id="cc1cd-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="cc1cd-160">TabMerge</span></span>
- <span data-ttu-id="cc1cd-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="cc1cd-161">TabGrayscale</span></span>
- <span data-ttu-id="cc1cd-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="cc1cd-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="cc1cd-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="cc1cd-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="cc1cd-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="cc1cd-164">TabSlideMaster</span></span>
- <span data-ttu-id="cc1cd-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="cc1cd-165">TabHandoutMaster</span></span>
- <span data-ttu-id="cc1cd-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="cc1cd-166">TabNotesMaster</span></span>
- <span data-ttu-id="cc1cd-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="cc1cd-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="cc1cd-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="cc1cd-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="cc1cd-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="cc1cd-169">OneNote</span></span>

- <span data-ttu-id="cc1cd-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-170">**TabHome**</span></span>
- <span data-ttu-id="cc1cd-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-171">**TabInsert**</span></span>
- <span data-ttu-id="cc1cd-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="cc1cd-172">**TabView**</span></span>
- <span data-ttu-id="cc1cd-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="cc1cd-173">TabDeveloper</span></span>
- <span data-ttu-id="cc1cd-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="cc1cd-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="cc1cd-175">组</span><span class="sxs-lookup"><span data-stu-id="cc1cd-175">Group</span></span>

<span data-ttu-id="cc1cd-p104">选项卡中的一组 UI 扩展点。一组可以有多达六个控件。需要 **id** 属性且每个 **id** 在清单内必须是唯一的。**id** 是一个最大长度为 125 个字符的字符串。查看 [Group 元素](group.md)。</span><span class="sxs-lookup"><span data-stu-id="cc1cd-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="cc1cd-180">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="cc1cd-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
