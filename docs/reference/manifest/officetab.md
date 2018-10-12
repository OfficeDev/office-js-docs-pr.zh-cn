# <a name="officetab-element"></a><span data-ttu-id="2f2cd-101">OfficeTab 元素</span><span class="sxs-lookup"><span data-stu-id="2f2cd-101">OfficeTab element</span></span>

<span data-ttu-id="2f2cd-p101">定义显示加载项命令的功能区选项卡。这可以是默认的选项卡（**Home**、**Message** 或 **Meeting**），或是由加载项定义的自定义选项卡。此元素是必需的。</span><span class="sxs-lookup"><span data-stu-id="2f2cd-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="2f2cd-105">子元素</span><span class="sxs-lookup"><span data-stu-id="2f2cd-105">Child elements</span></span>

|  <span data-ttu-id="2f2cd-106">元素</span><span class="sxs-lookup"><span data-stu-id="2f2cd-106">Element</span></span> |  <span data-ttu-id="2f2cd-107">必需</span><span class="sxs-lookup"><span data-stu-id="2f2cd-107">Required</span></span>  |  <span data-ttu-id="2f2cd-108">说明</span><span class="sxs-lookup"><span data-stu-id="2f2cd-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2f2cd-109">Group</span><span class="sxs-lookup"><span data-stu-id="2f2cd-109">Group</span></span>      | <span data-ttu-id="2f2cd-110">是</span><span class="sxs-lookup"><span data-stu-id="2f2cd-110">Yes</span></span> |  <span data-ttu-id="2f2cd-p102">定义一组命令。只可以将每一加载项的一个组添加到默认选项卡。</span><span class="sxs-lookup"><span data-stu-id="2f2cd-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="2f2cd-113">下面是主机的有效选项卡 `id` 值。</span><span class="sxs-lookup"><span data-stu-id="2f2cd-113">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="2f2cd-114">在桌面和联机状态中（例如，适用于 Windows 的 Word 2016 或更高版本和 Word Online）都支持以**加粗**显示的值。</span><span class="sxs-lookup"><span data-stu-id="2f2cd-114">Values in **bold** are supported in both desktop and online (for example, Word 2016 for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="2f2cd-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="2f2cd-115">Outlook</span></span>

- <span data-ttu-id="2f2cd-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="2f2cd-117">Word</span><span class="sxs-lookup"><span data-stu-id="2f2cd-117">Word</span></span>

- <span data-ttu-id="2f2cd-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-118">**TabHome**</span></span>
- <span data-ttu-id="2f2cd-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-119">**TabInsert**</span></span>
- <span data-ttu-id="2f2cd-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="2f2cd-120">TabWordDesign</span></span>
- <span data-ttu-id="2f2cd-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="2f2cd-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="2f2cd-122">TabReferences</span></span>
- <span data-ttu-id="2f2cd-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="2f2cd-123">TabMailings</span></span>
- <span data-ttu-id="2f2cd-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="2f2cd-124">TabReviewWord</span></span>
- <span data-ttu-id="2f2cd-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-125">**TabView**</span></span>
- <span data-ttu-id="2f2cd-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="2f2cd-126">TabDeveloper</span></span>
- <span data-ttu-id="2f2cd-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="2f2cd-127">TabAddIns</span></span>
- <span data-ttu-id="2f2cd-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="2f2cd-128">TabBlogPost</span></span>
- <span data-ttu-id="2f2cd-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="2f2cd-129">TabBlogInsert</span></span>
- <span data-ttu-id="2f2cd-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="2f2cd-130">TabPrintPreview</span></span>
- <span data-ttu-id="2f2cd-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="2f2cd-131">TabOutlining</span></span>
- <span data-ttu-id="2f2cd-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="2f2cd-132">TabConflicts</span></span>
- <span data-ttu-id="2f2cd-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="2f2cd-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="2f2cd-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="2f2cd-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="2f2cd-135">Excel</span><span class="sxs-lookup"><span data-stu-id="2f2cd-135">Excel</span></span>

- <span data-ttu-id="2f2cd-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-136">**TabHome**</span></span>
- <span data-ttu-id="2f2cd-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-137">**TabInsert**</span></span>
- <span data-ttu-id="2f2cd-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="2f2cd-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="2f2cd-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="2f2cd-139">TabFormulas</span></span>
- <span data-ttu-id="2f2cd-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-140">**TabData**</span></span>
- <span data-ttu-id="2f2cd-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-141">**TabReview**</span></span>
- <span data-ttu-id="2f2cd-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-142">**TabView**</span></span>
- <span data-ttu-id="2f2cd-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="2f2cd-143">TabDeveloper</span></span>
- <span data-ttu-id="2f2cd-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="2f2cd-144">TabAddIns</span></span>
- <span data-ttu-id="2f2cd-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="2f2cd-145">TabPrintPreview</span></span>
- <span data-ttu-id="2f2cd-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="2f2cd-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="2f2cd-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="2f2cd-147">PowerPoint</span></span>

- <span data-ttu-id="2f2cd-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-148">**TabHome**</span></span>
- <span data-ttu-id="2f2cd-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-149">**TabInsert**</span></span>
- <span data-ttu-id="2f2cd-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-150">**TabDesign**</span></span>
- <span data-ttu-id="2f2cd-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-151">**TabTransitions**</span></span>
- <span data-ttu-id="2f2cd-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-152">**TabAnimations**</span></span>
- <span data-ttu-id="2f2cd-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="2f2cd-153">TabSlideShow</span></span>
- <span data-ttu-id="2f2cd-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="2f2cd-154">TabReview</span></span>
- <span data-ttu-id="2f2cd-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-155">**TabView**</span></span>
- <span data-ttu-id="2f2cd-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="2f2cd-156">TabDeveloper</span></span>
- <span data-ttu-id="2f2cd-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="2f2cd-157">TabAddIns</span></span>
- <span data-ttu-id="2f2cd-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="2f2cd-158">TabPrintPreview</span></span>
- <span data-ttu-id="2f2cd-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="2f2cd-159">TabMerge</span></span>
- <span data-ttu-id="2f2cd-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="2f2cd-160">TabGrayscale</span></span>
- <span data-ttu-id="2f2cd-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="2f2cd-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="2f2cd-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="2f2cd-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="2f2cd-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="2f2cd-163">TabSlideMaster</span></span>
- <span data-ttu-id="2f2cd-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="2f2cd-164">TabHandoutMaster</span></span>
- <span data-ttu-id="2f2cd-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="2f2cd-165">TabNotesMaster</span></span>
- <span data-ttu-id="2f2cd-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="2f2cd-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="2f2cd-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="2f2cd-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="2f2cd-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="2f2cd-168">OneNote</span></span>

- <span data-ttu-id="2f2cd-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-169">**TabHome**</span></span>
- <span data-ttu-id="2f2cd-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-170">**TabInsert**</span></span>
- <span data-ttu-id="2f2cd-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="2f2cd-171">**TabView**</span></span>
- <span data-ttu-id="2f2cd-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="2f2cd-172">TabDeveloper</span></span>
- <span data-ttu-id="2f2cd-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="2f2cd-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="2f2cd-174">Group</span><span class="sxs-lookup"><span data-stu-id="2f2cd-174">Group</span></span>

<span data-ttu-id="2f2cd-p104">选项卡中的一组 UI 扩展点。一组最多可以有六个控件。需要 **id** 属性且每个 **id** 在清单内必须是唯一的。**id** 是一个最大长度为 125 个字符的字符串。请参阅 [Group 元素](group.md)。</span><span class="sxs-lookup"><span data-stu-id="2f2cd-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="2f2cd-179">OfficeTab 示例</span><span class="sxs-lookup"><span data-stu-id="2f2cd-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
