---
title: 清单文件中的 CustomTab 元素
description: 在功能区上，可以为它们的外接程序命令指定使用哪种选项卡和组。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: d74859d1326d29517b5a8226a86f901322957933
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173925"
---
# <a name="customtab-element"></a><span data-ttu-id="f04d8-103">CustomTab 元素</span><span class="sxs-lookup"><span data-stu-id="f04d8-103">CustomTab element</span></span>

<span data-ttu-id="f04d8-104">在功能区上，指定外接程序命令的选项卡和组。</span><span class="sxs-lookup"><span data-stu-id="f04d8-104">On the ribbon, specify the tab and group for your add-in commands.</span></span> <span data-ttu-id="f04d8-105">这可能位于默认选项卡（“主页”、“邮件”或“会议”）上，或位于外接程序定义的自定义选项卡上。</span><span class="sxs-lookup"><span data-stu-id="f04d8-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="f04d8-106">在自定义选项卡上，加载项可以具有自定义组或内置组。</span><span class="sxs-lookup"><span data-stu-id="f04d8-106">On custom tabs, the add-in can have custom or built-in groups.</span></span> <span data-ttu-id="f04d8-107">外接程序限定到一个自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="f04d8-107">Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="f04d8-108">**id** 属性在清单中必须是唯一的。</span><span class="sxs-lookup"><span data-stu-id="f04d8-108">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f04d8-109">在 Mac 上的 Outlook 中，该元素 `CustomTab` 不可用，因此您必须改为使用[OfficeTab。](officetab.md)</span><span class="sxs-lookup"><span data-stu-id="f04d8-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="f04d8-110">子元素</span><span class="sxs-lookup"><span data-stu-id="f04d8-110">Child elements</span></span>

|  <span data-ttu-id="f04d8-111">元素</span><span class="sxs-lookup"><span data-stu-id="f04d8-111">Element</span></span> |  <span data-ttu-id="f04d8-112">必需</span><span class="sxs-lookup"><span data-stu-id="f04d8-112">Required</span></span>  |  <span data-ttu-id="f04d8-113">说明</span><span class="sxs-lookup"><span data-stu-id="f04d8-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f04d8-114">Group</span><span class="sxs-lookup"><span data-stu-id="f04d8-114">Group</span></span>](group.md)      | <span data-ttu-id="f04d8-115">否</span><span class="sxs-lookup"><span data-stu-id="f04d8-115">No</span></span> |  <span data-ttu-id="f04d8-116">定义一组命令。</span><span class="sxs-lookup"><span data-stu-id="f04d8-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="f04d8-117">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="f04d8-117">OfficeGroup</span></span>](#officegroup)      | <span data-ttu-id="f04d8-118">否</span><span class="sxs-lookup"><span data-stu-id="f04d8-118">No</span></span> |  <span data-ttu-id="f04d8-119">代表内置的 Office 控件组。</span><span class="sxs-lookup"><span data-stu-id="f04d8-119">Represents a built-in Office control group.</span></span> <span data-ttu-id="f04d8-120">**重要** 提示：在 Outlook 中不可用。</span><span class="sxs-lookup"><span data-stu-id="f04d8-120">**Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="f04d8-121">Label</span><span class="sxs-lookup"><span data-stu-id="f04d8-121">Label</span></span>](#label-tab)      | <span data-ttu-id="f04d8-122">是</span><span class="sxs-lookup"><span data-stu-id="f04d8-122">Yes</span></span> |  <span data-ttu-id="f04d8-123">CustomTab 或组的标签。</span><span class="sxs-lookup"><span data-stu-id="f04d8-123">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="f04d8-124">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="f04d8-124">InsertAfter</span></span>](#insertafter)      | <span data-ttu-id="f04d8-125">否</span><span class="sxs-lookup"><span data-stu-id="f04d8-125">No</span></span> |  <span data-ttu-id="f04d8-126">指定自定义选项卡应紧接在指定的内置 Office 选项卡之后。 **重要** 说明：在 Outlook 中不可用。</span><span class="sxs-lookup"><span data-stu-id="f04d8-126">Specifies that the custom tab should be immediately after a specified built-in Office tab. **Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="f04d8-127">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="f04d8-127">InsertBefore</span></span>](#insertbefore)      | <span data-ttu-id="f04d8-128">否</span><span class="sxs-lookup"><span data-stu-id="f04d8-128">No</span></span> |  <span data-ttu-id="f04d8-129">指定自定义选项卡应紧接在指定的内置 Office 选项卡之前。 **重要** 说明：在 Outlook 中不可用。</span><span class="sxs-lookup"><span data-stu-id="f04d8-129">Specifies that the custom tab should be immediately before a specified built-in Office tab. **Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="f04d8-130">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="f04d8-130">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="f04d8-131">否</span><span class="sxs-lookup"><span data-stu-id="f04d8-131">No</span></span> |  <span data-ttu-id="f04d8-132">指定自定义选项卡是否应该显示在支持自定义上下文选项卡的应用程序和平台组合上。</span><span class="sxs-lookup"><span data-stu-id="f04d8-132">Specifies whether the custom tab should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="f04d8-133">**重要** 提示：在 Outlook 中不可用。</span><span class="sxs-lookup"><span data-stu-id="f04d8-133">**Important**: Not available in Outlook.</span></span> |

### <a name="group"></a><span data-ttu-id="f04d8-134">Group</span><span class="sxs-lookup"><span data-stu-id="f04d8-134">Group</span></span>

<span data-ttu-id="f04d8-135">可选，但如果不存在，则必须至少有一 **个 OfficeGroup** 元素。</span><span class="sxs-lookup"><span data-stu-id="f04d8-135">Optional, but if not present there must be at least one **OfficeGroup** element.</span></span> <span data-ttu-id="f04d8-136">查看 [Group 元素](group.md)。</span><span class="sxs-lookup"><span data-stu-id="f04d8-136">See [Group element](group.md).</span></span> <span data-ttu-id="f04d8-137">清单中 **组** 和 **OfficeGroup** 的顺序应该是您希望它们显示在自定义选项卡上的顺序。如果有多个元素，它们可能会同时存在，但所有元素都必须在 **Label 元素** 上方。</span><span class="sxs-lookup"><span data-stu-id="f04d8-137">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="officegroup"></a><span data-ttu-id="f04d8-138">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="f04d8-138">OfficeGroup</span></span>

<span data-ttu-id="f04d8-139">可选，但如果不存在，则必须至少有一 **个 Group** 元素。</span><span class="sxs-lookup"><span data-stu-id="f04d8-139">Optional, but if not present there must be at least one **Group** element.</span></span> <span data-ttu-id="f04d8-140">代表内置的 Office 控件组。</span><span class="sxs-lookup"><span data-stu-id="f04d8-140">Represents a built-in Office control group.</span></span> <span data-ttu-id="f04d8-141">**id** 属性指定内置 Office 组的 ID。</span><span class="sxs-lookup"><span data-stu-id="f04d8-141">The **id** attribute specifies the ID of the built-in Office group.</span></span> <span data-ttu-id="f04d8-142">若要查找内置组的 ID，请参阅"查找控件和[控件组的 ID"。](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="f04d8-142">To find the ID of a built-in group, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="f04d8-143">清单中 **组** 和 **OfficeGroup** 的顺序应该是您希望它们显示在自定义选项卡上的顺序。如果有多个元素，它们可能会同时存在，但所有元素都必须在 **Label 元素** 上方。</span><span class="sxs-lookup"><span data-stu-id="f04d8-143">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f04d8-144">`OfficeGroup`该元素在 Outlook 中不可用。</span><span class="sxs-lookup"><span data-stu-id="f04d8-144">The `OfficeGroup` element is not available in Outlook.</span></span>

### <a name="label-tab"></a><span data-ttu-id="f04d8-145">标签（选项卡）</span><span class="sxs-lookup"><span data-stu-id="f04d8-145">Label (Tab)</span></span>

<span data-ttu-id="f04d8-146">必需。</span><span class="sxs-lookup"><span data-stu-id="f04d8-146">Required.</span></span> <span data-ttu-id="f04d8-147">自定义选项卡的标签。**resid** 属性的长度不能超过 32 个字符，必须设置为 Resources 元素 **中 ShortStrings** 元素 **中 String** 元素的 **id** [属性值。](resources.md)</span><span class="sxs-lookup"><span data-stu-id="f04d8-147">The label of the custom tab. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="insertafter"></a><span data-ttu-id="f04d8-148">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="f04d8-148">InsertAfter</span></span>

<span data-ttu-id="f04d8-149">可选。</span><span class="sxs-lookup"><span data-stu-id="f04d8-149">Optional.</span></span> <span data-ttu-id="f04d8-150">指定自定义选项卡应紧接在指定的内置 Office 选项卡之后。元素的值是内置选项卡的 ID，如"TabHome"或"TabReview"。</span><span class="sxs-lookup"><span data-stu-id="f04d8-150">Specifies that the custom tab should be immediately after a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="f04d8-151"> ([查找控件和](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)控件组的标识。) 如果存在，则必须在 **Label 元素** 之后。</span><span class="sxs-lookup"><span data-stu-id="f04d8-151">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present, must be after the **Label** element.</span></span> <span data-ttu-id="f04d8-152">不能同时具有 **InsertAfter 和** **InsertBefore。**</span><span class="sxs-lookup"><span data-stu-id="f04d8-152">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f04d8-153">`InsertAfter`该元素在 Outlook 中不可用。</span><span class="sxs-lookup"><span data-stu-id="f04d8-153">The `InsertAfter` element is not available in Outlook.</span></span>

### <a name="insertbefore"></a><span data-ttu-id="f04d8-154">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="f04d8-154">InsertBefore</span></span>

<span data-ttu-id="f04d8-155">可选。</span><span class="sxs-lookup"><span data-stu-id="f04d8-155">Optional.</span></span> <span data-ttu-id="f04d8-156">指定自定义选项卡应紧接在指定的内置 Office 选项卡之前。元素的值是内置选项卡的 ID，如"TabHome"或"TabReview"。</span><span class="sxs-lookup"><span data-stu-id="f04d8-156">Specifies that the custom tab should be immediately before a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="f04d8-157"> ([查找控件和](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)控件组的标识。) 如果存在，则必须在 **Label 元素** 之后。</span><span class="sxs-lookup"><span data-stu-id="f04d8-157">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  If present, must be after the **Label** element.</span></span> <span data-ttu-id="f04d8-158">不能同时具有 **InsertAfter 和** **InsertBefore。**</span><span class="sxs-lookup"><span data-stu-id="f04d8-158">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f04d8-159">`InsertBefore`该元素在 Outlook 中不可用。</span><span class="sxs-lookup"><span data-stu-id="f04d8-159">The `InsertBefore` element is not available in Outlook.</span></span>

### <a name="overriddenbyribbonapi"></a><span data-ttu-id="f04d8-160">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="f04d8-160">OverriddenByRibbonApi</span></span>

<span data-ttu-id="f04d8-161">可选 (布尔) 。</span><span class="sxs-lookup"><span data-stu-id="f04d8-161">Optional (boolean).</span></span> <span data-ttu-id="f04d8-162">指定在支持 API 的应用程序和平台组合上是否隐藏 **CustomTab，** 该 API 在运行时在功能区上安装自定义上下文选项卡。</span><span class="sxs-lookup"><span data-stu-id="f04d8-162">Specifies whether the **CustomTab** will be hidden on application and platform combinations that support an API that installs a custom contextual tab on the ribbon at runtime.</span></span> <span data-ttu-id="f04d8-163">默认值（如果不存在）为 `false` 。</span><span class="sxs-lookup"><span data-stu-id="f04d8-163">The default value, if not present, is `false`.</span></span> <span data-ttu-id="f04d8-164">如果使用 **，OverriddenByRibbonApi** 必须是 **CustomTab 的第一个子级**。 </span><span class="sxs-lookup"><span data-stu-id="f04d8-164">If used, **OverriddenByRibbonApi** must be the *first* child of **CustomTab**.</span></span> <span data-ttu-id="f04d8-165">有关详细信息，请参阅 [OverriddenByRibbonApi](overriddenbyribbonapi.md)。</span><span class="sxs-lookup"><span data-stu-id="f04d8-165">For more information, see [OverriddenByRibbonApi](overriddenbyribbonapi.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f04d8-166">`OverriddenByRibbonApi`该元素在 Outlook 中不可用。</span><span class="sxs-lookup"><span data-stu-id="f04d8-166">The `OverriddenByRibbonApi` element is not available in Outlook.</span></span>

## <a name="customtab-example"></a><span data-ttu-id="f04d8-167">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="f04d8-167">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
