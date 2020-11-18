---
title: 清单文件中的 CustomTab 元素
description: 在功能区上，可以为它们的外接程序命令指定使用哪种选项卡和组。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 99670b27d963060a008899a8808ca967cfd710a6
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087936"
---
# <a name="customtab-element"></a><span data-ttu-id="f0789-103">CustomTab 元素</span><span class="sxs-lookup"><span data-stu-id="f0789-103">CustomTab element</span></span>

<span data-ttu-id="f0789-104">在功能区上，为您的外接程序命令指定选项卡和组。</span><span class="sxs-lookup"><span data-stu-id="f0789-104">On the ribbon, specify the tab and group for your add-in commands.</span></span> <span data-ttu-id="f0789-105">这可能位于默认选项卡（“主页”、“邮件”或“会议”）上，或位于外接程序定义的自定义选项卡上。</span><span class="sxs-lookup"><span data-stu-id="f0789-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="f0789-106">在自定义选项卡上，加载项可以具有自定义或内置组。</span><span class="sxs-lookup"><span data-stu-id="f0789-106">On custom tabs, the add-in can have custom or built-in groups.</span></span> <span data-ttu-id="f0789-107">外接程序限定到一个自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="f0789-107">Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="f0789-108">**Id** 属性在清单中必须是唯一的。</span><span class="sxs-lookup"><span data-stu-id="f0789-108">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f0789-109">在 Mac 上的 Outlook 中，该元素不可用， `CustomTab` 因此您必须改用 [OfficeTab](officetab.md) 。</span><span class="sxs-lookup"><span data-stu-id="f0789-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="f0789-110">子元素</span><span class="sxs-lookup"><span data-stu-id="f0789-110">Child elements</span></span>

|  <span data-ttu-id="f0789-111">元素</span><span class="sxs-lookup"><span data-stu-id="f0789-111">Element</span></span> |  <span data-ttu-id="f0789-112">必需</span><span class="sxs-lookup"><span data-stu-id="f0789-112">Required</span></span>  |  <span data-ttu-id="f0789-113">说明</span><span class="sxs-lookup"><span data-stu-id="f0789-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f0789-114">Group</span><span class="sxs-lookup"><span data-stu-id="f0789-114">Group</span></span>](group.md)      | <span data-ttu-id="f0789-115">否</span><span class="sxs-lookup"><span data-stu-id="f0789-115">No</span></span> |  <span data-ttu-id="f0789-116">定义一组命令。</span><span class="sxs-lookup"><span data-stu-id="f0789-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="f0789-117">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="f0789-117">OfficeGroup</span></span>](#officegroup)      | <span data-ttu-id="f0789-118">否</span><span class="sxs-lookup"><span data-stu-id="f0789-118">No</span></span> |  <span data-ttu-id="f0789-119">代表内置 Office 控件组。</span><span class="sxs-lookup"><span data-stu-id="f0789-119">Represents a built-in Office control group.</span></span>  |
|  [<span data-ttu-id="f0789-120">Label</span><span class="sxs-lookup"><span data-stu-id="f0789-120">Label</span></span>](#label-tab)      | <span data-ttu-id="f0789-121">是</span><span class="sxs-lookup"><span data-stu-id="f0789-121">Yes</span></span> |  <span data-ttu-id="f0789-122">CustomTab 或组的标签。</span><span class="sxs-lookup"><span data-stu-id="f0789-122">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="f0789-123">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="f0789-123">InsertAfter</span></span>](#insertafter)      | <span data-ttu-id="f0789-124">否</span><span class="sxs-lookup"><span data-stu-id="f0789-124">No</span></span> |  <span data-ttu-id="f0789-125">指定自定义选项卡应紧跟在指定的内置 "Office" 选项卡之后。</span><span class="sxs-lookup"><span data-stu-id="f0789-125">Specifies that the custom tab should be immediately after a specified built-in Office tab.</span></span>  |
|  [<span data-ttu-id="f0789-126">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="f0789-126">InsertBefore</span></span>](#insertbefore)      | <span data-ttu-id="f0789-127">否</span><span class="sxs-lookup"><span data-stu-id="f0789-127">No</span></span> |  <span data-ttu-id="f0789-128">指定自定义选项卡应紧跟在指定的内置 "Office" 选项卡之前。</span><span class="sxs-lookup"><span data-stu-id="f0789-128">Specifies that the custom tab should be immediately before a specified built-in Office tab.</span></span>  |

### <a name="group"></a><span data-ttu-id="f0789-129">Group</span><span class="sxs-lookup"><span data-stu-id="f0789-129">Group</span></span>

<span data-ttu-id="f0789-130">可选，但如果不存在，则必须至少有一个 **OfficeGroup** 元素。</span><span class="sxs-lookup"><span data-stu-id="f0789-130">Optional, but if not present there must be at least one **OfficeGroup** element.</span></span> <span data-ttu-id="f0789-131">查看 [Group 元素](group.md)。</span><span class="sxs-lookup"><span data-stu-id="f0789-131">See [Group element](group.md).</span></span> <span data-ttu-id="f0789-132">清单和 **OfficeGroup** 在清单 **中的顺序** 应是您希望它们出现在 "自定义" 选项卡上的顺序。如果有多个元素，则可以是混合的，但所有元素都必须位于 **Label** 元素的上方。</span><span class="sxs-lookup"><span data-stu-id="f0789-132">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="officegroup"></a><span data-ttu-id="f0789-133">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="f0789-133">OfficeGroup</span></span>

<span data-ttu-id="f0789-134">可选，但如果不存在，则必须至少有一个 **Group** 元素。</span><span class="sxs-lookup"><span data-stu-id="f0789-134">Optional, but if not present there must be at least one **Group** element.</span></span> <span data-ttu-id="f0789-135">代表内置 Office 控件组。</span><span class="sxs-lookup"><span data-stu-id="f0789-135">Represents a built-in Office control group.</span></span> <span data-ttu-id="f0789-136">**Id** 属性指定内置 Office 组的 id。</span><span class="sxs-lookup"><span data-stu-id="f0789-136">The **id** attribute specifies the ID of the built-in Office group.</span></span> <span data-ttu-id="f0789-137">若要查找内置组的 ID，请参阅 [查找控件和控件组的 id](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。</span><span class="sxs-lookup"><span data-stu-id="f0789-137">To find the ID of a built-in group, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="f0789-138">清单和 **OfficeGroup** 在清单 **中的顺序** 应是您希望它们出现在 "自定义" 选项卡上的顺序。如果有多个元素，则可以是混合的，但所有元素都必须位于 **Label** 元素的上方。</span><span class="sxs-lookup"><span data-stu-id="f0789-138">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="label-tab"></a><span data-ttu-id="f0789-139">标签（选项卡）</span><span class="sxs-lookup"><span data-stu-id="f0789-139">Label (Tab)</span></span>

<span data-ttu-id="f0789-140">必需。</span><span class="sxs-lookup"><span data-stu-id="f0789-140">Required.</span></span> <span data-ttu-id="f0789-141">自定义选项卡的标签。**Resid** 属性必须设置为 [Resources](resources.md)元素中的 **ShortStrings** 元素中 **String** 元素的 **id** 属性的值。</span><span class="sxs-lookup"><span data-stu-id="f0789-141">The label of the custom tab. The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="insertafter"></a><span data-ttu-id="f0789-142">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="f0789-142">InsertAfter</span></span>

<span data-ttu-id="f0789-143">可选。</span><span class="sxs-lookup"><span data-stu-id="f0789-143">Optional.</span></span> <span data-ttu-id="f0789-144">指定自定义选项卡应紧跟在指定的内置 "Office" 选项卡之后。元素的值是内置选项卡的 ID，如 "TabHome" 或 "TabReview"。</span><span class="sxs-lookup"><span data-stu-id="f0789-144">Specifies that the custom tab should be immediately after a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="f0789-145"> (请参阅 [查找控件和控件组的 id](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 ) 如果存在，则必须位于 **Label** 元素之后。</span><span class="sxs-lookup"><span data-stu-id="f0789-145">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present, must be after the **Label** element.</span></span> <span data-ttu-id="f0789-146">您不能同时具有 **InsertAfter** 和 **InsertBefore**。</span><span class="sxs-lookup"><span data-stu-id="f0789-146">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

### <a name="insertbefore"></a><span data-ttu-id="f0789-147">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="f0789-147">InsertBefore</span></span>

<span data-ttu-id="f0789-148">可选。</span><span class="sxs-lookup"><span data-stu-id="f0789-148">Optional.</span></span> <span data-ttu-id="f0789-149">指定自定义选项卡应紧跟在指定的内置 "Office" 选项卡之前。元素的值是内置选项卡的 ID，如 "TabHome" 或 "TabReview"。</span><span class="sxs-lookup"><span data-stu-id="f0789-149">Specifies that the custom tab should be immediately before a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="f0789-150"> (请参阅 [查找控件和控件组的 id](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 ) 如果存在，则必须位于 **Label** 元素之后。</span><span class="sxs-lookup"><span data-stu-id="f0789-150">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  If present, must be after the **Label** element.</span></span> <span data-ttu-id="f0789-151">您不能同时具有 **InsertAfter** 和 **InsertBefore**。</span><span class="sxs-lookup"><span data-stu-id="f0789-151">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

## <a name="customtab-example"></a><span data-ttu-id="f0789-152">CustomTab 示例</span><span class="sxs-lookup"><span data-stu-id="f0789-152">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
