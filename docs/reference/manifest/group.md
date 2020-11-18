---
title: 清单文件中的 Group 元素
description: 定义选项卡中的一组 UI 控件。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 6ee8d499767eccb95b4fdf9ceb91dd2cd12bce95
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087943"
---
# <a name="group-element"></a><span data-ttu-id="c4cd3-103">Group 元素</span><span class="sxs-lookup"><span data-stu-id="c4cd3-103">Group element</span></span>

<span data-ttu-id="c4cd3-104">定义选项卡中的一组 UI 控件。在自定义选项卡上，加载项可以创建多个组。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="c4cd3-105">外接程序限定到一个自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="c4cd3-106">属性</span><span class="sxs-lookup"><span data-stu-id="c4cd3-106">Attributes</span></span>

|  <span data-ttu-id="c4cd3-107">属性</span><span class="sxs-lookup"><span data-stu-id="c4cd3-107">Attribute</span></span>  |  <span data-ttu-id="c4cd3-108">必需</span><span class="sxs-lookup"><span data-stu-id="c4cd3-108">Required</span></span>  |  <span data-ttu-id="c4cd3-109">说明</span><span class="sxs-lookup"><span data-stu-id="c4cd3-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c4cd3-110">id</span><span class="sxs-lookup"><span data-stu-id="c4cd3-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="c4cd3-111">是</span><span class="sxs-lookup"><span data-stu-id="c4cd3-111">Yes</span></span>  | <span data-ttu-id="c4cd3-112">组的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="c4cd3-113">id attribute</span><span class="sxs-lookup"><span data-stu-id="c4cd3-113">id attribute</span></span>

<span data-ttu-id="c4cd3-p102">必需。 组的唯一标识符。 是一个最多为 125 个字符的字符串。 该字符串在清单内必须是唯一的，否则组将不能呈现。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c4cd3-118">子元素</span><span class="sxs-lookup"><span data-stu-id="c4cd3-118">Child elements</span></span>

|  <span data-ttu-id="c4cd3-119">元素</span><span class="sxs-lookup"><span data-stu-id="c4cd3-119">Element</span></span> |  <span data-ttu-id="c4cd3-120">必需</span><span class="sxs-lookup"><span data-stu-id="c4cd3-120">Required</span></span>  |  <span data-ttu-id="c4cd3-121">说明</span><span class="sxs-lookup"><span data-stu-id="c4cd3-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c4cd3-122">Label</span><span class="sxs-lookup"><span data-stu-id="c4cd3-122">Label</span></span>](#label)      | <span data-ttu-id="c4cd3-123">是</span><span class="sxs-lookup"><span data-stu-id="c4cd3-123">Yes</span></span> |  <span data-ttu-id="c4cd3-124">CustomTab 或组的标签。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="c4cd3-125">Icon</span><span class="sxs-lookup"><span data-stu-id="c4cd3-125">Icon</span></span>](icon.md)      | <span data-ttu-id="c4cd3-126">是</span><span class="sxs-lookup"><span data-stu-id="c4cd3-126">Yes</span></span> |  <span data-ttu-id="c4cd3-127">组的图像。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="c4cd3-128">Control</span><span class="sxs-lookup"><span data-stu-id="c4cd3-128">Control</span></span>](#control)    | <span data-ttu-id="c4cd3-129">否</span><span class="sxs-lookup"><span data-stu-id="c4cd3-129">No</span></span> |  <span data-ttu-id="c4cd3-130">表示控件对象。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-130">Represents a Control object.</span></span> <span data-ttu-id="c4cd3-131">可以是零个或多个。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-131">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="c4cd3-132">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="c4cd3-132">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="c4cd3-133">否</span><span class="sxs-lookup"><span data-stu-id="c4cd3-133">No</span></span> | <span data-ttu-id="c4cd3-134">表示内置 Office 控件之一。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-134">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="c4cd3-135">可以是零个或多个。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-135">Can be zero or more.</span></span> |

### <a name="label"></a><span data-ttu-id="c4cd3-136">标签</span><span class="sxs-lookup"><span data-stu-id="c4cd3-136">Label</span></span>

<span data-ttu-id="c4cd3-137">必需。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-137">Required.</span></span> <span data-ttu-id="c4cd3-138">组的标签。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-138">The label of the group.</span></span> <span data-ttu-id="c4cd3-139">**Resid** 属性必须设置为 [Resources](resources.md)元素中的 **ShortStrings** 元素中 **String** 元素的 **id** 属性的值。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-139">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="c4cd3-140">Icon</span><span class="sxs-lookup"><span data-stu-id="c4cd3-140">Icon</span></span>

<span data-ttu-id="c4cd3-141">必需。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-141">Required.</span></span> <span data-ttu-id="c4cd3-142">如果某个选项卡包含大量组，并且该程序窗口已调整大小，则会改为显示指定的图像。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-142">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="c4cd3-143">控制</span><span class="sxs-lookup"><span data-stu-id="c4cd3-143">Control</span></span>

<span data-ttu-id="c4cd3-144">可选，但如果不存在，则必须至少有一个 **OfficeControl**。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-144">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="c4cd3-145">有关受支持的控件类型的详细信息，请参阅 [Control](control.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-145">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="c4cd3-146">在清单中， **Control** 和 **OfficeControl** 的顺序是可互换的，如果存在多个元素，则可以是混合的，但所有元素都必须位于 **Icon** 元素的下面。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-146">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```

### <a name="officecontrol"></a><span data-ttu-id="c4cd3-147">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="c4cd3-147">OfficeControl</span></span>

<span data-ttu-id="c4cd3-148">可选，但如果不存在，则必须至少有一个 **控件**。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-148">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="c4cd3-149">在带元素的组中包含一个或多个内置 Office 控件 `<OfficeControl>` 。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-149">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="c4cd3-150">`id`属性指定内置 Office 控件的 ID。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-150">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="c4cd3-151">若要查找控件的 ID，请参阅 [查找控件和控件组的 id](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-151">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="c4cd3-152">在清单中， **Control** 和 **OfficeControl** 的顺序是可互换的，如果存在多个元素，则可以是混合的，但所有元素都必须位于 **Icon** 元素的下面。</span><span class="sxs-lookup"><span data-stu-id="c4cd3-152">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```
