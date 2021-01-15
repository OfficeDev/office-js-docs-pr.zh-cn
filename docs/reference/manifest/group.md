---
title: 清单文件中 Group 元素
description: 在选项卡中定义一组 UI 控件。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 3872ece926cc399ed2b30d4dabaacfb741e060ab
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771395"
---
# <a name="group-element"></a><span data-ttu-id="ea7c3-103">Group 元素</span><span class="sxs-lookup"><span data-stu-id="ea7c3-103">Group element</span></span>

<span data-ttu-id="ea7c3-104">在选项卡中定义一组 UI 控件。在自定义选项卡上，加载项可以创建多个组。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="ea7c3-105">外接程序限定到一个自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="ea7c3-106">属性</span><span class="sxs-lookup"><span data-stu-id="ea7c3-106">Attributes</span></span>

|  <span data-ttu-id="ea7c3-107">属性</span><span class="sxs-lookup"><span data-stu-id="ea7c3-107">Attribute</span></span>  |  <span data-ttu-id="ea7c3-108">必需</span><span class="sxs-lookup"><span data-stu-id="ea7c3-108">Required</span></span>  |  <span data-ttu-id="ea7c3-109">说明</span><span class="sxs-lookup"><span data-stu-id="ea7c3-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ea7c3-110">id</span><span class="sxs-lookup"><span data-stu-id="ea7c3-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="ea7c3-111">是</span><span class="sxs-lookup"><span data-stu-id="ea7c3-111">Yes</span></span>  | <span data-ttu-id="ea7c3-112">组的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="ea7c3-113">id attribute</span><span class="sxs-lookup"><span data-stu-id="ea7c3-113">id attribute</span></span>

<span data-ttu-id="ea7c3-p102">必需。 组的唯一标识符。 是一个最多为 125 个字符的字符串。 该字符串在清单内必须是唯一的，否则组将不能呈现。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="ea7c3-118">子元素</span><span class="sxs-lookup"><span data-stu-id="ea7c3-118">Child elements</span></span>

|  <span data-ttu-id="ea7c3-119">元素</span><span class="sxs-lookup"><span data-stu-id="ea7c3-119">Element</span></span> |  <span data-ttu-id="ea7c3-120">必需</span><span class="sxs-lookup"><span data-stu-id="ea7c3-120">Required</span></span>  |  <span data-ttu-id="ea7c3-121">说明</span><span class="sxs-lookup"><span data-stu-id="ea7c3-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ea7c3-122">Label</span><span class="sxs-lookup"><span data-stu-id="ea7c3-122">Label</span></span>](#label)      | <span data-ttu-id="ea7c3-123">是</span><span class="sxs-lookup"><span data-stu-id="ea7c3-123">Yes</span></span> |  <span data-ttu-id="ea7c3-124">CustomTab 或组的标签。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="ea7c3-125">Icon</span><span class="sxs-lookup"><span data-stu-id="ea7c3-125">Icon</span></span>](icon.md)      | <span data-ttu-id="ea7c3-126">是</span><span class="sxs-lookup"><span data-stu-id="ea7c3-126">Yes</span></span> |  <span data-ttu-id="ea7c3-127">组的图像。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="ea7c3-128">Control</span><span class="sxs-lookup"><span data-stu-id="ea7c3-128">Control</span></span>](#control)    | <span data-ttu-id="ea7c3-129">否</span><span class="sxs-lookup"><span data-stu-id="ea7c3-129">No</span></span> |  <span data-ttu-id="ea7c3-130">代表一个 Control 对象。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-130">Represents a Control object.</span></span> <span data-ttu-id="ea7c3-131">可以是零个或多个。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-131">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="ea7c3-132">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="ea7c3-132">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="ea7c3-133">否</span><span class="sxs-lookup"><span data-stu-id="ea7c3-133">No</span></span> | <span data-ttu-id="ea7c3-134">表示其中一个内置的 Office 控件。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-134">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="ea7c3-135">可以是零个或多个。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-135">Can be zero or more.</span></span> |

### <a name="label"></a><span data-ttu-id="ea7c3-136">标签</span><span class="sxs-lookup"><span data-stu-id="ea7c3-136">Label</span></span>

<span data-ttu-id="ea7c3-137">必需。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-137">Required.</span></span> <span data-ttu-id="ea7c3-138">组的标签。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-138">The label of the group.</span></span> <span data-ttu-id="ea7c3-139">**resid** 属性的长度不能超过 32 个字符，必须设置为 Resources 元素 **中 ShortStrings** 元素 **中 String** 元素的 **id** [属性值。](resources.md)</span><span class="sxs-lookup"><span data-stu-id="ea7c3-139">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="ea7c3-140">Icon</span><span class="sxs-lookup"><span data-stu-id="ea7c3-140">Icon</span></span>

<span data-ttu-id="ea7c3-141">必需。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-141">Required.</span></span> <span data-ttu-id="ea7c3-142">如果选项卡包含大量组，并且程序窗口调整了大小，则可能会改为显示指定的图像。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-142">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="ea7c3-143">控件</span><span class="sxs-lookup"><span data-stu-id="ea7c3-143">Control</span></span>

<span data-ttu-id="ea7c3-144">可选，但如果不存在，则必须至少有一 **个 OfficeControl**。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-144">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="ea7c3-145">有关支持的控件类型的详细信息，请参阅 [Control](control.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-145">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="ea7c3-146">清单中的 **Control** 和 **OfficeControl** 顺序是可互换的，如果有多个元素，它们可以相互交集，但所有元素都必须位于 **Icon** 元素下方。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-146">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

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

### <a name="officecontrol"></a><span data-ttu-id="ea7c3-147">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="ea7c3-147">OfficeControl</span></span>

<span data-ttu-id="ea7c3-148">可选，但如果不存在，则必须至少有一个 **控件**。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-148">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="ea7c3-149">在包含元素的组中包括一个或多个内置 Office `<OfficeControl>` 控件。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-149">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="ea7c3-150">`id`该属性指定内置 Office 控件的 ID。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-150">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="ea7c3-151">若要查找控件的 ID，请参阅["查找控件和控件组的 ID"。](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="ea7c3-151">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="ea7c3-152">清单中的 **Control** 和 **OfficeControl** 顺序是可互换的，如果有多个元素，它们可以相互交集，但所有元素都必须位于 **Icon** 元素下方。</span><span class="sxs-lookup"><span data-stu-id="ea7c3-152">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

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
