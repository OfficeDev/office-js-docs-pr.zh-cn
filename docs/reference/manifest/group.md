---
title: 清单文件中的 Group 元素
description: 定义选项卡中的一组 UI 控件。
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 6fe07497e98bd77aad7ad296850a0b9f9e9bf9a4
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718179"
---
# <a name="group-element"></a><span data-ttu-id="cb84e-103">Group 元素</span><span class="sxs-lookup"><span data-stu-id="cb84e-103">Group element</span></span>

<span data-ttu-id="cb84e-p101">在选项卡中定义 UI 控件组在自定义选项卡上，外接程序可以创建最多 10 个组。每个组限制为 6 个控件，不论它显示在哪个选项卡上。外接程序限定到一个自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="cb84e-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="cb84e-107">属性</span><span class="sxs-lookup"><span data-stu-id="cb84e-107">Attributes</span></span>

|  <span data-ttu-id="cb84e-108">属性</span><span class="sxs-lookup"><span data-stu-id="cb84e-108">Attribute</span></span>  |  <span data-ttu-id="cb84e-109">必需</span><span class="sxs-lookup"><span data-stu-id="cb84e-109">Required</span></span>  |  <span data-ttu-id="cb84e-110">说明</span><span class="sxs-lookup"><span data-stu-id="cb84e-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cb84e-111">id</span><span class="sxs-lookup"><span data-stu-id="cb84e-111">id</span></span>](#id-attribute)  |  <span data-ttu-id="cb84e-112">是</span><span class="sxs-lookup"><span data-stu-id="cb84e-112">Yes</span></span>  | <span data-ttu-id="cb84e-113">组的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="cb84e-113">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="cb84e-114">id attribute</span><span class="sxs-lookup"><span data-stu-id="cb84e-114">id attribute</span></span>

<span data-ttu-id="cb84e-p102">必需。 组的唯一标识符。 是一个最多为 125 个字符的字符串。 该字符串在清单内必须是唯一的，否则组将不能呈现。</span><span class="sxs-lookup"><span data-stu-id="cb84e-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="cb84e-119">子元素</span><span class="sxs-lookup"><span data-stu-id="cb84e-119">Child elements</span></span>
|  <span data-ttu-id="cb84e-120">元素</span><span class="sxs-lookup"><span data-stu-id="cb84e-120">Element</span></span> |  <span data-ttu-id="cb84e-121">必需</span><span class="sxs-lookup"><span data-stu-id="cb84e-121">Required</span></span>  |  <span data-ttu-id="cb84e-122">说明</span><span class="sxs-lookup"><span data-stu-id="cb84e-122">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cb84e-123">Label</span><span class="sxs-lookup"><span data-stu-id="cb84e-123">Label</span></span>](#label)      | <span data-ttu-id="cb84e-124">是</span><span class="sxs-lookup"><span data-stu-id="cb84e-124">Yes</span></span> |  <span data-ttu-id="cb84e-125">CustomTab 或组的标签。</span><span class="sxs-lookup"><span data-stu-id="cb84e-125">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="cb84e-126">Icon</span><span class="sxs-lookup"><span data-stu-id="cb84e-126">Icon</span></span>](icon.md)      | <span data-ttu-id="cb84e-127">是</span><span class="sxs-lookup"><span data-stu-id="cb84e-127">Yes</span></span> |  <span data-ttu-id="cb84e-128">组的图像。</span><span class="sxs-lookup"><span data-stu-id="cb84e-128">The image for a group.</span></span>  |
|  [<span data-ttu-id="cb84e-129">Control</span><span class="sxs-lookup"><span data-stu-id="cb84e-129">Control</span></span>](#control)    | <span data-ttu-id="cb84e-130">是</span><span class="sxs-lookup"><span data-stu-id="cb84e-130">Yes</span></span> |  <span data-ttu-id="cb84e-131">一个或多个控件对象的集合。</span><span class="sxs-lookup"><span data-stu-id="cb84e-131">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="cb84e-132">标签</span><span class="sxs-lookup"><span data-stu-id="cb84e-132">Label</span></span> 

<span data-ttu-id="cb84e-133">必需。</span><span class="sxs-lookup"><span data-stu-id="cb84e-133">Required.</span></span> <span data-ttu-id="cb84e-134">组的标签。</span><span class="sxs-lookup"><span data-stu-id="cb84e-134">The label of the group.</span></span> <span data-ttu-id="cb84e-135">**Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="cb84e-135">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="cb84e-136">Icon</span><span class="sxs-lookup"><span data-stu-id="cb84e-136">Icon</span></span>

<span data-ttu-id="cb84e-137">必需。</span><span class="sxs-lookup"><span data-stu-id="cb84e-137">Required.</span></span> <span data-ttu-id="cb84e-138">如果某个选项卡包含大量组，并且该程序窗口已调整大小，则会改为显示指定的图像。</span><span class="sxs-lookup"><span data-stu-id="cb84e-138">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="cb84e-139">Control</span><span class="sxs-lookup"><span data-stu-id="cb84e-139">Control</span></span>
<span data-ttu-id="cb84e-140">一个组需要至少一个控件。</span><span class="sxs-lookup"><span data-stu-id="cb84e-140">A group requires at least one control.</span></span> <span data-ttu-id="cb84e-141">有关受支持的控件类型的详细信息，请参阅[Control](control.md)元素。</span><span class="sxs-lookup"><span data-stu-id="cb84e-141">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span>

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
