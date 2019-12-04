---
title: 清单文件中的 Group 元素
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 35db4829b40078e97fbfc007e2fb552e00875f9c
ms.sourcegitcommit: 164b11b1e9d2ae20b3d816092025b32a9070450f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/04/2019
ms.locfileid: "39818725"
---
# <a name="group-element"></a><span data-ttu-id="9ea23-102">Group 元素</span><span class="sxs-lookup"><span data-stu-id="9ea23-102">Group element</span></span>

<span data-ttu-id="9ea23-p101">在选项卡中定义 UI 控件组在自定义选项卡上，外接程序可以创建最多 10 个组。每个组限制为 6 个控件，不论它显示在哪个选项卡上。外接程序限定到一个自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="9ea23-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="9ea23-106">属性</span><span class="sxs-lookup"><span data-stu-id="9ea23-106">Attributes</span></span>

|  <span data-ttu-id="9ea23-107">属性</span><span class="sxs-lookup"><span data-stu-id="9ea23-107">Attribute</span></span>  |  <span data-ttu-id="9ea23-108">必需</span><span class="sxs-lookup"><span data-stu-id="9ea23-108">Required</span></span>  |  <span data-ttu-id="9ea23-109">说明</span><span class="sxs-lookup"><span data-stu-id="9ea23-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9ea23-110">id</span><span class="sxs-lookup"><span data-stu-id="9ea23-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="9ea23-111">是</span><span class="sxs-lookup"><span data-stu-id="9ea23-111">Yes</span></span>  | <span data-ttu-id="9ea23-112">组的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="9ea23-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="9ea23-113">id attribute</span><span class="sxs-lookup"><span data-stu-id="9ea23-113">id attribute</span></span>

<span data-ttu-id="9ea23-p102">必需。 组的唯一标识符。 是一个最多为 125 个字符的字符串。 该字符串在清单内必须是唯一的，否则组将不能呈现。</span><span class="sxs-lookup"><span data-stu-id="9ea23-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="9ea23-118">子元素</span><span class="sxs-lookup"><span data-stu-id="9ea23-118">Child elements</span></span>
|  <span data-ttu-id="9ea23-119">元素</span><span class="sxs-lookup"><span data-stu-id="9ea23-119">Element</span></span> |  <span data-ttu-id="9ea23-120">必需</span><span class="sxs-lookup"><span data-stu-id="9ea23-120">Required</span></span>  |  <span data-ttu-id="9ea23-121">说明</span><span class="sxs-lookup"><span data-stu-id="9ea23-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9ea23-122">Label</span><span class="sxs-lookup"><span data-stu-id="9ea23-122">Label</span></span>](#label)      | <span data-ttu-id="9ea23-123">是</span><span class="sxs-lookup"><span data-stu-id="9ea23-123">Yes</span></span> |  <span data-ttu-id="9ea23-124">CustomTab 或组的标签。</span><span class="sxs-lookup"><span data-stu-id="9ea23-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="9ea23-125">Icon</span><span class="sxs-lookup"><span data-stu-id="9ea23-125">Icon</span></span>](icon.md)      | <span data-ttu-id="9ea23-126">是</span><span class="sxs-lookup"><span data-stu-id="9ea23-126">Yes</span></span> |  <span data-ttu-id="9ea23-127">组的图像。</span><span class="sxs-lookup"><span data-stu-id="9ea23-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="9ea23-128">Control</span><span class="sxs-lookup"><span data-stu-id="9ea23-128">Control</span></span>](#control)    | <span data-ttu-id="9ea23-129">是</span><span class="sxs-lookup"><span data-stu-id="9ea23-129">Yes</span></span> |  <span data-ttu-id="9ea23-130">一个或多个控件对象的集合。</span><span class="sxs-lookup"><span data-stu-id="9ea23-130">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="9ea23-131">标签</span><span class="sxs-lookup"><span data-stu-id="9ea23-131">Label</span></span> 

<span data-ttu-id="9ea23-p103">必需。组的标签。 **resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="9ea23-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="9ea23-135">Icon</span><span class="sxs-lookup"><span data-stu-id="9ea23-135">Icon</span></span>

<span data-ttu-id="9ea23-136">必需。</span><span class="sxs-lookup"><span data-stu-id="9ea23-136">Required.</span></span> <span data-ttu-id="9ea23-137">如果某个选项卡包含大量组，并且该程序窗口已调整大小，则会改为显示指定的图像。</span><span class="sxs-lookup"><span data-stu-id="9ea23-137">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="9ea23-138">Control</span><span class="sxs-lookup"><span data-stu-id="9ea23-138">Control</span></span>
<span data-ttu-id="9ea23-139">一个组需要至少一个控件。</span><span class="sxs-lookup"><span data-stu-id="9ea23-139">A group requires at least one control.</span></span> <span data-ttu-id="9ea23-140">有关受支持的控件类型的详细信息，请参阅[Control](control.md)元素。</span><span class="sxs-lookup"><span data-stu-id="9ea23-140">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span>

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
