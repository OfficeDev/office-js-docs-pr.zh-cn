---
title: 清单文件中的 Group 元素
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: ad1a566e259188ed20032bc5a3004736474e1f01
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670130"
---
# <a name="group-element"></a><span data-ttu-id="6a789-102">Group 元素</span><span class="sxs-lookup"><span data-stu-id="6a789-102">Group element</span></span>

<span data-ttu-id="6a789-p101">在选项卡中定义 UI 控件组在自定义选项卡上，外接程序可以创建最多 10 个组。每个组限制为 6 个控件，不论它显示在哪个选项卡上。外接程序限定到一个自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="6a789-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="6a789-106">属性</span><span class="sxs-lookup"><span data-stu-id="6a789-106">Attributes</span></span>

|  <span data-ttu-id="6a789-107">属性</span><span class="sxs-lookup"><span data-stu-id="6a789-107">Attribute</span></span>  |  <span data-ttu-id="6a789-108">必需</span><span class="sxs-lookup"><span data-stu-id="6a789-108">Required</span></span>  |  <span data-ttu-id="6a789-109">说明</span><span class="sxs-lookup"><span data-stu-id="6a789-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="6a789-110">id</span><span class="sxs-lookup"><span data-stu-id="6a789-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="6a789-111">是</span><span class="sxs-lookup"><span data-stu-id="6a789-111">Yes</span></span>  | <span data-ttu-id="6a789-112">组的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="6a789-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="6a789-113">id attribute</span><span class="sxs-lookup"><span data-stu-id="6a789-113">id attribute</span></span>

<span data-ttu-id="6a789-p102">必需。 组的唯一标识符。 是一个最多为 125 个字符的字符串。 该字符串在清单内必须是唯一的，否则组将不能呈现。</span><span class="sxs-lookup"><span data-stu-id="6a789-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="6a789-118">子元素</span><span class="sxs-lookup"><span data-stu-id="6a789-118">Child elements</span></span>
|  <span data-ttu-id="6a789-119">元素</span><span class="sxs-lookup"><span data-stu-id="6a789-119">Element</span></span> |  <span data-ttu-id="6a789-120">必需</span><span class="sxs-lookup"><span data-stu-id="6a789-120">Required</span></span>  |  <span data-ttu-id="6a789-121">说明</span><span class="sxs-lookup"><span data-stu-id="6a789-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="6a789-122">Label</span><span class="sxs-lookup"><span data-stu-id="6a789-122">Label</span></span>](#label)      | <span data-ttu-id="6a789-123">是</span><span class="sxs-lookup"><span data-stu-id="6a789-123">Yes</span></span> |  <span data-ttu-id="6a789-124">CustomTab 或组的标签。</span><span class="sxs-lookup"><span data-stu-id="6a789-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="6a789-125">Control</span><span class="sxs-lookup"><span data-stu-id="6a789-125">Control</span></span>](#control)    | <span data-ttu-id="6a789-126">是</span><span class="sxs-lookup"><span data-stu-id="6a789-126">Yes</span></span> |  <span data-ttu-id="6a789-127">一个或多个控件对象的集合。</span><span class="sxs-lookup"><span data-stu-id="6a789-127">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="6a789-128">标签</span><span class="sxs-lookup"><span data-stu-id="6a789-128">Label</span></span> 

<span data-ttu-id="6a789-p103">必需。组的标签。 **resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="6a789-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="6a789-132">Control</span><span class="sxs-lookup"><span data-stu-id="6a789-132">Control</span></span>
<span data-ttu-id="6a789-133">一个组需要至少一个控件。</span><span class="sxs-lookup"><span data-stu-id="6a789-133">A group requires at least one control.</span></span> <span data-ttu-id="6a789-134">有关受支持的控件类型的详细信息，请参阅[Control](control.md)元素。</span><span class="sxs-lookup"><span data-stu-id="6a789-134">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```
