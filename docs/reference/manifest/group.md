---
title: 清单文件中的 Group 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7cc1f4c398eeb013eb6033b207b395466f7d72ca
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450707"
---
# <a name="group-element"></a><span data-ttu-id="cddf5-102">Group 元素</span><span class="sxs-lookup"><span data-stu-id="cddf5-102">Group element</span></span>

<span data-ttu-id="cddf5-p101">在选项卡中定义 UI 控件组在自定义选项卡上，外接程序可以创建最多 10 个组。每个组限制为 6 个控件，不论它显示在哪个选项卡上。外接程序限定到一个自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="cddf5-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="cddf5-106">属性</span><span class="sxs-lookup"><span data-stu-id="cddf5-106">Attributes</span></span>

|  <span data-ttu-id="cddf5-107">属性</span><span class="sxs-lookup"><span data-stu-id="cddf5-107">Attribute</span></span>  |  <span data-ttu-id="cddf5-108">必需</span><span class="sxs-lookup"><span data-stu-id="cddf5-108">Required</span></span>  |  <span data-ttu-id="cddf5-109">说明</span><span class="sxs-lookup"><span data-stu-id="cddf5-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cddf5-110">id</span><span class="sxs-lookup"><span data-stu-id="cddf5-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="cddf5-111">是</span><span class="sxs-lookup"><span data-stu-id="cddf5-111">Yes</span></span>  | <span data-ttu-id="cddf5-112">组的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="cddf5-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="cddf5-113">id attribute</span><span class="sxs-lookup"><span data-stu-id="cddf5-113">id attribute</span></span>

<span data-ttu-id="cddf5-p102">必需。 组的唯一标识符。 是一个最多为 125 个字符的字符串。 该字符串在清单内必须是唯一的，否则组将不能呈现。</span><span class="sxs-lookup"><span data-stu-id="cddf5-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="cddf5-118">子元素</span><span class="sxs-lookup"><span data-stu-id="cddf5-118">Child elements</span></span>
|  <span data-ttu-id="cddf5-119">元素</span><span class="sxs-lookup"><span data-stu-id="cddf5-119">Element</span></span> |  <span data-ttu-id="cddf5-120">必需</span><span class="sxs-lookup"><span data-stu-id="cddf5-120">Required</span></span>  |  <span data-ttu-id="cddf5-121">说明</span><span class="sxs-lookup"><span data-stu-id="cddf5-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cddf5-122">Label</span><span class="sxs-lookup"><span data-stu-id="cddf5-122">Label</span></span>](#label)      | <span data-ttu-id="cddf5-123">是</span><span class="sxs-lookup"><span data-stu-id="cddf5-123">Yes</span></span> |  <span data-ttu-id="cddf5-124">CustomTab 或组的标签。</span><span class="sxs-lookup"><span data-stu-id="cddf5-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="cddf5-125">Control</span><span class="sxs-lookup"><span data-stu-id="cddf5-125">Control</span></span>](#control)    | <span data-ttu-id="cddf5-126">是</span><span class="sxs-lookup"><span data-stu-id="cddf5-126">Yes</span></span> |  <span data-ttu-id="cddf5-127">一个或多个控件对象的集合。</span><span class="sxs-lookup"><span data-stu-id="cddf5-127">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="cddf5-128">标签</span><span class="sxs-lookup"><span data-stu-id="cddf5-128">Label</span></span> 

<span data-ttu-id="cddf5-p103">必需。组的标签。 **resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="cddf5-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="cddf5-132">Control</span><span class="sxs-lookup"><span data-stu-id="cddf5-132">Control</span></span>
<span data-ttu-id="cddf5-133">一个组需要至少一个控件。</span><span class="sxs-lookup"><span data-stu-id="cddf5-133">A group requires at least one control.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```
