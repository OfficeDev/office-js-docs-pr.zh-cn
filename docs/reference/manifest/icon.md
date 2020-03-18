---
title: 清单文件中的 Icon 元素
description: 定义“按钮”或“菜单”控件的 Image 元素。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a17f43352b306850c853c230f6a3617eb165ca14
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718088"
---
# <a name="icon-element"></a><span data-ttu-id="de05d-103">Icon 元素</span><span class="sxs-lookup"><span data-stu-id="de05d-103">Icon element</span></span>

<span data-ttu-id="de05d-104">定义“[按钮](control.md#button-control)”或“[菜单](control.md#menu-dropdown-button-controls)”控件的 **Image** 元素。</span><span class="sxs-lookup"><span data-stu-id="de05d-104">Defines **Image** elements for [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="de05d-105">属性</span><span class="sxs-lookup"><span data-stu-id="de05d-105">Attributes</span></span>

|  <span data-ttu-id="de05d-106">属性</span><span class="sxs-lookup"><span data-stu-id="de05d-106">Attribute</span></span>  |  <span data-ttu-id="de05d-107">必需</span><span class="sxs-lookup"><span data-stu-id="de05d-107">Required</span></span>  |  <span data-ttu-id="de05d-108">说明</span><span class="sxs-lookup"><span data-stu-id="de05d-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="de05d-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="de05d-109">**xsi:type**</span></span>  |  <span data-ttu-id="de05d-110">否</span><span class="sxs-lookup"><span data-stu-id="de05d-110">No</span></span>  | <span data-ttu-id="de05d-p101">正在定义的图标类型。这仅适用于移动外形规格中的图标。[MobileFormFactor](mobileformfactor.md) 元素中所包含的 **Icon** 元素必须将此属性设置为 `bt:MobileIconList`。</span><span class="sxs-lookup"><span data-stu-id="de05d-p101">The type of icon being defined. This is only applicable to icons in mobile form factors. **Icon** elements contained within a [MobileFormFactor](mobileformfactor.md) element must have this attribute set to `bt:MobileIconList`.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="de05d-114">子元素</span><span class="sxs-lookup"><span data-stu-id="de05d-114">Child elements</span></span>

|  <span data-ttu-id="de05d-115">元素</span><span class="sxs-lookup"><span data-stu-id="de05d-115">Element</span></span> |  <span data-ttu-id="de05d-116">必需</span><span class="sxs-lookup"><span data-stu-id="de05d-116">Required</span></span>  |  <span data-ttu-id="de05d-117">说明</span><span class="sxs-lookup"><span data-stu-id="de05d-117">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="de05d-118">Image</span><span class="sxs-lookup"><span data-stu-id="de05d-118">Image</span></span>](#image)        | <span data-ttu-id="de05d-119">是</span><span class="sxs-lookup"><span data-stu-id="de05d-119">Yes</span></span> |   <span data-ttu-id="de05d-120">要使用的图像的 resid</span><span class="sxs-lookup"><span data-stu-id="de05d-120">resid of an image to use</span></span>         |

### <a name="image"></a><span data-ttu-id="de05d-121">图像</span><span class="sxs-lookup"><span data-stu-id="de05d-121">Image</span></span>

<span data-ttu-id="de05d-122">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="de05d-122">An image for the button.</span></span> <span data-ttu-id="de05d-123">**resid** 属性必须设置为 **Images** 元素（位于 [Resources](resources.md) 元素）中 **Image** 元素的 **id** 属性的值。</span><span class="sxs-lookup"><span data-stu-id="de05d-123">The **resid** attribute must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](resources.md) element.</span></span> <span data-ttu-id="de05d-124">The **size** attribute indicates the size in pixels of the image.</span><span class="sxs-lookup"><span data-stu-id="de05d-124">The **size** attribute indicates the size in pixels of the image.</span></span> <span data-ttu-id="de05d-125">需要三个图像大小（16、32和80像素），而支持五个其他大小（20、24、40、48和64像素）。 |</span><span class="sxs-lookup"><span data-stu-id="de05d-125">Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|</span></span>

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

## <a name="additional-requirements-for-mobile-form-factors"></a><span data-ttu-id="de05d-126">移动外形规格的其他要求</span><span class="sxs-lookup"><span data-stu-id="de05d-126">Additional requirements for mobile form factors</span></span>

<span data-ttu-id="de05d-p103">当父 **Icon** 元素是 [MobileFormFactor](mobileformfactor.md) 元素的后代时，所要求的最小大小会略有不同。清单必须至少提供 25、32 和 48 像素大小。所提供的每个大小必须出现三次，并将 `scale` 属性设置为 `1`、`2` 或 `3`。</span><span class="sxs-lookup"><span data-stu-id="de05d-p103">When the parent **Icon** element is a descendant of a [MobileFormFactor](mobileformfactor.md) element, the minimum required sizes are slightly different. The manifest must minimally provide 25, 32, and 48 pixel sizes. Each size provided must appear three times, with a `scale` attribute set to `1`, `2`, or `3`.</span></span>

```xml
<Icon xsi:type="bt:MobileIconList">
  <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
  <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
  <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
  <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
  <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
  <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
  <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
  <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
  <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
</Icon>
```
