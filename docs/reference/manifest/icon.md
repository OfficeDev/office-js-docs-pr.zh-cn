---
title: 清单文件中的 Icon 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 45f3dcda8e74430cf70aa765efc6b3aae0e2b448
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450616"
---
# <a name="icon-element"></a><span data-ttu-id="9213d-102">Icon 元素</span><span class="sxs-lookup"><span data-stu-id="9213d-102">Icon element</span></span>

<span data-ttu-id="9213d-103">定义“[按钮](control.md#button-control)”或“[菜单](control.md#menu-dropdown-button-controls)”控件的 **Image** 元素。</span><span class="sxs-lookup"><span data-stu-id="9213d-103">Defines **Image** elements for [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="9213d-104">属性</span><span class="sxs-lookup"><span data-stu-id="9213d-104">Attributes</span></span>

|  <span data-ttu-id="9213d-105">属性</span><span class="sxs-lookup"><span data-stu-id="9213d-105">Attribute</span></span>  |  <span data-ttu-id="9213d-106">必需</span><span class="sxs-lookup"><span data-stu-id="9213d-106">Required</span></span>  |  <span data-ttu-id="9213d-107">说明</span><span class="sxs-lookup"><span data-stu-id="9213d-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="9213d-108">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="9213d-108">**xsi:type**</span></span>  |  <span data-ttu-id="9213d-109">否</span><span class="sxs-lookup"><span data-stu-id="9213d-109">No</span></span>  | <span data-ttu-id="9213d-p101">正在定义的图标类型。这仅适用于移动外形规格中的图标。[MobileFormFactor](mobileformfactor.md) 元素中所包含的 **Icon** 元素必须将此属性设置为 `bt:MobileIconList`。</span><span class="sxs-lookup"><span data-stu-id="9213d-p101">The type of icon being defined. This is only applicable to icons in mobile form factors. **Icon** elements contained within a [MobileFormFactor](mobileformfactor.md) element must have this attribute set to `bt:MobileIconList`.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="9213d-113">子元素</span><span class="sxs-lookup"><span data-stu-id="9213d-113">Child elements</span></span>

|  <span data-ttu-id="9213d-114">元素</span><span class="sxs-lookup"><span data-stu-id="9213d-114">Element</span></span> |  <span data-ttu-id="9213d-115">必需</span><span class="sxs-lookup"><span data-stu-id="9213d-115">Required</span></span>  |  <span data-ttu-id="9213d-116">说明</span><span class="sxs-lookup"><span data-stu-id="9213d-116">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9213d-117">Image</span><span class="sxs-lookup"><span data-stu-id="9213d-117">Image</span></span>](#image)        | <span data-ttu-id="9213d-118">是</span><span class="sxs-lookup"><span data-stu-id="9213d-118">Yes</span></span> |   <span data-ttu-id="9213d-119">要使用的图像的 resid</span><span class="sxs-lookup"><span data-stu-id="9213d-119">resid of an image to use</span></span>         |

### <a name="image"></a><span data-ttu-id="9213d-120">图像</span><span class="sxs-lookup"><span data-stu-id="9213d-120">Image</span></span>

<span data-ttu-id="9213d-p102">按钮的图像。**resid** 属性必须设置为 **Images** 元素（位于 **Resources** 元素）中 **Image** 元素的 [id](resources.md) 属性的值。**size** 属性指示图像的大小，以像素为单位。有三个图像大小是必需的（16、32 和 80 像素），此外还支持五个其他大小（20、24、40、48 和 64 像素）。|</span><span class="sxs-lookup"><span data-stu-id="9213d-p102">An image for the button. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](resources.md) element. The **size** attribute indicates the size in pixels of the image. Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|</span></span>

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

## <a name="additional-requirements-for-mobile-form-factors"></a><span data-ttu-id="9213d-125">移动外形规格的其他要求</span><span class="sxs-lookup"><span data-stu-id="9213d-125">Additional requirements for mobile form factors</span></span>

<span data-ttu-id="9213d-p103">当父 **Icon** 元素是 [MobileFormFactor](mobileformfactor.md) 元素的后代时，所要求的最小大小会略有不同。清单必须至少提供 25、32 和 48 像素大小。所提供的每个大小必须出现三次，并将 `scale` 属性设置为 `1`、`2` 或 `3`。</span><span class="sxs-lookup"><span data-stu-id="9213d-p103">When the parent **Icon** element is a descendant of a [MobileFormFactor](mobileformfactor.md) element, the minimum required sizes are slightly different. The manifest must minimally provide 25, 32, and 48 pixel sizes. Each size provided must appear three times, with a `scale` attribute set to `1`, `2`, or `3`.</span></span>

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
