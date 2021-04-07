---
title: Office 外接程序的单声道样式图标指南
description: 有关在 Office 外接程序中使用单声道样式图标的指南。
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: b74b89b2d622a6166fa111ef92bd8b2fffe79f8a
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604671"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a><span data-ttu-id="5ed2c-103">Office 外接程序的单声道样式图标指南</span><span class="sxs-lookup"><span data-stu-id="5ed2c-103">Monoline style icon guidelines for Office Add-ins</span></span>

<span data-ttu-id="5ed2c-104">单声道样式图标在 Office 应用中使用。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-104">Monoline style iconography are used in Office apps.</span></span> <span data-ttu-id="5ed2c-105">如果您希望您的图标与非订阅 Office 2013+的新鲜样式匹配，请参阅 Office 外接程序的新鲜样式 [图标指南](add-in-icons-fresh.md)。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-105">If you would prefer that your icons match the Fresh style of non-subscription Office 2013+, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).</span></span>

## <a name="office-monoline-visual-style"></a><span data-ttu-id="5ed2c-106">Office 单声道视觉样式</span><span class="sxs-lookup"><span data-stu-id="5ed2c-106">Office Monoline visual style</span></span>

<span data-ttu-id="5ed2c-107">Monoline 样式的目标是具有一致、清晰且可访问的图标，以通过简单的视觉效果传达操作和功能，确保图标可供所有用户访问，并且具有与 Windows 中其他位置使用的样式一致的样式。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-107">The goal of the Monoline style to have consistent, clear, and accessible iconography to communicate action and features with simple visuals, ensure the icons are accessible to all users, and have a style that is consistent with those used elsewhere in Windows.</span></span>

<span data-ttu-id="5ed2c-108">以下指南适用于第三方开发人员，他们希望为功能创建图标，这些图标与已有的 Office 产品图标一致。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-108">The following guidelines are for 3rd party developers who want to create icons for features that will be consistent with the icons already present Office products.</span></span>

### <a name="design-principles"></a><span data-ttu-id="5ed2c-109">设计原则</span><span class="sxs-lookup"><span data-stu-id="5ed2c-109">Design principles</span></span>

- <span data-ttu-id="5ed2c-110">简单、干净、清晰。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-110">Simple, clean, clear.</span></span>
- <span data-ttu-id="5ed2c-111">仅包含必要的元素。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-111">Contain only necessary elements.</span></span>
- <span data-ttu-id="5ed2c-112">受 Windows 图标样式启发。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-112">Inspired by Windows icon style.</span></span>
- <span data-ttu-id="5ed2c-113">可供所有用户访问。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-113">Accessible to all users.</span></span>

#### <a name="conveying-meaning"></a><span data-ttu-id="5ed2c-114">传达含义</span><span class="sxs-lookup"><span data-stu-id="5ed2c-114">Conveying meaning</span></span>

- <span data-ttu-id="5ed2c-115">使用描述性元素（如页面）表示文档或信封来表示邮件。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-115">Use descriptive elements such as a page to represent a document or an envelope to represent mail.</span></span>
- <span data-ttu-id="5ed2c-116">使用相同的元素来表示同一概念，即邮件始终由信封而不是标记表示。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-116">Use the same element to represent the same concept, i.e., mail is always represented by an envelope, not a stamp.</span></span>
- <span data-ttu-id="5ed2c-117">在概念开发过程中使用核心隐喻。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-117">Use a core metaphor during concept development.</span></span>

#### <a name="reduction-of-elements"></a><span data-ttu-id="5ed2c-118">减少元素</span><span class="sxs-lookup"><span data-stu-id="5ed2c-118">Reduction of Elements</span></span>

- <span data-ttu-id="5ed2c-119">将图标缩小到核心含义，仅使用对隐喻至关重要的元素。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-119">Reduce the icon to its core meaning, using only elements that are essential to the metaphor.</span></span>
- <span data-ttu-id="5ed2c-120">将图标中的元素数限制为两个，无论图标大小如何。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-120">Limit the number of elements in an icon to two, regardless of icon size.</span></span>

#### <a name="consistency"></a><span data-ttu-id="5ed2c-121">一致性</span><span class="sxs-lookup"><span data-stu-id="5ed2c-121">Consistency</span></span>

<span data-ttu-id="5ed2c-122">图标的大小、排列和颜色应保持一致。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-122">Sizes, arrangement, and color of icons should be consistent.</span></span>

#### <a name="styling"></a><span data-ttu-id="5ed2c-123">样式设置</span><span class="sxs-lookup"><span data-stu-id="5ed2c-123">Styling</span></span>

##### <a name="perspective"></a><span data-ttu-id="5ed2c-124">Perspective</span><span class="sxs-lookup"><span data-stu-id="5ed2c-124">Perspective</span></span>

<span data-ttu-id="5ed2c-125">默认情况下，单声道图标是向前的。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-125">Monoline icons are forward-facing by default.</span></span> <span data-ttu-id="5ed2c-126">允许使用需要透视和/或旋转的某些元素，如立方体，但例外应保持最少。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-126">Certain elements that require perspective and/or rotation, such as a cube, are allowed, but exceptions should be kept to a minimum.</span></span>

##### <a name="embellishment"></a><span data-ttu-id="5ed2c-127">装饰</span><span class="sxs-lookup"><span data-stu-id="5ed2c-127">Embellishment</span></span>

<span data-ttu-id="5ed2c-128">单声道是一种简洁的样式。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-128">Monoline is a clean minimal style.</span></span> <span data-ttu-id="5ed2c-129">所有内容都使用平面颜色，这意味着没有渐变、纹理或光源。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-129">Everything uses flat color, which means there are no gradients, textures, or light sources.</span></span>

## <a name="designing"></a><span data-ttu-id="5ed2c-130">正在设计</span><span class="sxs-lookup"><span data-stu-id="5ed2c-130">Designing</span></span>

### <a name="sizes"></a><span data-ttu-id="5ed2c-131">大小</span><span class="sxs-lookup"><span data-stu-id="5ed2c-131">Sizes</span></span>

<span data-ttu-id="5ed2c-132">我们建议你生成所有这些大小的每个图标，以支持高 DPI 设备。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-132">We recommend that you produce each icon in all these sizes to support high DPI devices.</span></span> <span data-ttu-id="5ed2c-133">绝对 *必需的大小* 为 16 像素、20 像素和 32 像素，因为大小为 100%。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-133">The absolutely *required* sizes are 16 px, 20 px, and 32 px, as those are the 100% sizes.</span></span>

<span data-ttu-id="5ed2c-134">**16 像素、20 像素、24 像素、32 像素、40 像素、48 像素、64 像素、80 像素、96 像素**</span><span class="sxs-lookup"><span data-stu-id="5ed2c-134">**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5ed2c-135">有关作为加载项代表图标的图像，请参阅在 [AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) 和 Office 内创建有效列表，了解大小和其他要求。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-135">For an image that is your add-in's representative icon, see [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) for size and other requirements.</span></span>

### <a name="layout"></a><span data-ttu-id="5ed2c-136">布局</span><span class="sxs-lookup"><span data-stu-id="5ed2c-136">Layout</span></span>

<span data-ttu-id="5ed2c-137">下面是一个包含修饰符的图标布局示例。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-137">The following is an example of icon layout with a modifier.</span></span>

![右下角具有修饰符的图标关系图](../images/monolineicon1.png)  ![包含添加网格背景的相同图标的图示，以及基本、修饰符、填充和剪切的标注](../images/monolineicon2.png)

#### <a name="elements"></a><span data-ttu-id="5ed2c-140">元素</span><span class="sxs-lookup"><span data-stu-id="5ed2c-140">Elements</span></span>

- <span data-ttu-id="5ed2c-141">**基本**：图标表示的主要概念。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-141">**Base**: The main concept that the icon represents.</span></span> <span data-ttu-id="5ed2c-142">这通常是图标所需的唯一视觉对象，但有时可以使用辅助元素（修饰符）增强主要概念。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-142">This is usually the only visual needed for the icon, but sometimes the main concept can be enhanced with a secondary element, a modifier.</span></span>

- <span data-ttu-id="5ed2c-143">**修饰符** 覆盖基本元素的任何元素;即，通常表示操作或状态的修饰符。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-143">**Modifier** Any element that overlays the base; that is, a modifier that typically represents an action or a status.</span></span> <span data-ttu-id="5ed2c-144">它通过充当添加、更改或描述符来修改基本元素。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-144">It modifies the base element by acting as an addition, alteration, or a descriptor.</span></span>

![包含已调用基本和修饰符区域网格的关系图](../images/monolineicon3.png)

### <a name="construction"></a><span data-ttu-id="5ed2c-146">建造</span><span class="sxs-lookup"><span data-stu-id="5ed2c-146">Construction</span></span>

#### <a name="element-placement"></a><span data-ttu-id="5ed2c-147">元素放置</span><span class="sxs-lookup"><span data-stu-id="5ed2c-147">Element placement</span></span>

<span data-ttu-id="5ed2c-148">基元素放置在填充内图标的中心。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-148">Base elements are placed in the center of the icon within the padding.</span></span> <span data-ttu-id="5ed2c-149">如果无法完全居中放置，则基数应位于右上方。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-149">If it can't be placed perfectly centered, then the base should err to the top right.</span></span> <span data-ttu-id="5ed2c-150">在下面的示例中，图标完全居中。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-150">In the following example, the icon is perfectly centered.</span></span>

![显示完全居中的图标的关系图](../images/monolineicon4.png)

<span data-ttu-id="5ed2c-152">在下面的示例中，图标在左侧出错。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-152">In the following example, the icon is erring to the left.</span></span>

![显示左错误 1 像素的图标的图表](../images/monolineicon5.png)

<span data-ttu-id="5ed2c-154">修饰符几乎总是放置在图标画布的右下角。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-154">Modifiers are almost always placed in the bottom right corner of the icon canvas.</span></span> <span data-ttu-id="5ed2c-155">在极少数情况下，修饰符放置在不同的角。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-155">In some rare cases, modifiers are placed in a different corner.</span></span> <span data-ttu-id="5ed2c-156">例如，如果修改器无法识别右下角的基元素，请考虑将其放在左上角。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-156">For example, if the base element would be unrecognizable with the modifier in the bottom right corner, then consider placing it in the upper left corner.</span></span>

![显示四个图标的图示，其中修饰符位于右下角，一个图标的修饰符位于左上角](../images/monolineicon6.png)

#### <a name="padding"></a><span data-ttu-id="5ed2c-158">Padding</span><span class="sxs-lookup"><span data-stu-id="5ed2c-158">Padding</span></span>

<span data-ttu-id="5ed2c-159">每个大小图标在图标周围都有指定数量的填充。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-159">Each size icon has a specified amount of padding around the icon.</span></span> <span data-ttu-id="5ed2c-160">基本元素保留在填充内，但修饰符应向上扩展到画布边缘，在填充之外扩展到图标边框的边缘。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-160">The base element stays within the padding, but the modifier should butt up to the edge of the canvas, extending outside of the padding to the edge of the icon border.</span></span> <span data-ttu-id="5ed2c-161">下图显示了用于每个图标大小的推荐填充。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-161">The following images show the recommended padding to use for each of the icon sizes.</span></span>

|<span data-ttu-id="5ed2c-162">**16px**</span><span class="sxs-lookup"><span data-stu-id="5ed2c-162">**16px**</span></span>|<span data-ttu-id="5ed2c-163">**20px**</span><span class="sxs-lookup"><span data-stu-id="5ed2c-163">**20px**</span></span>|<span data-ttu-id="5ed2c-164">**24px**</span><span class="sxs-lookup"><span data-stu-id="5ed2c-164">**24px**</span></span>|<span data-ttu-id="5ed2c-165">**32px**</span><span class="sxs-lookup"><span data-stu-id="5ed2c-165">**32px**</span></span>|<span data-ttu-id="5ed2c-166">**40px**</span><span class="sxs-lookup"><span data-stu-id="5ed2c-166">**40px**</span></span>|<span data-ttu-id="5ed2c-167">**48px**</span><span class="sxs-lookup"><span data-stu-id="5ed2c-167">**48px**</span></span>|<span data-ttu-id="5ed2c-168">**64px**</span><span class="sxs-lookup"><span data-stu-id="5ed2c-168">**64px**</span></span>|<span data-ttu-id="5ed2c-169">**80px**</span><span class="sxs-lookup"><span data-stu-id="5ed2c-169">**80px**</span></span>|<span data-ttu-id="5ed2c-170">**96px**</span><span class="sxs-lookup"><span data-stu-id="5ed2c-170">**96px**</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![具有 0px 填充的 16 像素图标](../images/monolineicon7.png)|![具有 1px 填充的 20 像素图标](../images/monolineicon8.png)|![具有 1px 填充的 24 像素图标](../images/monolineicon9.png)|![具有 2px 填充的 32 像素图标](../images/monolineicon10.png)|![具有 2px 填充的 40 像素图标](../images/monolineicon11.png)|![具有 3px 填充的 48 像素图标](../images/monolineicon12.png)|![具有 4px 填充的 64 像素图标](../images/monolineicon13.png)|![具有 5px 填充的 80 像素图标](../images/monolineicon14.png)|![具有 6px 填充的 96 像素图标](../images/monolineicon15.png)|

#### <a name="line-weights"></a><span data-ttu-id="5ed2c-180">线条粗细</span><span class="sxs-lookup"><span data-stu-id="5ed2c-180">Line weights</span></span>

<span data-ttu-id="5ed2c-181">单声道是线条和轮廓形状的样式控制。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-181">Monoline is a style dominated by line and outlined shapes.</span></span> <span data-ttu-id="5ed2c-182">根据你生成图标的大小，应该使用以下行权重。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-182">Depending on what size you are producing the icon should use the following line weights.</span></span>

|<span data-ttu-id="5ed2c-183">图标大小：</span><span class="sxs-lookup"><span data-stu-id="5ed2c-183">Icon Size:</span></span>|<span data-ttu-id="5ed2c-184">16px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-184">16px</span></span>|<span data-ttu-id="5ed2c-185">20px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-185">20px</span></span>|<span data-ttu-id="5ed2c-186">24px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-186">24px</span></span>|<span data-ttu-id="5ed2c-187">32px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-187">32px</span></span>|<span data-ttu-id="5ed2c-188">40px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-188">40px</span></span>|<span data-ttu-id="5ed2c-189">48px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-189">48px</span></span>|<span data-ttu-id="5ed2c-190">64px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-190">64px</span></span>|<span data-ttu-id="5ed2c-191">80px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-191">80px</span></span>|<span data-ttu-id="5ed2c-192">96px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-192">96px</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|<span data-ttu-id="5ed2c-193">**线条粗细：**</span><span class="sxs-lookup"><span data-stu-id="5ed2c-193">**Line Weight:**</span></span>|<span data-ttu-id="5ed2c-194">1px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-194">1px</span></span>|<span data-ttu-id="5ed2c-195">1px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-195">1px</span></span>|<span data-ttu-id="5ed2c-196">1px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-196">1px</span></span>|<span data-ttu-id="5ed2c-197">1px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-197">1px</span></span>|<span data-ttu-id="5ed2c-198">2px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-198">2px</span></span>|<span data-ttu-id="5ed2c-199">2px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-199">2px</span></span>|<span data-ttu-id="5ed2c-200">2px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-200">2px</span></span>|<span data-ttu-id="5ed2c-201">2px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-201">2px</span></span>|<span data-ttu-id="5ed2c-202">3px</span><span class="sxs-lookup"><span data-stu-id="5ed2c-202">3px</span></span>|
|<span data-ttu-id="5ed2c-203">**示例图标：**</span><span class="sxs-lookup"><span data-stu-id="5ed2c-203">**Example icon:**</span></span>|![16 像素图标](../images/monolineicon16.png)|![20 像素图标](../images/monolineicon17.png)|![24 像素图标](../images/monolineicon18.png)|![32 像素图标](../images/monolineicon19.png)|![40 像素图标](../images/monolineicon20.png)|![48 像素图标](../images/monolineicon21.png)|![64 像素图标](../images/monolineicon22.png)|![80 像素图标](../images/monolineicon23.png)|![96 像素图标](../images/monolineicon24.png)|

#### <a name="cutouts"></a><span data-ttu-id="5ed2c-213">剪切</span><span class="sxs-lookup"><span data-stu-id="5ed2c-213">Cutouts</span></span>

<span data-ttu-id="5ed2c-214">当图标元素放置在另一个元素的顶部时， (元素) 的剪切线用于在两个元素之间提供空间，主要用于可读性。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-214">When an icon element is placed on top of another element, a cutout (of the bottom element) is used to provide space between the two elements, mainly for readability purposes.</span></span> <span data-ttu-id="5ed2c-215">当修饰符放置在基元素的顶部时，通常会发生这种情况，但在某些情况下，这两个元素都不是修饰符。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-215">This usually happens when a modifier is placed on top of a base element, but there are also cases where neither of the elements is a modifier.</span></span> <span data-ttu-id="5ed2c-216">这两个元素之间的这些切口有时称为"间隙"。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-216">These cutouts between the two elements is sometimes referred to as a "gap".</span></span>

<span data-ttu-id="5ed2c-217">间隙的大小应该与用于该大小的线粗细的宽度相同。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-217">The size of the gap should be the same width as the line weight used on that size.</span></span> <span data-ttu-id="5ed2c-218">如果制作 16 像素的图标，间隙宽度为 1px，如果是 48 像素的图标，间隙应为 2px。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-218">If making a 16 px icon, the gap width would be 1px and if it is a 48 px icon then the gap should be 2px.</span></span> <span data-ttu-id="5ed2c-219">以下示例显示一个 32 像素的图标，该图标的修饰符和基础基底之间的间隙为 1px。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-219">The following example shows a 32 px icon with a gap of 1px between the modifier and the underlying base.</span></span>

![修饰符和基础基之间的间隙为 1px 的 32 像素图标](../images/monolineicon25.png)

<span data-ttu-id="5ed2c-221">在某些情况下，如果修饰符有对角或曲线边缘且标准间隙未提供足够的分离，则间隙可能会增加 1/2 像素。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-221">In some cases, the gap can be increase by a 1/2 px if the modifier has a diagonal or curved edge and the standard gap doesn't provide enough separation.</span></span> <span data-ttu-id="5ed2c-222">这很可能只影响线条粗细为 1px 的图标：16 像素、20 像素、24 像素和 32 像素。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-222">This will likely only affect the icons with 1px line weight: 16 px, 20 px, 24 px, and 32 px.</span></span>

#### <a name="background-fills"></a><span data-ttu-id="5ed2c-223">背景填充</span><span class="sxs-lookup"><span data-stu-id="5ed2c-223">Background fills</span></span>

<span data-ttu-id="5ed2c-224">Monoline 图标集内大多数图标都需要背景填充。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-224">Most icons in the Monoline icon set require background fills.</span></span> <span data-ttu-id="5ed2c-225">但是，在某些情况下，对象自然没有填充，因此不应应用填充。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-225">However, there are cases where the object would not naturally have a fill, so no fill should be applied.</span></span> <span data-ttu-id="5ed2c-226">以下图标具有白色填充。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-226">The following icons have a white fill.</span></span>

![使用白色填充编译五个图标](../images/monolineicon26.png)

<span data-ttu-id="5ed2c-228">以下图标没有填充。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-228">The following icons have no fill.</span></span> <span data-ttu-id="5ed2c-229"> (包括齿轮图标，以显示中洞未填充。) </span><span class="sxs-lookup"><span data-stu-id="5ed2c-229">(The gear icon is included to show that the center hole is not filled.)</span></span>

![五个无填充图标的编译](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a><span data-ttu-id="5ed2c-231">填充最佳做法</span><span class="sxs-lookup"><span data-stu-id="5ed2c-231">Best practices for fills</span></span>

###### <a name="dos"></a><span data-ttu-id="5ed2c-232">Dos：</span><span class="sxs-lookup"><span data-stu-id="5ed2c-232">Dos:</span></span>

- <span data-ttu-id="5ed2c-233">填充具有已定义边界的任何元素，并且自然具有填充。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-233">Fill any element that has a defined boundary, and would naturally have a fill.</span></span>
- <span data-ttu-id="5ed2c-234">使用单独的形状创建背景填充。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-234">Use a separate shape to create the background fill.</span></span>
- <span data-ttu-id="5ed2c-235">使用 **调色板 中的** 背景 [填充](#color)。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-235">Use **Background Fill** from the [color palette](#color).</span></span>
- <span data-ttu-id="5ed2c-236">保持重叠元素之间的像素分隔。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-236">Maintain the pixel separation between overlapping elements.</span></span>
- <span data-ttu-id="5ed2c-237">在多个对象之间填充。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-237">Fill between multiple objects.</span></span>

###### <a name="donts"></a><span data-ttu-id="5ed2c-238">请勿：</span><span class="sxs-lookup"><span data-stu-id="5ed2c-238">Don'ts:</span></span>

- <span data-ttu-id="5ed2c-239">不要填充无法自然填充的对象;例如，一个平剪纸。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-239">Don't fill objects that would not naturally be filled; for example, a paperclip.</span></span>
- <span data-ttu-id="5ed2c-240">不要填充方括号。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-240">Don't fill brackets.</span></span>
- <span data-ttu-id="5ed2c-241">请勿在数字或 alpha 字符后面填充。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-241">Don't fill behind numbers or alpha characters.</span></span>

### <a name="color"></a><span data-ttu-id="5ed2c-242">颜色</span><span class="sxs-lookup"><span data-stu-id="5ed2c-242">Color</span></span>

<span data-ttu-id="5ed2c-243">调色板专为简单和辅助功能设计。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-243">The color palette has been designed for simplicity and accessibility.</span></span> <span data-ttu-id="5ed2c-244">它包含 4 种中性颜色以及蓝色、绿色、黄色、红色和紫色的两种变体。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-244">It contains 4 neutral colors and two variations for blue, green, yellow, red, and purple.</span></span> <span data-ttu-id="5ed2c-245">橙色有意不包含在单声道图标调色板中。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-245">Orange is intentionally not included in the Monoline icon color palette.</span></span> <span data-ttu-id="5ed2c-246">每种颜色旨在以本节所述的特定方式使用。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-246">Each color is intended to be used in specific ways as outlined in this section.</span></span>

#### <a name="palette"></a><span data-ttu-id="5ed2c-247">调色板</span><span class="sxs-lookup"><span data-stu-id="5ed2c-247">Palette</span></span>

![单色灰色的四种底纹：独立或大纲的深灰色、大纲或内容的中灰色、背景填充的浅灰色和浅灰色的填充](../images/monoline-grayshades.png)

![单行调色板包括独立、大纲和填充的蓝色、绿色、黄色、红色和紫色底纹](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a><span data-ttu-id="5ed2c-250">如何使用颜色</span><span class="sxs-lookup"><span data-stu-id="5ed2c-250">How to use color</span></span>

<span data-ttu-id="5ed2c-251">在单声道调色板中，所有颜色都有独立、大纲和填充变体。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-251">In the Monoline color palette, all colors have Standalone, Outline, and Fill variations.</span></span> <span data-ttu-id="5ed2c-252">通常，使用填充和边框构造元素。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-252">Generally, elements are constructed with a fill and a border.</span></span> <span data-ttu-id="5ed2c-253">颜色以下列模式之一应用：</span><span class="sxs-lookup"><span data-stu-id="5ed2c-253">The colors are applied in one of the following patterns:</span></span>

- <span data-ttu-id="5ed2c-254">对于没有填充的对象，单独的独立颜色。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-254">The Standalone color alone for objects that have no fill.</span></span>
- <span data-ttu-id="5ed2c-255">边框使用"边框"颜色，填充使用"填充"颜色。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-255">The border uses the Outline color and the fill uses the Fill color.</span></span>
- <span data-ttu-id="5ed2c-256">边框使用独立颜色，填充使用背景填充颜色。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-256">The border uses the Standalone color and the fill uses the Background Fill color.</span></span>

<span data-ttu-id="5ed2c-257">以下是使用颜色的示例。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-257">The following are examples of using color.</span></span>

![编译具有边框或填充颜色或两者同时具有颜色的三个图标](../images/monolineicon28.png)

<span data-ttu-id="5ed2c-259">最常见情况是让元素将深灰色独立版与背景填充一同使用。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-259">The most common situation will be to have an element use Dark Gray Standalone with Background Fill.</span></span>

<span data-ttu-id="5ed2c-260">使用彩色 Fill 时，它应始终具有相应的"轮廓"颜色。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-260">When using a colored Fill, it should always be with its corresponding Outline color.</span></span> <span data-ttu-id="5ed2c-261">例如，蓝色填充只能与蓝色边框一同使用。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-261">For example, Blue Fill should only be used with Blue Outline.</span></span> <span data-ttu-id="5ed2c-262">但是此一般规则有两个例外：</span><span class="sxs-lookup"><span data-stu-id="5ed2c-262">But there are two exceptions to this general rule:</span></span>

- <span data-ttu-id="5ed2c-263">背景填充可以与任何单独的颜色一同使用。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-263">Background Fill can be used with any color Standalone.</span></span>
- <span data-ttu-id="5ed2c-264">浅灰色填充可以与两种不同的大纲颜色一同使用：深灰色或中灰色。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-264">Light Gray Fill can be used with two different Outline colors: Dark Gray or Medium Gray.</span></span>

#### <a name="when-to-use-color"></a><span data-ttu-id="5ed2c-265">何时使用颜色</span><span class="sxs-lookup"><span data-stu-id="5ed2c-265">When to use color</span></span>

<span data-ttu-id="5ed2c-266">颜色应该用于传达图标的含义，而不是用于修饰。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-266">Color should be used to convey the meaning of the icon rather than for embellishment.</span></span> <span data-ttu-id="5ed2c-267">它 **应突出显示给用户** 的操作。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-267">It should **highlight the action** to the user.</span></span> <span data-ttu-id="5ed2c-268">将修饰符添加到具有颜色的基本元素时，基元素通常转换为深灰色和背景填充，以便修饰符可以是颜色元素，如以下示例，将"X"修饰符添加到下一组最左侧图标的图片基础中。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-268">When a modifier is added to a base element that has color, the base element is typically turned into Dark Gray and Background Fill so that the modifier can be the element of color, such as the case below with the "X" modifier being added to the picture base in the leftmost icon of the following set.</span></span>

![使用颜色的五个图标的编译](../images/monolineicon29.png)

<span data-ttu-id="5ed2c-270">除了上面提到的"轮廓"和"填充"外，你应当将图标限制为一种其他颜色。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-270">You should limit your icons to **one** additional color, other than the Outline and Fill mentioned above.</span></span> <span data-ttu-id="5ed2c-271">但是，如果它对于其隐喻至关重要，可以使用更多颜色，但除了灰色外，还有两种其他颜色。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-271">However, more colors can be used if it is vital for its metaphor, with a limit of two additional colors other than gray.</span></span> <span data-ttu-id="5ed2c-272">在极少数情况下，当需要更多颜色时，会存在例外情况。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-272">In rare cases, there are exceptions when more colors are needed.</span></span> <span data-ttu-id="5ed2c-273">以下是仅使用一种颜色的图标的很好示例。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-273">The following are good examples of icons that use just one color.</span></span>

  ![编译五个图标，每个图标使用一种颜色](../images/monolineicon30.png)

<span data-ttu-id="5ed2c-275">但以下图标使用的颜色过多。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-275">But the following icons use too many colors.</span></span>

  ![编译五个图标，每个图标都使用多个颜色](../images/monolineicon31.png)

<span data-ttu-id="5ed2c-277">对 **内部"** 内容"使用中灰色，如电子表格图标中的网格线。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-277">Use **Medium Gray** for interior "content", such as grid lines in an icon of a spreadsheet.</span></span> <span data-ttu-id="5ed2c-278">当内容需要显示控件的行为时，会使用其他内部颜色。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-278">Additional interior colors are used when the content needs to show the behavior of the control.</span></span>

![使用中灰色内部元素编译五个图标](../images/monolineicon32.png)

#### <a name="text-lines"></a><span data-ttu-id="5ed2c-280">文本行</span><span class="sxs-lookup"><span data-stu-id="5ed2c-280">Text lines</span></span>

<span data-ttu-id="5ed2c-281">当文本行位于"容器"中时 (例如，文档中的文本) 中灰色。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-281">When text lines are in a "container" (for example, text on a document), use medium gray.</span></span> <span data-ttu-id="5ed2c-282">不在容器中的文本行应为 **深灰色**。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-282">Text lines not in a container should be **Dark Gray**.</span></span>

### <a name="text"></a><span data-ttu-id="5ed2c-283">文本</span><span class="sxs-lookup"><span data-stu-id="5ed2c-283">Text</span></span>

<span data-ttu-id="5ed2c-284">避免在图标中使用文本字符。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-284">Avoid using text characters in icons.</span></span> <span data-ttu-id="5ed2c-285">由于 Office 产品已全球使用，因此我们希望尽可能将图标保持中性语言。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-285">Since Office products are used around the world, we want to keep icons as language neutral as possible.</span></span>

## <a name="production"></a><span data-ttu-id="5ed2c-286">生产</span><span class="sxs-lookup"><span data-stu-id="5ed2c-286">Production</span></span>

### <a name="icon-file-format"></a><span data-ttu-id="5ed2c-287">图标文件格式</span><span class="sxs-lookup"><span data-stu-id="5ed2c-287">Icon file format</span></span>

<span data-ttu-id="5ed2c-288">最终图标应另存为 .png 图像文件。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-288">The final icons should be saved as .png image files.</span></span> <span data-ttu-id="5ed2c-289">将 PNG 格式与透明背景一同使用，并且具有 32 位深度。</span><span class="sxs-lookup"><span data-stu-id="5ed2c-289">Use PNG format with a transparent background and have 32-bit depth.</span></span>

## <a name="see-also"></a><span data-ttu-id="5ed2c-290">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5ed2c-290">See also</span></span>

- [<span data-ttu-id="5ed2c-291">图标清单元素</span><span class="sxs-lookup"><span data-stu-id="5ed2c-291">Icon manifest element</span></span>](../reference/manifest/icon.md)
- [<span data-ttu-id="5ed2c-292">IconUrl 清单元素</span><span class="sxs-lookup"><span data-stu-id="5ed2c-292">IconUrl manifest element</span></span>](../reference/manifest/iconurl.md)
- [<span data-ttu-id="5ed2c-293">HighResolutionIconUrl 清单元素</span><span class="sxs-lookup"><span data-stu-id="5ed2c-293">HighResolutionIconUrl manifest element</span></span>](../reference/manifest/highresolutioniconurl.md)
- [<span data-ttu-id="5ed2c-294">创建加载项图标</span><span class="sxs-lookup"><span data-stu-id="5ed2c-294">Create an icon for your add-in</span></span>](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in)
