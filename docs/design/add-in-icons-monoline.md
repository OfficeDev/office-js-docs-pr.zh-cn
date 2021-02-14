---
title: Office 加载项的单行样式图标指南
description: 获取有关在 Office 外接程序中使用单行样式图标的指南。
ms.date: 2/09/2021
localization_priority: Normal
ms.openlocfilehash: 262cde129c7f7d3dd3f32b32e0a8e750cf016ef8
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237950"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a><span data-ttu-id="0846a-103">Office 加载项的单行样式图标指南</span><span class="sxs-lookup"><span data-stu-id="0846a-103">Monoline style icon guidelines for Office Add-ins</span></span>

<span data-ttu-id="0846a-104">单行样式图标在 Office 应用中使用。</span><span class="sxs-lookup"><span data-stu-id="0846a-104">Monoline style iconography are used in Office apps.</span></span> <span data-ttu-id="0846a-105">如果希望图标与非订阅 Office 2013+的新鲜样式匹配，请参阅 Office 外接程序的"全新样式 ["图标指南](add-in-icons-fresh.md)。</span><span class="sxs-lookup"><span data-stu-id="0846a-105">If you would prefer that your icons match the Fresh style of non-subscription Office 2013+, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).</span></span>

## <a name="office-monoline-visual-style"></a><span data-ttu-id="0846a-106">Office 单声道视觉样式</span><span class="sxs-lookup"><span data-stu-id="0846a-106">Office Monoline visual style</span></span>

<span data-ttu-id="0846a-107">单声道样式的目标是具有一致、清晰且可访问的图标，以通过简单的视觉效果传达操作和功能，确保图标可供所有用户访问，并且具有与 Windows 中其他位置使用的样式一致的样式。</span><span class="sxs-lookup"><span data-stu-id="0846a-107">The goal of the Monoline style to have consistent, clear, and accessible iconography to communicate action and features with simple visuals, ensure the icons are accessible to all users, and have a style that is consistent with those used elsewhere in Windows.</span></span>

<span data-ttu-id="0846a-108">以下指南适用于第三方开发人员，他们希望为功能创建图标，这些图标与已有的 Office 产品图标一致。</span><span class="sxs-lookup"><span data-stu-id="0846a-108">The following guidelines are for 3rd party developers who want to create icons for features that will be consistent with the icons already present Office products.</span></span>

### <a name="design-principles"></a><span data-ttu-id="0846a-109">设计原则</span><span class="sxs-lookup"><span data-stu-id="0846a-109">Design principles</span></span>

- <span data-ttu-id="0846a-110">简单、干净、清晰。</span><span class="sxs-lookup"><span data-stu-id="0846a-110">Simple, clean, clear.</span></span>
- <span data-ttu-id="0846a-111">仅包含必要的元素。</span><span class="sxs-lookup"><span data-stu-id="0846a-111">Contain only necessary elements.</span></span>
- <span data-ttu-id="0846a-112">受 Windows 图标样式启发。</span><span class="sxs-lookup"><span data-stu-id="0846a-112">Inspired by Windows icon style.</span></span>
- <span data-ttu-id="0846a-113">所有用户均可访问。</span><span class="sxs-lookup"><span data-stu-id="0846a-113">Accessible to all users.</span></span>

#### <a name="conveying-meaning"></a><span data-ttu-id="0846a-114">传达含义</span><span class="sxs-lookup"><span data-stu-id="0846a-114">Conveying meaning</span></span>

- <span data-ttu-id="0846a-115">使用描述性元素（如页面）表示文档或信封来表示邮件。</span><span class="sxs-lookup"><span data-stu-id="0846a-115">Use descriptive elements such as a page to represent a document or an envelope to represent mail.</span></span>
- <span data-ttu-id="0846a-116">使用相同的元素来表示同一概念，即邮件始终用信封而不是标记表示。</span><span class="sxs-lookup"><span data-stu-id="0846a-116">Use the same element to represent the same concept, i.e., mail is always represented by an envelope, not a stamp.</span></span>
- <span data-ttu-id="0846a-117">在概念开发过程中使用核心隐喻。</span><span class="sxs-lookup"><span data-stu-id="0846a-117">Use a core metaphor during concept development.</span></span>

#### <a name="reduction-of-elements"></a><span data-ttu-id="0846a-118">元素减少</span><span class="sxs-lookup"><span data-stu-id="0846a-118">Reduction of Elements</span></span>

- <span data-ttu-id="0846a-119">将图标缩小到其核心含义，仅使用对隐喻至关重要的元素。</span><span class="sxs-lookup"><span data-stu-id="0846a-119">Reduce the icon to its core meaning, using only elements that are essential to the metaphor.</span></span>
- <span data-ttu-id="0846a-120">将图标中的元素数限制为两个，无论图标大小如何。</span><span class="sxs-lookup"><span data-stu-id="0846a-120">Limit the number of elements in an icon to two, regardless of icon size.</span></span>

#### <a name="consistency"></a><span data-ttu-id="0846a-121">一致性</span><span class="sxs-lookup"><span data-stu-id="0846a-121">Consistency</span></span>

<span data-ttu-id="0846a-122">图标的大小、排列和颜色应一致。</span><span class="sxs-lookup"><span data-stu-id="0846a-122">Sizes, arrangement, and color of icons should be consistent.</span></span>

#### <a name="styling"></a><span data-ttu-id="0846a-123">样式设置</span><span class="sxs-lookup"><span data-stu-id="0846a-123">Styling</span></span>

##### <a name="perspective"></a><span data-ttu-id="0846a-124">Perspective</span><span class="sxs-lookup"><span data-stu-id="0846a-124">Perspective</span></span>

<span data-ttu-id="0846a-125">默认情况下，单声道图标是向前的。</span><span class="sxs-lookup"><span data-stu-id="0846a-125">Monoline icons are forward-facing by default.</span></span> <span data-ttu-id="0846a-126">允许使用需要透视和/或旋转的某些元素，如立方体，但例外应保持最少。</span><span class="sxs-lookup"><span data-stu-id="0846a-126">Certain elements that require perspective and/or rotation, such as a cube, are allowed, but exceptions should be kept to a minimum.</span></span>

##### <a name="embellishment"></a><span data-ttu-id="0846a-127">Embellishment</span><span class="sxs-lookup"><span data-stu-id="0846a-127">Embellishment</span></span>

<span data-ttu-id="0846a-128">单声道是一种简洁的最少样式。</span><span class="sxs-lookup"><span data-stu-id="0846a-128">Monoline is a clean minimal style.</span></span> <span data-ttu-id="0846a-129">所有内容都使用平面颜色，这意味着没有渐变、纹理或光源。</span><span class="sxs-lookup"><span data-stu-id="0846a-129">Everything uses flat color, which means there are no gradients, textures, or light sources.</span></span>

## <a name="designing"></a><span data-ttu-id="0846a-130">设计</span><span class="sxs-lookup"><span data-stu-id="0846a-130">Designing</span></span>

### <a name="sizes"></a><span data-ttu-id="0846a-131">大小</span><span class="sxs-lookup"><span data-stu-id="0846a-131">Sizes</span></span>

<span data-ttu-id="0846a-132">我们建议你以所有这些大小生成每个图标，以支持高 DPI 设备。</span><span class="sxs-lookup"><span data-stu-id="0846a-132">We recommend that you produce each icon in all these sizes to support high DPI devices.</span></span> <span data-ttu-id="0846a-133">绝对 *必需的大小* 为 16 像素、20 像素和 32 像素，因为大小为 100%。</span><span class="sxs-lookup"><span data-stu-id="0846a-133">The absolutely *required* sizes are 16 px, 20 px, and 32 px, as those are the 100% sizes.</span></span>

<span data-ttu-id="0846a-134">**16 像素、20 像素、24 像素、32 像素、40 像素、48 像素、64 像素、80 像素、96 像素**</span><span class="sxs-lookup"><span data-stu-id="0846a-134">**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**</span></span>

### <a name="layout"></a><span data-ttu-id="0846a-135">布局</span><span class="sxs-lookup"><span data-stu-id="0846a-135">Layout</span></span>

<span data-ttu-id="0846a-136">下面是一个包含修饰符的图标布局示例。</span><span class="sxs-lookup"><span data-stu-id="0846a-136">The following is an example of icon layout with a modifier.</span></span>

![右下角带修饰符的图标关系图](../images/monolineicon1.png)  ![包含基、修饰符、填充和标注的网格背景和标注的相同图标的图示](../images/monolineicon2.png)

#### <a name="elements"></a><span data-ttu-id="0846a-139">元素</span><span class="sxs-lookup"><span data-stu-id="0846a-139">Elements</span></span>

- <span data-ttu-id="0846a-140">**基本**：图标表示的主要概念。</span><span class="sxs-lookup"><span data-stu-id="0846a-140">**Base**: The main concept that the icon represents.</span></span> <span data-ttu-id="0846a-141">这通常是图标所需的唯一视觉对象，但有时可以使用辅助元素（修饰符）增强主要概念。</span><span class="sxs-lookup"><span data-stu-id="0846a-141">This is usually the only visual needed for the icon, but sometimes the main concept can be enhanced with a secondary element, a modifier.</span></span>

- <span data-ttu-id="0846a-142">**修饰符** 覆盖基本元素的任何元素;即，通常表示操作或状态的修饰符。</span><span class="sxs-lookup"><span data-stu-id="0846a-142">**Modifier** Any element that overlays the base; that is, a modifier that typically represents an action or a status.</span></span> <span data-ttu-id="0846a-143">它通过充当添加、更改或描述符来修改基元素。</span><span class="sxs-lookup"><span data-stu-id="0846a-143">It modifies the base element by acting as an addition, alteration, or a descriptor.</span></span>

![已调用基本和修饰符区域网格关系图](../images/monolineicon3.png)

### <a name="construction"></a><span data-ttu-id="0846a-145">建造</span><span class="sxs-lookup"><span data-stu-id="0846a-145">Construction</span></span>

#### <a name="element-placement"></a><span data-ttu-id="0846a-146">元素放置</span><span class="sxs-lookup"><span data-stu-id="0846a-146">Element placement</span></span>

<span data-ttu-id="0846a-147">基元素放置在填充内图标的中心。</span><span class="sxs-lookup"><span data-stu-id="0846a-147">Base elements are placed in the center of the icon within the padding.</span></span> <span data-ttu-id="0846a-148">如果无法完全居中放置，则基点应位于右上方。</span><span class="sxs-lookup"><span data-stu-id="0846a-148">If it can't be placed perfectly centered, then the base should err to the top right.</span></span> <span data-ttu-id="0846a-149">在下面的示例中，图标完全居中。</span><span class="sxs-lookup"><span data-stu-id="0846a-149">In the following example, the icon is perfectly centered.</span></span>

![显示完全居中的图标的图表](../images/monolineicon4.png)

<span data-ttu-id="0846a-151">在下面的示例中，图标在左侧出错。</span><span class="sxs-lookup"><span data-stu-id="0846a-151">In the following example, the icon is erring to the left.</span></span>

![显示左误 1 像素的图标的图表](../images/monolineicon5.png)

<span data-ttu-id="0846a-153">修饰符几乎总是放置在图标画布的右下角。</span><span class="sxs-lookup"><span data-stu-id="0846a-153">Modifiers are almost always placed in the bottom right corner of the icon canvas.</span></span> <span data-ttu-id="0846a-154">在极少数情况下，修饰符放置在不同的角。</span><span class="sxs-lookup"><span data-stu-id="0846a-154">In some rare cases, modifiers are placed in a different corner.</span></span> <span data-ttu-id="0846a-155">例如，如果右下角的修饰符无法识别基元素，请考虑将其放在左上角。</span><span class="sxs-lookup"><span data-stu-id="0846a-155">For example, if the base element would be unrecognizable with the modifier in the bottom right corner, then consider placing it in the upper left corner.</span></span>

![显示右下角有修饰符的四个图标，以及左上方有一个修饰符的图标的图表](../images/monolineicon6.png)

#### <a name="padding"></a><span data-ttu-id="0846a-157">填充</span><span class="sxs-lookup"><span data-stu-id="0846a-157">Padding</span></span>

<span data-ttu-id="0846a-158">每个大小图标在图标周围都有指定数量的填充。</span><span class="sxs-lookup"><span data-stu-id="0846a-158">Each size icon has a specified amount of padding around the icon.</span></span> <span data-ttu-id="0846a-159">基本元素保留在填充内，但修饰符应向上扩展到画布边缘，在填充外扩展到图标边框的边缘。</span><span class="sxs-lookup"><span data-stu-id="0846a-159">The base element stays within the padding, but the modifier should butt up to the edge of the canvas, extending outside of the padding to the edge of the icon border.</span></span> <span data-ttu-id="0846a-160">下图显示了用于每个图标大小的推荐填充。</span><span class="sxs-lookup"><span data-stu-id="0846a-160">The following images show the recommended padding to use for each of the icon sizes.</span></span>

|<span data-ttu-id="0846a-161">**16px**</span><span class="sxs-lookup"><span data-stu-id="0846a-161">**16px**</span></span>|<span data-ttu-id="0846a-162">**20px**</span><span class="sxs-lookup"><span data-stu-id="0846a-162">**20px**</span></span>|<span data-ttu-id="0846a-163">**24px**</span><span class="sxs-lookup"><span data-stu-id="0846a-163">**24px**</span></span>|<span data-ttu-id="0846a-164">**32px**</span><span class="sxs-lookup"><span data-stu-id="0846a-164">**32px**</span></span>|<span data-ttu-id="0846a-165">**40px**</span><span class="sxs-lookup"><span data-stu-id="0846a-165">**40px**</span></span>|<span data-ttu-id="0846a-166">**48px**</span><span class="sxs-lookup"><span data-stu-id="0846a-166">**48px**</span></span>|<span data-ttu-id="0846a-167">**64px**</span><span class="sxs-lookup"><span data-stu-id="0846a-167">**64px**</span></span>|<span data-ttu-id="0846a-168">**80px**</span><span class="sxs-lookup"><span data-stu-id="0846a-168">**80px**</span></span>|<span data-ttu-id="0846a-169">**96px**</span><span class="sxs-lookup"><span data-stu-id="0846a-169">**96px**</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![具有 0px 填充的 16 像素图标](../images/monolineicon7.png)|![具有 1px 填充的 20 像素图标](../images/monolineicon8.png)|![具有 1px 填充的 24 像素图标](../images/monolineicon9.png)|![具有 2px 填充的 32 像素图标](../images/monolineicon10.png)|![具有 2px 填充的 40 像素图标](../images/monolineicon11.png)|![具有 3px 填充的 48 像素图标](../images/monolineicon12.png)|![具有 4px 填充的 64 像素图标](../images/monolineicon13.png)|![具有 5px 填充的 80 像素图标](../images/monolineicon14.png)|![具有 6px 填充的 96 像素图标](../images/monolineicon15.png)|

#### <a name="line-weights"></a><span data-ttu-id="0846a-179">线条粗细</span><span class="sxs-lookup"><span data-stu-id="0846a-179">Line weights</span></span>

<span data-ttu-id="0846a-180">单声道是一种样式，由线条和轮廓形状控制。</span><span class="sxs-lookup"><span data-stu-id="0846a-180">Monoline is a style dominated by line and outlined shapes.</span></span> <span data-ttu-id="0846a-181">根据你生成图标的大小，应使用以下行粗细。</span><span class="sxs-lookup"><span data-stu-id="0846a-181">Depending on what size you are producing the icon should use the following line weights.</span></span>

|<span data-ttu-id="0846a-182">图标大小：</span><span class="sxs-lookup"><span data-stu-id="0846a-182">Icon Size:</span></span>|<span data-ttu-id="0846a-183">16px</span><span class="sxs-lookup"><span data-stu-id="0846a-183">16px</span></span>|<span data-ttu-id="0846a-184">20px</span><span class="sxs-lookup"><span data-stu-id="0846a-184">20px</span></span>|<span data-ttu-id="0846a-185">24px</span><span class="sxs-lookup"><span data-stu-id="0846a-185">24px</span></span>|<span data-ttu-id="0846a-186">32px</span><span class="sxs-lookup"><span data-stu-id="0846a-186">32px</span></span>|<span data-ttu-id="0846a-187">40px</span><span class="sxs-lookup"><span data-stu-id="0846a-187">40px</span></span>|<span data-ttu-id="0846a-188">48px</span><span class="sxs-lookup"><span data-stu-id="0846a-188">48px</span></span>|<span data-ttu-id="0846a-189">64px</span><span class="sxs-lookup"><span data-stu-id="0846a-189">64px</span></span>|<span data-ttu-id="0846a-190">80px</span><span class="sxs-lookup"><span data-stu-id="0846a-190">80px</span></span>|<span data-ttu-id="0846a-191">96px</span><span class="sxs-lookup"><span data-stu-id="0846a-191">96px</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|<span data-ttu-id="0846a-192">**线条粗细：**</span><span class="sxs-lookup"><span data-stu-id="0846a-192">**Line Weight:**</span></span>|<span data-ttu-id="0846a-193">1px</span><span class="sxs-lookup"><span data-stu-id="0846a-193">1px</span></span>|<span data-ttu-id="0846a-194">1px</span><span class="sxs-lookup"><span data-stu-id="0846a-194">1px</span></span>|<span data-ttu-id="0846a-195">1px</span><span class="sxs-lookup"><span data-stu-id="0846a-195">1px</span></span>|<span data-ttu-id="0846a-196">1px</span><span class="sxs-lookup"><span data-stu-id="0846a-196">1px</span></span>|<span data-ttu-id="0846a-197">2px</span><span class="sxs-lookup"><span data-stu-id="0846a-197">2px</span></span>|<span data-ttu-id="0846a-198">2px</span><span class="sxs-lookup"><span data-stu-id="0846a-198">2px</span></span>|<span data-ttu-id="0846a-199">2px</span><span class="sxs-lookup"><span data-stu-id="0846a-199">2px</span></span>|<span data-ttu-id="0846a-200">2px</span><span class="sxs-lookup"><span data-stu-id="0846a-200">2px</span></span>|<span data-ttu-id="0846a-201">3px</span><span class="sxs-lookup"><span data-stu-id="0846a-201">3px</span></span>|
|<span data-ttu-id="0846a-202">**示例图标：**</span><span class="sxs-lookup"><span data-stu-id="0846a-202">**Example icon:**</span></span>|![16 像素图标](../images/monolineicon16.png)|![20 像素图标](../images/monolineicon17.png)|![24 像素图标](../images/monolineicon18.png)|![32 像素图标](../images/monolineicon19.png)|![40 像素图标](../images/monolineicon20.png)|![48 像素图标](../images/monolineicon21.png)|![64 像素图标](../images/monolineicon22.png)|![80 像素图标](../images/monolineicon23.png)|![96 像素图标](../images/monolineicon24.png)|

#### <a name="cutouts"></a><span data-ttu-id="0846a-212">剪切线</span><span class="sxs-lookup"><span data-stu-id="0846a-212">Cutouts</span></span>

<span data-ttu-id="0846a-213">当图标元素放置在另一个元素的顶部时， (元素) 的剪切线用于提供两个元素之间的空间，主要用于可读性。</span><span class="sxs-lookup"><span data-stu-id="0846a-213">When an icon element is placed on top of another element, a cutout (of the bottom element) is used to provide space between the two elements, mainly for readability purposes.</span></span> <span data-ttu-id="0846a-214">当修饰符放置在基元素的顶部时，通常会发生这种情况，但在某些情况下，这两个元素都不是修饰符。</span><span class="sxs-lookup"><span data-stu-id="0846a-214">This usually happens when a modifier is placed on top of a base element, but there are also cases where neither of the elements is a modifier.</span></span> <span data-ttu-id="0846a-215">这两个元素之间的这些切口有时称为"间隙"。</span><span class="sxs-lookup"><span data-stu-id="0846a-215">These cutouts between the two elements is sometimes referred to as a "gap".</span></span>

<span data-ttu-id="0846a-216">间隙的大小应该与用于该大小的线条粗细的宽度相同。</span><span class="sxs-lookup"><span data-stu-id="0846a-216">The size of the gap should be the same width as the line weight used on that size.</span></span> <span data-ttu-id="0846a-217">如果创建 16 像素图标，间隙宽度为 1px，如果是 48 像素图标，间隙应为 2px。</span><span class="sxs-lookup"><span data-stu-id="0846a-217">If making a 16 px icon, the gap width would be 1px and if it is a 48 px icon then the gap should be 2px.</span></span> <span data-ttu-id="0846a-218">以下示例显示一个 32 像素图标，修饰符和基础基之间的间隙为 1px。</span><span class="sxs-lookup"><span data-stu-id="0846a-218">The following example shows a 32 px icon with a gap of 1px between the modifier and the underlying base.</span></span>

![修饰符和基础基之间的间隙为 1px 的 32 像素图标](../images/monolineicon25.png)

<span data-ttu-id="0846a-220">在某些情况下，如果修饰符具有对角线或曲线边缘，并且标准间隙没有提供足够间隔，则间隙可能会增加 1/2 像素。</span><span class="sxs-lookup"><span data-stu-id="0846a-220">In some cases, the gap can be increase by a 1/2 px if the modifier has a diagonal or curved edge and the standard gap doesn't provide enough separation.</span></span> <span data-ttu-id="0846a-221">这很可能只影响线条粗细为 1px 的图标：16 像素、20 像素、24 像素和 32 像素。</span><span class="sxs-lookup"><span data-stu-id="0846a-221">This will likely only affect the icons with 1px line weight: 16 px, 20 px, 24 px, and 32 px.</span></span>

#### <a name="background-fills"></a><span data-ttu-id="0846a-222">背景填充</span><span class="sxs-lookup"><span data-stu-id="0846a-222">Background fills</span></span>

<span data-ttu-id="0846a-223">单声道图标集中的大多数图标都需要背景填充。</span><span class="sxs-lookup"><span data-stu-id="0846a-223">Most icons in the Monoline icon set require background fills.</span></span> <span data-ttu-id="0846a-224">但是，在某些情况下，对象自然没有填充，因此不应应用填充。</span><span class="sxs-lookup"><span data-stu-id="0846a-224">However, there are cases where the object would not naturally have a fill, so no fill should be applied.</span></span> <span data-ttu-id="0846a-225">以下图标具有白色填充。</span><span class="sxs-lookup"><span data-stu-id="0846a-225">The following icons have a white fill.</span></span>

![使用白色填充编译五个图标](../images/monolineicon26.png)

<span data-ttu-id="0846a-227">以下图标没有填充。</span><span class="sxs-lookup"><span data-stu-id="0846a-227">The following icons have no fill.</span></span> <span data-ttu-id="0846a-228"> (包括齿轮图标，以显示中心内没有填充。) </span><span class="sxs-lookup"><span data-stu-id="0846a-228">(The gear icon is included to show that the center hole is not filled.)</span></span>

![无填充的五个图标的编译](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a><span data-ttu-id="0846a-230">填充最佳做法</span><span class="sxs-lookup"><span data-stu-id="0846a-230">Best practices for fills</span></span>

###### <a name="dos"></a><span data-ttu-id="0846a-231">Dos：</span><span class="sxs-lookup"><span data-stu-id="0846a-231">Dos:</span></span>

- <span data-ttu-id="0846a-232">填充具有定义边界且自然具有填充的任何元素。</span><span class="sxs-lookup"><span data-stu-id="0846a-232">Fill any element that has a defined boundary, and would naturally have a fill.</span></span>
- <span data-ttu-id="0846a-233">使用单独的形状创建背景填充。</span><span class="sxs-lookup"><span data-stu-id="0846a-233">Use a separate shape to create the background fill.</span></span>
- <span data-ttu-id="0846a-234">使用 **调色板** 中的 [背景填充](#color)。</span><span class="sxs-lookup"><span data-stu-id="0846a-234">Use **Background Fill** from the [color palette](#color).</span></span>
- <span data-ttu-id="0846a-235">保持重叠元素之间的像素分隔。</span><span class="sxs-lookup"><span data-stu-id="0846a-235">Maintain the pixel separation between overlapping elements.</span></span>
- <span data-ttu-id="0846a-236">在多个对象之间填充。</span><span class="sxs-lookup"><span data-stu-id="0846a-236">Fill between multiple objects.</span></span>

###### <a name="donts"></a><span data-ttu-id="0846a-237">请勿：</span><span class="sxs-lookup"><span data-stu-id="0846a-237">Don'ts:</span></span>

- <span data-ttu-id="0846a-238">不要填充无法自然填充的对象;例如，一个纸条。</span><span class="sxs-lookup"><span data-stu-id="0846a-238">Don't fill objects that would not naturally be filled; for example, a paperclip.</span></span>
- <span data-ttu-id="0846a-239">请勿填充方括号。</span><span class="sxs-lookup"><span data-stu-id="0846a-239">Don't fill brackets.</span></span>
- <span data-ttu-id="0846a-240">请勿在数字或 alpha 字符后面填充。</span><span class="sxs-lookup"><span data-stu-id="0846a-240">Don't fill behind numbers or alpha characters.</span></span>

### <a name="color"></a><span data-ttu-id="0846a-241">颜色</span><span class="sxs-lookup"><span data-stu-id="0846a-241">Color</span></span>

<span data-ttu-id="0846a-242">调色板专为简化和辅助功能设计。</span><span class="sxs-lookup"><span data-stu-id="0846a-242">The color palette has been designed for simplicity and accessibility.</span></span> <span data-ttu-id="0846a-243">它包含 4 种中性颜色以及蓝色、绿色、黄色、红色和紫色的两种变体。</span><span class="sxs-lookup"><span data-stu-id="0846a-243">It contains 4 neutral colors and two variations for blue, green, yellow, red, and purple.</span></span> <span data-ttu-id="0846a-244">橙色有意不包含在单声道图标调色板中。</span><span class="sxs-lookup"><span data-stu-id="0846a-244">Orange is intentionally not included in the Monoline icon color palette.</span></span> <span data-ttu-id="0846a-245">每种颜色都旨在以本节中概述的特定方式使用。</span><span class="sxs-lookup"><span data-stu-id="0846a-245">Each color is intended to be used in specific ways as outlined in this section.</span></span>

#### <a name="palette"></a><span data-ttu-id="0846a-246">调色板</span><span class="sxs-lookup"><span data-stu-id="0846a-246">Palette</span></span>

![单色灰色的四种阴影：独立或大纲的深灰色、大纲或内容的中灰色、背景填充的浅灰色和浅灰色的填充](../images/monoline-grayshades.png)

![单行调色板包括独立、大纲和填充的蓝色、绿色、黄色、红色和紫色底纹](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a><span data-ttu-id="0846a-249">如何使用颜色</span><span class="sxs-lookup"><span data-stu-id="0846a-249">How to use color</span></span>

<span data-ttu-id="0846a-250">在单声道调色板中，所有颜色都有独立、大纲和填充变体。</span><span class="sxs-lookup"><span data-stu-id="0846a-250">In the Monoline color palette, all colors have Standalone, Outline, and Fill variations.</span></span> <span data-ttu-id="0846a-251">通常，使用填充和边框构造元素。</span><span class="sxs-lookup"><span data-stu-id="0846a-251">Generally, elements are constructed with a fill and a border.</span></span> <span data-ttu-id="0846a-252">颜色采用以下模式之一：</span><span class="sxs-lookup"><span data-stu-id="0846a-252">The colors are applied in one of the following patterns:</span></span>

- <span data-ttu-id="0846a-253">没有填充的对象的独立颜色。</span><span class="sxs-lookup"><span data-stu-id="0846a-253">The Standalone color alone for objects that have no fill.</span></span>
- <span data-ttu-id="0846a-254">边框使用"轮廓"颜色，填充使用"填充"颜色。</span><span class="sxs-lookup"><span data-stu-id="0846a-254">The border uses the Outline color and the fill uses the Fill color.</span></span>
- <span data-ttu-id="0846a-255">边框使用独立颜色，填充使用背景填充颜色。</span><span class="sxs-lookup"><span data-stu-id="0846a-255">The border uses the Standalone color and the fill uses the Background Fill color.</span></span>

<span data-ttu-id="0846a-256">下面是使用颜色的示例。</span><span class="sxs-lookup"><span data-stu-id="0846a-256">The following are examples of using color.</span></span>

![在边框或填充中编译三个彩色图标，或同时编译两者](../images/monolineicon28.png)

<span data-ttu-id="0846a-258">最常见情况是让元素使用"深灰色独立"和"背景填充"。</span><span class="sxs-lookup"><span data-stu-id="0846a-258">The most common situation will be to have an element use Dark Gray Standalone with Background Fill.</span></span>

<span data-ttu-id="0846a-259">使用彩色填充时，它应始终带有相应的"轮廓"颜色。</span><span class="sxs-lookup"><span data-stu-id="0846a-259">When using a colored Fill, it should always be with its corresponding Outline color.</span></span> <span data-ttu-id="0846a-260">例如，蓝色填充只能与蓝色轮廓一同使用。</span><span class="sxs-lookup"><span data-stu-id="0846a-260">For example, Blue Fill should only be used with Blue Outline.</span></span> <span data-ttu-id="0846a-261">但此一般规则有两个例外：</span><span class="sxs-lookup"><span data-stu-id="0846a-261">But there are two exceptions to this general rule:</span></span>

- <span data-ttu-id="0846a-262">背景填充可以与任意颜色独立使用。</span><span class="sxs-lookup"><span data-stu-id="0846a-262">Background Fill can be used with any color Standalone.</span></span>
- <span data-ttu-id="0846a-263">浅灰色填充可用于两种不同的大纲颜色：深灰色或中灰色。</span><span class="sxs-lookup"><span data-stu-id="0846a-263">Light Gray Fill can be used with two different Outline colors: Dark Gray or Medium Gray.</span></span>

#### <a name="when-to-use-color"></a><span data-ttu-id="0846a-264">何时使用颜色</span><span class="sxs-lookup"><span data-stu-id="0846a-264">When to use color</span></span>

<span data-ttu-id="0846a-265">颜色应该用于传达图标的含义，而不是修饰。</span><span class="sxs-lookup"><span data-stu-id="0846a-265">Color should be used to convey the meaning of the icon rather than for embellishment.</span></span> <span data-ttu-id="0846a-266">它 **应突出显示给用户** 的操作。</span><span class="sxs-lookup"><span data-stu-id="0846a-266">It should **highlight the action** to the user.</span></span> <span data-ttu-id="0846a-267">当将修饰符添加到具有颜色的基元素中时，基元素通常会变为深灰色和背景填充，以便修饰符可以是颜色元素，如下面的情况，将"X"修饰符添加到以下集最左侧图标的图片基础。</span><span class="sxs-lookup"><span data-stu-id="0846a-267">When a modifier is added to a base element that has color, the base element is typically turned into Dark Gray and Background Fill so that the modifier can be the element of color, such as the case below with the "X" modifier being added to the picture base in the leftmost icon of the following set.</span></span>

![使用颜色的五个图标的编译](../images/monolineicon29.png)

<span data-ttu-id="0846a-269">除了上面提到的"轮廓"和"填充"外，你应当将图标限制为一种其他颜色。</span><span class="sxs-lookup"><span data-stu-id="0846a-269">You should limit your icons to **one** additional color, other than the Outline and Fill mentioned above.</span></span> <span data-ttu-id="0846a-270">但是，如果它对于其隐喻至关重要，可以使用更多颜色，但除了灰色外，其他两种颜色的限制。</span><span class="sxs-lookup"><span data-stu-id="0846a-270">However, more colors can be used if it is vital for its metaphor, with a limit of two additional colors other than gray.</span></span> <span data-ttu-id="0846a-271">在极少数情况下，当需要更多颜色时，会存在例外情况。</span><span class="sxs-lookup"><span data-stu-id="0846a-271">In rare cases, there are exceptions when more colors are needed.</span></span> <span data-ttu-id="0846a-272">以下是仅使用一种颜色的图标的不错示例。</span><span class="sxs-lookup"><span data-stu-id="0846a-272">The following are good examples of icons that use just one color.</span></span>

  ![编译五个图标，每个图标使用一种颜色](../images/monolineicon30.png)

<span data-ttu-id="0846a-274">但以下图标使用的颜色过多。</span><span class="sxs-lookup"><span data-stu-id="0846a-274">But the following icons use too many colors.</span></span>

  ![编译五个图标，每个图标都使用多个颜色](../images/monolineicon31.png)

<span data-ttu-id="0846a-276">对 **内部"** 内容"使用中灰色，如电子表格图标中的网格线。</span><span class="sxs-lookup"><span data-stu-id="0846a-276">Use **Medium Gray** for interior "content", such as grid lines in an icon of a spreadsheet.</span></span> <span data-ttu-id="0846a-277">当内容需要显示控件的行为时，会使用其他内部颜色。</span><span class="sxs-lookup"><span data-stu-id="0846a-277">Additional interior colors are used when the content needs to show the behavior of the control.</span></span>

![使用中灰色内部元素编译五个图标](../images/monolineicon32.png)

#### <a name="text-lines"></a><span data-ttu-id="0846a-279">文本行</span><span class="sxs-lookup"><span data-stu-id="0846a-279">Text lines</span></span>

<span data-ttu-id="0846a-280">当文本行位于"容器"中时 (例如，文档中的文本) 中灰色。</span><span class="sxs-lookup"><span data-stu-id="0846a-280">When text lines are in a "container" (for example, text on a document), use medium gray.</span></span> <span data-ttu-id="0846a-281">不在容器中的文本行应为 **深灰色**。</span><span class="sxs-lookup"><span data-stu-id="0846a-281">Text lines not in a container should be **Dark Gray**.</span></span>

### <a name="text"></a><span data-ttu-id="0846a-282">文本</span><span class="sxs-lookup"><span data-stu-id="0846a-282">Text</span></span>

<span data-ttu-id="0846a-283">避免在图标中使用文本字符。</span><span class="sxs-lookup"><span data-stu-id="0846a-283">Avoid using text characters in icons.</span></span> <span data-ttu-id="0846a-284">由于 Office 产品已全球使用，因此我们希望尽可能使图标保持中性语言。</span><span class="sxs-lookup"><span data-stu-id="0846a-284">Since Office products are used around the world, we want to keep icons as language neutral as possible.</span></span>

## <a name="production"></a><span data-ttu-id="0846a-285">生产</span><span class="sxs-lookup"><span data-stu-id="0846a-285">Production</span></span>

### <a name="icon-file-format"></a><span data-ttu-id="0846a-286">图标文件格式</span><span class="sxs-lookup"><span data-stu-id="0846a-286">Icon file format</span></span>

<span data-ttu-id="0846a-287">最终图标应另存为 .png 图像文件。</span><span class="sxs-lookup"><span data-stu-id="0846a-287">The final icons should be saved as .png image files.</span></span> <span data-ttu-id="0846a-288">使用具有透明背景且深度为 32 位的 PNG 格式。</span><span class="sxs-lookup"><span data-stu-id="0846a-288">Use PNG format with a transparent background and have 32-bit depth.</span></span>
