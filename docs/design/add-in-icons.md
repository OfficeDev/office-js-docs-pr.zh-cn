---
title: Office 外接程序的图标准则
description: ''
ms.date: 03/02/2019
localization_priority: Priority
ms.openlocfilehash: 8e741f70327584ddd1b6f51f19b276e072862229
ms.sourcegitcommit: 7ebd383f16ae5809bb6980a5f213b695d410e62c
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/06/2019
ms.locfileid: "30413634"
---
# <a name="icons"></a><span data-ttu-id="54c86-102">图标</span><span class="sxs-lookup"><span data-stu-id="54c86-102">Icons</span></span>
<span data-ttu-id="54c86-103">图标是行为或概念的可视化表示形式。</span><span class="sxs-lookup"><span data-stu-id="54c86-103">Icons are the visual representation of a behavior or concept.</span></span> <span data-ttu-id="54c86-104">它们通常用于为控件和命令添加含义。</span><span class="sxs-lookup"><span data-stu-id="54c86-104">They are often used to add meaning to controls and commands.</span></span> <span data-ttu-id="54c86-105">实际或符号化的视觉对象使用户能够以与标记帮助用户浏览其环境的相同方式浏览 UI。</span><span class="sxs-lookup"><span data-stu-id="54c86-105">Visuals, either realistic or symbolic, enable the user to navigate the UI the same way signs help users navigate their environment.</span></span> <span data-ttu-id="54c86-106">这些视觉对象应简单明了，并且只包含所需的详细信息，以使客户能够快速分析他们在选择控件时将会发生的操作。</span><span class="sxs-lookup"><span data-stu-id="54c86-106">They should be simple, clear, and contain only the necessary details to enable customers to quickly parse what action will occur when they choose a control.</span></span>

<span data-ttu-id="54c86-107">Office 功能区界面具有标准的视觉样式。</span><span class="sxs-lookup"><span data-stu-id="54c86-107">Office ribbon interfaces have a standard visual style.</span></span> <span data-ttu-id="54c86-108">这可以确保一致性并熟悉各个 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="54c86-108">This ensures consistency and familiarity across Office apps.</span></span> <span data-ttu-id="54c86-109">这些准则将有助于你为解决方案设计一组适合作为 Office 固有部分的 PNG 资产。</span><span class="sxs-lookup"><span data-stu-id="54c86-109">The guidelines will help you design a set of PNG assets for your solution that fit in as a natural part of Office.</span></span>

<span data-ttu-id="54c86-p103">许多 HTML 容器包含带有插图的控件。使用 Office UI Fabric 的自定义字体在外接程序中呈现 Office 样式图标。Fabric 的图标字体包含很多针对可缩放的常见 Office 隐喻、颜色和样式的字形以满足你的需要。如果你有带自己图标集的现有视觉语言，则可在 HTML 画布中随意使用。构建自己带标准图标集的品牌的连续性是任何设计语言的重要组成部分。请注意避免与 Office 隐喻产生冲突导致客户混淆。</span><span class="sxs-lookup"><span data-stu-id="54c86-p103">Many HTML containers contain controls with iconography. Use Office UI Fabric’s custom font to render Office styled icons in your add-in. Fabric’s icon font contains many glyphs for common Office metaphors that you can scale, color, and style to suit your needs. If you have an existing visual language with your own set of icons, feel free to use it in your HTML canvases. Building continuity with your own brand with a standard set of icons is an important part of any design language. Be careful to avoid creating confusion for customers by conflicting with Office metaphors.</span></span>


## <a name="design-icons-for-add-in-commands"></a><span data-ttu-id="54c86-116">加载项命令的设计图标</span><span class="sxs-lookup"><span data-stu-id="54c86-116">Design icons for add-in commands</span></span>

<span data-ttu-id="54c86-p104">[外接程序命令](add-in-commands.md)添加按钮、文本和 Office UI 图标。外接程序命令按钮应提供有意义的图标和标签，以便清楚地标识用户在使用命令时执行的操作。本文提供了样式和生产准则，可帮助你设计与 Office 无缝集成的图标。</span><span class="sxs-lookup"><span data-stu-id="54c86-p104">[Add-in commands](add-in-commands.md) add buttons, text, and icons to the Office UI. Your add-in command buttons should provide meaningful icons and labels that clearly identify the action the user is taking when they use a command. This article provides stylistic and production guidelines that help you design icons that integrate seamlessly with Office.</span></span> 

## <a name="office-icon-design-principles"></a><span data-ttu-id="54c86-120">Office 图标设计原则</span><span class="sxs-lookup"><span data-stu-id="54c86-120">Office icon design principles</span></span>

<span data-ttu-id="54c86-p105">Office 桌面客户端的 Office 2013 版本包括刷新的图标。替代样式更改已缩减。新图标仅包括必需通信元素。包括透视、渐变和光源的非必需元素均被删除。简化后的图标可支持对命令和控件的快速解析。请按照此样式设计最适合 Office 的图标。</span><span class="sxs-lookup"><span data-stu-id="54c86-p105">The Office 2013 release of the Office desktop clients includes refreshed iconography. The overriding stylistic change is reduction. The new icons include only essential communicative elements. Non-essential elements including perspective, gradients, and light source are removed. The simplified icons support faster parsing of commands and controls. Follow this style to best fit with Office.</span></span>

<span data-ttu-id="54c86-127">Office 图标均基于以下设计原则完成：</span><span class="sxs-lookup"><span data-stu-id="54c86-127">Office icons are based on the following design principles:</span></span> 

- <span data-ttu-id="54c86-128">以现代方式阐释 Office 图标集合</span><span class="sxs-lookup"><span data-stu-id="54c86-128">Modern interpretation of Office icon collection</span></span> 
- <span data-ttu-id="54c86-129">全新设计但又不陌生</span><span class="sxs-lookup"><span data-stu-id="54c86-129">Fresh yet familiar</span></span>  
- <span data-ttu-id="54c86-130">简单、清楚和直接</span><span class="sxs-lookup"><span data-stu-id="54c86-130">Simple, clear, and direct</span></span> 

<span data-ttu-id="54c86-131">下图显示了应用现代设计原则的图标。</span><span class="sxs-lookup"><span data-stu-id="54c86-131">The following image shows icons that apply the modern design principles.</span></span>

![显示 Office 旧图标的图像和刷新的以现代方式阐释的图标](../images/icons-images.png)

## <a name="best-practices"></a><span data-ttu-id="54c86-133">最佳实践</span><span class="sxs-lookup"><span data-stu-id="54c86-133">Best practices</span></span>

<span data-ttu-id="54c86-134">创建图标时，请遵循以下准则：</span><span class="sxs-lookup"><span data-stu-id="54c86-134">Follow these guidelines when you create your icons:</span></span> 

|<span data-ttu-id="54c86-135">允许事项</span><span class="sxs-lookup"><span data-stu-id="54c86-135">Do</span></span>|<span data-ttu-id="54c86-136">禁止事项</span><span class="sxs-lookup"><span data-stu-id="54c86-136">Don't</span></span>|
|:---|:---|
|<span data-ttu-id="54c86-137">让视觉对象保持简单明了，注重通信的关键元素。</span><span class="sxs-lookup"><span data-stu-id="54c86-137">Keep visuals simple and clear, focusing on the key element(s) of the communication.</span></span>| <span data-ttu-id="54c86-138">不要使用使图标显得杂乱的项目。</span><span class="sxs-lookup"><span data-stu-id="54c86-138">Don't use artifacts that make your icon look messy.</span></span>|
|<span data-ttu-id="54c86-139">使用 Office 图标语言来表示行为或概念。</span><span class="sxs-lookup"><span data-stu-id="54c86-139">Use the Office icon language to represent behaviors or concepts.</span></span>|<span data-ttu-id="54c86-p106">请勿在 Office 功能区或关联菜单中改变外接程序命令的 Office UI Fabric 用途。Fabric 图标风格不同，不能匹配。</span><span class="sxs-lookup"><span data-stu-id="54c86-p106">Don’t repurpose Office UI Fabric glyphs for add-in commands in the Office ribbon or contextual menus. Fabric icons are stylistically different and will not match.</span></span>|
|<span data-ttu-id="54c86-142">将画笔等公用 Office 视觉隐喻重用于格式或用于查找的放大镜。</span><span class="sxs-lookup"><span data-stu-id="54c86-142">Reuse common Office visual metaphors such as paintbrush for format or magnifying glass for find.</span></span>|<span data-ttu-id="54c86-143">不要对不同的命令重复使用视觉隐喻。</span><span class="sxs-lookup"><span data-stu-id="54c86-143">Don't reuse visual metaphors for different commands.</span></span> <span data-ttu-id="54c86-144">对不同的行为和概念使用同一图标可能会引起混淆。</span><span class="sxs-lookup"><span data-stu-id="54c86-144">Using the same icon for different behaviors and concepts can cause confusion.</span></span> |
|<span data-ttu-id="54c86-145">重绘图标，使其更大或更小。</span><span class="sxs-lookup"><span data-stu-id="54c86-145">Redraw your icons to make them small or larger.</span></span> <span data-ttu-id="54c86-146">请花时间重绘切割区、角和圆边，以最大化线条的清晰度。</span><span class="sxs-lookup"><span data-stu-id="54c86-146">Take the time to redraw cutouts, corners, and rounded edges to maximize line clarity.</span></span> |<span data-ttu-id="54c86-147">切勿通过缩小或扩大尺寸来调整图标大小。</span><span class="sxs-lookup"><span data-stu-id="54c86-147">Don't resize your icons by shrinking or enlarging in size.</span></span> <span data-ttu-id="54c86-148">这可能会导致视觉对象质量不佳和操作不清晰。</span><span class="sxs-lookup"><span data-stu-id="54c86-148">This can lead to poor visual quality and unclear actions.</span></span> <span data-ttu-id="54c86-149">对于较大尺寸的复杂图标，如果不是通过重绘来使其变小，则可能会降低清晰度。</span><span class="sxs-lookup"><span data-stu-id="54c86-149">Complex icons created at a larger size may lose clarity if resized to be smaller without redraw.</span></span> |
|<span data-ttu-id="54c86-p110">为辅助功能使用白色填充。图标中的大部分对象都需使用白色背景，以使其在 Office UI 主题中以及高对比度模式下清晰可辨。</span><span class="sxs-lookup"><span data-stu-id="54c86-p110">Use a white fill for accessibility. Most objects in your icons will require a white background to be legible across Office UI themes and in high-contrast modes.</span></span>  ||
|<span data-ttu-id="54c86-152">使用具有透明背景的 PNG 格式。</span><span class="sxs-lookup"><span data-stu-id="54c86-152">Use the PNG format with a transparent background.</span></span> ||
|<span data-ttu-id="54c86-153">避免在图标中使用可本地化的内容，包括印刷字符、段落标记指示和问号。</span><span class="sxs-lookup"><span data-stu-id="54c86-153">Avoid localizable content in your icons, including typographic characters, indications of paragraph rags, and question marks.</span></span> ||



## <a name="icon-size-recommendations-and-requirements"></a><span data-ttu-id="54c86-154">图标大小的建议和要求</span><span class="sxs-lookup"><span data-stu-id="54c86-154">Icon size recommendations and requirements</span></span>

<span data-ttu-id="54c86-155">Office 桌面图标是位图图像。</span><span class="sxs-lookup"><span data-stu-id="54c86-155">Office desktop icons are bitmap images.</span></span> <span data-ttu-id="54c86-156">根据用户的 DPI 设置和触摸模式将呈现不同的大小。</span><span class="sxs-lookup"><span data-stu-id="54c86-156">Different sizes will render depending on the user's DPI setting and touch mode.</span></span> <span data-ttu-id="54c86-157">包括所有八种支持的大小，可在所有受支持的解决方案和上下文中创建最佳体验。</span><span class="sxs-lookup"><span data-stu-id="54c86-157">Include all eight supported sizes to create the best experience in all supported resolutions and contexts.</span></span> <span data-ttu-id="54c86-158">以下是受支持的大小 - 三种是必需的：</span><span class="sxs-lookup"><span data-stu-id="54c86-158">The following are the supported sizes - three are required:</span></span>

- <span data-ttu-id="54c86-159">16 像素（必需）</span><span class="sxs-lookup"><span data-stu-id="54c86-159">16 px (Required)</span></span>
- <span data-ttu-id="54c86-160">20 像素</span><span class="sxs-lookup"><span data-stu-id="54c86-160">20 px</span></span>
- <span data-ttu-id="54c86-161">24 像素</span><span class="sxs-lookup"><span data-stu-id="54c86-161">24 px</span></span>
- <span data-ttu-id="54c86-162">32 像素（必需）</span><span class="sxs-lookup"><span data-stu-id="54c86-162">32 px (Required)</span></span>
- <span data-ttu-id="54c86-163">40 像素</span><span class="sxs-lookup"><span data-stu-id="54c86-163">40 px</span></span>
- <span data-ttu-id="54c86-164">48 像素</span><span class="sxs-lookup"><span data-stu-id="54c86-164">48 px</span></span>
- <span data-ttu-id="54c86-165">64 像素（建议，最适用于 Mac）</span><span class="sxs-lookup"><span data-stu-id="54c86-165">64 px (Recommended, best for Mac)</span></span>
- <span data-ttu-id="54c86-166">80 像素（必需）</span><span class="sxs-lookup"><span data-stu-id="54c86-166">80 px (Required)</span></span>  

<span data-ttu-id="54c86-167">确保根据每个尺寸重新绘制你的图标，而非将其缩小。</span><span class="sxs-lookup"><span data-stu-id="54c86-167">Make sure to redraw your icons for each size rather than shrink them to fit.</span></span>

![显示调整图标大小而非缩小图标的建议的图示](../images/icon-resizing.png)

<!--
The following table shows the icon sizes that render for different modes at different DPI settings.

|DPI |**Small**||**Medium**||**Large**||**Extra large**|
|:---|:---|:---|:---|:---|:---|:---|:---|
|    |**Mouse**|**Touch**|**Mouse**|**Touch**|**Mouse**|**Touch**|-|
|100%|16px|20px|24px||32px|40px|48px|
|125%|20px|24px|||40px|48px|60px|
|150%|24px|24px|36px||48px|48px|72px|
|200%|32px|40px|48px||64px|80px|96px|
|250%|40px||||80px||120px|
|300%|48px||||96px||144px

> [!NOTE]
> At DPI settings of 150% or greater, the icon does not get swapped out for a larger size when Touch mode is engaged. At DPI settings greater than 250%, Touch mode is turned off by default.

The following table lists the locations for certain icon sizes.

|Location|100% DPI|200% DPI|250% DPI|
|:-------|:-------|:-------|:-------|
|Small ribbon button|16px|32px|40px|
|Contextual menu|16px|32px|40px|
|Quick access toolbar (QAT)|16px|32px|40px|
|Large ribbon icon|32px|64px|80px|

-->

## <a name="icon-anatomy-and-layout"></a><span data-ttu-id="54c86-169">图标分析和布局</span><span class="sxs-lookup"><span data-stu-id="54c86-169">Icon anatomy and layout</span></span>

<span data-ttu-id="54c86-p112">Office 图标通常是由具有操作和概念修饰符的基本元素构成的。 操作修饰符表示诸如添加、打开、新建或关闭等的概念。概念修饰符表示图标的状态、更改或说明。</span><span class="sxs-lookup"><span data-stu-id="54c86-p112">Office icons are typically comprised of a base element with action and conceptual modifiers overlayed. Action modifiers represent concepts such as add, open, new, or close. Conceptual modifiers represent status, alteration, or a description of the icon.</span></span> 

<span data-ttu-id="54c86-p113">若要创建与 Office UI 相符的命令，请遵循基本元素和修饰符的布局准则。这将确保命令看起来具有专业性，且客户将信任你的外接程序。如果出现未按这些准则进行操作的情况，则这些操作应该是有意为之。</span><span class="sxs-lookup"><span data-stu-id="54c86-p113">To create commands that align with the Office UI, follow layout guidelines for the base element and modifiers. This ensures that your commands look professional and that your customers will trust your add-in. If you make exceptions to these guidelines, do so intentionally.</span></span>

<span data-ttu-id="54c86-176">以下图像显示 Office 图标中的基本元素和修饰符的布局。</span><span class="sxs-lookup"><span data-stu-id="54c86-176">The following image shows the layout of base elements and modifiers in an Office icon.</span></span>

![显示处于中间位置的图标基本元素的图像，其中修饰符位于右下方，操作修饰符位于左上方](../images/icon-layouts.png)

- <span data-ttu-id="54c86-178">将基本元素置于像素框架的中间位置，并在其周围填充空白。</span><span class="sxs-lookup"><span data-stu-id="54c86-178">Center base elements in the pixel frame with empty padding all around.</span></span>
- <span data-ttu-id="54c86-179">将操作修饰符置于左上方。</span><span class="sxs-lookup"><span data-stu-id="54c86-179">Place action modifiers on the top left.</span></span> 
- <span data-ttu-id="54c86-180">将概念修饰符置于右下方。</span><span class="sxs-lookup"><span data-stu-id="54c86-180">Place conceptual modifiers on the bottom right.</span></span>
- <span data-ttu-id="54c86-p114">限制图标中的元素数。在 32 像素中，将修饰符数限制为最多两个。在 16 像素中，将修饰符数限制为一个。</span><span class="sxs-lookup"><span data-stu-id="54c86-p114">Limit the number of elements in your icons. At 32px, limit the number of modifiers to a maximum of two. At 16px, limit the number of modifiers to one.</span></span>

###<a name="base-element-padding"></a><span data-ttu-id="54c86-184">基准元素填充</span><span class="sxs-lookup"><span data-stu-id="54c86-184">Base element padding</span></span>
<span data-ttu-id="54c86-p115">放置与大小相一致的基本元素。如果基本元素不能在框架居中显示，则将其对齐到左上方，并将多余的像素保留在右下方。为了获得最佳效果，请应用下表中列出的填充准则。</span><span class="sxs-lookup"><span data-stu-id="54c86-p115">Place base elements consistently across sizes. If base elements can't be centered in the frame, align them to the top left, leaving the extra pixels on the bottom right. For best results, apply the padding guidelines listed in the following table.</span></span>

###<a name="modifiers"></a><span data-ttu-id="54c86-188">修饰符</span><span class="sxs-lookup"><span data-stu-id="54c86-188">Modifiers</span></span>
<span data-ttu-id="54c86-p116">所有修饰符在每个元素间都应具有 1 像素的透明切割区，包括背景。元素不应直接重叠。在规则和边缘之间创建空白空间。修饰符在大小上可能略有不同，但会将这些尺寸作为起点使用。</span><span class="sxs-lookup"><span data-stu-id="54c86-p116">All modifiers should have a 1px transparent cutout between each element, including the background. Elements should not directly overlap. Create whitespace between rules and edges. Modifiers can vary slightly in size, but use these dimensions as a starting point.</span></span>


|<span data-ttu-id="54c86-193">**图标大小**</span><span class="sxs-lookup"><span data-stu-id="54c86-193">**Icon size**</span></span>|<span data-ttu-id="54c86-194">**在基本元素周围填充**</span><span class="sxs-lookup"><span data-stu-id="54c86-194">**Padding around base element**</span></span>|<span data-ttu-id="54c86-195">**修饰符大小**</span><span class="sxs-lookup"><span data-stu-id="54c86-195">**Modifier size**</span></span>|
|:---|:---|:---|
|<span data-ttu-id="54c86-196">16px</span><span class="sxs-lookup"><span data-stu-id="54c86-196">16px</span></span>|<span data-ttu-id="54c86-197">0</span><span class="sxs-lookup"><span data-stu-id="54c86-197">0%</span></span>|<span data-ttu-id="54c86-198">9px</span><span class="sxs-lookup"><span data-stu-id="54c86-198">9px</span></span>|
|<span data-ttu-id="54c86-199">20px</span><span class="sxs-lookup"><span data-stu-id="54c86-199">20px</span></span>|<span data-ttu-id="54c86-200">1px</span><span class="sxs-lookup"><span data-stu-id="54c86-200">1px</span></span>|<span data-ttu-id="54c86-201">10px</span><span class="sxs-lookup"><span data-stu-id="54c86-201">10px</span></span>|
|<span data-ttu-id="54c86-202">24px</span><span class="sxs-lookup"><span data-stu-id="54c86-202">24px</span></span>|<span data-ttu-id="54c86-203">1px</span><span class="sxs-lookup"><span data-stu-id="54c86-203">1px</span></span>|<span data-ttu-id="54c86-204">12px</span><span class="sxs-lookup"><span data-stu-id="54c86-204">12px</span></span>|
|<span data-ttu-id="54c86-205">32px</span><span class="sxs-lookup"><span data-stu-id="54c86-205">32px</span></span>|<span data-ttu-id="54c86-206">2px</span><span class="sxs-lookup"><span data-stu-id="54c86-206">2px</span></span>|<span data-ttu-id="54c86-207">14px</span><span class="sxs-lookup"><span data-stu-id="54c86-207">14px</span></span>|
|<span data-ttu-id="54c86-208">40px</span><span class="sxs-lookup"><span data-stu-id="54c86-208">40px</span></span>|<span data-ttu-id="54c86-209">2px</span><span class="sxs-lookup"><span data-stu-id="54c86-209">2px</span></span>|<span data-ttu-id="54c86-210">20px</span><span class="sxs-lookup"><span data-stu-id="54c86-210">20px</span></span>|
|<span data-ttu-id="54c86-211">48px</span><span class="sxs-lookup"><span data-stu-id="54c86-211">48px</span></span>|<span data-ttu-id="54c86-212">3px</span><span class="sxs-lookup"><span data-stu-id="54c86-212">3px</span></span>|<span data-ttu-id="54c86-213">22px</span><span class="sxs-lookup"><span data-stu-id="54c86-213">22px</span></span>|
|<span data-ttu-id="54c86-214">64px</span><span class="sxs-lookup"><span data-stu-id="54c86-214">64px</span></span>|<span data-ttu-id="54c86-215">5px</span><span class="sxs-lookup"><span data-stu-id="54c86-215">5px</span></span>|<span data-ttu-id="54c86-216">29px</span><span class="sxs-lookup"><span data-stu-id="54c86-216">29px</span></span>|
|<span data-ttu-id="54c86-217">80px</span><span class="sxs-lookup"><span data-stu-id="54c86-217">80px</span></span>|<span data-ttu-id="54c86-218">5px</span><span class="sxs-lookup"><span data-stu-id="54c86-218">5px</span></span>|<span data-ttu-id="54c86-219">38px</span><span class="sxs-lookup"><span data-stu-id="54c86-219">38px</span></span>|


## <a name="icon-colors"></a><span data-ttu-id="54c86-220">图标颜色</span><span class="sxs-lookup"><span data-stu-id="54c86-220">Icon colors</span></span>

> [!NOTE]
> <span data-ttu-id="54c86-221">这些颜色指南适用于[外接程序命令](add-in-commands.md)中使用的功能区图标。</span><span class="sxs-lookup"><span data-stu-id="54c86-221">These color guidelines are for ribbon icons used in [Add-in commands](add-in-commands.md).</span></span> <span data-ttu-id="54c86-222">这些图标不使用 Microsoft UI Fabric 呈现，调色板与 [Microsoft UI Fabric | 颜色 | 共享](https://fluentfabric.azurewebsites.net/#/color/shared)中描述的调色板不同。</span><span class="sxs-lookup"><span data-stu-id="54c86-222">These icons are not rendered with Microsoft UI Fabric and the color palette is different from the palette described at [Microsoft UI Fabric | Colors | Shared](https://fluentfabric.azurewebsites.net/#/color/shared).</span></span>

<span data-ttu-id="54c86-p118">Office 图标具有一个有限的调色板。使用下表中列出的颜色确保与 Office UI 无缝集成。对颜色使用应用以下准则：</span><span class="sxs-lookup"><span data-stu-id="54c86-p118">Office icons have a limited color palette. Use the colors listed in the following table to guarantee seamless integration with the Office UI. Apply the following guidelines to the use of color:</span></span> 

- <span data-ttu-id="54c86-p119">使用颜色传达图标含义，而非只是用作修饰。图标颜色应突出显示或强调操作、状态或明确区分标记的元素。</span><span class="sxs-lookup"><span data-stu-id="54c86-p119">Use color to communicate meaning rather than for embellishment. It should highlight or emphasize an action, status, or an element that explicitly differentiates the mark.</span></span>  
- <span data-ttu-id="54c86-p120">如有可能，除灰色外仅使用其他一种颜色。将附加颜色限制为最多两种。</span><span class="sxs-lookup"><span data-stu-id="54c86-p120">If possible, use only one additional color beyond gray. Limit additional colors to two at the most.</span></span>
- <span data-ttu-id="54c86-p121">所有图标大小中的颜色应具有一致的外观。Office 图标针对不同的图标大小具有略微不同的调色板。16 像素和更小的图标稍暗，而与 32 像素和更大的图标相比更亮。除了这些细微的调整以外，颜色的差别体现在大小上。</span><span class="sxs-lookup"><span data-stu-id="54c86-p121">Colors should have a consistent appearance in all icon sizes. Office icons have slightly different color palettes for different icon sizes. 16px and smaller icons are slightly darker and more vibrant than 32px and larger icons. Without these subtle adjustments, colors appear to vary across sizes.</span></span>   

|<span data-ttu-id="54c86-234">**颜色名称**</span><span class="sxs-lookup"><span data-stu-id="54c86-234">**Color name**</span></span>|<span data-ttu-id="54c86-235">**RGB**</span><span class="sxs-lookup"><span data-stu-id="54c86-235">**RGB**</span></span>|<span data-ttu-id="54c86-236">**十六进制**</span><span class="sxs-lookup"><span data-stu-id="54c86-236">**Hex**</span></span>|<span data-ttu-id="54c86-237">**颜色**</span><span class="sxs-lookup"><span data-stu-id="54c86-237">**Color**</span></span>|<span data-ttu-id="54c86-238">**类别**</span><span class="sxs-lookup"><span data-stu-id="54c86-238">**Category**</span></span>|
|:---|:---|:---|:---|:---|
|<span data-ttu-id="54c86-239">文本灰色 (80)</span><span class="sxs-lookup"><span data-stu-id="54c86-239">Text Gray (80)</span></span>|<span data-ttu-id="54c86-240">80、80、80</span><span class="sxs-lookup"><span data-stu-id="54c86-240">80, 80, 80</span></span>|<span data-ttu-id="54c86-241">#505050</span><span class="sxs-lookup"><span data-stu-id="54c86-241">#505050</span></span>| ![文本灰色 80 彩色图像](../images/color-text-gray-80.png) |<span data-ttu-id="54c86-243">文本</span><span class="sxs-lookup"><span data-stu-id="54c86-243">Text</span></span>|
|<span data-ttu-id="54c86-244">文本灰色 (95)</span><span class="sxs-lookup"><span data-stu-id="54c86-244">Text Gray (95)</span></span>|<span data-ttu-id="54c86-245">95、95、95</span><span class="sxs-lookup"><span data-stu-id="54c86-245">95, 95, 95</span></span>|<span data-ttu-id="54c86-246">#5F5F5F</span><span class="sxs-lookup"><span data-stu-id="54c86-246">#5F5F5F</span></span>| ![文本灰色 95 彩色图像](../images/color-text-gray-95.png) |<span data-ttu-id="54c86-248">文本</span><span class="sxs-lookup"><span data-stu-id="54c86-248">Text</span></span>|
|<span data-ttu-id="54c86-249">文本灰色 (105)</span><span class="sxs-lookup"><span data-stu-id="54c86-249">Text Gray (105)</span></span>|<span data-ttu-id="54c86-250">105、105、105</span><span class="sxs-lookup"><span data-stu-id="54c86-250">105, 105, 105</span></span>|<span data-ttu-id="54c86-251">#696969</span><span class="sxs-lookup"><span data-stu-id="54c86-251">#696969</span></span>| ![文本灰色 105 彩色图像](../images/color-text-gray-105.png) |<span data-ttu-id="54c86-253">文本</span><span class="sxs-lookup"><span data-stu-id="54c86-253">Text</span></span>|
|<span data-ttu-id="54c86-254">深灰色 32</span><span class="sxs-lookup"><span data-stu-id="54c86-254">Dark Gray 32</span></span>|<span data-ttu-id="54c86-255">128、128、128</span><span class="sxs-lookup"><span data-stu-id="54c86-255">128, 128, 128</span></span>|<span data-ttu-id="54c86-256">#808080</span><span class="sxs-lookup"><span data-stu-id="54c86-256">#808080</span></span>| ![深灰色 32 彩色图像](../images/color-dark-gray-32.png) |<span data-ttu-id="54c86-258">32 及以上</span><span class="sxs-lookup"><span data-stu-id="54c86-258">32 and above</span></span>|
|<span data-ttu-id="54c86-259">中灰色 32</span><span class="sxs-lookup"><span data-stu-id="54c86-259">Medium Gray 32</span></span>|<span data-ttu-id="54c86-260">158、158、158</span><span class="sxs-lookup"><span data-stu-id="54c86-260">158, 158, 158</span></span>|<span data-ttu-id="54c86-261">#9E9E9E</span><span class="sxs-lookup"><span data-stu-id="54c86-261">#9E9E9E</span></span>| ![中灰色 32 彩色图像](../images/color-medium-gray-32.png) |<span data-ttu-id="54c86-263">32 及以上</span><span class="sxs-lookup"><span data-stu-id="54c86-263">32 and above</span></span>|
|<span data-ttu-id="54c86-264">浅灰色所有</span><span class="sxs-lookup"><span data-stu-id="54c86-264">Light Gray ALL</span></span>|<span data-ttu-id="54c86-265">179、179、179</span><span class="sxs-lookup"><span data-stu-id="54c86-265">179, 179, 179</span></span>|<span data-ttu-id="54c86-266">#B3B3B3</span><span class="sxs-lookup"><span data-stu-id="54c86-266">#B3B3B3</span></span>| ![浅灰色所有彩色图像](../images/color-light-gray-all.png) |<span data-ttu-id="54c86-268">所有大小</span><span class="sxs-lookup"><span data-stu-id="54c86-268">All sizes</span></span>|
|<span data-ttu-id="54c86-269">深灰色 16</span><span class="sxs-lookup"><span data-stu-id="54c86-269">Dark Gray 16</span></span>|<span data-ttu-id="54c86-270">114、114、114</span><span class="sxs-lookup"><span data-stu-id="54c86-270">114, 114, 114</span></span>|<span data-ttu-id="54c86-271">#727272</span><span class="sxs-lookup"><span data-stu-id="54c86-271">#727272</span></span>| ![深灰色 16 彩色图像](../images/color-dark-gray-16.png) |<span data-ttu-id="54c86-273">16 及以下</span><span class="sxs-lookup"><span data-stu-id="54c86-273">16 and below</span></span>|
|<span data-ttu-id="54c86-274">中灰色 16</span><span class="sxs-lookup"><span data-stu-id="54c86-274">Medium Gray 16</span></span>|<span data-ttu-id="54c86-275">144、144、144</span><span class="sxs-lookup"><span data-stu-id="54c86-275">144, 144, 144</span></span>|<span data-ttu-id="54c86-276">#909090</span><span class="sxs-lookup"><span data-stu-id="54c86-276">#909090</span></span>| ![中灰色 16 彩色图像](../images/color-medium-gray-16.png) |<span data-ttu-id="54c86-278">16 及以下</span><span class="sxs-lookup"><span data-stu-id="54c86-278">16 and below</span></span>|
|<span data-ttu-id="54c86-279">蓝色 32</span><span class="sxs-lookup"><span data-stu-id="54c86-279">Blue 32</span></span>|<span data-ttu-id="54c86-280">77、130、184</span><span class="sxs-lookup"><span data-stu-id="54c86-280">77, 130, 184</span></span>|<span data-ttu-id="54c86-281">#4d82B8</span><span class="sxs-lookup"><span data-stu-id="54c86-281">#4d82B8</span></span>| ![蓝色 32 彩色图像](../images/color-blue-32.png) |<span data-ttu-id="54c86-283">32 及以上</span><span class="sxs-lookup"><span data-stu-id="54c86-283">32 and above</span></span>|
|<span data-ttu-id="54c86-284">蓝色 16</span><span class="sxs-lookup"><span data-stu-id="54c86-284">Blue 16</span></span>|<span data-ttu-id="54c86-285">74、125、177</span><span class="sxs-lookup"><span data-stu-id="54c86-285">74, 125, 177</span></span>|<span data-ttu-id="54c86-286">#4A7DB1</span><span class="sxs-lookup"><span data-stu-id="54c86-286">#4A7DB1</span></span>| ![蓝色 16 彩色图像](../images/color-blue-16.png) |<span data-ttu-id="54c86-288">16 及以下</span><span class="sxs-lookup"><span data-stu-id="54c86-288">16 and below</span></span>|
|<span data-ttu-id="54c86-289">黄色所有</span><span class="sxs-lookup"><span data-stu-id="54c86-289">Yellow ALL</span></span>|<span data-ttu-id="54c86-290">234、194、130</span><span class="sxs-lookup"><span data-stu-id="54c86-290">234, 194, 130</span></span>|<span data-ttu-id="54c86-291">#EAC282</span><span class="sxs-lookup"><span data-stu-id="54c86-291">#EAC282</span></span>| ![黄色所有彩色图像](../images/color-yellow-all.png) |<span data-ttu-id="54c86-293">所有大小</span><span class="sxs-lookup"><span data-stu-id="54c86-293">All sizes</span></span>|
|<span data-ttu-id="54c86-294">橙色 32</span><span class="sxs-lookup"><span data-stu-id="54c86-294">Orange 32</span></span>|<span data-ttu-id="54c86-295">231、142、70</span><span class="sxs-lookup"><span data-stu-id="54c86-295">231, 142, 70</span></span>|<span data-ttu-id="54c86-296">#E78E46</span><span class="sxs-lookup"><span data-stu-id="54c86-296">#E78E46</span></span>| ![橙色 32 彩色图像](../images/color-orange-32.png) |<span data-ttu-id="54c86-298">32 及以上</span><span class="sxs-lookup"><span data-stu-id="54c86-298">32 and above</span></span>|
|<span data-ttu-id="54c86-299">橙色 16</span><span class="sxs-lookup"><span data-stu-id="54c86-299">Orange 16</span></span>|<span data-ttu-id="54c86-300">227、142、70</span><span class="sxs-lookup"><span data-stu-id="54c86-300">227, 142, 70</span></span>|<span data-ttu-id="54c86-301">#E3751C</span><span class="sxs-lookup"><span data-stu-id="54c86-301">#E3751C</span></span>| ![橙色 16 彩色图像](../images/color-orange-16.png) |<span data-ttu-id="54c86-303">16 及以下</span><span class="sxs-lookup"><span data-stu-id="54c86-303">16 and below</span></span>|
|<span data-ttu-id="54c86-304">粉色所有</span><span class="sxs-lookup"><span data-stu-id="54c86-304">Pink ALL</span></span>|<span data-ttu-id="54c86-305">230、132、151</span><span class="sxs-lookup"><span data-stu-id="54c86-305">230, 132, 151</span></span>|<span data-ttu-id="54c86-306">#E68497</span><span class="sxs-lookup"><span data-stu-id="54c86-306">#E68497</span></span>| ![粉色所有彩色图像](../images/color-pink-all.png) |<span data-ttu-id="54c86-308">所有大小</span><span class="sxs-lookup"><span data-stu-id="54c86-308">All sizes</span></span>|
|<span data-ttu-id="54c86-309">绿色 32</span><span class="sxs-lookup"><span data-stu-id="54c86-309">Green 32</span></span>|<span data-ttu-id="54c86-310">118、167、151</span><span class="sxs-lookup"><span data-stu-id="54c86-310">118, 167, 151</span></span>|<span data-ttu-id="54c86-311">#76A797</span><span class="sxs-lookup"><span data-stu-id="54c86-311">#76A797</span></span>| ![绿色 32 彩色图像](../images/color-green-32.png) |<span data-ttu-id="54c86-313">32 及以上</span><span class="sxs-lookup"><span data-stu-id="54c86-313">32 and above</span></span>|
|<span data-ttu-id="54c86-314">绿色 16</span><span class="sxs-lookup"><span data-stu-id="54c86-314">Green 16</span></span>|<span data-ttu-id="54c86-315">104、164、144</span><span class="sxs-lookup"><span data-stu-id="54c86-315">104, 164, 144</span></span>|<span data-ttu-id="54c86-316">#68A490</span><span class="sxs-lookup"><span data-stu-id="54c86-316">#68A490</span></span>| ![绿色 16 彩色图像](../images/color-green-16.png) |<span data-ttu-id="54c86-318">16 及以下</span><span class="sxs-lookup"><span data-stu-id="54c86-318">16 and below</span></span>|
|<span data-ttu-id="54c86-319">红色 32</span><span class="sxs-lookup"><span data-stu-id="54c86-319">Red 32</span></span>|<span data-ttu-id="54c86-320">216、99、68</span><span class="sxs-lookup"><span data-stu-id="54c86-320">216, 99, 68</span></span>|<span data-ttu-id="54c86-321">#D86344</span><span class="sxs-lookup"><span data-stu-id="54c86-321">#D86344</span></span>| ![红色 32 彩色图像](../images/color-red-32.png) |<span data-ttu-id="54c86-323">32 及以上</span><span class="sxs-lookup"><span data-stu-id="54c86-323">32 and above</span></span>|
|<span data-ttu-id="54c86-324">红色 16</span><span class="sxs-lookup"><span data-stu-id="54c86-324">Red 16</span></span>|<span data-ttu-id="54c86-325">214、85、50</span><span class="sxs-lookup"><span data-stu-id="54c86-325">214, 85, 50</span></span>|<span data-ttu-id="54c86-326">#D65532</span><span class="sxs-lookup"><span data-stu-id="54c86-326">#D65532</span></span>| ![红色 16 彩色图像](../images/color-red-16.png) |<span data-ttu-id="54c86-328">16 及以下</span><span class="sxs-lookup"><span data-stu-id="54c86-328">16 and below</span></span>|
|<span data-ttu-id="54c86-329">紫色 32</span><span class="sxs-lookup"><span data-stu-id="54c86-329">Purple 32</span></span>|<span data-ttu-id="54c86-330">152、104、185</span><span class="sxs-lookup"><span data-stu-id="54c86-330">152, 104, 185</span></span>|<span data-ttu-id="54c86-331">#9868B9</span><span class="sxs-lookup"><span data-stu-id="54c86-331">#9868B9</span></span>| ![紫色 32 彩色图像](../images/color-purple-32.png) |<span data-ttu-id="54c86-333">32 及以上</span><span class="sxs-lookup"><span data-stu-id="54c86-333">32 and above</span></span>|
|<span data-ttu-id="54c86-334">紫色 16</span><span class="sxs-lookup"><span data-stu-id="54c86-334">Purple 16</span></span>|<span data-ttu-id="54c86-335">137、89、171</span><span class="sxs-lookup"><span data-stu-id="54c86-335">137, 89, 171</span></span>|<span data-ttu-id="54c86-336">#8959AB</span><span class="sxs-lookup"><span data-stu-id="54c86-336">#8959AB</span></span>| ![紫色 16 彩色图像](../images/color-purple-16.png) |<span data-ttu-id="54c86-338">16 及以下</span><span class="sxs-lookup"><span data-stu-id="54c86-338">16 and below</span></span>|


## <a name="icons-in-high-contrast-modes"></a><span data-ttu-id="54c86-339">高对比度模式下的图标</span><span class="sxs-lookup"><span data-stu-id="54c86-339">Icons in high contrast modes</span></span>

<span data-ttu-id="54c86-p122">Office 图标设计为在高对比度模式中完美呈现。前景元素与最大化易读性和启用重新着色的背景明显不同。在高对比度模式下，Office 会使用小于 190 的红色、绿色或蓝色值直到全黑，为任何像素的图标重新着色。其他所有像素都将是白色的。换言之，每个评估的 RGB 通道中的 0-189 值表示为黑色，而 190-255 值表示为白色。其他高对比度主题则使用相同的 190 阈值但不同的规则进行重新着色。例如，高对比度白色主题会将所有大于 190 的像素重新着色为不透明，而将所有其他像素重新着色为透明。应用下面的规则以最大化高对比度设置中的可读性。</span><span class="sxs-lookup"><span data-stu-id="54c86-p122">Office icons are designed to render well in high contrast modes. Foreground elements are well differentiated from backgrounds to maximize legibility and enable recoloring. In high contrast modes, Office will recolor any pixel of your icon with a red, green, or blue value less than 190 to full black. All other pixels will be white. In other words, each RGB channel is assessed where 0-189 values are black and 190-255 values are white. Other high-contrast themes recolor using the same 190 value threshold but with different rules. For example, the high-contrast white theme will recolor all pixels greater than 190 opaque but all other pixels as transparent. Apply the following guidelines to maximize legibility in high-contrast settings:</span></span>

- <span data-ttu-id="54c86-348">旨在以 190 阈值区分前景和背景元素。</span><span class="sxs-lookup"><span data-stu-id="54c86-348">Aim to differentiate foreground and background elements along the 190 value threshold.</span></span>
- <span data-ttu-id="54c86-349">遵循 Office 图标视觉样式。</span><span class="sxs-lookup"><span data-stu-id="54c86-349">Follow Office icon visual styles.</span></span>
- <span data-ttu-id="54c86-350">使用图标调色板中的颜色。</span><span class="sxs-lookup"><span data-stu-id="54c86-350">Use colors from our icon palette.</span></span>
- <span data-ttu-id="54c86-351">避免使用渐变。</span><span class="sxs-lookup"><span data-stu-id="54c86-351">Avoid the use of gradients.</span></span>
- <span data-ttu-id="54c86-352">避免使用值相似的颜色块。</span><span class="sxs-lookup"><span data-stu-id="54c86-352">Avoid large blocks of color with similar values.</span></span>

## <a name="see-also"></a><span data-ttu-id="54c86-353">另请参阅</span><span class="sxs-lookup"><span data-stu-id="54c86-353">See also</span></span>

- [<span data-ttu-id="54c86-354">加载项开发最佳做法</span><span class="sxs-lookup"><span data-stu-id="54c86-354">Add-in development best practices</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="54c86-355">Excel、Word 和 PowerPoint 的加载项命令</span><span class="sxs-lookup"><span data-stu-id="54c86-355">Add-in commands for Excel, Word, and PowerPoint</span></span>](../design/add-in-commands.md)




- <span data-ttu-id="54c86-p123">避免依赖徽标或品牌传达外接程序命令应起到的作用。品牌标志在较小的图标尺寸上和应用很多修饰符后并非总具有识别性。品牌标志经常与 Office 功能区图标样式冲突，并可能在饱和的环境中过度吸引用户的注意力。</span><span class="sxs-lookup"><span data-stu-id="54c86-p123">Avoid relying on your logo or brand to communicate what an add-in command does. Brand marks aren't always recognizable at smaller icon sizes and when modifiers are applied. Brand marks often conflict with Office ribbon icon styles, and can compete for user attention in a saturated environment.</span></span>


