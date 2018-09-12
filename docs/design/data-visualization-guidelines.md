---
title: Office 加载项的数据可视化样式指南
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 27de6b6b2f4352488ad8f63c3b6e1250cbfbb324
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945790"
---
# <a name="data-visualization-style-guidelines-for-office-add-ins"></a><span data-ttu-id="f6257-102">Office 加载项的数据可视化样式指南</span><span class="sxs-lookup"><span data-stu-id="f6257-102">Data visualization style guidelines for Office Add-ins</span></span>

<span data-ttu-id="f6257-p101">良好的数据可视化效果可帮助用户找到数据见解。他们可以使用这些见解来讲述具有说服力的故事。本文提供了准则，以帮助你在适用于 Excel 和其他 Office 应用的外接程序中设计有效的数据可视化。</span><span class="sxs-lookup"><span data-stu-id="f6257-p101">Good data visualizations help users find insights in their data. They can use those insights to tell stories that inform and persuade. This article provides guidelines to help you design effective data visualizations in your add-ins for Excel and other Office apps.</span></span>

<span data-ttu-id="f6257-p102">我们建议使用 [Office UI Fabric](https://developer.microsoft.com/fabric) 来创建数据可视化的镶边。Office UI Fabric 包含可与 Office 外观无缝集成的样式和组件。</span><span class="sxs-lookup"><span data-stu-id="f6257-p102">We recommend that you use [Office UI Fabric](https://developer.microsoft.com/fabric) to create the chrome for your data visualizations. Office UI Fabric includes styles and components that integrate seamlessly with the Office look and feel.</span></span> 

<!--The following figure shows a data visualization in an add-in that uses Fabric.

![Image of a data visualization with Fabric elements applied**](../images/fabric-data-visualization.png) 

-->

## <a name="data-visualization-elements"></a><span data-ttu-id="f6257-108">数据可视化元素</span><span class="sxs-lookup"><span data-stu-id="f6257-108">Data visualization elements</span></span>

<span data-ttu-id="f6257-109">数据可视化共享一个通用框架、常见的视觉对象和交互式元素，包括标题、标签和数据绘图，如下图所示。</span><span class="sxs-lookup"><span data-stu-id="f6257-109">Data visualizations share a general framework and common visual and interactive elements, including titles, labels, and data plots, as shown in the following figures.</span></span>

<span data-ttu-id="f6257-110">![标记了标题、轴、图例和绘图区的折线图的图像](../images/data-visualization-line-chart.png)
![标记了轴、网格线、图例和数据绘图的柱形图的图像](../images/data-visualization-column-chart.png)</span><span class="sxs-lookup"><span data-stu-id="f6257-110">![Image of a line chart with title, axes, legend, and plot area labeled](../images/data-visualization-line-chart.png)
![Image of a column chart with axes, gridlines, legend, and data plot labeled](../images/data-visualization-column-chart.png)</span></span>

### <a name="chart-titles"></a><span data-ttu-id="f6257-111">图表标题</span><span class="sxs-lookup"><span data-stu-id="f6257-111">Chart titles</span></span>

<span data-ttu-id="f6257-112">请遵循图表标题的以下准则：</span><span class="sxs-lookup"><span data-stu-id="f6257-112">Follow these guidelines for chart titles:</span></span>

- <span data-ttu-id="f6257-p103">使图表标题便于阅读。设定其位置以创建相对于其余图表的清晰视觉对象层次结构。</span><span class="sxs-lookup"><span data-stu-id="f6257-p103">Make your chart titles easily readable. Position them to create a clear visual hierarchy in relation to the rest of the chart.</span></span>
- <span data-ttu-id="f6257-p104">一般情况下，使用句子大写（大写第一个字词）。若要创建对比度或强化层次结构，可以全部使用大写，但应谨慎使用全部大写。</span><span class="sxs-lookup"><span data-stu-id="f6257-p104">In general, use sentence capitalization (capitalize the first word). To create contrast or to reinforce hierarchies, you can use all caps, but all caps should be used sparingly.</span></span>
- <span data-ttu-id="f6257-p105">纳入 [Office UI Fabric 类型校正](https://developer.microsoft.com/fabric#/styles/typography)使图表与使用 Segoe 的 Office UI 保持一致。你还可以使用不同的字样来区分图表内容和 UI。</span><span class="sxs-lookup"><span data-stu-id="f6257-p105">Incorporate the [Office UI Fabric type ramp](https://developer.microsoft.com/fabric#/styles/typography) to make your charts consistent with the Office UI, which uses Segoe. You can also use a different typeface to differentiate chart content from the UI.</span></span>
- <span data-ttu-id="f6257-119">使用带有大型计数器的 sans-serif 字样。</span><span class="sxs-lookup"><span data-stu-id="f6257-119">Use sans-serif typefaces with large counters.</span></span>

<span data-ttu-id="f6257-p106">下面的示例显示图表标题中使用的 serif 和 sans-serif 字样。请留意如何通过缩放对比度和空白的有效使用来构建强大的可视化层次结构。</span><span class="sxs-lookup"><span data-stu-id="f6257-p106">The following examples show serif and sans-serif typefaces used in chart titles. Notice how the scale contrast and effective use of white space create a strong visual hierarchy.</span></span>

<span data-ttu-id="f6257-122">![采用 serif 字体的数据可视化的图像](../images/data-visualization-serif.png)
![采用 sans-serif 字体的数据可视化的图像](../images/data-visualization-sans-serif.png)</span><span class="sxs-lookup"><span data-stu-id="f6257-122">![Image of a data visualization with serif font](../images/data-visualization-serif.png)
![Image of a data visualization with sans-serif font](../images/data-visualization-sans-serif.png)</span></span>

### <a name="axis-labels"></a><span data-ttu-id="f6257-123">轴标签</span><span class="sxs-lookup"><span data-stu-id="f6257-123">Axis labels</span></span>

<span data-ttu-id="f6257-p107">请确保轴标签颜色足够深，以便可以清楚地阅读，并且具有足够的文本和背景色对比度。请确保颜色不要过深，避免比数据墨迹更加突出。</span><span class="sxs-lookup"><span data-stu-id="f6257-p107">Make your axis labels dark enough to read clearly, with adequate contrast ratios between the text and background colors. Make sure that they are not so dark that they compete with data ink.</span></span>

<span data-ttu-id="f6257-p108">浅灰色轴标签效果最佳。如果使用的是 Fabric，请参阅[中性色调色板](https://developer.microsoft.com/fabric#/styles/colors)。</span><span class="sxs-lookup"><span data-stu-id="f6257-p108">Light grays are most effective for axis labels. If you’re using Fabric, see the [Neutral Colors palette](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

### <a name="data-ink"></a><span data-ttu-id="f6257-128">数据墨迹</span><span class="sxs-lookup"><span data-stu-id="f6257-128">Data ink</span></span>

<span data-ttu-id="f6257-p109">表示图表中的实际数据的像素被称为数据墨迹。这应该是可视化的中心焦点。避免使用投影、过粗边框或不必要的使数据失真或影响数据显示效果的设计元素。仅当数据值与颜色值关联时使用渐变。避免使用三维图表，除非可测量的目标值绑定到第三维度。</span><span class="sxs-lookup"><span data-stu-id="f6257-p109">The pixels that represent the actual data in a chart are referred to as data ink. This should be the central focus of the visualization. Avoid the use of drop shadows, heavy outlines, or unnecessary design elements that distort or compete with the data. Use gradients only when data values are tied to color values. Avoid three-dimensional charts unless a measurable, objective value is bound to a third dimension.</span></span>

### <a name="color"></a><span data-ttu-id="f6257-134">颜色</span><span class="sxs-lookup"><span data-stu-id="f6257-134">Color</span></span>

<span data-ttu-id="f6257-p110">选择遵循操作系统或应用程序主题的颜色，而不是硬编码的颜色。同时，确保所应用的颜色不会使数据失真。数据可视化中的颜色滥用可能会导致数据失真和信息读取不正确。</span><span class="sxs-lookup"><span data-stu-id="f6257-p110">Choose colors that follow operating system or application themes rather than hardcoded colors. At the same time, make sure that the colors you apply do not distort the data. Misuse of color in data visualizations can result in data distortion and incorrect reading of information.</span></span>

<span data-ttu-id="f6257-138">有关在数据可视化中使用颜色的最佳做法，请参阅以下内容：</span><span class="sxs-lookup"><span data-stu-id="f6257-138">For best practices for use of color in data visualizations, see the following:</span></span>


- [<span data-ttu-id="f6257-139">为什么彩虹色不是数据可视化的最佳选择</span><span class="sxs-lookup"><span data-stu-id="f6257-139">Why rainbow colors aren't the best option for data visualizations</span></span>](http://www.poynter.org/2013/why-rainbow-colors-arent-always-the-best-options-for-data-visualizations/224413/)
- [<span data-ttu-id="f6257-140">Color Brewer 2.0：制图的颜色建议</span><span class="sxs-lookup"><span data-stu-id="f6257-140">Color Brewer 2.0: Color Advice for Cartography</span></span>](http://colorbrewer2.org/)
- [<span data-ttu-id="f6257-141">我想要的色调</span><span class="sxs-lookup"><span data-stu-id="f6257-141">I Want Hue</span></span>](http://tools.medialab.sciences-po.fr/iwanthue/)

### <a name="gridlines"></a><span data-ttu-id="f6257-142">网格线</span><span class="sxs-lookup"><span data-stu-id="f6257-142">Gridlines</span></span>

<span data-ttu-id="f6257-p111">要准确读取图表，通常网格线是必不可少的，但应显示为辅助可视元素，用于增强数据墨迹效果，但不会影响数据显示。确保静态网格线较细且颜色较淡，除非专门将其设计用于高对比度的情况。还可以使用交互作用创建在用户与图表交互时上下文中显示的动态、实时网格线。</span><span class="sxs-lookup"><span data-stu-id="f6257-p111">Gridlines are often necessary for accurately reading a chart, but should be presented as a secondary visual element, enhancing the data ink, not competing with it. Make static gridlines thin and light, unless they are designed specifically for high contrast. You can also use interaction to create dynamic, just-in-time gridlines that appear in context when a user interacts with a chart.</span></span>

<span data-ttu-id="f6257-p112">浅灰色网格线效果最佳。如果使用的是 Fabric，请参阅[中性色调色板](https://developer.microsoft.com/fabric#/styles/colors)。</span><span class="sxs-lookup"><span data-stu-id="f6257-p112">Light grays are most effective for gridlines. If you’re using Fabric, see the [Neutral Colors palette](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

<span data-ttu-id="f6257-148">下图显示了带有网格线的数据可视化。</span><span class="sxs-lookup"><span data-stu-id="f6257-148">The following image shows a data visualization with gridlines.</span></span>

![带有网格线的数据可视化的图像](../images/data-visualization-gridlines.png)

### <a name="legends"></a><span data-ttu-id="f6257-150">图例</span><span class="sxs-lookup"><span data-stu-id="f6257-150">Legends</span></span>

<span data-ttu-id="f6257-151">如果需要，请添加图例：</span><span class="sxs-lookup"><span data-stu-id="f6257-151">Add legends if necessary to:</span></span>

- <span data-ttu-id="f6257-152">区分系列</span><span class="sxs-lookup"><span data-stu-id="f6257-152">Distinguish between series</span></span>
- <span data-ttu-id="f6257-153">存在缩放或值的更改</span><span class="sxs-lookup"><span data-stu-id="f6257-153">Present scale or value changes</span></span>

<span data-ttu-id="f6257-p113">请确保图例增强数据墨迹，但不会影响其显示效果。放置图例：</span><span class="sxs-lookup"><span data-stu-id="f6257-p113">Make sure that your legends enhance the data ink and do not compete with it. Place legends:</span></span>


- <span data-ttu-id="f6257-156">如果图表上方的所有图例项大小合适，则默认情况下会在绘图区上方左对齐。</span><span class="sxs-lookup"><span data-stu-id="f6257-156">Flush left above the plot area by default, if all legend items fit above the chart.</span></span>
- <span data-ttu-id="f6257-157">在绘图区的右上角，如果图表上方的所有图例项大小均不合适，请在必要时确保其可滚动。</span><span class="sxs-lookup"><span data-stu-id="f6257-157">On the upper right side of the plot area, if all legend items do not fit above the chart, and make it scrollable, if necessary.</span></span>

<span data-ttu-id="f6257-p114">为了优化可读性和可访问性，将图例标记映射到相关图表形状。例如，将圆形图例标记用于散点图和气泡图图例。将线段图例标记用于折线图。</span><span class="sxs-lookup"><span data-stu-id="f6257-p114">To optimize for readability and accessibility, map legend markers to the relevant chart shape. For example, use circle legend markers for scatter plot and bubble chart legends. Use line segment legend markers for line charts.</span></span>

### <a name="data-labels-and-tooltips"></a><span data-ttu-id="f6257-161">数据标签和工具提示</span><span class="sxs-lookup"><span data-stu-id="f6257-161">Data labels and tooltips</span></span>

<span data-ttu-id="f6257-p115">确保数据标签和工具提示拥有足够的空白和类型变体。使用算法来最小化封闭和冲突。例如，默认情况下，工具提示可能出现在数据点的右侧，但如果检测到右侧边缘，则会出现在左侧。</span><span class="sxs-lookup"><span data-stu-id="f6257-p115">Ensure that data labels and tooltips have adequate white space and type variation. Use algorithms to minimize occlusion and collision. For example, a tooltip might surface to the right of a data point by default, but surface to the left if right edges are detected.</span></span>

## <a name="design-principles"></a><span data-ttu-id="f6257-165">设计原则</span><span class="sxs-lookup"><span data-stu-id="f6257-165">Design principles</span></span>

<span data-ttu-id="f6257-166">Office Design 团队创建了以下设计原则集，我们可在为 Office 产品套件设计新的数据可视化时使用这些原则。</span><span class="sxs-lookup"><span data-stu-id="f6257-166">The Office Design team created the following set of design principles, which we use when designing new data visualizations for the Office product suite.</span></span>

### <a name="visual-design-principles"></a><span data-ttu-id="f6257-167">视觉对象设计原则</span><span class="sxs-lookup"><span data-stu-id="f6257-167">Visual design principles</span></span>

- <span data-ttu-id="f6257-p116">可视化效果应忠于数据并增强数据，使其易于理解。突出显示数据，仅在需要提供上下文时添加支持元素。避免不必要的装饰（投影、边框等）、图表垃圾或数据失真。</span><span class="sxs-lookup"><span data-stu-id="f6257-p116">Visualizations should honor and enhance the data, making it easy to understand. Highlight the data, adding supporting elements only as needed to provide context. Avoid unnecessary embellishments (drop shadows, outlines, etc), chart junk, or data distortion.</span></span>
- <span data-ttu-id="f6257-p117">可视化效果应通过提供丰富的视觉反馈吸引用户进行浏览。使用成熟的交互模式、接口控件，并清除系统反馈。</span><span class="sxs-lookup"><span data-stu-id="f6257-p117">Visualizations should encourage exploration by providing rich visual feedback. Use well-established interaction patterns, interface controls, and clear system feedback.</span></span>
- <span data-ttu-id="f6257-p118">体现久负盛名的设计原则。使用已制定的版式和可视通信设计原则来增强表单、可读性和含义。</span><span class="sxs-lookup"><span data-stu-id="f6257-p118">Embody time-honored design principles. Use established typographic and visual communication design principles to enhance form, readability, and meaning.</span></span>

### <a name="interaction-design-principles"></a><span data-ttu-id="f6257-175">交互设计原则</span><span class="sxs-lookup"><span data-stu-id="f6257-175">Interaction design principles</span></span>

- <span data-ttu-id="f6257-176">设计为允许进行浏览。</span><span class="sxs-lookup"><span data-stu-id="f6257-176">Design to allow for exploration.</span></span>
- <span data-ttu-id="f6257-177">允许与对象进行直接交互，以展示新见解（例如，通过拖动进行排序）。</span><span class="sxs-lookup"><span data-stu-id="f6257-177">Allow for direct interactions with objects that reveal new insights (sorting via drag, for example).</span></span>
- <span data-ttu-id="f6257-178">使用简单、直接、熟悉的交互模型。</span><span class="sxs-lookup"><span data-stu-id="f6257-178">Use simple, direct, familiar interaction models.</span></span>

<span data-ttu-id="f6257-179">有关如何设计用户友好交互式数据可视化的详细信息，请参阅 [UI 原则和陷阱](http://uitraps.com/)。</span><span class="sxs-lookup"><span data-stu-id="f6257-179">For more information about how to design user-friendly interactive data visualizations, see [UI Tenets and Traps](http://uitraps.com/).</span></span>

### <a name="motion-design-principles"></a><span data-ttu-id="f6257-180">动作设计原则</span><span class="sxs-lookup"><span data-stu-id="f6257-180">Motion design principles</span></span>

<span data-ttu-id="f6257-p119">动作随刺激而产生。视觉元素应以相同的速率朝相同的方向运动。这适用于：</span><span class="sxs-lookup"><span data-stu-id="f6257-p119">Motion follows stimulus. Visual elements should move in the same direction at the same rate. This applies to:</span></span>

- <span data-ttu-id="f6257-184">创建图表</span><span class="sxs-lookup"><span data-stu-id="f6257-184">Chart creation</span></span>
- <span data-ttu-id="f6257-185">从一种图表类型转换到另一种图表类型</span><span class="sxs-lookup"><span data-stu-id="f6257-185">Transition from one chart type to another chart type</span></span>
- <span data-ttu-id="f6257-186">筛选</span><span class="sxs-lookup"><span data-stu-id="f6257-186">Filtering</span></span>
- <span data-ttu-id="f6257-187">排序</span><span class="sxs-lookup"><span data-stu-id="f6257-187">Sorting</span></span>
- <span data-ttu-id="f6257-188">添加或减少数据</span><span class="sxs-lookup"><span data-stu-id="f6257-188">Adding or subtracting data</span></span>
- <span data-ttu-id="f6257-189">对数据进行刷新或切片</span><span class="sxs-lookup"><span data-stu-id="f6257-189">Brushing or slicing data</span></span>
- <span data-ttu-id="f6257-190">重设图表大小</span><span class="sxs-lookup"><span data-stu-id="f6257-190">Resizing a chart</span></span>

<span data-ttu-id="f6257-p120">创建因果关系感知。在暂存动画时：</span><span class="sxs-lookup"><span data-stu-id="f6257-p120">Create a perception of causality. When staging animations:</span></span>

- <span data-ttu-id="f6257-193">一次暂存一个。</span><span class="sxs-lookup"><span data-stu-id="f6257-193">Stage one thing at a time.</span></span> 
- <span data-ttu-id="f6257-194">在更改数据墨迹前，将更改暂存到轴中。</span><span class="sxs-lookup"><span data-stu-id="f6257-194">Stage changes to axes before changes to data ink.</span></span>
- <span data-ttu-id="f6257-195">如果对象以相同的速度朝相同的方向移动，那么可以暂存对象并将其制作成动画组。</span><span class="sxs-lookup"><span data-stu-id="f6257-195">Stage and animate objects as a group if they are moving at the same speed in the same direction.</span></span>
- <span data-ttu-id="f6257-p121">在只有 4-5 个对象的组中暂存数据元素。查看器很难独立跟踪数量超过 4-5 个的对象。</span><span class="sxs-lookup"><span data-stu-id="f6257-p121">Stage data elements in groups of no more than 4-5 objects. Viewers have difficulty tracking more than 4-5 objects independently.</span></span>

<span data-ttu-id="f6257-198">动作赋予涵义。</span><span class="sxs-lookup"><span data-stu-id="f6257-198">Motion adds meaning.</span></span>

- <span data-ttu-id="f6257-199">动画可帮助用户理解对数据的更改，提供上下文，并作为非语言注释层发挥作用。</span><span class="sxs-lookup"><span data-stu-id="f6257-199">Animations increase user comprehension of changes to the data, provide context, and act as a non-verbal annotation layer.</span></span>
- <span data-ttu-id="f6257-200">动作应发生在可视化效果具有含义的坐标空间中。</span><span class="sxs-lookup"><span data-stu-id="f6257-200">Motion should occur in a meaningful coordinate space of the visualization.</span></span>
- <span data-ttu-id="f6257-201">为视觉对象定制动画。</span><span class="sxs-lookup"><span data-stu-id="f6257-201">Tailor the animation to the visual.</span></span> 
- <span data-ttu-id="f6257-202">避免不必要的动画效果。</span><span class="sxs-lookup"><span data-stu-id="f6257-202">Avoid gratuitous animations.</span></span>

<span data-ttu-id="f6257-203">随数据运动。</span><span class="sxs-lookup"><span data-stu-id="f6257-203">Motion follows data.</span></span>

- <span data-ttu-id="f6257-p122">保留数据映射。如果某个区域与度量值关联，请使该区域保持在过渡状态。</span><span class="sxs-lookup"><span data-stu-id="f6257-p122">Preserve data mappings. If an area is tied to a measure, maintain that area in transition.</span></span>
- <span data-ttu-id="f6257-p123">保持统一的动画设计语言。如有可能，请将数据可视化动画映射到现有的 Office 动作设计语言。为类似的图表类型使用相似的动画。</span><span class="sxs-lookup"><span data-stu-id="f6257-p123">Maintain a consistent animation design language. Where possible, map data visualization animation to existing Office motion design language. Use similar animations for similar chart types.</span></span>

## <a name="accessibility-in-data-visualizations"></a><span data-ttu-id="f6257-209">数据可视化中的辅助功能</span><span class="sxs-lookup"><span data-stu-id="f6257-209">Accessibility in data visualizations</span></span>

- <span data-ttu-id="f6257-p124">请勿将颜色用作传达信息的唯一方式。色盲者将无法解读结果。在可以传达信息的前提下，除使用颜色外，还使用形状、大小和纹理。</span><span class="sxs-lookup"><span data-stu-id="f6257-p124">Do not use color as the only way to communicate information. People who are color blind will not be able to interpret the results. Use shape, size and texture in addition to color when possible to communicate information.</span></span>
- <span data-ttu-id="f6257-213">确保所有交互式元素（如按钮或选择列表）均可通过键盘访问。</span><span class="sxs-lookup"><span data-stu-id="f6257-213">Make all interactive elements, such as push buttons or pick lists, accessible from a keyboard.</span></span>
- <span data-ttu-id="f6257-214">将辅助功能事件发送到屏幕阅读器，以通知焦点更改、工具提示等。</span><span class="sxs-lookup"><span data-stu-id="f6257-214">Send accessibility events to screen readers to announce focus changes, tooltips, and so on.</span></span>

## <a name="see-also"></a><span data-ttu-id="f6257-215">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f6257-215">See also</span></span> 

- [<span data-ttu-id="f6257-216">数据和设计：关于准备和可视化信息的简单介绍</span><span class="sxs-lookup"><span data-stu-id="f6257-216">Data + Design: A Simple Introduction to Preparing and Visualizing Information</span></span>](https://infoactive.co/data-design)
- [<span data-ttu-id="f6257-217">构建数据可视化效果的五个最佳库</span><span class="sxs-lookup"><span data-stu-id="f6257-217">The Five Best Libraries for Building Data Visualizations</span></span>](http://www.fastcompany.com/3029760/the-five-best-libraries-for-building-data-vizualizations)
- [<span data-ttu-id="f6257-218">定量信息的视觉显示</span><span class="sxs-lookup"><span data-stu-id="f6257-218">The Visual Display of Quantitative Information</span></span>](https://www.edwardtufte.com/tufte/books_vdqi)
