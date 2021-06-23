---
title: Office 加载项的数据可视化样式指南
description: 获取有关如何在加载项中可视化数据Office一些好的做法。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: aebd0ea8731d099615141e203cc03b2972128c9a
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076348"
---
# <a name="data-visualization-style-guidelines-for-office-add-ins"></a><span data-ttu-id="15021-103">Office 加载项的数据可视化样式指南</span><span class="sxs-lookup"><span data-stu-id="15021-103">Data visualization style guidelines for Office Add-ins</span></span>

<span data-ttu-id="15021-p101">良好的数据可视化效果可帮助用户找到数据见解。他们可以使用这些见解来讲述具有说服力的故事。本文提供了准则，以帮助你在适用于 Excel 和其他 Office 应用的外接程序中设计有效的数据可视化。</span><span class="sxs-lookup"><span data-stu-id="15021-p101">Good data visualizations help users find insights in their data. They can use those insights to tell stories that inform and persuade. This article provides guidelines to help you design effective data visualizations in your add-ins for Excel and other Office apps.</span></span>

<span data-ttu-id="15021-107">我们建议你使用 Fluent [UI](../design/add-in-design.md)为数据可视化创建部件版式。</span><span class="sxs-lookup"><span data-stu-id="15021-107">We recommend that you use [Fluent UI](../design/add-in-design.md) to create the chrome for your data visualizations.</span></span> <span data-ttu-id="15021-108">FluentUI 包括样式和组件，这些样式和组件与Office体验无缝集成。</span><span class="sxs-lookup"><span data-stu-id="15021-108">Fluent UI includes styles and components that integrate seamlessly with the Office look and feel.</span></span>

## <a name="data-visualization-elements"></a><span data-ttu-id="15021-109">数据可视化元素</span><span class="sxs-lookup"><span data-stu-id="15021-109">Data visualization elements</span></span>

<span data-ttu-id="15021-110">数据可视化共享常规框架和常见的视觉和交互式元素，包括标题、标签和数据绘图，如下图所示。</span><span class="sxs-lookup"><span data-stu-id="15021-110">Data visualizations share a general framework and common visual and interactive elements, including titles, labels, and data plots, as shown in the following figure.</span></span>

![带标题、坐标轴、图例和绘图区标签的线图。](../images/excel-charts-visualization.png)

### <a name="chart-titles"></a><span data-ttu-id="15021-112">图表标题</span><span class="sxs-lookup"><span data-stu-id="15021-112">Chart titles</span></span>

<span data-ttu-id="15021-113">请遵循图表标题的以下准则：</span><span class="sxs-lookup"><span data-stu-id="15021-113">Follow these guidelines for chart titles:</span></span>

- <span data-ttu-id="15021-p103">使图表标题便于阅读。设定其位置以创建相对于其余图表的清晰视觉对象层次结构。</span><span class="sxs-lookup"><span data-stu-id="15021-p103">Make your chart titles easily readable. Position them to create a clear visual hierarchy in relation to the rest of the chart.</span></span>
- <span data-ttu-id="15021-p104">一般情况下，使用句子大写（大写第一个字词）。若要创建对比度或强化层次结构，可以全部使用大写，但应谨慎使用全部大写。</span><span class="sxs-lookup"><span data-stu-id="15021-p104">In general, use sentence capitalization (capitalize the first word). To create contrast or to reinforce hierarchies, you can use all caps, but all caps should be used sparingly.</span></span>
- <span data-ttu-id="15021-118">合并[Fluent UI 类型渐变](https://developer.microsoft.com/fluentui#/styles/web/typography)，使图表与使用 Segoe 的 Office UI 保持一致。</span><span class="sxs-lookup"><span data-stu-id="15021-118">Incorporate the [Fluent UI type ramp](https://developer.microsoft.com/fluentui#/styles/web/typography) to make your charts consistent with the Office UI, which uses Segoe.</span></span> <span data-ttu-id="15021-119">你还可以使用不同的字样来区分图表内容和 UI。</span><span class="sxs-lookup"><span data-stu-id="15021-119">You can also use a different typeface to differentiate chart content from the UI.</span></span>
- <span data-ttu-id="15021-120">使用带有大型计数器的 sans-serif 字样。</span><span class="sxs-lookup"><span data-stu-id="15021-120">Use sans-serif typefaces with large counters.</span></span>

### <a name="axis-labels"></a><span data-ttu-id="15021-121">轴标签</span><span class="sxs-lookup"><span data-stu-id="15021-121">Axis labels</span></span>

<span data-ttu-id="15021-p106">请确保轴标签颜色足够深，以便可以清楚地阅读，并且具有足够的文本和背景色对比度。请确保颜色不要过深，避免比数据墨迹更加突出。</span><span class="sxs-lookup"><span data-stu-id="15021-p106">Make your axis labels dark enough to read clearly, with adequate contrast ratios between the text and background colors. Make sure that they are not so dark that they compete with data ink.</span></span>

<span data-ttu-id="15021-124">浅灰色轴标签效果最佳。</span><span class="sxs-lookup"><span data-stu-id="15021-124">Light grays are most effective for axis labels.</span></span> <span data-ttu-id="15021-125">如果你正在使用中性Fluent UI，请参阅中性[颜色调色板](https://developer.microsoft.com/fluentui#/styles/web/colors/neutrals)。</span><span class="sxs-lookup"><span data-stu-id="15021-125">If you're using Fluent UI, see the [Neutral Colors palette](https://developer.microsoft.com/fluentui#/styles/web/colors/neutrals).</span></span>

### <a name="data-ink"></a><span data-ttu-id="15021-126">数据墨迹</span><span class="sxs-lookup"><span data-stu-id="15021-126">Data ink</span></span>

<span data-ttu-id="15021-p108">表示图表中的实际数据的像素被称为数据墨迹。这应该是可视化的中心焦点。避免使用投影、过粗边框或不必要的使数据失真或影响数据显示效果的设计元素。仅当数据值与颜色值关联时使用渐变。避免使用三维图表，除非可测量的目标值绑定到第三维度。</span><span class="sxs-lookup"><span data-stu-id="15021-p108">The pixels that represent the actual data in a chart are referred to as data ink. This should be the central focus of the visualization. Avoid the use of drop shadows, heavy outlines, or unnecessary design elements that distort or compete with the data. Use gradients only when data values are tied to color values. Avoid three-dimensional charts unless a measurable, objective value is bound to a third dimension.</span></span>

### <a name="color"></a><span data-ttu-id="15021-132">颜色</span><span class="sxs-lookup"><span data-stu-id="15021-132">Color</span></span>

<span data-ttu-id="15021-p109">选择遵循操作系统或应用程序主题的颜色，而不是硬编码的颜色。同时，确保所应用的颜色不会使数据失真。数据可视化中的颜色滥用可能会导致数据失真和信息读取不正确。</span><span class="sxs-lookup"><span data-stu-id="15021-p109">Choose colors that follow operating system or application themes rather than hardcoded colors. At the same time, make sure that the colors you apply do not distort the data. Misuse of color in data visualizations can result in data distortion and incorrect reading of information.</span></span>

<span data-ttu-id="15021-136">有关在数据可视化中使用颜色的最佳做法，请参阅以下内容：</span><span class="sxs-lookup"><span data-stu-id="15021-136">For best practices for use of color in data visualizations, see the following:</span></span>

- [<span data-ttu-id="15021-137">为什么彩虹色不是数据可视化的最佳选择</span><span class="sxs-lookup"><span data-stu-id="15021-137">Why rainbow colors aren't the best option for data visualizations</span></span>](https://www.poynter.org/2013/why-rainbow-colors-arent-always-the-best-options-for-data-visualizations/224413/)
- [<span data-ttu-id="15021-138">Color Brewer 2.0：制图的颜色建议</span><span class="sxs-lookup"><span data-stu-id="15021-138">Color Brewer 2.0: Color Advice for Cartography</span></span>](https://colorbrewer2.org/)
- [<span data-ttu-id="15021-139">我想要的色调</span><span class="sxs-lookup"><span data-stu-id="15021-139">I Want Hue</span></span>](https://tools.medialab.sciences-po.fr/iwanthue/)

### <a name="gridlines"></a><span data-ttu-id="15021-140">网格线</span><span class="sxs-lookup"><span data-stu-id="15021-140">Gridlines</span></span>

<span data-ttu-id="15021-p110">要准确读取图表，通常网格线是必不可少的，但应显示为辅助可视元素，用于增强数据墨迹效果，但不会影响数据显示。确保静态网格线较细且颜色较淡，除非专门将其设计用于高对比度的情况。还可以使用交互作用创建在用户与图表交互时上下文中显示的动态、实时网格线。</span><span class="sxs-lookup"><span data-stu-id="15021-p110">Gridlines are often necessary for accurately reading a chart, but should be presented as a secondary visual element, enhancing the data ink, not competing with it. Make static gridlines thin and light, unless they are designed specifically for high contrast. You can also use interaction to create dynamic, just-in-time gridlines that appear in context when a user interacts with a chart.</span></span>

<span data-ttu-id="15021-144">浅灰色网格线效果最佳。</span><span class="sxs-lookup"><span data-stu-id="15021-144">Light grays are most effective for gridlines.</span></span> <span data-ttu-id="15021-145">如果你正在使用中性Fluent UI，请参阅中性[颜色调色板](https://developer.microsoft.com/fluentui#/styles/web/colors/neutrals)。</span><span class="sxs-lookup"><span data-stu-id="15021-145">If you're using Fluent UI, see the [Neutral Colors palette](https://developer.microsoft.com/fluentui#/styles/web/colors/neutrals).</span></span>

<span data-ttu-id="15021-146">下图显示了带有网格线的数据可视化。</span><span class="sxs-lookup"><span data-stu-id="15021-146">The following image shows a data visualization with gridlines.</span></span>

![带网格线的线型图表的数据可视化。](../images/data-visualization.png)

### <a name="legends"></a><span data-ttu-id="15021-148">图例</span><span class="sxs-lookup"><span data-stu-id="15021-148">Legends</span></span>

<span data-ttu-id="15021-149">如果需要，请添加图例：</span><span class="sxs-lookup"><span data-stu-id="15021-149">Add legends if necessary to:</span></span>

- <span data-ttu-id="15021-150">区分系列</span><span class="sxs-lookup"><span data-stu-id="15021-150">Distinguish between series</span></span>
- <span data-ttu-id="15021-151">存在缩放或值的更改</span><span class="sxs-lookup"><span data-stu-id="15021-151">Present scale or value changes</span></span>

<span data-ttu-id="15021-p112">请确保图例增强数据墨迹，但不会影响其显示效果。放置图例：</span><span class="sxs-lookup"><span data-stu-id="15021-p112">Make sure that your legends enhance the data ink and do not compete with it. Place legends:</span></span>


- <span data-ttu-id="15021-154">如果图表上方的所有图例项大小合适，则默认情况下会在绘图区上方左对齐。</span><span class="sxs-lookup"><span data-stu-id="15021-154">Flush left above the plot area by default, if all legend items fit above the chart.</span></span>
- <span data-ttu-id="15021-155">在绘图区的右上角，如果图表上方的所有图例项大小均不合适，请在必要时确保其可滚动。</span><span class="sxs-lookup"><span data-stu-id="15021-155">On the upper right side of the plot area, if all legend items do not fit above the chart, and make it scrollable, if necessary.</span></span>

<span data-ttu-id="15021-p113">为了优化可读性和可访问性，将图例标记映射到相关图表形状。例如，将圆形图例标记用于散点图和气泡图图例。将线段图例标记用于折线图。</span><span class="sxs-lookup"><span data-stu-id="15021-p113">To optimize for readability and accessibility, map legend markers to the relevant chart shape. For example, use circle legend markers for scatter plot and bubble chart legends. Use line segment legend markers for line charts.</span></span>

### <a name="data-labels-and-tooltips"></a><span data-ttu-id="15021-159">数据标签和工具提示</span><span class="sxs-lookup"><span data-stu-id="15021-159">Data labels and tooltips</span></span>

<span data-ttu-id="15021-p114">确保数据标签和工具提示拥有足够的空白和类型变体。使用算法来最小化封闭和冲突。例如，默认情况下，工具提示可能出现在数据点的右侧，但如果检测到右侧边缘，则会出现在左侧。</span><span class="sxs-lookup"><span data-stu-id="15021-p114">Ensure that data labels and tooltips have adequate white space and type variation. Use algorithms to minimize occlusion and collision. For example, a tooltip might surface to the right of a data point by default, but surface to the left if right edges are detected.</span></span>

## <a name="design-principles"></a><span data-ttu-id="15021-163">设计原则</span><span class="sxs-lookup"><span data-stu-id="15021-163">Design principles</span></span>

<span data-ttu-id="15021-164">Office Design 团队创建了以下设计原则集，我们可在为 Office 产品套件设计新的数据可视化时使用这些原则。</span><span class="sxs-lookup"><span data-stu-id="15021-164">The Office Design team created the following set of design principles, which we use when designing new data visualizations for the Office product suite.</span></span>

### <a name="visual-design-principles"></a><span data-ttu-id="15021-165">视觉对象设计原则</span><span class="sxs-lookup"><span data-stu-id="15021-165">Visual design principles</span></span>

- <span data-ttu-id="15021-p115">可视化效果应忠于数据并增强数据，使其易于理解。突出显示数据，仅在需要提供上下文时添加支持元素。避免不必要的装饰（投影、边框等）、图表垃圾或数据失真。</span><span class="sxs-lookup"><span data-stu-id="15021-p115">Visualizations should honor and enhance the data, making it easy to understand. Highlight the data, adding supporting elements only as needed to provide context. Avoid unnecessary embellishments (drop shadows, outlines, etc), chart junk, or data distortion.</span></span>
- <span data-ttu-id="15021-p116">可视化效果应通过提供丰富的视觉反馈吸引用户进行浏览。使用成熟的交互模式、接口控件，并清除系统反馈。</span><span class="sxs-lookup"><span data-stu-id="15021-p116">Visualizations should encourage exploration by providing rich visual feedback. Use well-established interaction patterns, interface controls, and clear system feedback.</span></span>
- <span data-ttu-id="15021-p117">体现久负盛名的设计原则。使用已制定的版式和可视通信设计原则来增强表单、可读性和含义。</span><span class="sxs-lookup"><span data-stu-id="15021-p117">Embody time-honored design principles. Use established typographic and visual communication design principles to enhance form, readability, and meaning.</span></span>

### <a name="interaction-design-principles"></a><span data-ttu-id="15021-173">交互设计原则</span><span class="sxs-lookup"><span data-stu-id="15021-173">Interaction design principles</span></span>

- <span data-ttu-id="15021-174">设计为允许进行浏览。</span><span class="sxs-lookup"><span data-stu-id="15021-174">Design to allow for exploration.</span></span>
- <span data-ttu-id="15021-175">允许与对象进行直接交互，以展示新见解（例如，通过拖动进行排序）。</span><span class="sxs-lookup"><span data-stu-id="15021-175">Allow for direct interactions with objects that reveal new insights (sorting via drag, for example).</span></span>
- <span data-ttu-id="15021-176">使用简单、直接、熟悉的交互模型。</span><span class="sxs-lookup"><span data-stu-id="15021-176">Use simple, direct, familiar interaction models.</span></span>

<span data-ttu-id="15021-177">有关如何设计用户友好交互式数据可视化的详细信息，请参阅 [UI 原则和陷阱](https://uitraps.com/)。</span><span class="sxs-lookup"><span data-stu-id="15021-177">For more information about how to design user-friendly interactive data visualizations, see [UI Tenets and Traps](https://uitraps.com/).</span></span>

### <a name="motion-design-principles"></a><span data-ttu-id="15021-178">动作设计原则</span><span class="sxs-lookup"><span data-stu-id="15021-178">Motion design principles</span></span>

<span data-ttu-id="15021-p118">动作随刺激而产生。视觉元素应以相同的速率朝相同的方向运动。这适用于：</span><span class="sxs-lookup"><span data-stu-id="15021-p118">Motion follows stimulus. Visual elements should move in the same direction at the same rate. This applies to:</span></span>

- <span data-ttu-id="15021-182">创建图表</span><span class="sxs-lookup"><span data-stu-id="15021-182">Chart creation</span></span>
- <span data-ttu-id="15021-183">从一种图表类型转换到另一种图表类型</span><span class="sxs-lookup"><span data-stu-id="15021-183">Transition from one chart type to another chart type</span></span>
- <span data-ttu-id="15021-184">筛选</span><span class="sxs-lookup"><span data-stu-id="15021-184">Filtering</span></span>
- <span data-ttu-id="15021-185">排序</span><span class="sxs-lookup"><span data-stu-id="15021-185">Sorting</span></span>
- <span data-ttu-id="15021-186">添加或减少数据</span><span class="sxs-lookup"><span data-stu-id="15021-186">Adding or subtracting data</span></span>
- <span data-ttu-id="15021-187">对数据进行刷新或切片</span><span class="sxs-lookup"><span data-stu-id="15021-187">Brushing or slicing data</span></span>
- <span data-ttu-id="15021-188">重设图表大小</span><span class="sxs-lookup"><span data-stu-id="15021-188">Resizing a chart</span></span>

<span data-ttu-id="15021-p119">创建因果关系感知。在暂存动画时：</span><span class="sxs-lookup"><span data-stu-id="15021-p119">Create a perception of causality. When staging animations:</span></span>

- <span data-ttu-id="15021-191">一次暂存一个。</span><span class="sxs-lookup"><span data-stu-id="15021-191">Stage one thing at a time.</span></span>
- <span data-ttu-id="15021-192">在更改数据墨迹前，将更改暂存到轴中。</span><span class="sxs-lookup"><span data-stu-id="15021-192">Stage changes to axes before changes to data ink.</span></span>
- <span data-ttu-id="15021-193">如果对象以相同的速度朝相同的方向移动，那么可以暂存对象并将其制作成动画组。</span><span class="sxs-lookup"><span data-stu-id="15021-193">Stage and animate objects as a group if they are moving at the same speed in the same direction.</span></span>
- <span data-ttu-id="15021-p120">在只有 4-5 个对象的组中暂存数据元素。查看器很难独立跟踪数量超过 4-5 个的对象。</span><span class="sxs-lookup"><span data-stu-id="15021-p120">Stage data elements in groups of no more than 4-5 objects. Viewers have difficulty tracking more than 4-5 objects independently.</span></span>

<span data-ttu-id="15021-196">动作赋予涵义。</span><span class="sxs-lookup"><span data-stu-id="15021-196">Motion adds meaning.</span></span>

- <span data-ttu-id="15021-197">动画可帮助用户理解对数据的更改，提供上下文，并作为非语言注释层发挥作用。</span><span class="sxs-lookup"><span data-stu-id="15021-197">Animations increase user comprehension of changes to the data, provide context, and act as a non-verbal annotation layer.</span></span>
- <span data-ttu-id="15021-198">动作应发生在可视化效果具有含义的坐标空间中。</span><span class="sxs-lookup"><span data-stu-id="15021-198">Motion should occur in a meaningful coordinate space of the visualization.</span></span>
- <span data-ttu-id="15021-199">为视觉对象定制动画。</span><span class="sxs-lookup"><span data-stu-id="15021-199">Tailor the animation to the visual.</span></span>
- <span data-ttu-id="15021-200">避免不必要的动画效果。</span><span class="sxs-lookup"><span data-stu-id="15021-200">Avoid gratuitous animations.</span></span>

<span data-ttu-id="15021-201">随数据运动。</span><span class="sxs-lookup"><span data-stu-id="15021-201">Motion follows data.</span></span>

- <span data-ttu-id="15021-p121">保留数据映射。如果某个区域与度量值关联，请使该区域保持在过渡状态。</span><span class="sxs-lookup"><span data-stu-id="15021-p121">Preserve data mappings. If an area is tied to a measure, maintain that area in transition.</span></span>
- <span data-ttu-id="15021-p122">保持统一的动画设计语言。如有可能，请将数据可视化动画映射到现有的 Office 动作设计语言。为类似的图表类型使用相似的动画。</span><span class="sxs-lookup"><span data-stu-id="15021-p122">Maintain a consistent animation design language. Where possible, map data visualization animation to existing Office motion design language. Use similar animations for similar chart types.</span></span>

## <a name="accessibility-in-data-visualizations"></a><span data-ttu-id="15021-207">数据可视化中的辅助功能</span><span class="sxs-lookup"><span data-stu-id="15021-207">Accessibility in data visualizations</span></span>

- <span data-ttu-id="15021-p123">请勿将颜色用作传达信息的唯一方式。色盲者将无法解读结果。在可以传达信息的前提下，除使用颜色外，还使用形状、大小和纹理。</span><span class="sxs-lookup"><span data-stu-id="15021-p123">Do not use color as the only way to communicate information. People who are color blind will not be able to interpret the results. Use shape, size and texture in addition to color when possible to communicate information.</span></span>
- <span data-ttu-id="15021-211">确保所有交互式元素（如按钮或选择列表）均可通过键盘访问。</span><span class="sxs-lookup"><span data-stu-id="15021-211">Make all interactive elements, such as push buttons or pick lists, accessible from a keyboard.</span></span>
- <span data-ttu-id="15021-212">将辅助功能事件发送到屏幕阅读器，以通知焦点更改、工具提示等。</span><span class="sxs-lookup"><span data-stu-id="15021-212">Send accessibility events to screen readers to announce focus changes, tooltips, and so on.</span></span>

## <a name="see-also"></a><span data-ttu-id="15021-213">另请参阅</span><span class="sxs-lookup"><span data-stu-id="15021-213">See also</span></span>

- [<span data-ttu-id="15021-214">构建数据可视化效果的五个最佳库</span><span class="sxs-lookup"><span data-stu-id="15021-214">The Five Best Libraries for Building Data Visualizations</span></span>](https://www.fastcompany.com/3029760/the-five-best-libraries-for-building-data-vizualizations)
- [<span data-ttu-id="15021-215">定量信息的视觉显示</span><span class="sxs-lookup"><span data-stu-id="15021-215">The Visual Display of Quantitative Information</span></span>](https://www.edwardtufte.com/tufte/books_vdqi)
