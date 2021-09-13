---
title: Office 加载项的数据可视化样式指南
description: 获取有关如何在加载项中可视化数据Office一些好的做法。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: cc3c743e3a793c4d4fdc2639313eb40a01923ada
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149371"
---
# <a name="data-visualization-style-guidelines-for-office-add-ins"></a>Office 加载项的数据可视化样式指南

良好的数据可视化效果可帮助用户找到数据见解。他们可以使用这些见解来讲述具有说服力的故事。本文提供了准则，以帮助你在适用于 Excel 和其他 Office 应用的外接程序中设计有效的数据可视化。

我们建议你使用 Fluent [UI](../design/add-in-design.md)为数据可视化创建部件版式。 FluentUI 包括样式和组件，这些样式和组件与Office外观无缝集成。

## <a name="data-visualization-elements"></a>数据可视化元素

数据可视化共享常规框架和常见的视觉和交互式元素，包括标题、标签和数据绘图，如下图所示。

![带标题、坐标轴、图例和绘图区标签的线图。](../images/excel-charts-visualization.png)

### <a name="chart-titles"></a>图表标题

遵循图表标题的以下指南。

- 使图表标题便于阅读。设定其位置以创建相对于其余图表的清晰视觉对象层次结构。
- 一般情况下，使用句子大写（大写第一个字词）。若要创建对比度或强化层次结构，可以全部使用大写，但应谨慎使用全部大写。
- 合并[Fluent UI 类型](https://developer.microsoft.com/fluentui#/styles/web/typography)渐变，使图表与使用 Segoe 的 Office UI 保持一致。 你还可以使用不同的字样来区分图表内容和 UI。
- 使用带有大型计数器的 sans-serif 字样。

### <a name="axis-labels"></a>轴标签

请确保轴标签颜色足够深，以便可以清楚地阅读，并且具有足够的文本和背景色对比度。请确保颜色不要过深，避免比数据墨迹更加突出。

浅灰色轴标签效果最佳。 如果你使用的是中性Fluent UI，请参阅中性[颜色调色板](https://developer.microsoft.com/fluentui#/styles/web/colors/neutrals)。

### <a name="data-ink"></a>数据墨迹

表示图表中的实际数据的像素被称为数据墨迹。这应该是可视化的中心焦点。避免使用投影、过粗边框或不必要的使数据失真或影响数据显示效果的设计元素。仅当数据值与颜色值关联时使用渐变。避免使用三维图表，除非可测量的目标值绑定到第三维度。

### <a name="color"></a>颜色

选择遵循操作系统或应用程序主题的颜色，而不是硬编码的颜色。同时，确保所应用的颜色不会使数据失真。数据可视化中的颜色滥用可能会导致数据失真和信息读取不正确。

有关在数据可视化中使用颜色的最佳做法，请参阅以下内容：

- [为什么彩虹色不是数据可视化的最佳选择](https://www.poynter.org/2013/why-rainbow-colors-arent-always-the-best-options-for-data-visualizations/224413/)
- [Color Brewer 2.0：制图的颜色建议](https://colorbrewer2.org/)
- [我想要的色调](https://tools.medialab.sciences-po.fr/iwanthue/)

### <a name="gridlines"></a>网格线

要准确读取图表，通常网格线是必不可少的，但应显示为辅助可视元素，用于增强数据墨迹效果，但不会影响数据显示。确保静态网格线较细且颜色较淡，除非专门将其设计用于高对比度的情况。还可以使用交互作用创建在用户与图表交互时上下文中显示的动态、实时网格线。

浅灰色网格线效果最佳。 如果你使用的是中性Fluent UI，请参阅中性[颜色调色板](https://developer.microsoft.com/fluentui#/styles/web/colors/neutrals)。

下图显示了带有网格线的数据可视化。

![带网格线的线型图表的数据可视化。](../images/data-visualization.png)

### <a name="legends"></a>图例

如果需要，请添加图例：

- 区分系列
- 存在缩放或值的更改

请确保图例增强数据墨迹，但不会影响其显示效果。放置图例：

- 如果图表上方的所有图例项大小合适，则默认情况下会在绘图区上方左对齐。
- 在绘图区的右上角，如果图表上方的所有图例项大小均不合适，请在必要时确保其可滚动。

为了优化可读性和可访问性，将图例标记映射到相关图表形状。例如，将圆形图例标记用于散点图和气泡图图例。将线段图例标记用于折线图。

### <a name="data-labels-and-tooltips"></a>数据标签和工具提示

确保数据标签和工具提示拥有足够的空白和类型变体。使用算法来最小化封闭和冲突。例如，默认情况下，工具提示可能出现在数据点的右侧，但如果检测到右侧边缘，则会出现在左侧。

## <a name="design-principles"></a>设计原则

Office Design 团队创建了以下设计原则集，我们可在为 Office 产品套件设计新的数据可视化时使用这些原则。

### <a name="visual-design-principles"></a>视觉对象设计原则

- 可视化效果应忠于数据并增强数据，使其易于理解。突出显示数据，仅在需要提供上下文时添加支持元素。避免不必要的装饰（投影、边框等）、图表垃圾或数据失真。
- 可视化效果应通过提供丰富的视觉反馈吸引用户进行浏览。使用成熟的交互模式、接口控件，并清除系统反馈。
- 体现久负盛名的设计原则。使用已制定的版式和可视通信设计原则来增强表单、可读性和含义。

### <a name="interaction-design-principles"></a>交互设计原则

- 设计为允许进行浏览。
- 允许与对象进行直接交互，以展示新见解（例如，通过拖动进行排序）。
- 使用简单、直接、熟悉的交互模型。

有关如何设计用户友好交互式数据可视化的详细信息，请参阅 [UI 原则和陷阱](https://uitraps.com/)。

### <a name="motion-design-principles"></a>动作设计原则

动作随刺激而产生。视觉元素应以相同的速率朝相同的方向运动。这适用于：

- 创建图表
- 从一种图表类型转换到另一种图表类型
- 筛选
- 排序
- 添加或减少数据
- 对数据进行刷新或切片
- 重设图表大小

创建因果关系感知。在暂存动画时：

- 一次暂存一个。
- 在更改数据墨迹前，将更改暂存到轴中。
- 如果对象以相同的速度朝相同的方向移动，那么可以暂存对象并将其制作成动画组。
- 在只有 4-5 个对象的组中暂存数据元素。查看器很难独立跟踪数量超过 4-5 个的对象。

动作赋予涵义。

- 动画可帮助用户理解对数据的更改，提供上下文，并作为非语言注释层发挥作用。
- 动作应发生在可视化效果具有含义的坐标空间中。
- 为视觉对象定制动画。
- 避免不必要的动画效果。

随数据运动。

- 保留数据映射。如果某个区域与度量值关联，请使该区域保持在过渡状态。
- 保持统一的动画设计语言。如有可能，请将数据可视化动画映射到现有的 Office 动作设计语言。为类似的图表类型使用相似的动画。

## <a name="accessibility-in-data-visualizations"></a>数据可视化中的辅助功能

- 请勿将颜色用作传达信息的唯一方式。色盲者将无法解读结果。在可以传达信息的前提下，除使用颜色外，还使用形状、大小和纹理。
- 确保所有交互式元素（如按钮或选择列表）均可通过键盘访问。
- 将辅助功能事件发送到屏幕阅读器，以通知焦点更改、工具提示等。

## <a name="see-also"></a>另请参阅

- [构建数据可视化效果的五个最佳库](https://www.fastcompany.com/3029760/the-five-best-libraries-for-building-data-vizualizations)
- [定量信息的视觉显示](https://www.edwardtufte.com/tufte/books_vdqi)
