---
title: 清单文件中的 GetStarted 元素
description: 提供在 Word、Excel、PowerPoint 和 OneNote 中安装加载项时出现的标注使用的信息。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0ad6196dc45e4ea06c2b43ac5da66a560ab0b899
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771412"
---
# <a name="getstarted-element"></a>GetStarted 元素

提供在 Word、Excel、PowerPoint 和 OneNote 中安装加载项时出现的标注使用的信息。 **GetStarted** 元素是 [DesktopFormFactor](desktopformfactor.md)的子元素。

## <a name="child-elements"></a>子元素

| 元素                       | 必需 | 说明                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [标题](#title)               | 是      | 定义外接程序公开功能的位置。     |
| [说明](#description)   | 是      | 包含 JavaScript 函数的文件的 URL。|
| [LearnMoreUrl](#learnmoreurl) | 是       | 指向详细说明外接程序的页面的 URL。   |

### <a name="title"></a>标题 

必需。 用于标注顶部的标题。 resid 属性引用"资源"部分 **ShortStrings** 元素中的 [](resources.md)有效 ID，并且不能超过 32 个字符。

### <a name="description"></a>说明

必需。 标注的说明/正文内容。 resid 属性引用"资源"部分 **LongStrings** 元素中的 [](resources.md)有效 ID，并且不能超过 32 个字符。

### <a name="learnmoreurl"></a>LearnMoreUrl

必需。 指向用户可以了解你的外接程序详细信息的页面 URL。 resid 属性引用 Resources 节 **的 Urls** 元素 [](resources.md)中的有效 ID，并且不能超过 32 个字符。

> [!NOTE]
> **LearnMoreUrl** 当前无法在 Word、Excel 或 PowerPoint 客户端中呈现。 我们建议为所有客户端添加此 URL，以便 URL 在可用时呈现。 

## <a name="see-also"></a>另请参阅

下面的代码示例使用 **GetStarted** 元素：

* [用于控制表和图表格式化的 Excel Web 外接程序](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [Word 外接程序 JavaScript SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
