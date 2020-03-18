---
title: Office 加载项中的 Office UI Fabric 
description: 概述如何在 Office 外接程序中使用 Office UI Fabric 组件。
ms.date: 12/04/2017
localization_priority: Normal
ms.openlocfilehash: 3e65e123d6195fc435b12c477985a10a3a2b0399
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718704"
---
# <a name="office-ui-fabric-in-office-add-ins"></a>Office 加载项中的 Office UI Fabric 

Office UI Fabric 是用于生成 Office 和 Office 365 用户体验的 JavaScript 前端框架。Fabric 提供了以视觉对象为中心的组件，可在 Office 外接程序中进行扩展、返工和使用。由于 Fabric 使用的是 Office 设计语言，因此 Fabric 的用户体验组件看起来像是 Office 的自然扩展。 

若要生成外接程序，我们建议使用 Office UI Fabric 生成用户体验。使用 Office UI Fabric 是可选的。

以下各节介绍如何开始使用 Fabric 以满足要求。 

## <a name="use-fabric-core-icons-fonts-colors"></a>使用 Fabric Core：图标、字体、颜色
Fabric Core 包含设计语言的基本元素，如图标、颜色、类型和网格等。 Fabric Core 与框架无关。 Fabric Core 供 Fabric React 使用并且包含其中。

开始使用 Fabric Core：

1. 向页面上的 HTML 添加 CDN 参考。  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```   
    
2. 使用 Fabric 图标和字体。 

    若要使用 Fabric 图标，在页面上包括“i”元素，然后引用适当的类。可以通过更改字号来控制图标的大小。例如，下面的代码展示了如何制作使用 themePrimary (#0078d7) 颜色的超大表图标。 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    若要查找 Office UI Fabric 中可用的更多图标，请在“[图标](https://developer.microsoft.com/fabric#/styles/icons)”页上使用搜索功能。找到要在外接程序中使用的图标后，请务必在图标名称前加上前缀 `ms-Icon--`。 

    若要了解 Office UI Fabric 中可用的字号和颜色，请参阅[版式](https://developer.microsoft.com/fabric#/styles/typography)和[颜色](https://developer.microsoft.com/fabric#/styles/colors)。
 
## <a name="use-fabric-components"></a>使用 Fabric 组件 
Fabric 提供了多种可用于生成外界程序的 UX 组件，包括以下类型的组件：

- 输入组件 - 如按钮、复选框和切换
- 导航组件 - 如透视、痕迹
- 通知组件 - 如消息栏和标注  

不建议在外接程序中使用所有 Fabric 组件。建议在外接程序中使用下列 Fabric React UX 组件：

- [痕迹导航栏](https://developer.microsoft.com/fabric#/components/breadcrumb)
- [按钮](https://developer.microsoft.com/fabric#/components/button)
- [复选框](https://developer.microsoft.com/fabric#/components/checkbox)
- [ChoiceGroup](https://developer.microsoft.com/fabric#/components/choicegroup)
- [下拉列表](https://developer.microsoft.com/fabric#/components/dropdown)
- [标签](https://developer.microsoft.com/fabric#/components/label)
- [列表](https://developer.microsoft.com/fabric#/components/list)
- [透视](https://developer.microsoft.com/fabric#/components/pivot)
- [TextField](https://developer.microsoft.com/fabric#/components/textfield)
- [切换](https://developer.microsoft.com/fabric#/components/toggle)

你可以使用不同的 JavaScript 框架（如 Angular 或 React）来生成外接程序。若要开始将 Fabric 组件与框架一起使用，请参阅以下资源。

|**框架**|**示例**|
|:------------|:----------|
|**React**|[在 Office 外接程序中使用 Office UI Fabric React](using-office-ui-fabric-react.md )|
|**Angular**| 请参阅包含 Angular 1.5 指令的社区项目 [ngOfficeUIFabric](http://ngofficeuifabric.com/)，以及[考虑使用 Angular 2 组件包装 Fabric 组件](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)|
