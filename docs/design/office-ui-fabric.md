---
title: Office 加载项中的 Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 7b1e4a9c377c9a60195a51115d7f275603f1ca5a
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944032"
---
# <a name="office-ui-fabric-in-office-add-ins"></a>Office 加载项中的 Office UI Fabric 

Office UI Fabric 是用于生成 Office 和 Office 365 用户体验的 JavaScript 前端框架。Fabric 提供了以视觉对象为中心的组件，可在 Office 外接程序中进行扩展、返工和使用。由于 Fabric 使用的是 Office 设计语言，因此 Fabric 的用户体验组件看起来像是 Office 的自然扩展。 

若要生成外接程序，我们建议使用 Office UI Fabric 生成用户体验。使用 Office UI Fabric 是可选的。

以下各节介绍如何开始使用 Fabric 以满足要求。 

## <a name="use-fabric-core-icons-fonts-colors"></a>使用 Fabric Core：图标、字体、颜色
Fabric Core 包含设计语言的基本元素，如图标、颜色、类型和网格等。Fabric core 与框架无关。Fabric React 和 Fabric JS 都使用 Fabric Core。

开始使用 Fabric Core：

1. 向页面上的 HTML 添加 CDN 参考。  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
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
- 导航组件 - 如透视和痕迹
- 通知组件 - 例如，消息栏和标注  

并非所有 Fabric 组件都推荐用于外接程序。以下是我们建议在外接程序中使用的 Fabric React UX 组件列表：

- [痕迹](https://developer.microsoft.com/fabric#/components/breadcrumb)
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
|**回应**|[在 Office 外接程序中使用 Office UI Fabric React](using-office-ui-fabric-react.md )|
|**角度**| 请参阅包含 Angular 1.5 指令的社区项目 [ngOfficeUIFabric](http://ngofficeuifabric.com/)，以及[考虑使用 Angular 2 组件包装 Fabric 组件](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)|
