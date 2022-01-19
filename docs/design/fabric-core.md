---
title: Office 加载项中的 Fabric Core
description: 大致了解如何在外接程序中Office Fabric Core 和 Fabric UI 组件。
ms.date: 01/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 3d10cc5d8f33c8dd66f4f988fdd5a082580b1aca
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/19/2022
ms.locfileid: "62074201"
---
# <a name="fabric-core-in-office-add-ins"></a>Office 加载项中的 Fabric Core

Fabric Core 是 CSS 类和 SASS mix 的开源集合，旨在用于非 React *Office* 外接程序。Fabric Core 包含 UI 设计Fluent的基本元素，如图标、颜色、字样和网格。 Fabric Core 与框架无关，因此可用于任何单页应用程序或任何服务器端 Web UI 框架。  (历史原因，它被称为"Fabric Core"，而不是"Fluent Core"。) 

如果外接程序的 UI 不是基于React的，则您还可以使用一组非React组件。 请参阅[使用 Office UI Fabric JS 组件](#use-office-ui-fabric-js-components)。

> [!NOTE]
> 本文介绍了 Fabric Core 在加载项Office的使用。但它还用于各种应用Microsoft 365扩展。 有关详细信息，请参阅[Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core)和开源存储库Office UI Fabric [Core。](https://github.com/OfficeDev/office-ui-fabric-core)

## <a name="use-fabric-core-icons-fonts-colors"></a>使用 Fabric Core：图标、字体、颜色

1. 添加内容传送网络 (CDN) 对页面上的 HTML 的引用。

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. 使用 Fabric Core 图标和字体。

    若要使用 Fabric Core 图标，请在你的页面上包括"i"元素，然后引用相应的类。 可以通过更改字号来控制图标的大小。 例如，下面的代码展示了如何制作使用 themePrimary (#0078d7) 颜色的超大表图标。

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    有关更详细的说明，请参阅Fluent [UI 图标](https://developer.microsoft.com/fluentui#/styles/web/icons)。 若要查找 Fabric Core 中可用的更多图标，请使用该页面上的搜索功能。 找到要在外接程序中使用的图标后，请务必在图标名称前加上前缀 `ms-Icon--`。

    有关 Fabric Core 中可用的字体大小和颜色的信息，请参阅 Colors[](https://developer.microsoft.com/fluentui#/styles/web/typography)中的版式和颜色[目录](https://developer.microsoft.com/fluentui#/styles/web/colors)。

示例包含在本文稍后 [的示例中](#samples) 。

## <a name="use-office-ui-fabric-js-components"></a>使用 Office UI Fabric JS 组件

具有非 React API 的外接程序还可使用[Office UI Fabric JS](https://github.com/OfficeDev/office-ui-fabric-js)中的任意组件，包括按钮、对话框、选取器等。 有关说明，请参阅存储库自述。

示例包含在本文稍后 [的示例中](#samples) 。

## <a name="samples"></a>示例

以下示例外接程序使用 Fabric Core 和/或 Office UI Fabric JS 组件。 其中一些资源已存档，这意味着不再使用 Bug 或安全修补程序更新它们，但你仍可以使用它们了解如何使用 Fabric Core 和 Fabric UI 组件。

- [Excel加载项 JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
- [Excel加载项 SalesLeads](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)
- [Excel外接程序 WoodGrove 支出趋势](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [Excel内容外接程序 Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)
- [Office外接程序 Fabric UI 示例](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Outlook外接程序 GifMe](https://github.com/OfficeDev/Outlook-Add-in-GifMe)
- [PowerPoint Microsoft 外接程序 Graph ASPNET 插入图](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Word 外接程序 Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)
- [Word 外接程序 JS 修订](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word 加载项 MarkdownConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion)
