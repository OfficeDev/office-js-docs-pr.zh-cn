---
title: 适用于 Office 外接程序的 UX 设计模式
description: 大致了解适用于外接程序的 UI 设计Office包括导航、身份验证、首次运行和品牌打造的模式。
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3515a5bf915b711f79aa328ba2cc50a3b03670a4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149178"
---
# <a name="ux-design-patterns-for-office-add-ins"></a>适用于 Office 外接程序的 UX 设计模式

在设计 Office 外接程序的用户体验时，应为 Office 用户提供具有吸引力的体验并通过在默认 Office UI 内无缝接入来扩展整体 Office 体验。  

我们的 UX 模式由组件组成。 组件是帮助客户与软件或服务元素进行交互的控件。 按钮、导航、和菜单是常见组件的示例，通常具有一致的样式和行为。

[Fluent UI React组件](using-office-ui-fabric-react.md)的外观和行为与 Office 的一部分类似，Office UI Fabric [JS 的中性框架组件一样](fabric-core.md)。 利用任一组组件与Office。 或者，如果您的外接程序具有自己的预先不存在的组件语言，则无需放弃它。 与 Office 集成的同时寻找保留该语言的机会。 想办法改变风格元素、消除冲突或采用可避免用户混淆的样式和行为。

提供的模式是基于常见客户方案和用户体验研究的最佳做法解决方案。 它们旨在提供设计和开发外接程序的快速入口点，以及指导在 Microsoft 品牌元素和您自己的元素之间实现平衡。 提供简洁的新式用户体验，在 Microsoft Fluent UI 设计语言和合作伙伴的独特品牌标识之间实现设计元素的平衡，这可帮助提高外接程序的用户保留率和采用率。

使用 UX 模式模板来实现以下目的：

* 将解决方案应用于常见的客户方案。
* 应用设计最佳实践。
* 合并[Fluent UI](https://developer.microsoft.com/fluentui#/get-started)组件和样式。
* 构建以可视方式与默认 Office UI 集成的外接程序。
* 形成 UX 概念并将其可视化。

## <a name="getting-started"></a>入门

该模式按照外接程序中的常见按键操作或体验来进行组织。 主要的组包括：

* [初次运行体验 (FRE)](../design/first-run-experience-patterns.md)
* [身份验证](../design/authentication-patterns.md)
* [导航](../design/navigation-patterns.md)
* [品牌设计](../design/branding-patterns.md)

浏览每个分组，了解如何使用最佳做法来设计外接程序。

> [!NOTE]
> 本文档中显示的所有示例屏幕均按 **1366x768** 的分辨率进行设计和显示。

## <a name="see-also"></a>另请参阅

* [设计工具包](design-toolkits.md)
* [Fluent UI](https://developer.microsoft.com/fluentui#)
* [Office 加载项开发最佳做法](../concepts/add-in-development-best-practices.md)
* [Office 加载项中的 Fluent UI React](using-office-ui-fabric-react.md)
