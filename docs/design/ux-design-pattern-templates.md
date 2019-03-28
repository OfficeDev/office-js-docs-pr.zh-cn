---
title: 适用于 Office 外接程序的 UX 设计模式
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 40b36fb138169bdf848e5f58569e6fc3dee8c09b
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871540"
---
# <a name="ux-design-patterns-for-office-add-ins"></a>适用于 Office 外接程序的 UX 设计模式

在设计 Office 外接程序的用户体验时，应为 Office 用户提供具有吸引力的体验并通过在默认 Office UI 内无缝接入来扩展整体 Office 体验。  

我们的 UX 模式由组件组成。 组件是帮助客户与软件或服务元素进行交互的控件。 按钮、导航、和菜单是常见组件的示例，通常具有一致的样式和行为。

Office UI Fabric 呈现外观和行为类似于 Office 部件的组件。 利用 Fabric 来轻松与 Office 集成。 如果外接程序有自己预先存在的组件语言，则不需要为支持 Fabric 而放弃它。 与 Office 集成的同时寻找保留该语言的机会。 想办法改变风格元素、消除冲突或采用可避免用户混淆的样式和行为。

提供的模式是基于常见客户方案和用户体验研究的最佳做法解决方案。 它们旨在提供设计和开发外接程序的快速切入点，以及提供在 Microsoft 和品牌元素之间实现平衡的指导。 提供整洁的新式用户体验，并在 Microsoft Fabric 设计语言的设计元素与合作伙伴的独特品牌标识之间保持平衡，可能有助于提高外接程序的用户保留率和采用率。

使用 UX 模式模板来实现以下目的：

* 将解决方案应用于常见的客户方案。
* 应用设计最佳实践。
* 纳入“[Office UI Fabric](https://developer.microsoft.com/fabric#/get-started)”组件和样式。
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
* [Office UI Fabric](https://developer.microsoft.com/fabric)
* [开发 Office 外接程序的最佳做法](/office/dev/add-ins/concepts/add-in-development-best-practices)
* [Fabric React 使用入门](/office/dev/add-ins/design/using-office-ui-fabric-react)
