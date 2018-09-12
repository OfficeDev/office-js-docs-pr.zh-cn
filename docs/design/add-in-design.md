---
title: 设计 Office 加载项
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: a864dc0fc6d30ecfcd7a71d1b2677d408519af0e
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944196"
---
# <a name="design-your-office-add-ins"></a>设计 Office 加载项

Office 外接程序可通过提供用户可在 Office 客户端内访问的上下文功能来扩展 Office 体验。通过外接程序，用户可以访问 Office 内的第三方功能以完成更多操作，而无需进行成本高昂的上下文切换。 

你的外接程序 UX 设计必须与 Office 无缝集成，为用户提供高效、自然的交互。利用[外接程序命令](add-in-commands.md)提供对外接程序的访问权限，并应用创建基于 HTML 的自定义 UI 时建议的最佳实践。

## <a name="office-design-principles"></a>Office 设计原则

Office 应用程序遵循一套常规交互原则。应用共享内容并具有外观和行为相似的元素。此通用性基于一套设计原则。这些原则帮助 Office 团队创建支持客户任务的界面。了解并遵循这些原则将有助于支持 Office 内部的客户目标。

若要打造积极的加载项体验，请遵循 Office 设计原则：

- **对 Office 进行明确设计。** 外接程序的功能、外观和感受必须和谐地完善 Office 体验。外接程序应该让人感觉就像安装在本机一样。它们应无缝融入 iPad 或 PowerPoint Online 上的 Word。设计良好的外接程序将恰当地融合体验、平台和 Office 应用程序。请考虑使用 Office UI Fabric 作为设计语言。在适当的位置应用文档和 UI 主题。

- **重点关注几个关键任务；好好完成。** 帮助客户在不影响其他工作的情况下完成一项工作。为客户提供真正的价值。与 Office 文档交互时，关注常见用例并认真挑选出用户最受益的。

- **使内容优先于 Chrome。** 使客户的页面、幻灯片或电子表格始终关注体验。外接程序是辅助界面。没有任何辅助 Chrome 应当与外接程序的内容和功能交互。请明智地品牌化你的体验。我们知道这对于向用户提供独特且可识别的功能但避免干扰十分重要。努力将重点集中于内容和任务完成，而非品牌关注。

- **为用户带来愉快体验并让用户得到控制。** 人们喜欢使用兼具功能和视觉吸引力的产品。 仔细打造你的体验。 通过考虑每个交互和视觉细节来打磨细节。 允许用户控制他们的体验。 完成任务的必要步骤必须清晰明确。 重要的决定应该易于理解。 操作应该能方便地撤消。 外接程序不是最终目标 - 它是对 Office 功能的增强。

- **针对所有平台和输入方法进行设计**。外接程序设计用于 Office 支持的所有平台，您的外接程序 UI 应该进行优化，以便跨平台和外形规格运行。支持鼠标/键盘和触摸输入设备，确保您的自定义 HTML UI 响应迅速，可适应不同的外形规格。有关详细信息，请参阅[触摸](../concepts/add-in-development-best-practices.md#optimize-for-touch)。 

## <a name="see-also"></a>另请参阅
- [Office UI Fabric](https://developer.microsoft.com/en-us/fabric) 
- [加载项开发最佳做法](../concepts/add-in-development-best-practices.md)

