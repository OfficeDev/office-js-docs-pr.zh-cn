---
title: 适用于 Office 的 JavaScript API
description: ''
ms.date: 05/13/2019
localization_priority: Priority
ms.openlocfilehash: 8d834aee4c21448210d9619fedd42d5ebb79e09d
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575322"
---
# <a name="javascript-api-for-office"></a>适用于 Office 的 JavaScript API

借助适用于 Office 的 JavaScript API，您可以创建可与 Office 主机应用程序中的对象模型进行交互的 Web 应用程序。 你的应用程序将引用 office.js 库中，该库是一个脚本加载程序。 Office.js 库加载适用于正在运行外接程序的 Office 应用程序的对象模型。 你可以使用以下 JavaScript 对象模型：

- **公用 API** - 与 **Office 2013** 一起引入的 API。 这为**所有 Office 主机应用程序**加载 API，并将外接程序应用程序与 Office 客户端应用程序相连接。 对象模型包含特定于 Office 客户端的 API 以及适用于多个 Office 客户端主机应用程序的 API。 所有这些内容位于**通用 API** 下。 此对象模型使用回调。 

  **Outlook** 还使用通用 API 语法。 代码中的别名 Office 下的全部内容包含可以用于编写与 Office 文档、工作簿、演示文稿、邮件项以及 Office 加载项中的项目中的内容交互的脚本的对象。如果加载项面向 Office 2013 及更高版本，则必须使用这些通用 API。 此对象模型使用回调。

- **特定于主机的 API** - 与 **Office 2016** 一起引入的 API。 此对象模型提供特定于主机的强类型对象，这些对象对应于使用 Office 客户端时所看到的熟悉对象，并表示 Office JavaScript API 的未来。 特定于主机的 JavaScript API 当前可用于 Excel、OneNote、PowerPoint 和 Word。

## <a name="supported-host-applications"></a>支持的主机应用程序

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [PowerPoint](overview/powerpoint-add-ins-reference-overview.md)
- [项目](overview/project-add-ins-reference-overview.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [通用 API](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [Project](overview/project-add-ins-reference-overview.md) 支持使用 JavaScript API 制作的加载项，但目前没有专为与 Project 交互而设计的 JavaScript API。 你可以使用通用 API 来创建 Project 加载项。

了解有关[支持的主机和其他要求](../concepts/requirements-for-running-office-add-ins.md)的详细信息。

## <a name="open-api-specifications"></a>开放 API 规范

在我们设计和开发新的 API 以用于 Office 外接程序时，我们将使它们适用于[开放 API 规范](openspec/openspec.md)页的反馈。了解管道中的新增功能，并提供您对我们的设计规范的宝贵意见。

## <a name="see-also"></a>另请参阅

- [Office JavaScript API 参考](/javascript/api/overview/office)
