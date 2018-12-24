---
title: 适用于 Office 的 JavaScript API
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: d1f57ec9e4420a17ef0997d8d293c484887d5d79
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432772"
---
# <a name="javascript-api-for-office"></a>适用于 Office 的 JavaScript API

借助适用于 Office 的 JavaScript API，您可以创建可与 Office 主机应用程序中的对象模型进行交互的 Web 应用程序。 你的应用程序将引用 office.js 库中，该库是一个脚本加载程序。 Office.js 库加载适用于正在运行外接程序的 Office 应用程序的对象模型。 你可以使用以下 JavaScript 对象模型：

- **公用 API** - 与 **Office 2013** 一起引入的 API。 这为**所有 Office 主机应用程序**加载 API，并将外接程序应用程序与 Office 客户端应用程序相连接。 对象模型包含特定于 Office 客户端的 API 以及适用于多个 Office 客户端主机应用程序的 API。 所有这些内容位于**共享 API** 下。 

  **Outlook** 还使用通用 API 语法。 代码中的别名 Office 下的全部内容包含可以用于编写与 Office 文档、工作簿、演示文稿、邮件项以及 Office 外接程序中的项目中的内容交互的脚本的对象。如果外接程序面向 Office 2013 及更高版本，则必须使用这些公用 API。 此对象模型使用回叫。

- **特定于主机的 API** - 与 **Office 2016** 一起引入的 API。 此对象模型提供特定于主机的强类型对象，这些对象对应于使用 Office 客户端时所看到的熟悉对象，并表示 Office JavaScript API 的未来。 特定于主机的 API 目前包括 Word JavaScript API 和 Excel JavaScript API。

## <a name="supported-host-applications"></a>支持的主机应用程序

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [共享 API](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint 和 Project](requirement-sets/powerpoint-and-project-note.md) 支持通过 JavaScript API 创建的外接程序。 但是，它们当前没有特定于主机的 API。 你可以通过共享 API 与这些主机交互。

了解有关[支持的主机和其他要求](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins)的详细信息。

## <a name="open-api-specifications"></a>开放 API 规范

在我们设计和开发新的 API 以用于 Office 外接程序时，我们将使它们适用于[开放 API 规范](openspec.md)页的反馈。了解管道中的新增功能，并提供您对我们的设计规范的宝贵意见。

## <a name="see-also"></a>另请参阅

- [Office JavaScript API 参考](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)