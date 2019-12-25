---
title: 开发 Office 加载项
description: Office 加载项开发简介。
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: 731226883e2bdea4b68d0720042010a0f0117098
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851686"
---
# <a name="develop-office-add-ins"></a>开发 Office 加载项

> [!TIP]
> 阅读本文之前，请查看[构建 Office 加载项](../overview/office-add-ins-fundamentals.md)。

所有 Office 加载项均基于 Office 加载项平台构建。 它们共享一个可实现某些功能的公共框架。 无论构建任何加载项，你都需要了解主机和平台可用性、Office JavaScript API 编程模式、如何在清单文件中指定加载项的设置和功能等重要概念。 本文档的“**核心概念**” > “**开发**”部分在此介绍了这类核心开发概念。 在浏览与所构建的加载项（例如 [Excel](../excel/index.md)）相对应的主机特定文档之前，请先查看此处的信息。

> [!NOTE]
> 本文档的“**核心概念**” > “**开发**” > “**操作方法**”部分包含侧重于具体开发概念或任务的文章。 例如，你将在此处找到诸如[使用 Visual Studio Code 开发加载项](develop-add-ins-vscode.md)、[随文档自动打开任务窗格](automatically-open-a-task-pane-with-a-document.md)、[创建加载项命令](create-addin-commands.md)以及[打开对话框](dialog-api-in-office-add-ins.md)等任务的信息。

## <a name="next-steps"></a>后续步骤

在熟悉此处介绍的核心概念之后，请浏览与所构建的加载项（例如 [Excel](../excel/index.md)）相对应的主机特定文档。 文档中每个主机特定的部分都包含关于为特定 Office 主机构建加载项的具体信息。

## <a name="see-also"></a>另请参阅

- [Office 加载项平台概述](../overview/office-add-ins.md)
- [构建 Office 加载项](../overview/office-add-ins-fundamentals.md)
- [Office 加载项的核心概念](../overview/core-concepts-office-add-ins.md)
- [设计 Office 加载项](../design/add-in-design.md)
- [测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)
- [发布 Office 加载项](../publish/publish.md)