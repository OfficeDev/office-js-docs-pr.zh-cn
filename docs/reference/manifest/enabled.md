---
title: 清单文件中 Enabled 元素
description: 了解如何指定外接程序启动时禁用外接程序命令。
ms.date: 03/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: fc635e91b005eb51c70e8517058fc03fa4f26c6c
ms.sourcegitcommit: 856f057a8c9b937bfb37e7d81a6b71dbed4b8ff4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/16/2022
ms.locfileid: "63511261"
---
# <a name="enabled-element"></a>Enabled 元素

指定在外接程序 [启动](control-button.md) 时 [是否](control-menu.md) 启用"按钮"控件或"菜单"控件。 **Enabled** 元素是 Control 的子 [元素](control.md)。 如果省略它，则默认值为 `true`。

**外接程序类型：** 任务窗格

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [RibbonApi 1.0](../requirement-sets/ribbon-api-requirement-sets.md)

此元素仅在 Excel、PowerPoint 和 Word `Name` 中有效;即，[当 Host](host.md) 元素的属性为"Workbook"、"Presentation"或"Document"时。

还可以以编程方式启用和禁用父控件。 有关详细信息，请参阅[启用和禁用加载项命令](../../design/disable-add-in-commands.md)。

## <a name="example"></a>示例

```xml
<Enabled>false</Enabled>
```
