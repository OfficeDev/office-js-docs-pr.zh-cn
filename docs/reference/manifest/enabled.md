---
title: 清单文件中 Enabled 元素
description: 了解如何指定外接程序启动时禁用外接程序命令。
ms.date: 11/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 4c0107daaf73aee6ba116553a8d01250e9c7d981
ms.sourcegitcommit: 997a20f9fb011b96a50ceb04a4b9943d92d6ecf4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/19/2021
ms.locfileid: "61081433"
---
# <a name="enabled-element"></a>Enabled 元素

指定在外接程序[启动](control.md#button-control)[时是否](control.md#menu-dropdown-button-controls)启用"按钮"或"菜单"控件。 **Enabled** 元素是 Control 的子 [元素](control.md)。 如果省略它，则默认值为 `true` 。

**外接程序类型：** 任务窗格

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [RibbonApi 1.0](../requirement-sets/ribbon-api-requirement-sets.md)

此元素仅在 Excel中有效;即，当 Host 元素的 属性为 `Name` "Workbook"时。 [](host.md)

还可以以编程方式启用和禁用父控件。 有关详细信息，请参阅[启用和禁用加载项命令](../../design/disable-add-in-commands.md)。

## <a name="example"></a>示例

```xml
<Enabled>false</Enabled>
```
