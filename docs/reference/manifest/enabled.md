---
title: 清单文件中 Enabled 元素
description: 了解如何指定外接程序命令在加载项启动时处于禁用状态。
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: be18767638af6f2be6352cea46739f6a01b7dd45
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771387"
---
# <a name="enabled-element"></a>Enabled 元素

指定[在加载项启动](control.md#button-control)[时是否](control.md#menu-dropdown-button-controls)启用"按钮"或"菜单"控件。 Enabled 元素是 Control 的子[元素](control.md)。 如果省略它，则默认值为 `true` 。

此元素仅在 Excel 中有效;即，当 `Name` [Host](host.md) 元素的属性为"Workbook"时。

也可以以编程方式启用和禁用父控件。 有关详细信息，请参阅[启用和禁用加载项命令](../../design/disable-add-in-commands.md)。

## <a name="example"></a>示例

```xml
<Enabled>false</Enabled>
```
