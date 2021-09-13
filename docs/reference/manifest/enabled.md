---
title: 清单文件中 Enabled 元素
description: 了解如何指定外接程序启动时禁用外接程序命令。
ms.date: 01/04/2021
ms.localizationpriority: medium
ms.openlocfilehash: a14385f7114eb3d35845b5d9873bdd718b46c0e9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152360"
---
# <a name="enabled-element"></a>Enabled 元素

指定在加载项[启动](control.md#button-control)[时是否](control.md#menu-dropdown-button-controls)启用"按钮"或"菜单"控件。 **Enabled** 元素是 Control 的子 [元素](control.md)。 如果省略它，则默认值为 `true` 。

此元素仅在 Excel;即，当 `Name` [Host](host.md)元素的 属性为"Workbook"时。

还可以以编程方式启用和禁用父控件。 有关详细信息，请参阅[启用和禁用加载项命令](../../design/disable-add-in-commands.md)。

## <a name="example"></a>示例

```xml
<Enabled>false</Enabled>
```
