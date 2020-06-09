---
title: 清单文件中已启用的元素
description: 了解如何指定在启动外接程序时禁用外接程序命令。
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 2849689fec99190c3a9b039c6c04069bc8194ee1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611566"
---
# <a name="enabled-element"></a>Enabled 元素

指定在启动外接端时是否启用[按钮](control.md#button-control)或[菜单](control.md#menu-dropdown-button-controls)控件。 **Enabled**元素是[Control](control.md)的子元素。 如果省略，则默认为 `true` 。

此外，还可以以编程方式启用和禁用父控件。 有关详细信息，请参阅[启用和禁用加载项命令](../../design/disable-add-in-commands.md)。

## <a name="example"></a>示例

```xml
<Enabled>false</Enabled>
```
