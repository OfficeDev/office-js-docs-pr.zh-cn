---
title: 清单文件中已启用的元素
description: 了解如何指定在启动外接程序时禁用外接程序命令。
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: a47ab97ff5a159c73bea52f130ce0c16efe2b6b6
ms.sourcegitcommit: 0e7ed44019d6564c79113639af831ea512fa0a13
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/09/2020
ms.locfileid: "42566189"
---
# <a name="enabled-element"></a>Enabled 元素

指定在启动外接端时是否启用[按钮](control.md#button-control)或[菜单](control.md#menu-dropdown-button-controls)控件。 **Enabled**元素是[Control](control.md)的子元素。 如果省略，则默认为`true`。 

此外，还可以以编程方式启用和禁用父控件。 有关详细信息，请参阅[Enable And Disable 外接程序命令](/office/dev/add-ins/design/disable-add-in-commands)。

## <a name="example"></a>示例

```xml
<Enabled>false</Enabled>
```

