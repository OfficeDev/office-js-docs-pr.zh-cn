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
# <a name="enabled-element"></a><span data-ttu-id="977b3-103">Enabled 元素</span><span class="sxs-lookup"><span data-stu-id="977b3-103">Enabled element</span></span>

<span data-ttu-id="977b3-104">指定在启动外接端时是否启用[按钮](control.md#button-control)或[菜单](control.md#menu-dropdown-button-controls)控件。</span><span class="sxs-lookup"><span data-stu-id="977b3-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="977b3-105">**Enabled**元素是[Control](control.md)的子元素。</span><span class="sxs-lookup"><span data-stu-id="977b3-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="977b3-106">如果省略，则默认为`true`。</span><span class="sxs-lookup"><span data-stu-id="977b3-106">If it is omitted, the default is `true`.</span></span> 

<span data-ttu-id="977b3-107">此外，还可以以编程方式启用和禁用父控件。</span><span class="sxs-lookup"><span data-stu-id="977b3-107">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="977b3-108">有关详细信息，请参阅[Enable And Disable 外接程序命令](/office/dev/add-ins/design/disable-add-in-commands)。</span><span class="sxs-lookup"><span data-stu-id="977b3-108">For more information, see [Enable and Disable Add-in Commands](/office/dev/add-ins/design/disable-add-in-commands).</span></span>

## <a name="example"></a><span data-ttu-id="977b3-109">示例</span><span class="sxs-lookup"><span data-stu-id="977b3-109">Example</span></span>

```xml
<Enabled>false</Enabled>
```

