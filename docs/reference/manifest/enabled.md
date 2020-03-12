---
title: 清单文件中已启用的元素
description: 了解如何指定在启动外接程序时禁用外接程序命令。
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 4c2c013c8e55966ba2678755536ce04ae3014ed0
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596898"
---
# <a name="enabled-element"></a><span data-ttu-id="42147-103">Enabled 元素</span><span class="sxs-lookup"><span data-stu-id="42147-103">Enabled element</span></span>

<span data-ttu-id="42147-104">指定在启动外接端时是否启用[按钮](control.md#button-control)或[菜单](control.md#menu-dropdown-button-controls)控件。</span><span class="sxs-lookup"><span data-stu-id="42147-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="42147-105">**Enabled**元素是[Control](control.md)的子元素。</span><span class="sxs-lookup"><span data-stu-id="42147-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="42147-106">如果省略，则默认为`true`。</span><span class="sxs-lookup"><span data-stu-id="42147-106">If it is omitted, the default is `true`.</span></span>

<span data-ttu-id="42147-107">此外，还可以以编程方式启用和禁用父控件。</span><span class="sxs-lookup"><span data-stu-id="42147-107">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="42147-108">有关详细信息，请参阅[启用和禁用加载项命令](../../design/disable-add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="42147-108">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

## <a name="example"></a><span data-ttu-id="42147-109">示例</span><span class="sxs-lookup"><span data-stu-id="42147-109">Example</span></span>

```xml
<Enabled>false</Enabled>
```
