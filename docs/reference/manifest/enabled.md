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
# <a name="enabled-element"></a><span data-ttu-id="f7c3b-103">Enabled 元素</span><span class="sxs-lookup"><span data-stu-id="f7c3b-103">Enabled element</span></span>

<span data-ttu-id="f7c3b-104">指定[在加载项启动](control.md#button-control)[时是否](control.md#menu-dropdown-button-controls)启用"按钮"或"菜单"控件。</span><span class="sxs-lookup"><span data-stu-id="f7c3b-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="f7c3b-105">Enabled 元素是 Control 的子[元素](control.md)。</span><span class="sxs-lookup"><span data-stu-id="f7c3b-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="f7c3b-106">如果省略它，则默认值为 `true` 。</span><span class="sxs-lookup"><span data-stu-id="f7c3b-106">If it is omitted, the default is `true`.</span></span>

<span data-ttu-id="f7c3b-107">此元素仅在 Excel 中有效;即，当 `Name` [Host](host.md) 元素的属性为"Workbook"时。</span><span class="sxs-lookup"><span data-stu-id="f7c3b-107">This element is only valid in Excel; that is, when the `Name` attribute of the [Host](host.md) element is "Workbook".</span></span>

<span data-ttu-id="f7c3b-108">也可以以编程方式启用和禁用父控件。</span><span class="sxs-lookup"><span data-stu-id="f7c3b-108">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="f7c3b-109">有关详细信息，请参阅[启用和禁用加载项命令](../../design/disable-add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="f7c3b-109">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

## <a name="example"></a><span data-ttu-id="f7c3b-110">示例</span><span class="sxs-lookup"><span data-stu-id="f7c3b-110">Example</span></span>

```xml
<Enabled>false</Enabled>
```
