---
title: 自定义函数和数据类型概述
description: 将 Excel 数据类型与自定义函数和 Office 加载项配合使用。
ms.date: 11/03/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 91d2fb21aae57ed7a5777136f3c4540925f339c8
ms.sourcegitcommit: 210251da940964b9eb28f1071977ea1fe80271b4
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/05/2021
ms.locfileid: "60793573"
---
# <a name="use-data-types-with-custom-functions-in-excel-preview"></a>在 Excel 中将数据类型与自定义函数配合使用（预览版）

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

数据类型扩展了 Excel JavaScript API，以支持四个原始数据类型（字符串、数字、布尔值和错误）以外的数据类型。 数据类型包括支持 Web 图像、带格式数字值、实体值和实体值中的数组。

这些数据类型放大了自定义函数的功能，因为自定义函数接受数据类型作为输入值和输出值。 可以通过自定义函数生成数据类型，或将现有数据类型作为函数参数引入计算。 设置数据类型的 JSON 架构后，将在整个自定义函数计算中维护此架构。

如果要详细了解如何将数据类型与 Excel 加载项配合使用，请参阅 [Excel 加载项中的数据类型概述](excel-data-types-overview.md)。如果要详细了解如何将自定义数据类型与自定义函数集成，请参阅 [自定义函数和数据类型核心概念](custom-functions-data-types-concepts.md)。

## <a name="see-also"></a>另请参阅

* [ Excel 加载项中的数据类型的概述](excel-data-types-overview.md)
* [Excel 数据类型核心概念](excel-data-types-concepts.md)
* [自定义函数和数据类型核心概念](custom-functions-data-types-concepts.md)
* [将 Office 加载项配置为使用共享 JavaScript 运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
