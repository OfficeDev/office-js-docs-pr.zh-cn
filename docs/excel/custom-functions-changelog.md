---
ms.date: 01/08/2019
description: 发现 Excel 自定义函数的最新更新。
title: 自定义函数更改日志（预览）
localization_priority: Normal
ms.openlocfilehash: 03e4dd922ac3895e11a508f97e7ac3fa3e7b1cb0
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449272"
---
# <a name="custom-functions-changelog-preview"></a>自定义函数更改日志（预览）

Excel 自定义函数仍处于预览状态，这意味着将会对该产品进行频繁更改，包括更改和发布新功能。 此更改日志提供了与产品所有更改相关的最新信息。

- **2017 年 11 月 7 日**：发布了*自定义函数（预览）和示例
- **2017 年 11 月 20 日**：修复了使用内部版本 8801 及更高版本的函数的兼容性问题
- **2017 年 11 月 28 日**：发布了*对取消异步函数的支持（需要对流式处理函数进行相应更改）
- **2018 年 5 月 7 日**：发布了*对 Mac、Excel Online 和在进程中运行的异步函数的支持
- **2018 年 9 月 20 日**：发布了对自定义函数 JavaScript 运行时的支持。 有关详细信息，请参阅 [Excel 自定义函数的运行时](custom-functions-runtime.md)。
- **2018 年 10 月 20 日**：随着 [10 月预览体验内部版本](https://support.office.com/en-us/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24)的推出，自定义函数现在需要适用于 Windows Desktop 和 Online 的[自定义函数元数据](custom-functions-json.md)中的“id”参数。 在 Mac 上，应忽略此参数。 自定义函数现也支持可选参数和 `any` 返回类型。
- **2018 年 12 月 12 日**：自定义函数中现在包括用于发现单元格地址的方法。 有关详细信息，请参阅[确定调用自定义函数的单元格](custom-functions-overview.md#determine-which-cell-invoked-your-custom-function)。
- **2019 年 1 月 8 日**：绑定方法 `CustomFunctionMapping()` 已更改为 `CustomFunctions.associate()`。 有关详细信息，请参阅[自定义函数最佳实践（预览）](custom-functions-best-practices.md)。

\* 转到 [Office 预览体验成员](https://products.office.com/office-insider)频道（以前称为“预览体验成员 - 快”）

有关产品的已知问题列表，请参阅[已知问题](custom-functions-overview.md#known-issues)。 

## <a name="see-also"></a>另请参阅

* [自定义函数概述](custom-functions-overview.md)
* [自定义函数元数据](custom-functions-json.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
* [自定义函数调试](custom-functions-debugging.md)