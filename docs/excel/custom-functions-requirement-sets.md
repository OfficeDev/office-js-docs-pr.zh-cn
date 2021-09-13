---
title: 自定义函数要求集
description: 有关 JavaScript API 的自定义函数Excel的详细信息。
ms.date: 09/14/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 0d29d56bb41d44ed8553e97c583e41510e83c132
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149411"
---
# <a name="custom-functions-requirement-sets"></a>自定义函数要求集

[自定义函数](custom-functions-overview.md)使用独立于核心 Excel JavaScript API 的要求集。 下表列出了自定义函数要求集、受支持的 Office 客户端应用程序，以及这些应用程序的版本或版本号。

|  要求集  |  Windows 版 Office<br>（关联至 Microsoft 365 订阅）  |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版 |
|:-----|-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.3 | 16.0.13127.20296 或更高版本 | 不支持 | 16.40.20081000 或更高版本 | 2020 年 7 月 |
| CustomFunctionsRuntime 1.2 | 16.0.12527.20194 或更高版本 | 不支持 | 16.34.20020900 或更高版本 | 2020 年 1 月 |
| CustomFunctionsRuntime 1.1 | 16.0.12527.20092 或更高版本 | 不支持 | 16.34 或更高版本 | 2019 年 5 月 |

> [!NOTE]
> Excel 2019 Office或更早版本不支持自定义 (一次购买) 。

## <a name="customfunctionsruntime-11-12-and-13"></a>CustomFunctionsRuntime 1.1、1.2 和 1.3

CustomFunctionsRuntime 1.1 是 API 的第一个版本。 要求集 1.2 添加 `CustomFunctions.Error` 对象以支持错误处理。 要求集 1.3[](make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions)向 `ErrorCode` [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error)对象添加了 XLL 流式处理支持和新选项。 

## <a name="see-also"></a>另请参阅

- [自定义函数参考文档](/javascript/api/custom-functions-runtime)
- [Excel JavaScript API 要求集](../reference/requirement-sets/excel-api-requirement-sets.md)
