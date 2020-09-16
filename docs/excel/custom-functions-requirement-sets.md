---
title: 自定义函数要求集
description: 有关 Excel JavaScript API 的自定义函数要求集的详细信息。
ms.date: 09/14/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0860dd2d1b55376a85eadf04898d288d83b0205d
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819523"
---
# <a name="custom-functions-requirement-sets"></a>自定义函数要求集

[自定义函数](custom-functions-overview.md)使用独立于核心 Excel JavaScript API 的要求集。 下表列出了自定义函数要求集、受支持的 Office 客户端应用程序，以及这些应用程序的内部版本或编号。

|  要求集  |  Windows 版 Office<br>（关联至 Microsoft 365 订阅）  |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版 |
|:-----|-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1。3 | 16.0.13127.20296 或更高版本 | 不支持 | 16.40.20081000 或更高版本 | 2020 年 7 月 |
| CustomFunctionsRuntime 1。2 | 16.0.12527.20194 或更高版本 | 不支持 | 16.34.20020900 或更高版本 | 2020 年 1 月 |
| CustomFunctionsRuntime 1.1 | 16.0.12527.20092 或更高版本 | 不支持 | 16.34 或更高版本 | 2019 年 5 月 |

> [!NOTE]
> Office 2019 或更早版本 (一次性购买) 不支持 Excel 自定义函数。

## <a name="customfunctionsruntime-11-12-and-13"></a>CustomFunctionsRuntime 1.1、1.2 和1。3

CustomFunctionsRuntime 1.1 是 API 的第一个版本。 要求集1.2 添加了 `CustomFunctions.Error` 支持错误处理的对象。 要求集1.3 将 [XLL 流](make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) 支持和新 `ErrorCode` 选项添加到 [customfunctions.js](/javascript/api/custom-functions-runtime/customfunctions.error) 对象。 

## <a name="see-also"></a>另请参阅

- [自定义函数参考文档](/javascript/api/custom-functions-runtime)
- [Excel JavaScript API 要求集](../reference/requirement-sets/excel-api-requirement-sets.md)
