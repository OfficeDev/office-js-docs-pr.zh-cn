---
title: 自定义函数要求集
description: 有关 JavaScript API 的自定义函数Excel集的详细信息。
ms.date: 02/15/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 7558035b6b151977e985ec04ed1fa84c116f0886
ms.sourcegitcommit: 789545a81bd61ec2e7adef2bc24c06b5be113b00
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/18/2022
ms.locfileid: "62892515"
---
# <a name="custom-functions-requirement-sets"></a>自定义函数要求集

[自定义函数](../../excel/custom-functions-overview.md)使用独立于核心 Excel JavaScript API 的要求集。 下表列出了自定义函数要求集、受支持的 Office 客户端应用程序，以及这些应用程序的版本或版本号。

|  要求集  |  Office 2021 或更高版本Windows<br>（一次性购买）  |  Windows 版 Office<br>（关联至 Microsoft 365 订阅）  |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br> (两个订阅<br> 和一次购买 Office Mac 2021 及更高版本)   | Office 网页版 |
|:-----|:-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.3 | 16.0.14326.20454 或更高版本 | 16.0.13127.20296 或更高版本 | 不支持 | 16.40.20081000 或更高版本 | 2020 年 7 月 |
| CustomFunctionsRuntime 1.2 | 16.0.14326.20454 或更高版本 | 16.0.12527.20194 或更高版本 | 不支持 | 16.34.20020900 或更高版本 | 2020 年 1 月 |
| CustomFunctionsRuntime 1.1 | 16.0.14326.20454 或更高版本 | 16.0.12527.20092 或更高版本 | 不支持 | 16.34 或更高版本 | 2019 年 5 月 |

## <a name="customfunctionsruntime-11-12-and-13"></a>CustomFunctionsRuntime 1.1、1.2 和 1.3

CustomFunctionsRuntime 1.1 是 API 的第一个版本。 要求集 1.2 添加对象 `CustomFunctions.Error` 以支持错误处理。 要求集 1.3 向 [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) 对象添加了 [XLL](../../excel/make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) `ErrorCode` 流式处理支持和新选项。

## <a name="see-also"></a>另请参阅

- [自定义函数参考文档](/javascript/api/custom-functions-runtime)
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md)
