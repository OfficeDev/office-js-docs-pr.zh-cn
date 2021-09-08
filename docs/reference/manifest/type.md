---
title: 清单文件中的类型元素
description: Type 元素指定等效加载项是 COM 加载项还是 XLL。
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 5af3359c232e91b097311bfc06fc9b1c932b0703
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938895"
---
# <a name="type-element"></a>Type 元素

指定等效加载项是 COM 加载项还是 XLL。

**外接程序类型：** 任务窗格、自定义函数

## <a name="syntax"></a>语法

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>包含于

[EquivalentAddin](equivalentaddin.md)

## <a name="add-in-type-values"></a>外接程序类型值

必须为 元素指定下列值之 `Type` 一。

- COM：指定等效加载项是 COM 加载项。
- XLL：指定等效加载项是一个Excel XLL。

## <a name="see-also"></a>另请参阅

- [让自定义功能与 XLL 用户定义的功能兼容](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [让 Office 加载项与现有 COM 加载项兼容](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)