---
title: 清单文件中的 Type 元素
description: Type 元素指定等效加载项是 COM 加载项还是 XLL。
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: 9eeab172ed4ebf06fc93e42f56f8d33f5e7a92db
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720314"
---
# <a name="type-element"></a>Type 元素

指定等效的外接程序是 COM 加载项还是 XLL。

**外接类型：** 任务窗格，自定义函数

## <a name="syntax"></a>语法

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>包含于

[EquivalentAdd-in](equivalentaddin.md)

## <a name="add-in-type-values"></a>外接类型值

必须为`Type`元素指定下列值之一。

- COM：指定等效的加载项是 COM 加载项。
- XLL：指定等效的外接程序是 Excel XLL。

## <a name="see-also"></a>另请参阅

- [让自定义功能与 XLL 用户定义的功能兼容](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [使 Excel 外接程序与现有 COM 外接程序兼容](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)