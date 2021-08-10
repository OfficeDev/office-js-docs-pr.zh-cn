---
title: 清单文件中的类型元素
description: Type 元素指定等效加载项是 COM 加载项还是 XLL。
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: ca6fa7183727870593dd3e726abc72fdc0d6f0b518fdb8451ec80c6b590f8c83
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092473"
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