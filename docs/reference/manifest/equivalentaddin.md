---
title: 清单文件中的 EquivalentAddin 元素
description: ''
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 33cfb8b73e050fad7e392e0234962d346e903713
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059921"
---
# <a name="equivalentaddin-element"></a>EquivalentAddin 元素

为等效的 COM 外接程序或 XLL 指定向后兼容性。

**外接类型:** 任务窗格, 自定义函数

## <a name="syntax"></a>语法

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>包含于

[EquivalentAdd-ins](equivalentaddins.md)

## <a name="must-contain"></a>必须包含

[Type](type.md)

## <a name="can-contain"></a>可以包含

[ProgId](progid.md)
[文件名](filename.md)

## <a name="remarks"></a>说明

若要将 COM 加载项指定为等效的`ProgId`加载项, 请同时提供和`Type`元素。 若要将 XLL 指定为等效的外接程序, 请同时`FileName`提供`Type`和元素。

## <a name="see-also"></a>另请参阅

- [让自定义功能与 XLL 用户定义的功能兼容](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [使 Excel 外接程序与现有 COM 外接程序兼容](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)