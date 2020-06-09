---
title: 清单文件中的 EquivalentAddin 元素
description: 为等效的 COM 外接程序或 XLL 指定向后兼容性。
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: e14fe91bf7a5fe321019acf205ddb1753fedd569
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611559"
---
# <a name="equivalentaddin-element"></a>EquivalentAddin 元素

为等效的 COM 外接程序或 XLL 指定向后兼容性。

**外接类型：** 任务窗格，自定义函数

## <a name="syntax"></a>语法

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>包含于

[EquivalentAdd-ins](equivalentaddins.md)

## <a name="must-contain"></a>必须包含

[类型](type.md)

## <a name="can-contain"></a>可以包含

[ProgId](progid.md) 
[FileName](filename.md)

## <a name="remarks"></a>备注

若要将 COM 加载项指定为等效的加载项，请同时提供 `ProgId` 和 `Type` 元素。 若要将 XLL 指定为等效的外接程序，请同时提供 `FileName` 和 `Type` 元素。

## <a name="see-also"></a>另请参阅

- [让自定义功能与 XLL 用户定义的功能兼容](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [使 Excel 外接程序与现有 COM 外接程序兼容](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)