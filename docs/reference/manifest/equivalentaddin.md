---
title: 清单文件中 EquivalentAddin 元素
description: 指定等效 COM 加载项或 XLL 的向后兼容性。
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: f77a70681c8a12674d9e22022276e511552861ad
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990689"
---
# <a name="equivalentaddin-element"></a>EquivalentAddin 元素

指定等效 COM 加载项或 XLL 的向后兼容性。

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**外接程序类型：** 任务窗格、邮件、自定义函数

## <a name="syntax"></a>语法

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>包含于

[EquivalentAddins](equivalentaddins.md)

## <a name="must-contain"></a>必须包含

[Type](type.md)

## <a name="can-contain"></a>可以包含

[ProgId](progid.md) 
[FileName](filename.md)

## <a name="remarks"></a>说明

若要将 COM 加载项指定为等效加载项，请提供 和 `ProgId` `Type` 元素。 若要将 XLL 指定为等效的外接程序，请提供 和 `FileName` `Type` 元素。

## <a name="see-also"></a>另请参阅

- [让自定义功能与 XLL 用户定义的功能兼容](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [让 Office 加载项与现有 COM 加载项兼容](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)