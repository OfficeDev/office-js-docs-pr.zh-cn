---
title: 清单文件中 EquivalentAddin 元素
description: 指定等效 COM 加载项或 XLL 的向后兼容性。
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: e318a9028ebefdeca9aaf5baac465a1ec1af0a73
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042131"
---
# <a name="equivalentaddin-element"></a>EquivalentAddin 元素

指定等效 COM 加载项或 XLL 的向后兼容性。

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**外接程序类型：** 任务窗格、邮件、自定义函数

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

## <a name="syntax"></a>语法

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>包含于

[EquivalentAddins](equivalentaddins.md)

## <a name="must-contain"></a>必须包含

[类型](type.md)

## <a name="can-contain"></a>可以包含

[ProgId](progid.md) 
[FileName](filename.md)

## <a name="remarks"></a>说明

若要将 COM 加载项指定为等效加载项，请提供 和 `ProgId` `Type` 元素。 若要将 XLL 指定为等效的外接程序，请提供 和 `FileName` `Type` 元素。

## <a name="see-also"></a>另请参阅

- [让自定义功能与 XLL 用户定义的功能兼容](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [让 Office 加载项与现有 COM 加载项兼容](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)