---
title: 清单文件中的 Requirements 元素
description: Requirements 元素指定外接程序激活Office所需的最低要求集和方法。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 319ddc59901c524ed1cee580a81cff749ad570db
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938451"
---
# <a name="requirements-element"></a>Requirements 元素

指定 JavaScript API 要求Office要求 (要求集和/或) 加载项[](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)Office需要的方法。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>可以包含

|元素|内容|邮件|任务窗格|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[方法](methods.md)|x||x|

## <a name="remarks"></a>注释

有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。
