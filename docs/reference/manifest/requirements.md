---
title: 清单文件中的 Requirements 元素
description: Requirements 元素指定外接程序激活Office要求集和方法。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3020037b48e3f759acf6a7e2758bb8c1fd2dd36429e0b21613e22fca33cacc1a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098101"
---
# <a name="requirements-element"></a>Requirements 元素

指定 JavaScript API 的最低Office要求 (要求集和/或) [](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)加载项Office需要激活的方法。

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
