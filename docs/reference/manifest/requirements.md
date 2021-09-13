---
title: 清单文件中的 Requirements 元素
description: Requirements 元素指定外接程序激活Office要求集和方法。
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 3a5a393485094b5cc830b5120c3abd8c211eff1e
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152384"
---
# <a name="requirements-element"></a>Requirements 元素

指定 JavaScript API 的最低Office要求 (要求集和/或) 加载项[](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)Office激活的方法。

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
