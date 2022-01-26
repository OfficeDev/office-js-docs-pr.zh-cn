---
title: 清单文件中的 Methods 元素
description: Methods 元素指定 Office 外接程序Office激活或替代基本清单设置Office JavaScript API 方法的列表。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4c39c6363cd33e103cf40c0f7f047fa694db1411
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222274"
---
# <a name="methods-element"></a>Methods 元素

此元素的含义取决于它在清单中的使用位置。

## <a name="in-the-base-manifest"></a>在基本清单中

在基本清单 (即，父 **Requirements** 元素是 [OfficeApp](officeapp.md)) 的直接子级时 **，Methods** 元素指定 Office 外接程序需要由 Office 激活的 Office JavaScript API 方法的列表。

**外接程序类型：** 内容、任务窗格

## <a name="as-a-grandchild-of-a-versionoverrides-element"></a>作为 VersionOverrides 元素的子级

指定 Office 版本和平台 (（如 Windows、Mac、Web 和 iOS 或 iPad) ）必须支持的 Office JavaScript API 方法的最小集，以便[VersionOverrides](versionoverrides.md)生效。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 与父 [Requirements](requirements.md) 元素相同。

**与以下要求集相关联**：

- 与父 [Requirements](requirements.md) 元素相同。

## <a name="syntax"></a>语法

```XML
<Methods>
   ...
</Methods>
```

## <a name="contained-in"></a>包含于

[要求](requirements.md)

## <a name="can-contain"></a>可以包含

[方法](method.md)

## <a name="remarks"></a>注解

在 **基本** 清单中使用时，Methods 和 **Method** 元素在邮件外接程序中不受支持。 有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。
