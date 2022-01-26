---
title: 清单文件中的 Sets 元素
description: Sets 元素 Office指定您的 Office 外接程序需要的最小 JavaScript API Office或替代基本清单设置。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: df0cf686fe213a51321595a000438ca2a411f2c7
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222141"
---
# <a name="sets-element"></a>Sets 元素

此元素的含义取决于它在清单中的使用位置。

## <a name="in-the-base-manifest"></a>在基本清单中

在基本清单 (中（即，父 **Requirements** 元素是 [OfficeApp](officeapp.md)) 的直接子级）中时 **，Sets** 元素指定 Office JavaScript API 要求 (要求集) ，Office [](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)外接程序需要这些要求集的最小子集，Office 才能激活这些要求。

**外接程序类型：** 内容、任务窗格、邮件

## <a name="as-a-grandchild-of-a-versionoverrides-element"></a>作为 VersionOverrides 元素的子级

指定 Office 版本和平台[ (（](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)如 Windows、Mac、Web 和 iOS 或 iPad) ）必须支持的 Office JavaScript API 要求 (要求集) ，以便[VersionOverrides](versionoverrides.md)生效。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 与父 [Requirements](requirements.md) 元素相同。

**与以下要求集相关联**：

- 与父 [Requirements](requirements.md) 元素相同。

## <a name="syntax"></a>语法

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>包含于

[要求](requirements.md)

## <a name="can-contain"></a>可以包含

[Set](set.md)

## <a name="attributes"></a>属性

|属性|类型|必需|说明|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|字符串|可选|指定所有子 **Set** 元素的默认 MinVersion [属性值。](set.md) 默认值为“1.1”。|

## <a name="remarks"></a>注释

有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

有关 **Set** 元素的 **MinVersion** 属性和 **Sets** 元素的 **DefaultMinVersion** 属性详细信息，请参阅指定哪些 Office 版本和平台可以托管 [您的外接程序](../../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)。

