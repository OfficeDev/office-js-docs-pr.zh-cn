---
title: 清单文件中的 Method 元素
description: Method 元素指定 Office JavaScript API 中的单个方法，Office 外接程序需要该方法才能由 Office 激活或替代基本清单设置。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 052fb41a7077781843ea7e63d9601a819058dfa6
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222267"
---
# <a name="method-element"></a>Method 元素

此元素的含义取决于它在清单中的使用位置。

## <a name="in-the-base-manifest"></a>在基本清单中

在基本清单 (即，当您在基本清单中使用的元素是 [OfficeApp](officeapp.md)) 的直接子级时 **，Method** 元素会指定 Office JavaScript API 中的单个方法，Office 外接程序需要该方法才能由 Office 激活。

**外接程序类型：** 内容、任务窗格

## <a name="as-a-great-grandchild-of-a-versionoverrides-element"></a>作为 VersionOverrides 元素的子级

指定 Office JavaScript API 中必须受 Office 版本和平台 (（如 Windows、Mac、Web 和 iOS 或 iPad) ）支持的单个方法，以便[VersionOverrides](versionoverrides.md)生效。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 与上一个 [Requirements](requirements.md) 元素相同。

**与以下要求集相关联**：

- 与上一个 [Requirements](requirements.md) 元素相同。

## <a name="syntax"></a>语法

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>包含于

[Methods](methods.md)

## <a name="attributes"></a>属性

|属性|类型|必需|说明|
|:-----|:-----|:-----|:-----|
|Name|字符串|必需|指定由其父对象限定的所需方法的名称。 例如，若要指定 `getSelectedDataAsync` 方法，必须指定 `"Document.getSelectedDataAsync"` 。|

## <a name="remarks"></a>备注

在 **基本** 清单中使用时，邮件外接程序不支持 Methods 和 **Method** 元素。 有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

> [!IMPORTANT]
> 因为无法指定各个方法的最低版本要求，所以为了确保在运行时提供可用的方法，当您在外接程序的脚本中调用该方法时，还应该使用 **if** 语句。 若要详细了解如何操作，请参阅了解 Office [JavaScript API。](../../develop/understanding-the-javascript-api-for-office.md)
