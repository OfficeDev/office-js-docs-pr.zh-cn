---
title: 清单文件中的 Requirements 元素
description: Requirements 元素指定外接程序所需的最低要求集和方法Office外接程序需要由 Office 或替代基本清单设置。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 85dcd08f3bfcffe34c4c479608f25ea0c2b6a134
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222281"
---
# <a name="requirements-element"></a>Requirements 元素

此元素的含义取决于它是在基本清单中使用，还是 [](#in-the-base-manifest)用作 [**VersionOverrides** 元素的子元素](#as-a-child-of-a-versionoverrides-element)。

> [!TIP]
> 使用此元素之前，请熟悉指定[Office和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)

## <a name="in-the-base-manifest"></a>在基本清单中

在基本清单 (（即 [OfficeApp](officeapp.md)) 的直接子级）中使用时 **，Requirements** 元素指定 Office JavaScript API 要求的最小集 (要求集和/或方法 [](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)) Office 外接程序需要由 Office 激活。 外接程序将不会在不支持指定方法和要求集的 Office 版本和平台 (（如 Windows、Mac、Web 和 iOS 或 iPad) ）的任意组合上激活。

**外接程序类型：** 任务窗格、邮件

## <a name="as-a-child-of-a-versionoverrides-element"></a>作为 VersionOverrides 元素的子元素

当用作 [VersionOverrides](versionoverrides.md)的子项时，指定 Office 版本和平台 (（如 Windows、Mac、Web 和 iOS 或 iPad) ）必须支持的 Office JavaScript API 要求 (要求集和/或方法) ，以便替代基本清单设置的 **VersionOverrides** 元素中的设置 [](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) 生效。

请考虑在基本清单中指定要求 A 的外接程序，并指定 **VersionOverrides** 中的要求 B。 

- 如果平台Office版本不支持 A，则外接程序不会激活，Office不分析清单的 **VersionOverrides** 部分。 
- 如果同时支持 A 和 B，则激活外接程序， **并且 VersionOverrides** 中的所有标记均生效。 
- 如果支持 A，但不支持 B，则激活外接程序，**并且 VersionOverrides** 中的某些标记将生效。  具体而言，不替代基本清单元素 **的 VersionOverrides** 的子元素将生效。 例如 **，WebApplicationInfo** 元素或 **EquivalentAddins** 生效。 但是 **，VersionOverrides** 的所有替代基本清单元素的子元素（如 **Hosts）** 不会生效。 相反，Office会使用原本已被替代的基本清单标记的值。 

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- 当父 **VersionOverrides** 的类型为 Taskpane [1.0 时，AddinCommands](../requirement-sets/add-in-commands-requirement-sets.md) 1.1。
- [当父](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) **VersionOverrides** 类型为 Mail 1.0 时，邮箱 1.3。
- [当父](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) **VersionOverrides** 类型为 Mail 1.1 时，邮箱 1.5。

### <a name="remarks"></a>备注

如果 **Requirements** 元素未指定在基本清单的 **Requirements** 中未指定的其他要求，则它在 **VersionOverrides** 中没有任何用途。 如果Office版本和平台不支持基本清单中的要求，则不激活外接程序，并且不会分析 **VersionOverrides** 元素。 因此，只有在满足以下两个条件时，才应在 **VersionOverrides** 中使用 **Requirements** 元素：

- 您的外接程序具有通过 **VersionOverrides** (（如外接程序命令) ）中的配置实现的附加功能，并且要求在基本清单的 **Requirements** 元素中未指定的方法或要求集。 
- 外接程序非常有用，应该 (但无需额外功能) ，即使平台和 Office 版本的组合不支持额外功能的要求。

> [!TIP]
> 不要重复 **VersionOverrides** 中基本清单中的 Requirement 元素。 这样做没有任何效果，并且可能会令人误解 **VersionOverrides** 中的 **Requirements** 元素的用途。

> [!WARNING]
> 在 **VersionOverrides** 中使用 **Requirements** 元素之前，请谨慎，因为在不支持该要求的平台和版本组合上，不会安装任何外接程序命令，甚至不会安装调用不需要要求的功能 *的命令。* 例如，请考虑具有两个自定义功能区按钮的外接程序。 其中一个Office调用要求集 **ExcelApi 1.4** (及更高版本中可用的 javaScript) 。 其他调用仅在 **ExcelApi 1.9** (及更高版本) 。 如果在 **VersionOverrides** 中对 **ExcelApi 1.9** 提出要求，*两个按钮* 都不会出现在功能区上。 此方案中更好的策略是使用运行时检查方法和要求集支持 [中所述的技术](../../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support)。 第二个按钮调用的代码首先 `isSetSupported` 用于检查是否支持 **ExcelApi 1.9**。 如果不支持此功能，则代码会向用户显示一条消息，指出外接程序的此功能不适用于其 Office。 

> [!NOTE]
> 在邮件外接程序中 **，VersionOverrides** 1.1 可以嵌套在 **VersionOverrides** 1.0 中。 Office始终使用平台支持的最高 **版本 VersionOverrides，Office** 版本。

## <a name="syntax"></a>语法

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md) 
[VersionOverrides](versionoverrides.md)

## <a name="can-contain"></a>可以包含

|元素|内容|邮件|任务窗格|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[方法](methods.md)|x||x|

## <a name="see-also"></a>另请参阅

有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。
