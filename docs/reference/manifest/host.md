---
title: 清单文件中的 Host 元素
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: debb4d59f75ce974ffb21d853c6b65a579c4e685
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127567"
---
# <a name="host-element"></a>Host 元素

指定应在其中激活外接程序的单个 Office 应用程序类型。

> [!IMPORTANT] 
> **Host** 元素的语法根据该元素是否在[基本清单](#basic-manifest)中或 [VersionOverrides](#versionoverrides-node) 节点中定义而不同。 但功能相同。  

## <a name="basic-manifest"></a>基本清单

在基本清单（在 [OfficeApp](officeapp.md) 下）中定义时，主机类型由 `Name` 属性决定。

### <a name="attributes"></a>属性

| 属性     | 类型   | 必需 | 说明                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [Name](#name) | string | 必需 | Office 主机应用程序的类型名称。 |

### <a name="name"></a>名称
指定此外接程序面向的主机类型。值必须为以下值之一：

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

### <a name="example"></a>示例
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a>VersionOverrides 节点
在 [VersionOverrides](versionoverrides.md) 中定义时，主机类型由 `xsi:type` 属性决定。 

### <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  是  | 描述这些设置适用的 Office 主机。|

### <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [DesktopFormFactor](desktopformfactor.md)    |  是   |  定义桌面外形规格的设置。 |
|  [MobileFormFactor](mobileformfactor.md)    |  否   |  定义移动设备规格的设置。 **注意:** 仅在 iOS 上的 Outlook 中支持此元素。 |
|  [AllFormFactors](allformfactors.md)    |  否   |  定义所有外形规格的设置。 仅用于 Excel 中的自定义函数。 |

### <a name="xsitype"></a>xsi:type

控制所包含的设置适用的 Office 主机类别（Word、Excel、PowerPoint、Outlook 和 OneNote）。 值必须为以下值之一：

- `Document` (Word)
- `MailHost` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## <a name="host-example"></a>主机示例 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
