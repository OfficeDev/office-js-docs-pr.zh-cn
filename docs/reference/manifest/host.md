---
title: 清单文件中的 Host 元素
description: 指定应在其中激活外接程序的单个 Office 应用程序类型。
ms.date: 11/05/2019
ms.localizationpriority: medium
ms.openlocfilehash: e262183f11361d02d1a710d1f442539a8e936bcd
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152596"
---
# <a name="host-element"></a>Host 元素

指定应在其中激活外接程序的单个 Office 应用程序类型。

> [!IMPORTANT]
> **Host** 元素的语法根据该元素是否在 [基本清单](#basic-manifest)中或 [VersionOverrides](#versionoverrides-node) 节点中定义而不同。 但功能相同。  

## <a name="basic-manifest"></a>基本清单

在基本清单（在 [OfficeApp](officeapp.md) 下）中定义时，主机类型由 `Name` 属性决定。

### <a name="attributes"></a>属性

| 属性     | 类型   | 必需 | 说明                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [Name](#name) | string | 必需 | 客户端应用程序Office的名称。 |

### <a name="name"></a>名称

指定此外接程序面向的主机类型。值必须为以下值之一：

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

> [!IMPORTANT]
> 我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。 作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。

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
|  [xsi:type](#xsitype)  |  是  | 介绍Office应用这些设置的应用程序。|

### <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [DesktopFormFactor](desktopformfactor.md)    |  是   |  定义桌面外形规格的设置。 |
|  [MobileFormFactor](mobileformfactor.md)    |  否   |  定义移动外形因素的设置。 **注意：** 此元素仅在 iOS Outlook Android 上的设备上受支持。 |
|  [AllFormFactors](allformfactors.md)    |  否   |  定义所有外形规格的设置。 仅用于 Excel 中的自定义函数。 |

### <a name="xsitype"></a>xsi:type

控制应用程序Office Word (、Excel、PowerPoint、Outlook、OneNote) 应用包含的设置。 值必须为以下值之一：

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
