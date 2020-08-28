---
title: 清单文件中的 Host 元素
description: 指定应在其中激活外接程序的单个 Office 应用程序类型。
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 5b6c6e6b5471b4117c28cf92e11eb0a99b512a97
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292284"
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
| [Name](#name) | string | 必需 | Office 客户端应用程序类型的名称。 |

### <a name="name"></a>名称

指定此外接程序面向的主机类型。 值必须是下列值之一。

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
|  [xsi:type](#xsitype)  |  是  | 介绍应用这些设置的 Office 应用程序。|

### <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [DesktopFormFactor](desktopformfactor.md)    |  是   |  定义桌面外形规格的设置。 |
|  [MobileFormFactor](mobileformfactor.md)    |  否   |  定义移动设备规格的设置。 **注意：** 仅在 iOS 和 Android 上的 Outlook 中支持此元素。 |
|  [AllFormFactors](allformfactors.md)    |  否   |  定义所有外形规格的设置。 仅用于 Excel 中的自定义函数。 |

### <a name="xsitype"></a>xsi:type

控制包含的设置应用 (Word、Excel、PowerPoint、Outlook、OneNote) 的 Office 应用程序。 值必须为以下值之一：

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
