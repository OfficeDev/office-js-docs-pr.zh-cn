---
title: 清单文件中的 VersionOverrides 元素
description: Office清单的 VersionOverrides 元素参考文档 (XML) 文件。
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 657bdebbc88993badd9d0e60946239edd55d5533
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042145"
---
# <a name="versionoverrides-element"></a>VersionOverrides 元素

此元素包含基本清单中不支持的功能的信息。 它的子标记可能会替代基本清单或父 **VersionOverrides** (中的某些) 。 **VersionOverrides** 是清单中的 [根 OfficeApp](officeapp.md) 元素或父 **VersionOverrides 元素的子** 元素。 此元素在清单架构 v1.1 及更高版本中受支持，但在单独的 VersionOverrides 架构中定义。

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **xmlns**       |  是  |  VersionOverrides 架构命名空间。 允许的值因此元素的 `<VersionOverrides>` **xsi：type** 值和父元素的 **xsi：type** 值 `<OfficeApp>` 而异。 请参阅 [下面的命名空间](#namespace-values) 值。|
|  **xsi:type**  |  是  | 架构版本。目前的唯一有效值为 `VersionOverridesV1_0` 和 `VersionOverridesV1_1`。 |

### <a name="namespace-values"></a>命名空间值

下面列出了 **xmlns** 属性的必需值，具体取决于根元素 **的 xsi：type** `<OfficeApp>` 值。

- **TaskPaneApp** 仅支持 VersionOverrides 的 1.0 版 **，xmlns** 必须为 `http://schemas.microsoft.com/office/taskpaneappversionoverrides` 。
- **ContentApp** 仅支持 VersionOverrides 的 1.0 版， **并且 xmlns** 必须为 `http://schemas.microsoft.com/office/contentappversionoverrides` 。
- **MailApp** 支持 VersionOverrides 的版本 1.0 和 1.1，因此 **xmlns** 的值因此元素的 `<VersionOverrides>` **xsi：type** 值而异：
  - 当 **xsi：type** 为 `VersionOverridesV1_0` 时 **，xmlns** 必须为 `http://schemas.microsoft.com/office/mailappversionoverrides` 。
  - 当 **xsi：type** 为 `VersionOverridesV1_1` 时 **，xmlns** 必须为 `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` 。

> [!NOTE]
> 当前仅Outlook 2016或更高版本支持 VersionOverrides v1.1 架构和 `VersionOverridesV1_1` 类型。

## <a name="variant-schemas"></a>Variant 架构

每个可能的 **xmlns** 值具有不同的架构，因此每个都有一个单独的引用页。

- [VersionOverrides 1.0 TaskPane](versionoverrides-1-0-taskpane.md)
- [VersionOverrides 1.0 内容](versionoverrides-1-0-content.md)
- [VersionOverrides 1.0 Mail](versionoverrides-1-0-mail.md)
- [VersionOverrides 1.1 邮件](versionoverrides-1-1-mail.md)
