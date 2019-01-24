---
title: 清单文件中的 VersionOverrides 元素
description: ''
ms.date: 01/15/2019
localization_priority: Normal
ms.openlocfilehash: 197a636169b7f00edd44019cee21686065845800
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387798"
---
# <a name="versionoverrides-element"></a>VersionOverrides 元素

根元素包含由外接程序实现的外接程序命令的信息。**VersionOverrides** 是清单中 [OfficeApp](./officeapp.md) 元素的子元素。此元素在清单架构 v1.1 及更高版本中受支持，但是在 VersionOverrides v1.0 或 v1.1 架构中进行定义。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **xmlns**       |  是  |  若 `xsi:type` 为 `VersionOverridesV1_0`，架构位置必须是 `http://schemas.microsoft.com/office/mailappversionoverrides`；若 `xsi:type` 为 `VersionOverridesV1_1`，架构位置必须是 `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`。|
|  **xsi:type**  |  是  | 架构版本。目前的唯一有效值为 `VersionOverridesV1_0` 和 `VersionOverridesV1_1`。 |

> [!NOTE]
> 当前只有 Outlook 2016 支持 VersionOverrides v1.1 架构和 `VersionOverridesV1_1` 类型。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **Description**    |  否   |  描述外接程序。这会替代清单中任何父级部分中的 `Description` 元素。说明文本包含在 [Rescources](./resources.md) 元素中的 **LongString** 元素的子元素中。**Description** 元素的 `resid` 属性被设置为包含文本的 `String` 元素的 `id` 属性的值。|
|  **Requirements**  |  否   |  指定外接程序要求的最低要求集和 Office.js 的版本。这会替代清单中父级部分中的 `Requirements` 元素。|
|  [Hosts](./hosts.md)                |  是  |  指定 Office 主机的集合。子级 Hosts 元素替代清单中父级部分中的 Hosts 元素。  |
|  [Resources](./resources.md)    |  是  | 定义其他清单元素引用一组的资源（字符串、URL 和图像）。|
|  **VersionOverrides**    |  否  | 在新版架构下定义外接程序命令。有关详细信息，请参阅[实现多个版本](#implementing-multiple-versions)。 |
|  **WebApplicationInfo**    |  否  | 指定加载项关联 Web 应用程序的详细信息。 |

### <a name="versionoverrides-example"></a>VersionOverrides 示例

下面是典型的示例`<VersionOverrides>`元素，包括不是必需的但通常使用的一些子元素。

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a>实现多个版本

清单可以实现 `VersionOverrides` 元素的多个版本，这些版本支持不同版本的 VersionOverrides 架构。为此，可以视情况支持新版架构中的新功能，同时仍支持不支持新功能的旧版客户端。

新版架构的 `VersionOverrides` 元素必须是旧版架构的 `VersionOverrides` 元素的子元素，才能实现多个版本。 `VersionOverrides` 子元素不会从父元素继承任何值。

若要实现 VersionOverrides v1.0 和 v1.1 架构，清单如以下示例所示：

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
  </VersionOverrides>
...
</OfficeApp>
```
