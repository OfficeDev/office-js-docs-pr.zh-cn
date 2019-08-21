---
title: 清单文件中的 VersionOverrides 元素
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: ce65cdced1b3cf885cee09732c2cda0081a53cfc
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477878"
---
# <a name="versionoverrides-element"></a>VersionOverrides 元素

根元素包含由外接程序实现的外接程序命令的信息。**VersionOverrides** 是清单中 [OfficeApp](./officeapp.md) 元素的子元素。此元素在清单架构 v1.1 及更高版本中受支持，但是在 VersionOverrides v1.0 或 v1.1 架构中进行定义。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **xmlns**       |  是  |  若 `http://schemas.microsoft.com/office/mailappversionoverrides` 为 `xsi:type`，架构位置必须是 `VersionOverridesV1_0`；若 `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` 为 `xsi:type`，架构位置必须是 `VersionOverridesV1_1`。|
|  **xsi:type**  |  是  | 架构版本。目前的唯一有效值为 `VersionOverridesV1_0` 和 `VersionOverridesV1_1`。 |

> [!NOTE]
> 目前, 只有 Outlook 2016 或更高版本支持 VersionOverrides v1.1 架构和`VersionOverridesV1_1`类型。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **说明**    |  否   |  描述外接程序。这会替代清单中任何父级部分中的 `Description` 元素。说明文本包含在 **Rescources** 元素中的 [LongString](./resources.md) 元素的子元素中。`resid` 元素的 **** 属性被设置为包含文本的 `id` 元素的 `String` 属性的值。|
| **EquivalentAddins** | 否 | 指定与等效的 COM 外接程序、XLL 或这两者的向后兼容性。 |
|  **Requirements**  |  否   |  指定外接程序要求的最低要求集和 Office.js 的版本。这会替代清单中父级部分中的 `Requirements` 元素。|
|  [Hosts](./hosts.md)                |  是  |  指定 Office 主机的集合。子级 Hosts 元素替代清单中父级部分中的 Hosts 元素。  |
|  [Resources](./resources.md)    |  是  | 定义其他清单元素引用的资源集合（字符串、URL 和图像）。|
|  [EquivalentAddins](./equivalentaddins.md)    |  否  | 指定与 web 外接程序等效的本机 (COM/XLL) 加载项。 如果安装了等效的本机加载项, 则不会激活 web 外接程序。|
|  **VersionOverrides**    |  否  | 在新版架构下定义外接程序命令。有关详细信息，请参阅[实现多个版本](#implementing-multiple-versions)。 |
|  [WebApplicationInfo](./webapplicationinfo.md)    |  否  | 指定有关使用安全令牌颁发者 (如 Azure Active Directory v2.0) 的加载项注册的详细信息。 |

### <a name="versionoverrides-example"></a>VersionOverrides 示例

下面是典型`<VersionOverrides>`元素的一个示例, 其中包括一些不需要但通常使用的子元素。

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
