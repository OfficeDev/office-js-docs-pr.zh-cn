---
title: 清单文件中的 VersionOverrides 元素
description: Office清单的 VersionOverrides 元素参考文档 (XML) 文件。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 787ba8e7d90900cc72d6c5e9370d68ced0faee2f
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348655"
---
# <a name="versionoverrides-element"></a>VersionOverrides 元素

根元素包含由外接程序实现的外接程序命令的信息。**VersionOverrides** 是清单中 [OfficeApp](officeapp.md) 元素的子元素。此元素在清单架构 v1.1 及更高版本中受支持，但是在 VersionOverrides v1.0 或 v1.1 架构中进行定义。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **xmlns**       |  是  |  VersionOverrides 架构命名空间。 允许的值因此元素的 `<VersionOverrides>` **xsi：type** 值和父元素的 **xsi：type** 值 `<OfficeApp>` 而异。 请参阅 [下面的命名空间](#namespace-values) 值。|
|  **xsi:type**  |  是  | 架构版本。目前的唯一有效值为 `VersionOverridesV1_0` 和 `VersionOverridesV1_1`。 |

### <a name="namespace-values"></a>命名空间值

下面列出了 **xmlns** 值的必需值，具体取决于 **父元素的 xsi：type** `<OfficeApp>` 值。

- **TaskPaneApp** 仅支持 VersionOverrides 的 1.0 版 **，xmlns** 应为 `http://schemas.microsoft.com/office/taskpaneappversionoverrides` 。
- **ContentApp** 仅支持 VersionOverrides 的版本 1.0，xmlns 应为 `http://schemas.microsoft.com/office/contentappversionoverrides` 。
- **MailApp** 支持 VersionOverrides 的版本 1.0 和 1.1，因此 **xmlns** 的值因此元素的 `<VersionOverrides>` **xsi：type** 值而异：
    - 当 **xsi：type** 为 `VersionOverridesV1_0` 时 **，xmlns** 必须为 `http://schemas.microsoft.com/office/mailappversionoverrides` 。
    - 当 **xsi：type** 为 `VersionOverridesV1_1` 时 **，xmlns** 必须为 `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` 。

> [!NOTE]
> 当前仅Outlook 2016或更高版本支持 VersionOverrides v1.1 架构和 `VersionOverridesV1_1` 类型。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **说明**    |  否   |  描述外接程序。 这会替代清单中任何父级部分中的 `Description` 元素。 说明文本包含在 **Rescources** 元素中的 [LongString](resources.md) 元素的子元素中。 Description 元素的 属性不能超过 32 个字符，并设置为包含文本的元素 `resid`  `id` 的 `String` 属性的值。|
|  **Requirements**  |  否   |  指定外接程序要求的最低要求集和 Office.js 的版本。这会替代清单中父级部分中的 `Requirements` 元素。|
|  [Hosts](hosts.md)                |  是  |  指定应用程序Office集合。 子 Hosts 元素替代清单的父部分中的 Hosts 元素。  |
|  [Resources](resources.md)    |  是  | 定义其他清单元素引用的资源集合（字符串、URL 和图像）。|
|  [EquivalentAddins](equivalentaddins.md)    |  否  | 指定与 web (等效) COM/XLL 加载项的本机属性。 如果安装了等效的本机外接程序，则不激活 Web 外接程序。|
|  **VersionOverrides**    |  否  | 在新版架构下定义外接程序命令。有关详细信息，请参阅[实现多个版本](#implementing-multiple-versions)。 |
|  [WebApplicationInfo](webapplicationinfo.md)    |  否  | 指定有关外接程序注册到安全令牌颁发者（如 Azure Active Directory V2.0）的详细信息。 |
|  [ExtendedPermissions](extendedpermissions.md) |  否  |  指定扩展权限的集合。 |

### <a name="versionoverrides-example"></a>VersionOverrides 示例

下面是典型元素的示例，包括一些不需要但 `<VersionOverrides>` 通常使用的子元素。

```xml
<OfficeApp ... xsi:type="MailApp">
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

若要同时实现 VersionOverrides v1.0 和 v1.1 架构，清单将类似于以下示例。

```xml
<OfficeApp ... xsi:type="MailApp">
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
