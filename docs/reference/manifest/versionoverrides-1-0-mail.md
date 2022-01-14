---
title: 邮件外接程序的清单文件中 VersionOverrides 1.0 元素
description: VersionOverrides 元素的参考文档 (XML) 文件Office外接程序 (邮件) 文档。
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: b433a52ad922fb3d397993a3861038f2f82ff165
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042167"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-mail-add-in"></a>邮件外接程序的清单文件中 VersionOverrides 1.0 元素

此元素包含基本清单中不支持的功能的信息。

> [!NOTE]
> 本文假定你熟悉 [VersionOverrides](versionoverrides.md)元素的概述，该元素包含有关元素的属性和变体的重要信息。

**外接程序类型：** 邮件

**仅在以下 VersionOverrides 架构中有效**：

- 邮件 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [Mailbox 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)
- 某些子元素可能与其他要求集相关联。

## <a name="child-elements"></a>子元素

下表仅适用于 **VersionOverrides** 元素的版本 1.0，仅适用于邮件外接程序。

> [!NOTE]
> 在 iOS 中， `<WebApplicationInfo>` 仅受支持。 **VersionOverrides 的所有其他子元素** 将被忽略。

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [说明](#description)    |  否   |  描述外接程序。 |
|  [Requirements](requirements.md)  |  否   |  指定为了使父项中的标记生效而必须支持的最低 `VersionOverrides` 要求集。 这应 *始终比* 清单的基本 `Requirements` 部分中的 元素更加严格。|
|  [Hosts](hosts.md)                |  是  |  指定应用程序Office集合。 子 Hosts 元素替代清单的父部分中的 Hosts 元素。  |
|  [Resources](resources.md)    |  是  | 定义其他清单元素引用一组的资源（字符串、URL 和图像）。|
|  **VersionOverrides**    |  否  | 在新版架构下定义外接程序命令。有关详细信息，请参阅[实现多个版本](#implementing-multiple-versions)。 |
|  [WebApplicationInfo](webapplicationinfo.md)    |  否  | 指定有关外接程序注册到安全令牌颁发者（如 Azure Active Directory V2.0）的详细信息。 |

### <a name="description"></a>Description

描述外接程序。 这会替代清单中任何父级部分中的 `Description` 元素。 说明文本包含在 **Rescources** 元素中的 [LongString](resources.md) 元素的子元素中。 Description 元素的 属性不能超过 32 个字符，并且必须匹配包含在 Resources 元素中的 `resid`  `id` **ShortString** 元素的子元素的 [属性的值](resources.md)。 

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- 当父级为 Taskpane [1.0 类型时，AddinCommands](../requirement-sets/add-in-commands-requirement-sets.md) `<VersionOverrides>` 1.1。
- [当父级为 Mail 1.0](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) `<VersionOverrides>` 类型时，邮箱 1.3。
- [当父级为](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) Mail `<VersionOverrides>` 1.1 类型时，邮箱 1.5。

## <a name="example"></a>示例

下面展示了一个非常简单的示例。 有关完整示例，请参阅外接程序代码示例中的示例Office[清单](https://github.com/OfficeDev/PnP-OfficeAddins)。

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
