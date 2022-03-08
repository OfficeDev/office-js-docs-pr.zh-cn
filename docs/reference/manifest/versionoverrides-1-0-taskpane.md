---
title: 任务窗格外接程序的清单文件中 VersionOverrides 1.0 元素
description: " (XML) 清单Office的 VersionOverrides 元素 (文档) 文档。"
ms.date: 02/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: f2d6867db8a8b35d4296b9907e4dbbb440ea28db
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340244"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-task-pane-add-in"></a>任务窗格外接程序的清单文件中 VersionOverrides 1.0 元素

此元素包含基本清单中不支持的功能的信息。

> [!NOTE]
> 本文假定你熟悉 [VersionOverrides](versionoverrides.md) 元素的概述，该元素包含有关该元素的属性和变体的重要信息。

**外接程序类型：** 任务窗格

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) (Excel、PowerPoint 和 Word.) 
- 某些子元素可能与其他要求集相关联。

## <a name="child-elements"></a>子元素

下表仅适用于 **VersionOverrides** 元素的版本 1.0，仅适用于任务窗格外接程序。

> [!NOTE]
> 在 iOS 中，仅 **支持 WebApplicationInfo** 。 **VersionOverrides 的所有其他子元素** 将被忽略。

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [说明](#description)    |  否   |  描述外接程序。 |
|  [Requirements](requirements.md)  |  否   |  指定为使父 **VersionOverrides** 中的标记生效而必须支持的最低要求集。 这应始终 *比* 清单的基 **部分中的 Requirements** 元素更加严格。|
|  [Hosts](hosts.md)                |  是  |  指定应用程序Office集合。 子 Hosts 元素替代清单的父部分中的 Hosts 元素。  |
|  [Resources](resources.md)    |  是  | 定义其他清单元素引用的资源集合（字符串、URL 和图像）。|
|  [EquivalentAddins](equivalentaddins.md)    |  否  | 指定与 web (等效) COM/XLL 加载项的本机属性。 如果安装了等效的本机外接程序，则不激活 Web 外接程序。|
|  **VersionOverrides**    |  否  | 当前在任务窗格加载项的 VersionOverrides 1.0 中不可用。 |
|  [WebApplicationInfo](webapplicationinfo.md)    |  否  | 指定有关外接程序注册到安全令牌颁发者（如 Azure Active Directory V2.0）的详细信息。 |

### <a name="description"></a>Description

描述外接程序。 这将覆盖清单任何父部分的 **Description** 元素。 说明文本包含在 **Rescources** 元素中的 [LongString](resources.md) 元素的子元素中。 `resid` Description 元素的 **属性** 不能超过 32 `id` 个字符，并且必须匹配包含在 Resources 元素中的 **ShortString** 元素的子元素的 [属性的值](resources.md)。

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

## <a name="example"></a>示例

下面展示了一个非常简单的示例。 有关更复杂的示例，请参阅外接程序代码示例中的示例Office[清单](https://github.com/OfficeDev/PnP-OfficeAddins)。

```xml
<OfficeApp ... xsi:type="Taskpane">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="Workbook">
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
