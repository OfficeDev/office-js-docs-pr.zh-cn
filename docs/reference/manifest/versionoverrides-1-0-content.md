---
title: VersionOverrides 内容外接程序的清单文件中 1.0 元素
description: VersionOverrides 元素的参考 (XML) Office清单 (内容) 文档。
ms.date: 02/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0ef083ef5df322c230292625576e36db8923d00c
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63341049"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-content-add-in"></a>VersionOverrides 内容外接程序的清单文件中 1.0 元素

此元素包含基本清单中不支持的功能的信息。

> [!NOTE]
> 本文假定你熟悉 [VersionOverrides](versionoverrides.md) 元素的概述，该元素包含有关该元素的属性和变体的重要信息。

## <a name="child-elements"></a>子元素

下表仅适用于 **VersionOverrides** 元素的版本 1.0，仅适用于内容外接程序。

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **VersionOverrides**    |  否  | 当前在 VersionOverrides 1.0 中对内容外接程序不可用。 |
|  [WebApplicationInfo](webapplicationinfo.md)    |  否  | 指定有关外接程序注册到安全令牌颁发者（如 Azure Active Directory V2.0）的详细信息。 |

## <a name="example"></a>示例

下面展示了一个非常简单的示例。 有关更复杂的示例，请参阅外接程序代码示例中的示例Office[清单](https://github.com/OfficeDev/PnP-OfficeAddins)。

```xml
<OfficeApp ... xsi:type="Content">
...
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/contentappversionoverrides" xsi:type="VersionOverridesV1_0">
        <WebApplicationInfo>
            <Id>$application_GUID here$</Id>
            <Resource>api://localhost:44355/$application_GUID here$</Resource>
            <Scopes>
                <Scope>Files.Read.All</Scope>
                <Scope>profile</Scope>
            </Scopes>
        </WebApplicationInfo>
    </VersionOverrides>
...
</OfficeApp>
```
