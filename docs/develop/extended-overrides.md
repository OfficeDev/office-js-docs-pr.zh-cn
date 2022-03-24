---
title: 使用清单的扩展替代
description: 了解如何使用清单的扩展替代来配置扩展性功能。
ms.date: 02/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 43a922f559100157dbdacbb401d38c4d9ba22010
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743792"
---
# <a name="work-with-extended-overrides-of-the-manifest"></a>使用清单的扩展替代

外接程序的一Office扩展性功能使用托管在服务器上的 JSON 文件（而不是外接程序的 XML 清单）进行配置。

> [!NOTE]
> 本文假定你熟悉Office清单及其在外接程序中的角色。如果您[Office，](add-in-manifests.md)请阅读外接程序 XML 清单。

下表指定了需要扩展替代的扩展性功能以及指向功能文档的链接。

| 功能 | 开发说明 |
| :----- | :----- |
| 键盘快捷方式 | [将自定义键盘快捷方式添加到Office加载项](../design/keyboard-shortcuts.md) |

定义 JSON 格式的架构是 [扩展清单架构](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。

> [!TIP]
> 本文有点抽象。 请考虑阅读表中的其中一篇文章，以明确概念。

## <a name="tell-office-where-to-find-the-json-file"></a>告知Office JSON 文件的位置

使用清单告知Office JSON 文件的位置。 在 *紧* (不在 `<VersionOverrides>`) 元素内，添加 [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 元素。 将 属性 `Url` 设置为 JSON 文件的完整 URL。 下面是最简单的可能元素的示例 `<ExtendedOverrides>` 。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

下面是一个非常简单的扩展覆盖 JSON 文件的示例。 它将键盘快捷方式 Ctrl+Shift+A 分配给 (打开加载项任务窗格) 定义的函数。

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+A"
            }
        }
    ]
}
```

## <a name="localize-the-extended-overrides-file"></a>本地化扩展替代文件

如果您的外接程序支持多个区域`ResourceUrl`设置，您可以使用 元素的 属性`<ExtendedOverrides>`将 Office指向本地化资源的文件。 示例如下。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

若要详细了解如何创建和使用资源文件、如何在扩展替代文件中引用资源，以及此处未讨论的其他选项，请参阅本地化 [扩展替代](localization.md#localize-extended-overrides)。
