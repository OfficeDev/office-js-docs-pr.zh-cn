---
title: 使用清单的扩展替代
description: 了解如何使用清单的扩展替代来配置扩展性功能。
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: 4eb8936e8a01b81a3883f848446d20ebf4ecf863
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505568"
---
# <a name="work-with-extended-overrides-of-the-manifest"></a>使用清单的扩展替代

Office 外接程序的一些扩展性功能使用托管在服务器上的 JSON 文件进行配置，而不是使用加载项的 XML 清单进行配置。

> [!NOTE]
> 本文假定你熟悉 Office 外接程序清单及其在外接程序中的角色。如果 [最近尚未阅读，](add-in-manifests.md)请阅读 Office 外接程序 XML 清单。

下表指定需要扩展覆盖的扩展性功能以及指向该功能文档的链接。

| 功能 | 开发说明 |
| :----- | :----- |
| 键盘快捷方式 | [将自定义键盘快捷方式添加到 Office 加载项](../design/keyboard-shortcuts.md) |

定义 JSON 格式的架构是 [扩展清单架构](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。

> [!TIP]
> 本文有点抽象。 请考虑阅读表格中的其中一篇文章，以明确概念。

## <a name="tell-office-where-to-find-the-json-file"></a>告知 Office 在哪里可以找到 JSON 文件

使用清单告诉 Office 在哪里可以找到 JSON 文件。 紧 *(* 清单) 元素的内部，添加 `<VersionOverrides>` [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 元素。 将 `Url` 属性设置为 JSON 文件的完整 URL。 下面是可能最简单的元素 `<ExtendedOverrides>` 的示例。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

下面是一个非常简单的扩展覆盖 JSON 文件的示例。 它将键盘快捷方式 Ctrl+Shift+A 分配给 (加载项任务窗格) 定义的函数。

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

如果加载项支持多个区域设置，可以使用元素的属性将 Office 指向 `ResourceUrl` `<ExtendedOverrides>` 本地化资源的文件。 示例如下。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

若要详细了解如何创建和使用资源文件、如何在扩展覆盖文件中引用资源，以及此处未讨论的其他选项，请参阅["本地化扩展覆盖"。](localization.md#localize-extended-overrides)
