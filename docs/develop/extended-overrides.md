---
title: 处理清单的扩展替代
description: 了解如何使用清单的扩展替代配置扩展性功能。
ms.date: 02/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 43e9820f54f2812130f7f86529c52b20b92811a0
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659948"
---
# <a name="work-with-extended-overrides-of-the-manifest"></a>使用清单的扩展替代

Office 外接程序的某些扩展性功能配置有托管在服务器上的 JSON 文件，而不是外接程序的 XML 清单。

> [!NOTE]
> 本文假定你熟悉 Office 加载项清单及其在加载项中的角色。如果最近没有，请阅读 [Office 加载项 XML 清单](add-in-manifests.md)。

下表指定了需要扩展替代的扩展性功能以及指向该功能文档的链接。

| 功能 | 开发说明 |
| :----- | :----- |
| 键盘快捷方式 | [将自定义键盘快捷方式添加到 Office 加载项](../design/keyboard-shortcuts.md) |

定义 JSON 格式的架构是 [扩展清单架构](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。

> [!TIP]
> 本文有些抽象。 请考虑阅读表中的一篇文章，以增加概念的清晰度。

## <a name="tell-office-where-to-find-the-json-file"></a>告诉 Office 在哪里可以找到 JSON 文件

使用清单告诉 Office 在哪里可以找到 JSON 文件。 紧 *接着* (不在清单中元素) **\<VersionOverrides\>** 内，添加 [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) 元素。 将属性 `Url` 设置为 JSON 文件的完整 URL。 下面是最 **\<ExtendedOverrides\>** 简单的元素的示例。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

下面是一个非常简单的扩展替代 JSON 文件的示例。 它将键盘快捷方式 CTRL+SHIFT+A 分配给 (在其他位置定义的函数，) 打开加载项的任务窗格。

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

## <a name="localize-the-extended-overrides-file"></a>本地化扩展重写文件

如果外接程序支持多个区域设置，则可以使用 `ResourceUrl` 元素的 **\<ExtendedOverrides\>** 属性将 Office 指向本地化资源的文件。 示例如下。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

有关如何创建和使用资源文件、如何在扩展替代文件中引用其资源以及此处未讨论的其他选项的更多详细信息，请参阅 [Localize 扩展替代](localization.md#localize-extended-overrides)。
