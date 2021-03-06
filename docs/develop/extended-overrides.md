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
# <a name="work-with-extended-overrides-of-the-manifest"></a><span data-ttu-id="1f390-103">使用清单的扩展替代</span><span class="sxs-lookup"><span data-stu-id="1f390-103">Work with Extended Overrides of the manifest</span></span>

<span data-ttu-id="1f390-104">Office 外接程序的一些扩展性功能使用托管在服务器上的 JSON 文件进行配置，而不是使用加载项的 XML 清单进行配置。</span><span class="sxs-lookup"><span data-stu-id="1f390-104">Some extensibility features of Office Add-ins are configured with JSON files that are hosted on your server, instead of with the add-in's XML manifest.</span></span>

> [!NOTE]
> <span data-ttu-id="1f390-105">本文假定你熟悉 Office 外接程序清单及其在外接程序中的角色。如果 [最近尚未阅读，](add-in-manifests.md)请阅读 Office 外接程序 XML 清单。</span><span class="sxs-lookup"><span data-stu-id="1f390-105">This article assumes that you're familiar with Office add-in manifests and their role in add-ins. Please read [Office Add-ins XML manifest](add-in-manifests.md), if you haven't recently.</span></span>

<span data-ttu-id="1f390-106">下表指定需要扩展覆盖的扩展性功能以及指向该功能文档的链接。</span><span class="sxs-lookup"><span data-stu-id="1f390-106">The following table specifies the extensibility features that require an extended override along with links to documentation of the feature.</span></span>

| <span data-ttu-id="1f390-107">功能</span><span class="sxs-lookup"><span data-stu-id="1f390-107">Feature</span></span> | <span data-ttu-id="1f390-108">开发说明</span><span class="sxs-lookup"><span data-stu-id="1f390-108">Development Instructions</span></span> |
| :----- | :----- |
| <span data-ttu-id="1f390-109">键盘快捷方式</span><span class="sxs-lookup"><span data-stu-id="1f390-109">Keyboard shortcuts</span></span> | [<span data-ttu-id="1f390-110">将自定义键盘快捷方式添加到 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="1f390-110">Add Custom keyboard shortcuts to your Office Add-ins</span></span>](../design/keyboard-shortcuts.md) |

<span data-ttu-id="1f390-111">定义 JSON 格式的架构是 [扩展清单架构](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="1f390-111">The schema that defines the JSON format is [extended-manifest schema](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!TIP]
> <span data-ttu-id="1f390-112">本文有点抽象。</span><span class="sxs-lookup"><span data-stu-id="1f390-112">This article is somewhat abstract.</span></span> <span data-ttu-id="1f390-113">请考虑阅读表格中的其中一篇文章，以明确概念。</span><span class="sxs-lookup"><span data-stu-id="1f390-113">Consider reading one of the articles in the table to add clarity to the concepts.</span></span>

## <a name="tell-office-where-to-find-the-json-file"></a><span data-ttu-id="1f390-114">告知 Office 在哪里可以找到 JSON 文件</span><span class="sxs-lookup"><span data-stu-id="1f390-114">Tell Office where to find the JSON file</span></span>

<span data-ttu-id="1f390-115">使用清单告诉 Office 在哪里可以找到 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="1f390-115">Use the manifest to tell Office where to find the JSON file.</span></span> <span data-ttu-id="1f390-116">紧 *(* 清单) 元素的内部，添加 `<VersionOverrides>` [ExtendedOverrides](../reference/manifest/extendedoverrides.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="1f390-116">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="1f390-117">将 `Url` 属性设置为 JSON 文件的完整 URL。</span><span class="sxs-lookup"><span data-stu-id="1f390-117">Set the `Url` attribute to the full URL of a JSON file.</span></span> <span data-ttu-id="1f390-118">下面是可能最简单的元素 `<ExtendedOverrides>` 的示例。</span><span class="sxs-lookup"><span data-stu-id="1f390-118">The following is an example of the simplest possible `<ExtendedOverrides>` element.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="1f390-119">下面是一个非常简单的扩展覆盖 JSON 文件的示例。</span><span class="sxs-lookup"><span data-stu-id="1f390-119">The following is an example of a very simple extended overrides JSON file.</span></span> <span data-ttu-id="1f390-120">它将键盘快捷方式 Ctrl+Shift+A 分配给 (加载项任务窗格) 定义的函数。</span><span class="sxs-lookup"><span data-stu-id="1f390-120">It assigns keyboard shortcut CTRL+SHIFT+A to a function (defined elsewhere) that opens the add-in's task pane.</span></span>

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

## <a name="localize-the-extended-overrides-file"></a><span data-ttu-id="1f390-121">本地化扩展替代文件</span><span class="sxs-lookup"><span data-stu-id="1f390-121">Localize the extended overrides file</span></span>

<span data-ttu-id="1f390-122">如果加载项支持多个区域设置，可以使用元素的属性将 Office 指向 `ResourceUrl` `<ExtendedOverrides>` 本地化资源的文件。</span><span class="sxs-lookup"><span data-stu-id="1f390-122">If your add-in supports multiple locales, you can use the `ResourceUrl` attribute of the `<ExtendedOverrides>` element to point Office to a file of localized resources.</span></span> <span data-ttu-id="1f390-123">示例如下。</span><span class="sxs-lookup"><span data-stu-id="1f390-123">The following is an example.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="1f390-124">若要详细了解如何创建和使用资源文件、如何在扩展覆盖文件中引用资源，以及此处未讨论的其他选项，请参阅["本地化扩展覆盖"。](localization.md#localize-extended-overrides)</span><span class="sxs-lookup"><span data-stu-id="1f390-124">For more details about how to create and use the resources file, how to refer to its resources in the extended overrides file, and for additional options not discussed here, see [Localize extended overrides](localization.md#localize-extended-overrides).</span></span>
