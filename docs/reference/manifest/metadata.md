---
title: 清单文件中的 Metadata 元素
description: Metadata 元素定义自定义函数在 Excel 中使用的元数据设置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8ea81818aa96b407ce386ec318495ec5ba773d05
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718067"
---
# <a name="metadata-element"></a><span data-ttu-id="c54ab-103">Metadata 元素</span><span class="sxs-lookup"><span data-stu-id="c54ab-103">Metadata element</span></span>

<span data-ttu-id="c54ab-104">定义 Excel 中的自定义函数所使用的元数据设置。</span><span class="sxs-lookup"><span data-stu-id="c54ab-104">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="c54ab-105">属性</span><span class="sxs-lookup"><span data-stu-id="c54ab-105">Attributes</span></span>

<span data-ttu-id="c54ab-106">无</span><span class="sxs-lookup"><span data-stu-id="c54ab-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="c54ab-107">子元素</span><span class="sxs-lookup"><span data-stu-id="c54ab-107">Child elements</span></span>

|  <span data-ttu-id="c54ab-108">元素</span><span class="sxs-lookup"><span data-stu-id="c54ab-108">Element</span></span>  |  <span data-ttu-id="c54ab-109">必需</span><span class="sxs-lookup"><span data-stu-id="c54ab-109">Required</span></span>  |  <span data-ttu-id="c54ab-110">说明</span><span class="sxs-lookup"><span data-stu-id="c54ab-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c54ab-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="c54ab-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="c54ab-112">是</span><span class="sxs-lookup"><span data-stu-id="c54ab-112">Yes</span></span>  | <span data-ttu-id="c54ab-113">包含自定义函数所使用的 JSON 文件的资源 ID 的字符串。</span><span class="sxs-lookup"><span data-stu-id="c54ab-113">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="c54ab-114">示例</span><span class="sxs-lookup"><span data-stu-id="c54ab-114">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
