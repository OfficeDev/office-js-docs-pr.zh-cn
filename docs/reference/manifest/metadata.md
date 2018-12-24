---
title: 清单文件中的 Metadata 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 79038fc13eba76176be19e484ffa57e64727bf94
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432660"
---
# <a name="metadata-element"></a><span data-ttu-id="d6f16-102">Metadata 元素</span><span class="sxs-lookup"><span data-stu-id="d6f16-102">MetaData element</span></span>

<span data-ttu-id="d6f16-103">定义 Excel 中的自定义函数所使用的元数据设置。</span><span class="sxs-lookup"><span data-stu-id="d6f16-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="d6f16-104">属性</span><span class="sxs-lookup"><span data-stu-id="d6f16-104">Attributes</span></span>

<span data-ttu-id="d6f16-105">无</span><span class="sxs-lookup"><span data-stu-id="d6f16-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="d6f16-106">子元素</span><span class="sxs-lookup"><span data-stu-id="d6f16-106">Child elements</span></span>

|  <span data-ttu-id="d6f16-107">元素</span><span class="sxs-lookup"><span data-stu-id="d6f16-107">Element</span></span>  |  <span data-ttu-id="d6f16-108">必需</span><span class="sxs-lookup"><span data-stu-id="d6f16-108">Required</span></span>  |  <span data-ttu-id="d6f16-109">说明</span><span class="sxs-lookup"><span data-stu-id="d6f16-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d6f16-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="d6f16-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="d6f16-111">是</span><span class="sxs-lookup"><span data-stu-id="d6f16-111">Yes</span></span>  | <span data-ttu-id="d6f16-112">包含自定义函数所使用的 JSON 文件的资源 ID 的字符串。</span><span class="sxs-lookup"><span data-stu-id="d6f16-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="d6f16-113">示例</span><span class="sxs-lookup"><span data-stu-id="d6f16-113">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
