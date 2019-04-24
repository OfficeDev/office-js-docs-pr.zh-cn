---
title: 清单文件中的 Metadata 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a3aecb1983905658f3a55fdb8bf0629a8d5ef474
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452044"
---
# <a name="metadata-element"></a><span data-ttu-id="a1cfb-102">Metadata 元素</span><span class="sxs-lookup"><span data-stu-id="a1cfb-102">Metadata element</span></span>

<span data-ttu-id="a1cfb-103">定义 Excel 中的自定义函数所使用的元数据设置。</span><span class="sxs-lookup"><span data-stu-id="a1cfb-103">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="a1cfb-104">属性</span><span class="sxs-lookup"><span data-stu-id="a1cfb-104">Attributes</span></span>

<span data-ttu-id="a1cfb-105">无</span><span class="sxs-lookup"><span data-stu-id="a1cfb-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="a1cfb-106">子元素</span><span class="sxs-lookup"><span data-stu-id="a1cfb-106">Child elements</span></span>

|  <span data-ttu-id="a1cfb-107">元素</span><span class="sxs-lookup"><span data-stu-id="a1cfb-107">Element</span></span>  |  <span data-ttu-id="a1cfb-108">必需</span><span class="sxs-lookup"><span data-stu-id="a1cfb-108">Required</span></span>  |  <span data-ttu-id="a1cfb-109">说明</span><span class="sxs-lookup"><span data-stu-id="a1cfb-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a1cfb-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="a1cfb-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="a1cfb-111">是</span><span class="sxs-lookup"><span data-stu-id="a1cfb-111">Yes</span></span>  | <span data-ttu-id="a1cfb-112">包含自定义函数所使用的 JSON 文件的资源 ID 的字符串。</span><span class="sxs-lookup"><span data-stu-id="a1cfb-112">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="a1cfb-113">示例</span><span class="sxs-lookup"><span data-stu-id="a1cfb-113">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
