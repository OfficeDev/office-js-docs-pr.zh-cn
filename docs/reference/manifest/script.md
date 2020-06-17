---
title: 清单文件中的 Script 元素
description: Script 元素定义自定义函数在 Excel 中使用的脚本设置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 791f49f15673a029b982e40946f8cc90f02ba887
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608088"
---
# <a name="script-element"></a><span data-ttu-id="063de-103">Script 元素</span><span class="sxs-lookup"><span data-stu-id="063de-103">Script element</span></span>

<span data-ttu-id="063de-104">定义 Excel 中的自定义函数所使用的脚本设置。</span><span class="sxs-lookup"><span data-stu-id="063de-104">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="063de-105">属性</span><span class="sxs-lookup"><span data-stu-id="063de-105">Attributes</span></span>

<span data-ttu-id="063de-106">无</span><span class="sxs-lookup"><span data-stu-id="063de-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="063de-107">子元素</span><span class="sxs-lookup"><span data-stu-id="063de-107">Child elements</span></span>

|<span data-ttu-id="063de-108">元素</span><span class="sxs-lookup"><span data-stu-id="063de-108">Elements</span></span>  |  <span data-ttu-id="063de-109">必需</span><span class="sxs-lookup"><span data-stu-id="063de-109">Required</span></span>  |  <span data-ttu-id="063de-110">Description</span><span class="sxs-lookup"><span data-stu-id="063de-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="063de-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="063de-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="063de-112">是</span><span class="sxs-lookup"><span data-stu-id="063de-112">Yes</span></span>  | <span data-ttu-id="063de-113">包含自定义函数所使用的 JavaScript 文件的资源 ID 的字符串。</span><span class="sxs-lookup"><span data-stu-id="063de-113">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="063de-114">示例</span><span class="sxs-lookup"><span data-stu-id="063de-114">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
