---
title: 清单文件中的 Script 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 95e4cbadc35302b4f76108e0ff2a51d31ca89aac
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433136"
---
# <a name="script-element"></a><span data-ttu-id="48bc6-102">Script 元素</span><span class="sxs-lookup"><span data-stu-id="48bc6-102">Script element</span></span>

<span data-ttu-id="48bc6-103">定义 Excel 中的自定义函数所使用的脚本设置。</span><span class="sxs-lookup"><span data-stu-id="48bc6-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="48bc6-104">属性</span><span class="sxs-lookup"><span data-stu-id="48bc6-104">Attributes</span></span>

<span data-ttu-id="48bc6-105">无</span><span class="sxs-lookup"><span data-stu-id="48bc6-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="48bc6-106">子元素</span><span class="sxs-lookup"><span data-stu-id="48bc6-106">Child elements</span></span>

|<span data-ttu-id="48bc6-107">元素</span><span class="sxs-lookup"><span data-stu-id="48bc6-107">Elements</span></span>  |  <span data-ttu-id="48bc6-108">必需</span><span class="sxs-lookup"><span data-stu-id="48bc6-108">Required</span></span>  |  <span data-ttu-id="48bc6-109">说明</span><span class="sxs-lookup"><span data-stu-id="48bc6-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="48bc6-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="48bc6-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="48bc6-111">是</span><span class="sxs-lookup"><span data-stu-id="48bc6-111">Yes</span></span>  | <span data-ttu-id="48bc6-112">包含自定义函数所使用的 JavaScript 文件的资源 ID 的字符串。</span><span class="sxs-lookup"><span data-stu-id="48bc6-112">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="48bc6-113">示例</span><span class="sxs-lookup"><span data-stu-id="48bc6-113">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
