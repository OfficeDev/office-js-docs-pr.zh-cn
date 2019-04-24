---
title: 清单文件中的 Script 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8352ada0eeb6af071d5f20f750dcdeaefe31e918
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450434"
---
# <a name="script-element"></a><span data-ttu-id="cdc05-102">Script 元素</span><span class="sxs-lookup"><span data-stu-id="cdc05-102">Script element</span></span>

<span data-ttu-id="cdc05-103">定义 Excel 中的自定义函数所使用的脚本设置。</span><span class="sxs-lookup"><span data-stu-id="cdc05-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="cdc05-104">属性</span><span class="sxs-lookup"><span data-stu-id="cdc05-104">Attributes</span></span>

<span data-ttu-id="cdc05-105">无</span><span class="sxs-lookup"><span data-stu-id="cdc05-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="cdc05-106">子元素</span><span class="sxs-lookup"><span data-stu-id="cdc05-106">Child elements</span></span>

|<span data-ttu-id="cdc05-107">元素</span><span class="sxs-lookup"><span data-stu-id="cdc05-107">Elements</span></span>  |  <span data-ttu-id="cdc05-108">必需</span><span class="sxs-lookup"><span data-stu-id="cdc05-108">Required</span></span>  |  <span data-ttu-id="cdc05-109">说明</span><span class="sxs-lookup"><span data-stu-id="cdc05-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cdc05-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="cdc05-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="cdc05-111">是</span><span class="sxs-lookup"><span data-stu-id="cdc05-111">Yes</span></span>  | <span data-ttu-id="cdc05-112">包含自定义函数所使用的 JavaScript 文件的资源 ID 的字符串。</span><span class="sxs-lookup"><span data-stu-id="cdc05-112">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="cdc05-113">示例</span><span class="sxs-lookup"><span data-stu-id="cdc05-113">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
