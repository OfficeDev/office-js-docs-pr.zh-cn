---
title: 清单文件中的 AllFormFactors 元素
description: 指定加载项的所有外观设置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f1285f92b5eb89993e7fcfe79aab2325b86aca3d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720713"
---
# <a name="allformfactors-element"></a><span data-ttu-id="36dc9-103">AllFormFactors 元素</span><span class="sxs-lookup"><span data-stu-id="36dc9-103">AllFormFactors element</span></span>

<span data-ttu-id="36dc9-104">指定加载项的所有外观设置。</span><span class="sxs-lookup"><span data-stu-id="36dc9-104">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="36dc9-105">目前，使用 **AllFormFactors** 的唯一功能是自定义函数。</span><span class="sxs-lookup"><span data-stu-id="36dc9-105">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="36dc9-106">使用自定义函数时，**AllFormFactors** 是必备元素。</span><span class="sxs-lookup"><span data-stu-id="36dc9-106">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="36dc9-107">子元素</span><span class="sxs-lookup"><span data-stu-id="36dc9-107">Child elements</span></span>

|  <span data-ttu-id="36dc9-108">元素</span><span class="sxs-lookup"><span data-stu-id="36dc9-108">Element</span></span> |  <span data-ttu-id="36dc9-109">必需</span><span class="sxs-lookup"><span data-stu-id="36dc9-109">Required</span></span>  |  <span data-ttu-id="36dc9-110">说明</span><span class="sxs-lookup"><span data-stu-id="36dc9-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="36dc9-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="36dc9-111">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="36dc9-112">是</span><span class="sxs-lookup"><span data-stu-id="36dc9-112">Yes</span></span> |  <span data-ttu-id="36dc9-113">定义加载项用于公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="36dc9-113">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="36dc9-114">AllFormFactors 示例</span><span class="sxs-lookup"><span data-stu-id="36dc9-114">AllFormFactors example</span></span>

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
