---
title: 清单文件中的 AllFormFactors 元素
description: 指定加载项的所有外观设置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 9dac322312c1dfd60f6deb4296413e12b55a6a49
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608794"
---
# <a name="allformfactors-element"></a><span data-ttu-id="d3f8f-103">AllFormFactors 元素</span><span class="sxs-lookup"><span data-stu-id="d3f8f-103">AllFormFactors element</span></span>

<span data-ttu-id="d3f8f-104">指定加载项的所有外观设置。</span><span class="sxs-lookup"><span data-stu-id="d3f8f-104">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="d3f8f-105">目前，使用 **AllFormFactors** 的唯一功能是自定义函数。</span><span class="sxs-lookup"><span data-stu-id="d3f8f-105">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="d3f8f-106">使用自定义函数时，**AllFormFactors** 是必备元素。</span><span class="sxs-lookup"><span data-stu-id="d3f8f-106">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="d3f8f-107">子元素</span><span class="sxs-lookup"><span data-stu-id="d3f8f-107">Child elements</span></span>

|  <span data-ttu-id="d3f8f-108">元素</span><span class="sxs-lookup"><span data-stu-id="d3f8f-108">Element</span></span> |  <span data-ttu-id="d3f8f-109">必需</span><span class="sxs-lookup"><span data-stu-id="d3f8f-109">Required</span></span>  |  <span data-ttu-id="d3f8f-110">Description</span><span class="sxs-lookup"><span data-stu-id="d3f8f-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d3f8f-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="d3f8f-111">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="d3f8f-112">是</span><span class="sxs-lookup"><span data-stu-id="d3f8f-112">Yes</span></span> |  <span data-ttu-id="d3f8f-113">定义加载项用于公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="d3f8f-113">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="d3f8f-114">AllFormFactors 示例</span><span class="sxs-lookup"><span data-stu-id="d3f8f-114">AllFormFactors example</span></span>

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
