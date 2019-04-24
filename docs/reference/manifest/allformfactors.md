---
title: 清单文件中的 AllFormFactors 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8059501f88f966b285398ac7cf243e6b0e4e44ea
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450735"
---
# <a name="allformfactors-element"></a><span data-ttu-id="4b471-102">AllFormFactors 元素</span><span class="sxs-lookup"><span data-stu-id="4b471-102">AllFormFactors element</span></span>

<span data-ttu-id="4b471-103">指定加载项的所有外观设置。</span><span class="sxs-lookup"><span data-stu-id="4b471-103">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="4b471-104">目前，使用 **AllFormFactors** 的唯一功能是自定义函数。</span><span class="sxs-lookup"><span data-stu-id="4b471-104">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="4b471-105">使用自定义函数时，**AllFormFactors** 是必备元素。</span><span class="sxs-lookup"><span data-stu-id="4b471-105">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="4b471-106">子元素</span><span class="sxs-lookup"><span data-stu-id="4b471-106">Child elements</span></span>

|  <span data-ttu-id="4b471-107">元素</span><span class="sxs-lookup"><span data-stu-id="4b471-107">Element</span></span> |  <span data-ttu-id="4b471-108">必需</span><span class="sxs-lookup"><span data-stu-id="4b471-108">Required</span></span>  |  <span data-ttu-id="4b471-109">说明</span><span class="sxs-lookup"><span data-stu-id="4b471-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4b471-110">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="4b471-110">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="4b471-111">是</span><span class="sxs-lookup"><span data-stu-id="4b471-111">Yes</span></span> |  <span data-ttu-id="4b471-112">定义加载项用于公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="4b471-112">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="4b471-113">AllFormFactors 示例</span><span class="sxs-lookup"><span data-stu-id="4b471-113">AllFormFactors example</span></span>

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
