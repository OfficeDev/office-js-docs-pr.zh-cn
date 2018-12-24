---
title: 清单文件中的 AllFormFactors 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: de7fcdce48e175d15ca6268f24082e37b2085b05
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433276"
---
# <a name="allformfactors-element"></a><span data-ttu-id="57510-102">AllFormFactors 元素</span><span class="sxs-lookup"><span data-stu-id="57510-102">AllFormFactors element</span></span>

<span data-ttu-id="57510-103">指定加载项的所有外观设置。</span><span class="sxs-lookup"><span data-stu-id="57510-103">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="57510-104">目前，使用 **AllFormFactors** 的唯一功能是自定义函数。</span><span class="sxs-lookup"><span data-stu-id="57510-104">Currently, the only feature using AllFormFactors is custom functions.</span></span> <span data-ttu-id="57510-105">使用自定义函数时，**AllFormFactors** 是必备元素。</span><span class="sxs-lookup"><span data-stu-id="57510-105">AllFormFactors is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="57510-106">子元素</span><span class="sxs-lookup"><span data-stu-id="57510-106">Child elements</span></span>

|  <span data-ttu-id="57510-107">元素</span><span class="sxs-lookup"><span data-stu-id="57510-107">Element</span></span> |  <span data-ttu-id="57510-108">必需</span><span class="sxs-lookup"><span data-stu-id="57510-108">Required</span></span>  |  <span data-ttu-id="57510-109">说明</span><span class="sxs-lookup"><span data-stu-id="57510-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="57510-110">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="57510-110">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="57510-111">是</span><span class="sxs-lookup"><span data-stu-id="57510-111">Yes</span></span> |  <span data-ttu-id="57510-112">定义加载项用于公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="57510-112">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="57510-113">AllFormFactors 示例</span><span class="sxs-lookup"><span data-stu-id="57510-113">AllFormFactors example</span></span>

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
