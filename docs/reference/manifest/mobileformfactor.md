---
title: 清单文件中的 MobileFormFactor 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 34106011cb855b6ac7c6d0fc21c16fd13e52b281
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324839"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="c4396-102">MobileFormFactor 元素</span><span class="sxs-lookup"><span data-stu-id="c4396-102">MobileFormFactor element</span></span>

<span data-ttu-id="c4396-p101">指定对移动外形规格的外接程序的设置。它包含移动外形规格的所有外接程序信息（**资源**节点的信息除外）。</span><span class="sxs-lookup"><span data-stu-id="c4396-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="c4396-105">每个**MobileFormFactor**定义都包含**FunctionFile**元素和一个或多个**ExtensionPoint**元素。</span><span class="sxs-lookup"><span data-stu-id="c4396-105">Each **MobileFormFactor** definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="c4396-106">有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。</span><span class="sxs-lookup"><span data-stu-id="c4396-106">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="c4396-p103">在 VersionOverrides 架构 1.1 中定义了 **MobileFormFactor** 元素。包含  [VersionOverrides](versionoverrides.md) 元素的 `VersionOverridesV1_1` 属性值必须为 `xsi:type`。</span><span class="sxs-lookup"><span data-stu-id="c4396-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c4396-109">子元素</span><span class="sxs-lookup"><span data-stu-id="c4396-109">Child elements</span></span>

| <span data-ttu-id="c4396-110">元素</span><span class="sxs-lookup"><span data-stu-id="c4396-110">Element</span></span>                               | <span data-ttu-id="c4396-111">必需</span><span class="sxs-lookup"><span data-stu-id="c4396-111">Required</span></span> | <span data-ttu-id="c4396-112">说明</span><span class="sxs-lookup"><span data-stu-id="c4396-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="c4396-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="c4396-113">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="c4396-114">是</span><span class="sxs-lookup"><span data-stu-id="c4396-114">Yes</span></span>      | <span data-ttu-id="c4396-115">定义外接程序公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="c4396-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="c4396-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="c4396-116">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="c4396-117">是</span><span class="sxs-lookup"><span data-stu-id="c4396-117">Yes</span></span>      | <span data-ttu-id="c4396-118">包含 JavaScript 函数的文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="c4396-118">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="c4396-119">MobileFormFactor 示例</span><span class="sxs-lookup"><span data-stu-id="c4396-119">MobileFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
