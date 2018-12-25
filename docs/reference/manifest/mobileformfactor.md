---
title: 清单文件中的 MobileFormFactor 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: f0a68c7127f7872207a58ed252def7a2977c33ed
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433690"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="21c54-102">MobileFormFactor 元素</span><span class="sxs-lookup"><span data-stu-id="21c54-102">MobileFormFactor element</span></span>

<span data-ttu-id="21c54-p101">指定对移动外形规格的外接程序的设置。它包含移动外形规格的所有外接程序信息（**资源**节点的信息除外）。</span><span class="sxs-lookup"><span data-stu-id="21c54-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="21c54-p102">每个 **MobileFormFactor** 定义均包含 **FunctionFile** 元素以及一个或多个 **ExtensionPoint** 元素。有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。</span><span class="sxs-lookup"><span data-stu-id="21c54-p102">Each **MobileFormFactor** definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="21c54-p103">在 VersionOverrides 架构 1.1 中定义了 **MobileFormFactor** 元素。包含  [VersionOverrides](versionoverrides.md) 元素的 `VersionOverridesV1_1` 属性值必须为 `xsi:type`。</span><span class="sxs-lookup"><span data-stu-id="21c54-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="21c54-109">子元素</span><span class="sxs-lookup"><span data-stu-id="21c54-109">Child elements</span></span>

| <span data-ttu-id="21c54-110">元素</span><span class="sxs-lookup"><span data-stu-id="21c54-110">Element</span></span>                               | <span data-ttu-id="21c54-111">必需</span><span class="sxs-lookup"><span data-stu-id="21c54-111">Required</span></span> | <span data-ttu-id="21c54-112">说明</span><span class="sxs-lookup"><span data-stu-id="21c54-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="21c54-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="21c54-113">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="21c54-114">是</span><span class="sxs-lookup"><span data-stu-id="21c54-114">Yes</span></span>      | <span data-ttu-id="21c54-115">定义外接程序公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="21c54-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="21c54-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="21c54-116">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="21c54-117">是</span><span class="sxs-lookup"><span data-stu-id="21c54-117">Yes</span></span>      | <span data-ttu-id="21c54-118">包含 JavaScript 函数的文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="21c54-118">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="21c54-119">MobileFormFactor 示例</span><span class="sxs-lookup"><span data-stu-id="21c54-119">MobileFormFactor example</span></span>

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
