---
title: 清单文件中的 MobileFormFactor 元素
description: MobileFormFactor 元素指定外接程序的移动外观设置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 5e52e66a2b97a32a19d42a4938dbeaed8f367478
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641471"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="792ed-103">MobileFormFactor 元素</span><span class="sxs-lookup"><span data-stu-id="792ed-103">MobileFormFactor element</span></span>

<span data-ttu-id="792ed-p101">指定对移动外形规格的外接程序的设置。它包含移动外形规格的所有外接程序信息（**资源**节点的信息除外）。</span><span class="sxs-lookup"><span data-stu-id="792ed-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="792ed-106">每个**MobileFormFactor**定义都包含**FunctionFile**元素和一个或多个**ExtensionPoint**元素。</span><span class="sxs-lookup"><span data-stu-id="792ed-106">Each **MobileFormFactor** definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="792ed-107">有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。</span><span class="sxs-lookup"><span data-stu-id="792ed-107">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="792ed-p103">在 VersionOverrides 架构 1.1 中定义了 **MobileFormFactor** 元素。包含  [VersionOverrides](versionoverrides.md) 元素的 `VersionOverridesV1_1` 属性值必须为 `xsi:type`。</span><span class="sxs-lookup"><span data-stu-id="792ed-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="792ed-110">子元素</span><span class="sxs-lookup"><span data-stu-id="792ed-110">Child elements</span></span>

| <span data-ttu-id="792ed-111">元素</span><span class="sxs-lookup"><span data-stu-id="792ed-111">Element</span></span>                             | <span data-ttu-id="792ed-112">必需</span><span class="sxs-lookup"><span data-stu-id="792ed-112">Required</span></span> | <span data-ttu-id="792ed-113">说明</span><span class="sxs-lookup"><span data-stu-id="792ed-113">Description</span></span>  |
|:------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="792ed-114">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="792ed-114">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="792ed-115">是</span><span class="sxs-lookup"><span data-stu-id="792ed-115">Yes</span></span>      | <span data-ttu-id="792ed-116">定义外接程序公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="792ed-116">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="792ed-117">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="792ed-117">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="792ed-118">是</span><span class="sxs-lookup"><span data-stu-id="792ed-118">Yes</span></span>      | <span data-ttu-id="792ed-119">包含 JavaScript 函数的文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="792ed-119">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="792ed-120">MobileFormFactor 示例</span><span class="sxs-lookup"><span data-stu-id="792ed-120">MobileFormFactor example</span></span>

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
