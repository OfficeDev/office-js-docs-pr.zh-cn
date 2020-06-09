---
title: 清单文件中的 DesktopFormFactor 元素
description: 指定对桌面外形规格的外接程序的设置。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 46de234f2d97a9e6c7645c17a0f0a61d0c3e1a80
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612281"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="77fee-103">DesktopFormFactor 元素</span><span class="sxs-lookup"><span data-stu-id="77fee-103">DesktopFormFactor element</span></span>

<span data-ttu-id="77fee-104">指定对桌面外形规格的外接程序的设置。</span><span class="sxs-lookup"><span data-stu-id="77fee-104">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="77fee-105">桌面外形规格包括 web、Windows 和 Mac 上的 Office。</span><span class="sxs-lookup"><span data-stu-id="77fee-105">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="77fee-106">除了 "**资源**" 节点外，它还包含桌面外形规格的所有外接程序信息。</span><span class="sxs-lookup"><span data-stu-id="77fee-106">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="77fee-107">每个 DesktopFormFactor 定义都包含**FunctionFile**元素和一个或多个**ExtensionPoint**元素。</span><span class="sxs-lookup"><span data-stu-id="77fee-107">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="77fee-108">有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。</span><span class="sxs-lookup"><span data-stu-id="77fee-108">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="77fee-109">子元素</span><span class="sxs-lookup"><span data-stu-id="77fee-109">Child elements</span></span>

| <span data-ttu-id="77fee-110">元素</span><span class="sxs-lookup"><span data-stu-id="77fee-110">Element</span></span>                               | <span data-ttu-id="77fee-111">必需</span><span class="sxs-lookup"><span data-stu-id="77fee-111">Required</span></span> | <span data-ttu-id="77fee-112">Description</span><span class="sxs-lookup"><span data-stu-id="77fee-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="77fee-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="77fee-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="77fee-114">是</span><span class="sxs-lookup"><span data-stu-id="77fee-114">Yes</span></span>      | <span data-ttu-id="77fee-115">定义外接程序公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="77fee-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="77fee-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="77fee-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="77fee-117">是</span><span class="sxs-lookup"><span data-stu-id="77fee-117">Yes</span></span>      | <span data-ttu-id="77fee-118">包含 JavaScript 函数的文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="77fee-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="77fee-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="77fee-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="77fee-120">否</span><span class="sxs-lookup"><span data-stu-id="77fee-120">No</span></span>       | <span data-ttu-id="77fee-121">定义在 Word、Excel 或 PowerPoint 主机中安装外接程序时将显示的标注。</span><span class="sxs-lookup"><span data-stu-id="77fee-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="77fee-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="77fee-122">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="77fee-123">否</span><span class="sxs-lookup"><span data-stu-id="77fee-123">No</span></span> | <span data-ttu-id="77fee-124">定义 Outlook 外接程序在代理应用场景中是否可用，默认设置为 *false*。</span><span class="sxs-lookup"><span data-stu-id="77fee-124">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="77fee-125">DesktopFormFactor 示例</span><span class="sxs-lookup"><span data-stu-id="77fee-125">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- Information on this extension point. -->
      </ExtensionPoint>
      <!-- Possibly more ExtensionPoint elements. -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
