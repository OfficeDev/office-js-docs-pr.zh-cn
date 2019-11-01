---
title: 清单文件中的 DesktopFormFactor 元素
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: bada3cd4cff7973517aedb83235a224ef6c273eb
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901960"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="c2d44-102">DesktopFormFactor 元素</span><span class="sxs-lookup"><span data-stu-id="c2d44-102">DesktopFormFactor element</span></span>

<span data-ttu-id="c2d44-103">指定对桌面外形规格的外接程序的设置。</span><span class="sxs-lookup"><span data-stu-id="c2d44-103">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="c2d44-104">桌面外形规格包括 web、Windows 和 Mac 上的 Office。</span><span class="sxs-lookup"><span data-stu-id="c2d44-104">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="c2d44-105">它包含该外形规格的所有外接程序信息（**资源**节点的信息除外）。</span><span class="sxs-lookup"><span data-stu-id="c2d44-105">It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="c2d44-p102">每个 DesktopFormFactor 定义均包含 **FunctionFile** 元素和一个或多个 **ExtensionPoint** 元素。有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。</span><span class="sxs-lookup"><span data-stu-id="c2d44-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="c2d44-108">子元素</span><span class="sxs-lookup"><span data-stu-id="c2d44-108">Child elements</span></span>

| <span data-ttu-id="c2d44-109">元素</span><span class="sxs-lookup"><span data-stu-id="c2d44-109">Element</span></span>                               | <span data-ttu-id="c2d44-110">必需</span><span class="sxs-lookup"><span data-stu-id="c2d44-110">Required</span></span> | <span data-ttu-id="c2d44-111">说明</span><span class="sxs-lookup"><span data-stu-id="c2d44-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="c2d44-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="c2d44-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="c2d44-113">是</span><span class="sxs-lookup"><span data-stu-id="c2d44-113">Yes</span></span>      | <span data-ttu-id="c2d44-114">定义外接程序公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="c2d44-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="c2d44-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="c2d44-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="c2d44-116">是</span><span class="sxs-lookup"><span data-stu-id="c2d44-116">Yes</span></span>      | <span data-ttu-id="c2d44-117">包含 JavaScript 函数的文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="c2d44-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="c2d44-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="c2d44-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="c2d44-119">否</span><span class="sxs-lookup"><span data-stu-id="c2d44-119">No</span></span>       | <span data-ttu-id="c2d44-120">定义在 Word、Excel 或 PowerPoint 主机中安装外接程序时将显示的标注。</span><span class="sxs-lookup"><span data-stu-id="c2d44-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="c2d44-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="c2d44-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="c2d44-122">否</span><span class="sxs-lookup"><span data-stu-id="c2d44-122">No</span></span> | <span data-ttu-id="c2d44-123">定义 Outlook 外接程序在代理应用场景中是否可用，默认设置为 *false*。</span><span class="sxs-lookup"><span data-stu-id="c2d44-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="c2d44-124">DesktopFormFactor 示例</span><span class="sxs-lookup"><span data-stu-id="c2d44-124">DesktopFormFactor example</span></span>

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
