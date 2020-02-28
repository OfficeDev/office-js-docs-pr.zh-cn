---
title: 清单文件中的 DesktopFormFactor 元素
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 2fe97d99ff5bdc9f23a5760824e241ee4dfb800f
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325274"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="daeeb-102">DesktopFormFactor 元素</span><span class="sxs-lookup"><span data-stu-id="daeeb-102">DesktopFormFactor element</span></span>

<span data-ttu-id="daeeb-103">指定对桌面外形规格的外接程序的设置。</span><span class="sxs-lookup"><span data-stu-id="daeeb-103">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="daeeb-104">桌面外形规格包括 web、Windows 和 Mac 上的 Office。</span><span class="sxs-lookup"><span data-stu-id="daeeb-104">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="daeeb-105">除了 "**资源**" 节点外，它还包含桌面外形规格的所有外接程序信息。</span><span class="sxs-lookup"><span data-stu-id="daeeb-105">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="daeeb-106">每个 DesktopFormFactor 定义都包含**FunctionFile**元素和一个或多个**ExtensionPoint**元素。</span><span class="sxs-lookup"><span data-stu-id="daeeb-106">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="daeeb-107">有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。</span><span class="sxs-lookup"><span data-stu-id="daeeb-107">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="daeeb-108">子元素</span><span class="sxs-lookup"><span data-stu-id="daeeb-108">Child elements</span></span>

| <span data-ttu-id="daeeb-109">元素</span><span class="sxs-lookup"><span data-stu-id="daeeb-109">Element</span></span>                               | <span data-ttu-id="daeeb-110">必需</span><span class="sxs-lookup"><span data-stu-id="daeeb-110">Required</span></span> | <span data-ttu-id="daeeb-111">说明</span><span class="sxs-lookup"><span data-stu-id="daeeb-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="daeeb-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="daeeb-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="daeeb-113">是</span><span class="sxs-lookup"><span data-stu-id="daeeb-113">Yes</span></span>      | <span data-ttu-id="daeeb-114">定义外接程序公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="daeeb-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="daeeb-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="daeeb-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="daeeb-116">是</span><span class="sxs-lookup"><span data-stu-id="daeeb-116">Yes</span></span>      | <span data-ttu-id="daeeb-117">包含 JavaScript 函数的文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="daeeb-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="daeeb-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="daeeb-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="daeeb-119">否</span><span class="sxs-lookup"><span data-stu-id="daeeb-119">No</span></span>       | <span data-ttu-id="daeeb-120">定义在 Word、Excel 或 PowerPoint 主机中安装外接程序时将显示的标注。</span><span class="sxs-lookup"><span data-stu-id="daeeb-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="daeeb-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="daeeb-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="daeeb-122">否</span><span class="sxs-lookup"><span data-stu-id="daeeb-122">No</span></span> | <span data-ttu-id="daeeb-123">定义 Outlook 外接程序在代理应用场景中是否可用，默认设置为 *false*。</span><span class="sxs-lookup"><span data-stu-id="daeeb-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="daeeb-124">DesktopFormFactor 示例</span><span class="sxs-lookup"><span data-stu-id="daeeb-124">DesktopFormFactor example</span></span>

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
