---
title: 清单文件中的 DesktopFormFactor 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: dea632f7f8afa5d9b69f257798022e9e520e9394
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433738"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="89aec-102">DesktopFormFactor 元素</span><span class="sxs-lookup"><span data-stu-id="89aec-102">DesktopFormFactor element</span></span>

<span data-ttu-id="89aec-p101">指定对桌面外形规格的外接程序的设置。桌面外形规格包括 Office for Windows、Office for Mac 和 Office Online。它包含该外形规格的所有外接程序信息（**资源**节点的信息除外）。</span><span class="sxs-lookup"><span data-stu-id="89aec-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="89aec-p102">每个 DesktopFormFactor 定义均包含 **FunctionFile** 元素和一个或多个 **ExtensionPoint** 元素。有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。</span><span class="sxs-lookup"><span data-stu-id="89aec-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="89aec-108">子元素</span><span class="sxs-lookup"><span data-stu-id="89aec-108">Child elements</span></span>

| <span data-ttu-id="89aec-109">元素</span><span class="sxs-lookup"><span data-stu-id="89aec-109">Element</span></span>                               | <span data-ttu-id="89aec-110">必需</span><span class="sxs-lookup"><span data-stu-id="89aec-110">Required</span></span> | <span data-ttu-id="89aec-111">说明</span><span class="sxs-lookup"><span data-stu-id="89aec-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="89aec-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="89aec-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="89aec-113">是</span><span class="sxs-lookup"><span data-stu-id="89aec-113">Yes</span></span>      | <span data-ttu-id="89aec-114">定义外接程序公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="89aec-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="89aec-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="89aec-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="89aec-116">是</span><span class="sxs-lookup"><span data-stu-id="89aec-116">Yes</span></span>      | <span data-ttu-id="89aec-117">包含 JavaScript 函数的文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="89aec-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="89aec-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="89aec-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="89aec-119">否</span><span class="sxs-lookup"><span data-stu-id="89aec-119">No</span></span>       | <span data-ttu-id="89aec-120">定义在 Word、Excel 或 PowerPoint 主机中安装外接程序时将显示的标注。</span><span class="sxs-lookup"><span data-stu-id="89aec-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="89aec-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="89aec-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="89aec-122">否</span><span class="sxs-lookup"><span data-stu-id="89aec-122">No</span></span> | <span data-ttu-id="89aec-123">定义 Outlook 外接程序在代理应用场景中是否可用，默认设置为 *false*。</span><span class="sxs-lookup"><span data-stu-id="89aec-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span><br><br><span data-ttu-id="89aec-124">**重要说明**：此元素仅适用于针对 Exchange Online 的 Outlook 外接程序预览要求集。</span><span class="sxs-lookup"><span data-stu-id="89aec-124">**Important**: This element is only available in the Outlook add-ins Preview requirement set against Exchange Online.</span></span> <span data-ttu-id="89aec-125">使用此元素的外接程序无法发布到 AppSource 或通过集中部署进行部署。</span><span class="sxs-lookup"><span data-stu-id="89aec-125">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="89aec-126">DesktopFormFactor 示例</span><span class="sxs-lookup"><span data-stu-id="89aec-126">DesktopFormFactor example</span></span>

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
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
