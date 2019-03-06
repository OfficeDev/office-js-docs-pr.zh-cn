---
title: 清单文件中的 DesktopFormFactor 元素
description: ''
ms.date: 03/01/2019
localization_priority: Normal
ms.openlocfilehash: cddf76af01ec9f3016b28a3f7692aa6dfeb9bd60
ms.sourcegitcommit: 7ebd383f16ae5809bb6980a5f213b695d410e62c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/06/2019
ms.locfileid: "30413620"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="05fbc-102">DesktopFormFactor 元素</span><span class="sxs-lookup"><span data-stu-id="05fbc-102">DesktopFormFactor element</span></span>

<span data-ttu-id="05fbc-p101">指定对桌面外形规格的外接程序的设置。桌面外形规格包括 Office for Windows、Office for Mac 和 Office Online。它包含该外形规格的所有外接程序信息（**资源**节点的信息除外）。</span><span class="sxs-lookup"><span data-stu-id="05fbc-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="05fbc-p102">每个 DesktopFormFactor 定义均包含 **FunctionFile** 元素和一个或多个 **ExtensionPoint** 元素。有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。</span><span class="sxs-lookup"><span data-stu-id="05fbc-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="05fbc-108">子元素</span><span class="sxs-lookup"><span data-stu-id="05fbc-108">Child elements</span></span>

| <span data-ttu-id="05fbc-109">元素</span><span class="sxs-lookup"><span data-stu-id="05fbc-109">Element</span></span>                               | <span data-ttu-id="05fbc-110">必需</span><span class="sxs-lookup"><span data-stu-id="05fbc-110">Required</span></span> | <span data-ttu-id="05fbc-111">说明</span><span class="sxs-lookup"><span data-stu-id="05fbc-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="05fbc-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="05fbc-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="05fbc-113">是</span><span class="sxs-lookup"><span data-stu-id="05fbc-113">Yes</span></span>      | <span data-ttu-id="05fbc-114">定义外接程序公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="05fbc-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="05fbc-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="05fbc-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="05fbc-116">是</span><span class="sxs-lookup"><span data-stu-id="05fbc-116">Yes</span></span>      | <span data-ttu-id="05fbc-117">包含 JavaScript 函数的文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="05fbc-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="05fbc-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="05fbc-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="05fbc-119">否</span><span class="sxs-lookup"><span data-stu-id="05fbc-119">No</span></span>       | <span data-ttu-id="05fbc-120">定义在 Word、Excel 或 PowerPoint 主机中安装外接程序时将显示的标注。</span><span class="sxs-lookup"><span data-stu-id="05fbc-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="05fbc-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="05fbc-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="05fbc-122">否</span><span class="sxs-lookup"><span data-stu-id="05fbc-122">No</span></span> | <span data-ttu-id="05fbc-123">定义 Outlook 外接程序在代理应用场景中是否可用，默认设置为 *false*。</span><span class="sxs-lookup"><span data-stu-id="05fbc-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span><br><br><span data-ttu-id="05fbc-124">**重要说明**: 由于 Outlook 外接程序的代理访问当前处于预览阶段, 使用`SupportSharedFolders`元素的外接程序不能发布到 AppSource 或通过集中部署进行部署。</span><span class="sxs-lookup"><span data-stu-id="05fbc-124">**Important**: Because delegate access for Outlook add-ins is currently in preview, add-ins that use the `SupportSharedFolders` element cannot be published to AppSource or deployed via centralized deployment.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="05fbc-125">DesktopFormFactor 示例</span><span class="sxs-lookup"><span data-stu-id="05fbc-125">DesktopFormFactor example</span></span>

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
