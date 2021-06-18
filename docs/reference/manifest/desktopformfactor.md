---
title: 清单文件中的 DesktopFormFactor 元素
description: 指定对桌面外形规格的外接程序的设置。
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: 66673d83fd8608a1ec10492d7a944b0515de61c0
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007787"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="c3568-103">DesktopFormFactor 元素</span><span class="sxs-lookup"><span data-stu-id="c3568-103">DesktopFormFactor element</span></span>

<span data-ttu-id="c3568-104">指定对桌面外形规格的外接程序的设置。</span><span class="sxs-lookup"><span data-stu-id="c3568-104">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="c3568-105">桌面设备包括 Office web 版、Windows 和 Mac。</span><span class="sxs-lookup"><span data-stu-id="c3568-105">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="c3568-106">它包含桌面设备类型的所有外接程序信息，"资源"节点 **除外** 。</span><span class="sxs-lookup"><span data-stu-id="c3568-106">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="c3568-107">每个 DesktopFormFactor 定义都包含 **FunctionFile** 元素和一个或多个 **ExtensionPoint** 元素。</span><span class="sxs-lookup"><span data-stu-id="c3568-107">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="c3568-108">有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。</span><span class="sxs-lookup"><span data-stu-id="c3568-108">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="c3568-109">子元素</span><span class="sxs-lookup"><span data-stu-id="c3568-109">Child elements</span></span>

| <span data-ttu-id="c3568-110">元素</span><span class="sxs-lookup"><span data-stu-id="c3568-110">Element</span></span>                               | <span data-ttu-id="c3568-111">必需</span><span class="sxs-lookup"><span data-stu-id="c3568-111">Required</span></span> | <span data-ttu-id="c3568-112">说明</span><span class="sxs-lookup"><span data-stu-id="c3568-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="c3568-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="c3568-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="c3568-114">是</span><span class="sxs-lookup"><span data-stu-id="c3568-114">Yes</span></span>      | <span data-ttu-id="c3568-115">定义外接程序公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="c3568-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="c3568-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="c3568-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="c3568-117">是</span><span class="sxs-lookup"><span data-stu-id="c3568-117">Yes</span></span>      | <span data-ttu-id="c3568-118">包含 JavaScript 函数的文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="c3568-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="c3568-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="c3568-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="c3568-120">否</span><span class="sxs-lookup"><span data-stu-id="c3568-120">No</span></span>       | <span data-ttu-id="c3568-121">定义在 Word、加载项或加载项中安装加载项时Excel标注PowerPoint。</span><span class="sxs-lookup"><span data-stu-id="c3568-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint.</span></span> |
| [<span data-ttu-id="c3568-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="c3568-122">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="c3568-123">否</span><span class="sxs-lookup"><span data-stu-id="c3568-123">No</span></span> | <span data-ttu-id="c3568-124">定义 Outlook 外接程序现在在预览版 (中是否可用) 以及共享文件夹 (即委派访问权限) 方案。</span><span class="sxs-lookup"><span data-stu-id="c3568-124">Defines whether the Outlook add-in is available in shared mailbox (now in preview) and shared folders (that is, delegate access) scenarios.</span></span> <span data-ttu-id="c3568-125">默认情况下设置为 *false。*</span><span class="sxs-lookup"><span data-stu-id="c3568-125">Set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="c3568-126">DesktopFormFactor 示例</span><span class="sxs-lookup"><span data-stu-id="c3568-126">DesktopFormFactor example</span></span>

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
