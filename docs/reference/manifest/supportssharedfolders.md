---
title: 清单文件中的 SupportsSharedFolders 元素
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 42fa1cf74634b183994e633d728d3be66e1e83f0
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902240"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="32c46-102">SupportsSharedFolders 元素</span><span class="sxs-lookup"><span data-stu-id="32c46-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="32c46-103">定义 Outlook 加载项在代理应用场景中是否可用。</span><span class="sxs-lookup"><span data-stu-id="32c46-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="32c46-104">**SupportsSharedFolders** 元素是 [DesktopFormFactor](desktopformfactor.md) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="32c46-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="32c46-105">默认情况下，此元素设置为 *false*。</span><span class="sxs-lookup"><span data-stu-id="32c46-105">It is set to *false* by default.</span></span>

<span data-ttu-id="32c46-106">以下是 **SupportsSharedFolders** 元素的示例。</span><span class="sxs-lookup"><span data-stu-id="32c46-106">The following is an example of the  **SupportsSharedFolders** element.</span></span>

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```
