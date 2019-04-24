---
title: 清单文件中的 SupportsSharedFolders 元素
description: ''
ms.date: 04/02/2019
localization_priority: Normal
ms.openlocfilehash: 976f8ba00f6ac9ac32def56933af1077527b7e9c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452037"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="a5c3e-102">SupportsSharedFolders 元素</span><span class="sxs-lookup"><span data-stu-id="a5c3e-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="a5c3e-103">定义 Outlook 加载项在代理应用场景中是否可用。</span><span class="sxs-lookup"><span data-stu-id="a5c3e-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="a5c3e-104">**SupportsSharedFolders** 元素是 [DesktopFormFactor](desktopformfactor.md) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="a5c3e-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="a5c3e-105">默认情况下，此元素设置为 *false*。</span><span class="sxs-lookup"><span data-stu-id="a5c3e-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a5c3e-106">Outlook 外接程序的委派访问权限当前[处于预览阶段](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview), 仅在对 Exchange Online 运行的客户端中受支持。</span><span class="sxs-lookup"><span data-stu-id="a5c3e-106">Delegate access for Outlook add-ins is currently [in preview](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview) and only supported in clients that run against Exchange Online.</span></span> <span data-ttu-id="a5c3e-107">使用此元素的外接程序无法发布到 AppSource 或通过集中部署进行部署。</span><span class="sxs-lookup"><span data-stu-id="a5c3e-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="a5c3e-108">以下是 **SupportsSharedFolders** 元素的示例。</span><span class="sxs-lookup"><span data-stu-id="a5c3e-108">The following is an example of the  **SupportsSharedFolders** element.</span></span>

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
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```
