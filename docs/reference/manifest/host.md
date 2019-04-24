---
title: 清单文件中的 Host 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f496e3e0c16f24d20e1d1db76208e61267235131
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450504"
---
# <a name="host-element"></a><span data-ttu-id="e2b9a-102">Host 元素</span><span class="sxs-lookup"><span data-stu-id="e2b9a-102">Host element</span></span>

<span data-ttu-id="e2b9a-103">指定应在其中激活外接程序的单个 Office 应用程序类型。</span><span class="sxs-lookup"><span data-stu-id="e2b9a-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="e2b9a-104">**Host** 元素的语法根据该元素是否在[基本清单](#basic-manifest)中或 [VersionOverrides](#versionoverrides-node) 节点中定义而不同。</span><span class="sxs-lookup"><span data-stu-id="e2b9a-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="e2b9a-105">但功能相同。</span><span class="sxs-lookup"><span data-stu-id="e2b9a-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="e2b9a-106">基本清单</span><span class="sxs-lookup"><span data-stu-id="e2b9a-106">Basic manifest</span></span>

<span data-ttu-id="e2b9a-107">在基本清单（在 [OfficeApp](officeapp.md) 下）中定义时，主机类型由 `Name` 属性决定。</span><span class="sxs-lookup"><span data-stu-id="e2b9a-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="e2b9a-108">属性</span><span class="sxs-lookup"><span data-stu-id="e2b9a-108">Attributes</span></span>

| <span data-ttu-id="e2b9a-109">属性</span><span class="sxs-lookup"><span data-stu-id="e2b9a-109">Attribute</span></span>     | <span data-ttu-id="e2b9a-110">类型</span><span class="sxs-lookup"><span data-stu-id="e2b9a-110">Type</span></span>   | <span data-ttu-id="e2b9a-111">必需</span><span class="sxs-lookup"><span data-stu-id="e2b9a-111">Required</span></span> | <span data-ttu-id="e2b9a-112">说明</span><span class="sxs-lookup"><span data-stu-id="e2b9a-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="e2b9a-113">Name</span><span class="sxs-lookup"><span data-stu-id="e2b9a-113">Name</span></span>](#name) | <span data-ttu-id="e2b9a-114">string</span><span class="sxs-lookup"><span data-stu-id="e2b9a-114">string</span></span> | <span data-ttu-id="e2b9a-115">必需</span><span class="sxs-lookup"><span data-stu-id="e2b9a-115">required</span></span> | <span data-ttu-id="e2b9a-116">Office 主机应用程序的类型名称。</span><span class="sxs-lookup"><span data-stu-id="e2b9a-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="e2b9a-117">名称</span><span class="sxs-lookup"><span data-stu-id="e2b9a-117">Name</span></span>
<span data-ttu-id="e2b9a-p102">指定此外接程序面向的主机类型。值必须为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="e2b9a-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="e2b9a-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="e2b9a-120">`Document` (Word)</span></span>
- <span data-ttu-id="e2b9a-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="e2b9a-121">`Database` (Access)</span></span>
- <span data-ttu-id="e2b9a-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="e2b9a-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="e2b9a-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="e2b9a-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="e2b9a-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="e2b9a-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="e2b9a-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="e2b9a-125">`Project` (Project)</span></span>
- <span data-ttu-id="e2b9a-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="e2b9a-126">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="e2b9a-127">示例</span><span class="sxs-lookup"><span data-stu-id="e2b9a-127">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="e2b9a-128">VersionOverrides 节点</span><span class="sxs-lookup"><span data-stu-id="e2b9a-128">VersionOverrides node</span></span>
<span data-ttu-id="e2b9a-129">在 [VersionOverrides](versionoverrides.md) 中定义时，主机类型由 `xsi:type` 属性决定。</span><span class="sxs-lookup"><span data-stu-id="e2b9a-129">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="e2b9a-130">属性</span><span class="sxs-lookup"><span data-stu-id="e2b9a-130">Attributes</span></span>

|  <span data-ttu-id="e2b9a-131">属性</span><span class="sxs-lookup"><span data-stu-id="e2b9a-131">Attribute</span></span>  |  <span data-ttu-id="e2b9a-132">必需</span><span class="sxs-lookup"><span data-stu-id="e2b9a-132">Required</span></span>  |  <span data-ttu-id="e2b9a-133">说明</span><span class="sxs-lookup"><span data-stu-id="e2b9a-133">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e2b9a-134">xsi:type</span><span class="sxs-lookup"><span data-stu-id="e2b9a-134">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="e2b9a-135">是</span><span class="sxs-lookup"><span data-stu-id="e2b9a-135">Yes</span></span>  | <span data-ttu-id="e2b9a-136">描述这些设置适用的 Office 主机。</span><span class="sxs-lookup"><span data-stu-id="e2b9a-136">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="e2b9a-137">子元素</span><span class="sxs-lookup"><span data-stu-id="e2b9a-137">Child elements</span></span>

|  <span data-ttu-id="e2b9a-138">元素</span><span class="sxs-lookup"><span data-stu-id="e2b9a-138">Element</span></span> |  <span data-ttu-id="e2b9a-139">必需</span><span class="sxs-lookup"><span data-stu-id="e2b9a-139">Required</span></span>  |  <span data-ttu-id="e2b9a-140">说明</span><span class="sxs-lookup"><span data-stu-id="e2b9a-140">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e2b9a-141">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="e2b9a-141">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="e2b9a-142">是</span><span class="sxs-lookup"><span data-stu-id="e2b9a-142">Yes</span></span>   |  <span data-ttu-id="e2b9a-143">定义桌面外形规格的设置。</span><span class="sxs-lookup"><span data-stu-id="e2b9a-143">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="e2b9a-144">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="e2b9a-144">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="e2b9a-145">否</span><span class="sxs-lookup"><span data-stu-id="e2b9a-145">No</span></span>   |  <span data-ttu-id="e2b9a-p103">定义移动外形规格的设置。**注意：** 仅在 Outlook for iOS 中支持此元素。</span><span class="sxs-lookup"><span data-stu-id="e2b9a-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="e2b9a-148">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="e2b9a-148">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="e2b9a-149">否</span><span class="sxs-lookup"><span data-stu-id="e2b9a-149">No</span></span>   |  <span data-ttu-id="e2b9a-150">定义所有外形规格的设置。</span><span class="sxs-lookup"><span data-stu-id="e2b9a-150">Defines the settings for all form factors.</span></span> <span data-ttu-id="e2b9a-151">仅用于 Excel 中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="e2b9a-151">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="e2b9a-152">xsi:type</span><span class="sxs-lookup"><span data-stu-id="e2b9a-152">xsi:type</span></span>

<span data-ttu-id="e2b9a-153">控制所包含的设置适用的 Office 主机类别（Word、Excel、PowerPoint、Outlook 和 OneNote）。</span><span class="sxs-lookup"><span data-stu-id="e2b9a-153">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="e2b9a-154">值必须为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="e2b9a-154">The value must be one of the following:</span></span>

- <span data-ttu-id="e2b9a-155">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="e2b9a-155">`Document` (Word)</span></span>
- <span data-ttu-id="e2b9a-156">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="e2b9a-156">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="e2b9a-157">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="e2b9a-157">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="e2b9a-158">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="e2b9a-158">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="e2b9a-159">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="e2b9a-159">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="e2b9a-160">主机示例</span><span class="sxs-lookup"><span data-stu-id="e2b9a-160">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
