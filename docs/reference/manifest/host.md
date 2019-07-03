---
title: 清单文件中的 Host 元素
description: ''
ms.date: 07/01/2019
localization_priority: Normal
ms.openlocfilehash: e7b557034f70b03ed57598b7ffb9f43878db7392
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454893"
---
# <a name="host-element"></a><span data-ttu-id="9bbd9-102">Host 元素</span><span class="sxs-lookup"><span data-stu-id="9bbd9-102">Host element</span></span>

<span data-ttu-id="9bbd9-103">指定应在其中激活外接程序的单个 Office 应用程序类型。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="9bbd9-104">**Host** 元素的语法根据该元素是否在[基本清单](#basic-manifest)中或 [VersionOverrides](#versionoverrides-node) 节点中定义而不同。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="9bbd9-105">但功能相同。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="9bbd9-106">基本清单</span><span class="sxs-lookup"><span data-stu-id="9bbd9-106">Basic manifest</span></span>

<span data-ttu-id="9bbd9-107">在基本清单（在 [OfficeApp](officeapp.md) 下）中定义时，主机类型由 `Name` 属性决定。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="9bbd9-108">属性</span><span class="sxs-lookup"><span data-stu-id="9bbd9-108">Attributes</span></span>

| <span data-ttu-id="9bbd9-109">属性</span><span class="sxs-lookup"><span data-stu-id="9bbd9-109">Attribute</span></span>     | <span data-ttu-id="9bbd9-110">类型</span><span class="sxs-lookup"><span data-stu-id="9bbd9-110">Type</span></span>   | <span data-ttu-id="9bbd9-111">必需</span><span class="sxs-lookup"><span data-stu-id="9bbd9-111">Required</span></span> | <span data-ttu-id="9bbd9-112">说明</span><span class="sxs-lookup"><span data-stu-id="9bbd9-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="9bbd9-113">Name</span><span class="sxs-lookup"><span data-stu-id="9bbd9-113">Name</span></span>](#name) | <span data-ttu-id="9bbd9-114">string</span><span class="sxs-lookup"><span data-stu-id="9bbd9-114">string</span></span> | <span data-ttu-id="9bbd9-115">必需</span><span class="sxs-lookup"><span data-stu-id="9bbd9-115">required</span></span> | <span data-ttu-id="9bbd9-116">Office 主机应用程序的类型名称。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="9bbd9-117">名称</span><span class="sxs-lookup"><span data-stu-id="9bbd9-117">Name</span></span>

<span data-ttu-id="9bbd9-118">指定此外接程序面向的主机类型。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-118">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="9bbd9-119">值必须是下列值之一。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-119">The value must be one of the following.</span></span>

- <span data-ttu-id="9bbd9-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="9bbd9-120">`Document` (Word)</span></span>
- <span data-ttu-id="9bbd9-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="9bbd9-121">`Database` (Access)</span></span>
- <span data-ttu-id="9bbd9-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="9bbd9-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="9bbd9-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="9bbd9-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="9bbd9-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="9bbd9-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="9bbd9-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="9bbd9-125">`Project` (Project)</span></span>
- <span data-ttu-id="9bbd9-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="9bbd9-126">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9bbd9-127">我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-127">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="9bbd9-128">作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-128">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="9bbd9-129">示例</span><span class="sxs-lookup"><span data-stu-id="9bbd9-129">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="9bbd9-130">VersionOverrides 节点</span><span class="sxs-lookup"><span data-stu-id="9bbd9-130">VersionOverrides node</span></span>

<span data-ttu-id="9bbd9-131">在 [VersionOverrides](versionoverrides.md) 中定义时，主机类型由 `xsi:type` 属性决定。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-131">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="9bbd9-132">属性</span><span class="sxs-lookup"><span data-stu-id="9bbd9-132">Attributes</span></span>

|  <span data-ttu-id="9bbd9-133">属性</span><span class="sxs-lookup"><span data-stu-id="9bbd9-133">Attribute</span></span>  |  <span data-ttu-id="9bbd9-134">必需</span><span class="sxs-lookup"><span data-stu-id="9bbd9-134">Required</span></span>  |  <span data-ttu-id="9bbd9-135">说明</span><span class="sxs-lookup"><span data-stu-id="9bbd9-135">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9bbd9-136">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9bbd9-136">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="9bbd9-137">是</span><span class="sxs-lookup"><span data-stu-id="9bbd9-137">Yes</span></span>  | <span data-ttu-id="9bbd9-138">描述这些设置适用的 Office 主机。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-138">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="9bbd9-139">子元素</span><span class="sxs-lookup"><span data-stu-id="9bbd9-139">Child elements</span></span>

|  <span data-ttu-id="9bbd9-140">元素</span><span class="sxs-lookup"><span data-stu-id="9bbd9-140">Element</span></span> |  <span data-ttu-id="9bbd9-141">必需</span><span class="sxs-lookup"><span data-stu-id="9bbd9-141">Required</span></span>  |  <span data-ttu-id="9bbd9-142">说明</span><span class="sxs-lookup"><span data-stu-id="9bbd9-142">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9bbd9-143">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="9bbd9-143">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="9bbd9-144">是</span><span class="sxs-lookup"><span data-stu-id="9bbd9-144">Yes</span></span>   |  <span data-ttu-id="9bbd9-145">定义桌面外形规格的设置。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-145">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="9bbd9-146">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="9bbd9-146">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="9bbd9-147">否</span><span class="sxs-lookup"><span data-stu-id="9bbd9-147">No</span></span>   |  <span data-ttu-id="9bbd9-148">定义移动设备规格的设置。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-148">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="9bbd9-149">**注意:** 仅在 iOS 上的 Outlook 中支持此元素。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-149">**Note:** This element is only supported in Outlook on iOS.</span></span> |
|  [<span data-ttu-id="9bbd9-150">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="9bbd9-150">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="9bbd9-151">否</span><span class="sxs-lookup"><span data-stu-id="9bbd9-151">No</span></span>   |  <span data-ttu-id="9bbd9-152">定义所有外形规格的设置。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-152">Defines the settings for all form factors.</span></span> <span data-ttu-id="9bbd9-153">仅用于 Excel 中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-153">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="9bbd9-154">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9bbd9-154">xsi:type</span></span>

<span data-ttu-id="9bbd9-155">控制所包含的设置适用的 Office 主机类别（Word、Excel、PowerPoint、Outlook 和 OneNote）。</span><span class="sxs-lookup"><span data-stu-id="9bbd9-155">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="9bbd9-156">值必须为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="9bbd9-156">The value must be one of the following:</span></span>

- <span data-ttu-id="9bbd9-157">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="9bbd9-157">`Document` (Word)</span></span>
- <span data-ttu-id="9bbd9-158">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="9bbd9-158">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="9bbd9-159">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="9bbd9-159">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="9bbd9-160">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="9bbd9-160">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="9bbd9-161">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="9bbd9-161">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="9bbd9-162">主机示例</span><span class="sxs-lookup"><span data-stu-id="9bbd9-162">Host example</span></span> 

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
