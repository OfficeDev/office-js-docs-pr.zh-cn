---
title: 清单文件中的 Host 元素
description: 指定应在其中激活外接程序的单个 Office 应用程序类型。
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: b9f03e6d6b028ca6f4616ae81b8fd76601256793
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718130"
---
# <a name="host-element"></a><span data-ttu-id="cf165-103">Host 元素</span><span class="sxs-lookup"><span data-stu-id="cf165-103">Host element</span></span>

<span data-ttu-id="cf165-104">指定应在其中激活外接程序的单个 Office 应用程序类型。</span><span class="sxs-lookup"><span data-stu-id="cf165-104">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cf165-105">**Host** 元素的语法根据该元素是否在[基本清单](#basic-manifest)中或 [VersionOverrides](#versionoverrides-node) 节点中定义而不同。</span><span class="sxs-lookup"><span data-stu-id="cf165-105">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="cf165-106">但功能相同。</span><span class="sxs-lookup"><span data-stu-id="cf165-106">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="cf165-107">基本清单</span><span class="sxs-lookup"><span data-stu-id="cf165-107">Basic manifest</span></span>

<span data-ttu-id="cf165-108">在基本清单（在 [OfficeApp](officeapp.md) 下）中定义时，主机类型由 `Name` 属性决定。</span><span class="sxs-lookup"><span data-stu-id="cf165-108">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="cf165-109">属性</span><span class="sxs-lookup"><span data-stu-id="cf165-109">Attributes</span></span>

| <span data-ttu-id="cf165-110">属性</span><span class="sxs-lookup"><span data-stu-id="cf165-110">Attribute</span></span>     | <span data-ttu-id="cf165-111">类型</span><span class="sxs-lookup"><span data-stu-id="cf165-111">Type</span></span>   | <span data-ttu-id="cf165-112">必需</span><span class="sxs-lookup"><span data-stu-id="cf165-112">Required</span></span> | <span data-ttu-id="cf165-113">说明</span><span class="sxs-lookup"><span data-stu-id="cf165-113">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="cf165-114">Name</span><span class="sxs-lookup"><span data-stu-id="cf165-114">Name</span></span>](#name) | <span data-ttu-id="cf165-115">string</span><span class="sxs-lookup"><span data-stu-id="cf165-115">string</span></span> | <span data-ttu-id="cf165-116">必需</span><span class="sxs-lookup"><span data-stu-id="cf165-116">required</span></span> | <span data-ttu-id="cf165-117">Office 主机应用程序的类型名称。</span><span class="sxs-lookup"><span data-stu-id="cf165-117">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="cf165-118">名称</span><span class="sxs-lookup"><span data-stu-id="cf165-118">Name</span></span>

<span data-ttu-id="cf165-119">指定此外接程序面向的主机类型。</span><span class="sxs-lookup"><span data-stu-id="cf165-119">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="cf165-120">值必须是下列值之一。</span><span class="sxs-lookup"><span data-stu-id="cf165-120">The value must be one of the following.</span></span>

- <span data-ttu-id="cf165-121">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="cf165-121">`Document` (Word)</span></span>
- <span data-ttu-id="cf165-122">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="cf165-122">`Database` (Access)</span></span>
- <span data-ttu-id="cf165-123">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="cf165-123">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="cf165-124">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="cf165-124">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="cf165-125">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="cf165-125">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="cf165-126">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="cf165-126">`Project` (Project)</span></span>
- <span data-ttu-id="cf165-127">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="cf165-127">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cf165-128">我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。</span><span class="sxs-lookup"><span data-stu-id="cf165-128">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="cf165-129">作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。</span><span class="sxs-lookup"><span data-stu-id="cf165-129">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="cf165-130">示例</span><span class="sxs-lookup"><span data-stu-id="cf165-130">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="cf165-131">VersionOverrides 节点</span><span class="sxs-lookup"><span data-stu-id="cf165-131">VersionOverrides node</span></span>

<span data-ttu-id="cf165-132">在 [VersionOverrides](versionoverrides.md) 中定义时，主机类型由 `xsi:type` 属性决定。</span><span class="sxs-lookup"><span data-stu-id="cf165-132">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="cf165-133">属性</span><span class="sxs-lookup"><span data-stu-id="cf165-133">Attributes</span></span>

|  <span data-ttu-id="cf165-134">属性</span><span class="sxs-lookup"><span data-stu-id="cf165-134">Attribute</span></span>  |  <span data-ttu-id="cf165-135">必需</span><span class="sxs-lookup"><span data-stu-id="cf165-135">Required</span></span>  |  <span data-ttu-id="cf165-136">说明</span><span class="sxs-lookup"><span data-stu-id="cf165-136">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cf165-137">xsi:type</span><span class="sxs-lookup"><span data-stu-id="cf165-137">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="cf165-138">是</span><span class="sxs-lookup"><span data-stu-id="cf165-138">Yes</span></span>  | <span data-ttu-id="cf165-139">描述这些设置适用的 Office 主机。</span><span class="sxs-lookup"><span data-stu-id="cf165-139">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="cf165-140">子元素</span><span class="sxs-lookup"><span data-stu-id="cf165-140">Child elements</span></span>

|  <span data-ttu-id="cf165-141">元素</span><span class="sxs-lookup"><span data-stu-id="cf165-141">Element</span></span> |  <span data-ttu-id="cf165-142">必需</span><span class="sxs-lookup"><span data-stu-id="cf165-142">Required</span></span>  |  <span data-ttu-id="cf165-143">说明</span><span class="sxs-lookup"><span data-stu-id="cf165-143">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cf165-144">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="cf165-144">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="cf165-145">是</span><span class="sxs-lookup"><span data-stu-id="cf165-145">Yes</span></span>   |  <span data-ttu-id="cf165-146">定义桌面外形规格的设置。</span><span class="sxs-lookup"><span data-stu-id="cf165-146">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="cf165-147">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="cf165-147">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="cf165-148">否</span><span class="sxs-lookup"><span data-stu-id="cf165-148">No</span></span>   |  <span data-ttu-id="cf165-149">定义移动设备规格的设置。</span><span class="sxs-lookup"><span data-stu-id="cf165-149">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="cf165-150">**注意：** 仅在 iOS 和 Android 上的 Outlook 中支持此元素。</span><span class="sxs-lookup"><span data-stu-id="cf165-150">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="cf165-151">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="cf165-151">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="cf165-152">否</span><span class="sxs-lookup"><span data-stu-id="cf165-152">No</span></span>   |  <span data-ttu-id="cf165-153">定义所有外形规格的设置。</span><span class="sxs-lookup"><span data-stu-id="cf165-153">Defines the settings for all form factors.</span></span> <span data-ttu-id="cf165-154">仅用于 Excel 中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="cf165-154">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="cf165-155">xsi:type</span><span class="sxs-lookup"><span data-stu-id="cf165-155">xsi:type</span></span>

<span data-ttu-id="cf165-156">控制所包含的设置适用的 Office 主机类别（Word、Excel、PowerPoint、Outlook 和 OneNote）。</span><span class="sxs-lookup"><span data-stu-id="cf165-156">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="cf165-157">值必须为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="cf165-157">The value must be one of the following:</span></span>

- <span data-ttu-id="cf165-158">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="cf165-158">`Document` (Word)</span></span>
- <span data-ttu-id="cf165-159">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="cf165-159">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="cf165-160">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="cf165-160">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="cf165-161">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="cf165-161">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="cf165-162">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="cf165-162">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="cf165-163">主机示例</span><span class="sxs-lookup"><span data-stu-id="cf165-163">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
