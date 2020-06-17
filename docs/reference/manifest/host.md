---
title: 清单文件中的 Host 元素
description: 指定应在其中激活外接程序的单个 Office 应用程序类型。
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 5db9df97c4ba558d54756b983a26cb7b71e049d5
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611811"
---
# <a name="host-element"></a><span data-ttu-id="c851a-103">Host 元素</span><span class="sxs-lookup"><span data-stu-id="c851a-103">Host element</span></span>

<span data-ttu-id="c851a-104">指定应在其中激活外接程序的单个 Office 应用程序类型。</span><span class="sxs-lookup"><span data-stu-id="c851a-104">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c851a-105">**Host** 元素的语法根据该元素是否在[基本清单](#basic-manifest)中或 [VersionOverrides](#versionoverrides-node) 节点中定义而不同。</span><span class="sxs-lookup"><span data-stu-id="c851a-105">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="c851a-106">但功能相同。</span><span class="sxs-lookup"><span data-stu-id="c851a-106">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="c851a-107">基本清单</span><span class="sxs-lookup"><span data-stu-id="c851a-107">Basic manifest</span></span>

<span data-ttu-id="c851a-108">在基本清单（在 [OfficeApp](officeapp.md) 下）中定义时，主机类型由 `Name` 属性决定。</span><span class="sxs-lookup"><span data-stu-id="c851a-108">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="c851a-109">属性</span><span class="sxs-lookup"><span data-stu-id="c851a-109">Attributes</span></span>

| <span data-ttu-id="c851a-110">属性</span><span class="sxs-lookup"><span data-stu-id="c851a-110">Attribute</span></span>     | <span data-ttu-id="c851a-111">类型</span><span class="sxs-lookup"><span data-stu-id="c851a-111">Type</span></span>   | <span data-ttu-id="c851a-112">必需</span><span class="sxs-lookup"><span data-stu-id="c851a-112">Required</span></span> | <span data-ttu-id="c851a-113">说明</span><span class="sxs-lookup"><span data-stu-id="c851a-113">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="c851a-114">Name</span><span class="sxs-lookup"><span data-stu-id="c851a-114">Name</span></span>](#name) | <span data-ttu-id="c851a-115">string</span><span class="sxs-lookup"><span data-stu-id="c851a-115">string</span></span> | <span data-ttu-id="c851a-116">必需</span><span class="sxs-lookup"><span data-stu-id="c851a-116">required</span></span> | <span data-ttu-id="c851a-117">Office 主机应用程序的类型名称。</span><span class="sxs-lookup"><span data-stu-id="c851a-117">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="c851a-118">名称</span><span class="sxs-lookup"><span data-stu-id="c851a-118">Name</span></span>

<span data-ttu-id="c851a-119">指定此外接程序面向的主机类型。</span><span class="sxs-lookup"><span data-stu-id="c851a-119">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="c851a-120">值必须是下列值之一。</span><span class="sxs-lookup"><span data-stu-id="c851a-120">The value must be one of the following.</span></span>

- <span data-ttu-id="c851a-121">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="c851a-121">`Document` (Word)</span></span>
- <span data-ttu-id="c851a-122">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="c851a-122">`Database` (Access)</span></span>
- <span data-ttu-id="c851a-123">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="c851a-123">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="c851a-124">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="c851a-124">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="c851a-125">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="c851a-125">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="c851a-126">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="c851a-126">`Project` (Project)</span></span>
- <span data-ttu-id="c851a-127">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="c851a-127">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c851a-128">我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。</span><span class="sxs-lookup"><span data-stu-id="c851a-128">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="c851a-129">作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。</span><span class="sxs-lookup"><span data-stu-id="c851a-129">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="c851a-130">示例</span><span class="sxs-lookup"><span data-stu-id="c851a-130">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="c851a-131">VersionOverrides 节点</span><span class="sxs-lookup"><span data-stu-id="c851a-131">VersionOverrides node</span></span>

<span data-ttu-id="c851a-132">在 [VersionOverrides](versionoverrides.md) 中定义时，主机类型由 `xsi:type` 属性决定。</span><span class="sxs-lookup"><span data-stu-id="c851a-132">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="c851a-133">属性</span><span class="sxs-lookup"><span data-stu-id="c851a-133">Attributes</span></span>

|  <span data-ttu-id="c851a-134">属性</span><span class="sxs-lookup"><span data-stu-id="c851a-134">Attribute</span></span>  |  <span data-ttu-id="c851a-135">必需</span><span class="sxs-lookup"><span data-stu-id="c851a-135">Required</span></span>  |  <span data-ttu-id="c851a-136">Description</span><span class="sxs-lookup"><span data-stu-id="c851a-136">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c851a-137">xsi:type</span><span class="sxs-lookup"><span data-stu-id="c851a-137">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="c851a-138">是</span><span class="sxs-lookup"><span data-stu-id="c851a-138">Yes</span></span>  | <span data-ttu-id="c851a-139">描述这些设置适用的 Office 主机。</span><span class="sxs-lookup"><span data-stu-id="c851a-139">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="c851a-140">子元素</span><span class="sxs-lookup"><span data-stu-id="c851a-140">Child elements</span></span>

|  <span data-ttu-id="c851a-141">元素</span><span class="sxs-lookup"><span data-stu-id="c851a-141">Element</span></span> |  <span data-ttu-id="c851a-142">必需</span><span class="sxs-lookup"><span data-stu-id="c851a-142">Required</span></span>  |  <span data-ttu-id="c851a-143">Description</span><span class="sxs-lookup"><span data-stu-id="c851a-143">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c851a-144">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="c851a-144">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="c851a-145">是</span><span class="sxs-lookup"><span data-stu-id="c851a-145">Yes</span></span>   |  <span data-ttu-id="c851a-146">定义桌面外形规格的设置。</span><span class="sxs-lookup"><span data-stu-id="c851a-146">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="c851a-147">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="c851a-147">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="c851a-148">否</span><span class="sxs-lookup"><span data-stu-id="c851a-148">No</span></span>   |  <span data-ttu-id="c851a-149">定义移动设备规格的设置。</span><span class="sxs-lookup"><span data-stu-id="c851a-149">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="c851a-150">**注意：** 仅在 iOS 和 Android 上的 Outlook 中支持此元素。</span><span class="sxs-lookup"><span data-stu-id="c851a-150">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="c851a-151">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="c851a-151">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="c851a-152">否</span><span class="sxs-lookup"><span data-stu-id="c851a-152">No</span></span>   |  <span data-ttu-id="c851a-153">定义所有外形规格的设置。</span><span class="sxs-lookup"><span data-stu-id="c851a-153">Defines the settings for all form factors.</span></span> <span data-ttu-id="c851a-154">仅用于 Excel 中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="c851a-154">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="c851a-155">xsi:type</span><span class="sxs-lookup"><span data-stu-id="c851a-155">xsi:type</span></span>

<span data-ttu-id="c851a-156">控制所包含的设置适用的 Office 主机类别（Word、Excel、PowerPoint、Outlook 和 OneNote）。</span><span class="sxs-lookup"><span data-stu-id="c851a-156">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="c851a-157">值必须为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="c851a-157">The value must be one of the following:</span></span>

- <span data-ttu-id="c851a-158">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="c851a-158">`Document` (Word)</span></span>
- <span data-ttu-id="c851a-159">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="c851a-159">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="c851a-160">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="c851a-160">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="c851a-161">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="c851a-161">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="c851a-162">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="c851a-162">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="c851a-163">主机示例</span><span class="sxs-lookup"><span data-stu-id="c851a-163">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
