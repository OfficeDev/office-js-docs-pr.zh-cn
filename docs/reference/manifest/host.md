---
title: 清单文件中的 Host 元素
description: 指定应在其中激活外接程序的单个 Office 应用程序类型。
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 45d4ed42946038699be235ff3912c071a92ff226
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348326"
---
# <a name="host-element"></a><span data-ttu-id="27141-103">Host 元素</span><span class="sxs-lookup"><span data-stu-id="27141-103">Host element</span></span>

<span data-ttu-id="27141-104">指定应在其中激活外接程序的单个 Office 应用程序类型。</span><span class="sxs-lookup"><span data-stu-id="27141-104">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="27141-105">**Host** 元素的语法根据该元素是否在 [基本清单](#basic-manifest)中或 [VersionOverrides](#versionoverrides-node) 节点中定义而不同。</span><span class="sxs-lookup"><span data-stu-id="27141-105">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="27141-106">但功能相同。</span><span class="sxs-lookup"><span data-stu-id="27141-106">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="27141-107">基本清单</span><span class="sxs-lookup"><span data-stu-id="27141-107">Basic manifest</span></span>

<span data-ttu-id="27141-108">在基本清单（在 [OfficeApp](officeapp.md) 下）中定义时，主机类型由 `Name` 属性决定。</span><span class="sxs-lookup"><span data-stu-id="27141-108">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="27141-109">属性</span><span class="sxs-lookup"><span data-stu-id="27141-109">Attributes</span></span>

| <span data-ttu-id="27141-110">属性</span><span class="sxs-lookup"><span data-stu-id="27141-110">Attribute</span></span>     | <span data-ttu-id="27141-111">类型</span><span class="sxs-lookup"><span data-stu-id="27141-111">Type</span></span>   | <span data-ttu-id="27141-112">必需</span><span class="sxs-lookup"><span data-stu-id="27141-112">Required</span></span> | <span data-ttu-id="27141-113">说明</span><span class="sxs-lookup"><span data-stu-id="27141-113">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="27141-114">Name</span><span class="sxs-lookup"><span data-stu-id="27141-114">Name</span></span>](#name) | <span data-ttu-id="27141-115">string</span><span class="sxs-lookup"><span data-stu-id="27141-115">string</span></span> | <span data-ttu-id="27141-116">必需</span><span class="sxs-lookup"><span data-stu-id="27141-116">required</span></span> | <span data-ttu-id="27141-117">客户端应用程序Office的名称。</span><span class="sxs-lookup"><span data-stu-id="27141-117">The name of the type of Office client application.</span></span> |

### <a name="name"></a><span data-ttu-id="27141-118">名称</span><span class="sxs-lookup"><span data-stu-id="27141-118">Name</span></span>

<span data-ttu-id="27141-p102">指定此外接程序面向的主机类型。值必须为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="27141-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="27141-121">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="27141-121">`Document` (Word)</span></span>
- <span data-ttu-id="27141-122">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="27141-122">`Database` (Access)</span></span>
- <span data-ttu-id="27141-123">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="27141-123">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="27141-124">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="27141-124">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="27141-125">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="27141-125">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="27141-126">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="27141-126">`Project` (Project)</span></span>
- <span data-ttu-id="27141-127">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="27141-127">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="27141-128">我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。</span><span class="sxs-lookup"><span data-stu-id="27141-128">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="27141-129">作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。</span><span class="sxs-lookup"><span data-stu-id="27141-129">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="27141-130">示例</span><span class="sxs-lookup"><span data-stu-id="27141-130">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="27141-131">VersionOverrides 节点</span><span class="sxs-lookup"><span data-stu-id="27141-131">VersionOverrides node</span></span>

<span data-ttu-id="27141-132">在 [VersionOverrides](versionoverrides.md) 中定义时，主机类型由 `xsi:type` 属性决定。</span><span class="sxs-lookup"><span data-stu-id="27141-132">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="27141-133">属性</span><span class="sxs-lookup"><span data-stu-id="27141-133">Attributes</span></span>

|  <span data-ttu-id="27141-134">属性</span><span class="sxs-lookup"><span data-stu-id="27141-134">Attribute</span></span>  |  <span data-ttu-id="27141-135">必需</span><span class="sxs-lookup"><span data-stu-id="27141-135">Required</span></span>  |  <span data-ttu-id="27141-136">说明</span><span class="sxs-lookup"><span data-stu-id="27141-136">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="27141-137">xsi:type</span><span class="sxs-lookup"><span data-stu-id="27141-137">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="27141-138">是</span><span class="sxs-lookup"><span data-stu-id="27141-138">Yes</span></span>  | <span data-ttu-id="27141-139">介绍Office应用这些设置的应用程序。</span><span class="sxs-lookup"><span data-stu-id="27141-139">Describes the Office application where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="27141-140">子元素</span><span class="sxs-lookup"><span data-stu-id="27141-140">Child elements</span></span>

|  <span data-ttu-id="27141-141">元素</span><span class="sxs-lookup"><span data-stu-id="27141-141">Element</span></span> |  <span data-ttu-id="27141-142">必需</span><span class="sxs-lookup"><span data-stu-id="27141-142">Required</span></span>  |  <span data-ttu-id="27141-143">说明</span><span class="sxs-lookup"><span data-stu-id="27141-143">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="27141-144">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="27141-144">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="27141-145">是</span><span class="sxs-lookup"><span data-stu-id="27141-145">Yes</span></span>   |  <span data-ttu-id="27141-146">定义桌面外形规格的设置。</span><span class="sxs-lookup"><span data-stu-id="27141-146">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="27141-147">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="27141-147">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="27141-148">否</span><span class="sxs-lookup"><span data-stu-id="27141-148">No</span></span>   |  <span data-ttu-id="27141-149">定义移动外形因素的设置。</span><span class="sxs-lookup"><span data-stu-id="27141-149">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="27141-150">**注意：** 此元素仅在 iOS Outlook Android 上的设备上受支持。</span><span class="sxs-lookup"><span data-stu-id="27141-150">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="27141-151">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="27141-151">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="27141-152">否</span><span class="sxs-lookup"><span data-stu-id="27141-152">No</span></span>   |  <span data-ttu-id="27141-153">定义所有外形规格的设置。</span><span class="sxs-lookup"><span data-stu-id="27141-153">Defines the settings for all form factors.</span></span> <span data-ttu-id="27141-154">仅用于 Excel 中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="27141-154">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="27141-155">xsi:type</span><span class="sxs-lookup"><span data-stu-id="27141-155">xsi:type</span></span>

<span data-ttu-id="27141-156">控制应用程序Office Word (、Excel、PowerPoint、Outlook OneNote) 应用所包含的设置。</span><span class="sxs-lookup"><span data-stu-id="27141-156">Controls which Office application (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="27141-157">值必须为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="27141-157">The value must be one of the following:</span></span>

- <span data-ttu-id="27141-158">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="27141-158">`Document` (Word)</span></span>
- <span data-ttu-id="27141-159">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="27141-159">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="27141-160">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="27141-160">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="27141-161">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="27141-161">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="27141-162">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="27141-162">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="27141-163">主机示例</span><span class="sxs-lookup"><span data-stu-id="27141-163">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
