---
title: 清单文件中的 OfficeApp 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 86f38ab77e98bb01370e40c8ada38bae171e0c2d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450455"
---
# <a name="officeapp-element"></a><span data-ttu-id="33176-102">OfficeApp 元素</span><span class="sxs-lookup"><span data-stu-id="33176-102">OfficeApp element</span></span>

<span data-ttu-id="33176-103">Office 外接程序清单中的根元素。</span><span class="sxs-lookup"><span data-stu-id="33176-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="33176-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="33176-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="33176-105">语法</span><span class="sxs-lookup"><span data-stu-id="33176-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="33176-106">包含于</span><span class="sxs-lookup"><span data-stu-id="33176-106">Contained in</span></span>

 <span data-ttu-id="33176-107">_none_</span><span class="sxs-lookup"><span data-stu-id="33176-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="33176-108">必须包含</span><span class="sxs-lookup"><span data-stu-id="33176-108">Must contain</span></span>

|<span data-ttu-id="33176-109">**元素**</span><span class="sxs-lookup"><span data-stu-id="33176-109">**Element**</span></span>|<span data-ttu-id="33176-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="33176-110">**Content**</span></span>|<span data-ttu-id="33176-111">**Mail**</span><span class="sxs-lookup"><span data-stu-id="33176-111">**Mail**</span></span>|<span data-ttu-id="33176-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="33176-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="33176-113">Id</span><span class="sxs-lookup"><span data-stu-id="33176-113">Id</span></span>](id.md)|<span data-ttu-id="33176-114">x</span><span class="sxs-lookup"><span data-stu-id="33176-114">x</span></span>|<span data-ttu-id="33176-115">x</span><span class="sxs-lookup"><span data-stu-id="33176-115">x</span></span>|<span data-ttu-id="33176-116">x</span><span class="sxs-lookup"><span data-stu-id="33176-116">x</span></span>|
|[<span data-ttu-id="33176-117">版本</span><span class="sxs-lookup"><span data-stu-id="33176-117">Version</span></span>](version.md)|<span data-ttu-id="33176-118">x</span><span class="sxs-lookup"><span data-stu-id="33176-118">x</span></span>|<span data-ttu-id="33176-119">x</span><span class="sxs-lookup"><span data-stu-id="33176-119">x</span></span>|<span data-ttu-id="33176-120">x</span><span class="sxs-lookup"><span data-stu-id="33176-120">x</span></span>|
|[<span data-ttu-id="33176-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="33176-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="33176-122">x</span><span class="sxs-lookup"><span data-stu-id="33176-122">x</span></span>|<span data-ttu-id="33176-123">x</span><span class="sxs-lookup"><span data-stu-id="33176-123">x</span></span>|<span data-ttu-id="33176-124">x</span><span class="sxs-lookup"><span data-stu-id="33176-124">x</span></span>|
|[<span data-ttu-id="33176-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="33176-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="33176-126">x</span><span class="sxs-lookup"><span data-stu-id="33176-126">x</span></span>|<span data-ttu-id="33176-127">x</span><span class="sxs-lookup"><span data-stu-id="33176-127">x</span></span>|<span data-ttu-id="33176-128">x</span><span class="sxs-lookup"><span data-stu-id="33176-128">x</span></span>|
|[<span data-ttu-id="33176-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="33176-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="33176-130">x</span><span class="sxs-lookup"><span data-stu-id="33176-130">x</span></span>||<span data-ttu-id="33176-131">x</span><span class="sxs-lookup"><span data-stu-id="33176-131">x</span></span>|
|[<span data-ttu-id="33176-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="33176-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="33176-133">x</span><span class="sxs-lookup"><span data-stu-id="33176-133">x</span></span>|<span data-ttu-id="33176-134">x</span><span class="sxs-lookup"><span data-stu-id="33176-134">x</span></span>|<span data-ttu-id="33176-135">x</span><span class="sxs-lookup"><span data-stu-id="33176-135">x</span></span>|
|[<span data-ttu-id="33176-136">说明</span><span class="sxs-lookup"><span data-stu-id="33176-136">Description</span></span>](description.md)|<span data-ttu-id="33176-137">x</span><span class="sxs-lookup"><span data-stu-id="33176-137">x</span></span>|<span data-ttu-id="33176-138">x</span><span class="sxs-lookup"><span data-stu-id="33176-138">x</span></span>|<span data-ttu-id="33176-139">x</span><span class="sxs-lookup"><span data-stu-id="33176-139">x</span></span>|
|[<span data-ttu-id="33176-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="33176-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="33176-141">x</span><span class="sxs-lookup"><span data-stu-id="33176-141">x</span></span>||
|[<span data-ttu-id="33176-142">Permissions</span><span class="sxs-lookup"><span data-stu-id="33176-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="33176-143">x</span><span class="sxs-lookup"><span data-stu-id="33176-143">x</span></span>||<span data-ttu-id="33176-144">x</span><span class="sxs-lookup"><span data-stu-id="33176-144">x</span></span>|
|[<span data-ttu-id="33176-145">Rule</span><span class="sxs-lookup"><span data-stu-id="33176-145">Rule</span></span>](rule.md)||<span data-ttu-id="33176-146">x</span><span class="sxs-lookup"><span data-stu-id="33176-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="33176-147">可以包含</span><span class="sxs-lookup"><span data-stu-id="33176-147">Can contain</span></span>

|<span data-ttu-id="33176-148">**Element**</span><span class="sxs-lookup"><span data-stu-id="33176-148">**Element**</span></span>|<span data-ttu-id="33176-149">**Content**</span><span class="sxs-lookup"><span data-stu-id="33176-149">**Content**</span></span>|<span data-ttu-id="33176-150">**Mail**</span><span class="sxs-lookup"><span data-stu-id="33176-150">**Mail**</span></span>|<span data-ttu-id="33176-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="33176-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="33176-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="33176-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="33176-153">x</span><span class="sxs-lookup"><span data-stu-id="33176-153">x</span></span>|<span data-ttu-id="33176-154">x</span><span class="sxs-lookup"><span data-stu-id="33176-154">x</span></span>|<span data-ttu-id="33176-155">x</span><span class="sxs-lookup"><span data-stu-id="33176-155">x</span></span>|
|[<span data-ttu-id="33176-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="33176-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="33176-157">x</span><span class="sxs-lookup"><span data-stu-id="33176-157">x</span></span>|<span data-ttu-id="33176-158">x</span><span class="sxs-lookup"><span data-stu-id="33176-158">x</span></span>|<span data-ttu-id="33176-159">x</span><span class="sxs-lookup"><span data-stu-id="33176-159">x</span></span>|
|[<span data-ttu-id="33176-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="33176-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="33176-161">x</span><span class="sxs-lookup"><span data-stu-id="33176-161">x</span></span>|<span data-ttu-id="33176-162">x</span><span class="sxs-lookup"><span data-stu-id="33176-162">x</span></span>|<span data-ttu-id="33176-163">x</span><span class="sxs-lookup"><span data-stu-id="33176-163">x</span></span>|
|[<span data-ttu-id="33176-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="33176-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="33176-165">x</span><span class="sxs-lookup"><span data-stu-id="33176-165">x</span></span>|<span data-ttu-id="33176-166">x</span><span class="sxs-lookup"><span data-stu-id="33176-166">x</span></span>|<span data-ttu-id="33176-167">x</span><span class="sxs-lookup"><span data-stu-id="33176-167">x</span></span>|
|[<span data-ttu-id="33176-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="33176-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="33176-169">x</span><span class="sxs-lookup"><span data-stu-id="33176-169">x</span></span>|<span data-ttu-id="33176-170">x</span><span class="sxs-lookup"><span data-stu-id="33176-170">x</span></span>|<span data-ttu-id="33176-171">x</span><span class="sxs-lookup"><span data-stu-id="33176-171">x</span></span>|
|[<span data-ttu-id="33176-172">Hosts</span><span class="sxs-lookup"><span data-stu-id="33176-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="33176-173">x</span><span class="sxs-lookup"><span data-stu-id="33176-173">x</span></span>|<span data-ttu-id="33176-174">x</span><span class="sxs-lookup"><span data-stu-id="33176-174">x</span></span>|<span data-ttu-id="33176-175">x</span><span class="sxs-lookup"><span data-stu-id="33176-175">x</span></span>|
|[<span data-ttu-id="33176-176">Requirements</span><span class="sxs-lookup"><span data-stu-id="33176-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="33176-177">x</span><span class="sxs-lookup"><span data-stu-id="33176-177">x</span></span>|<span data-ttu-id="33176-178">x</span><span class="sxs-lookup"><span data-stu-id="33176-178">x</span></span>|<span data-ttu-id="33176-179">x</span><span class="sxs-lookup"><span data-stu-id="33176-179">x</span></span>|
|[<span data-ttu-id="33176-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="33176-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="33176-181">x</span><span class="sxs-lookup"><span data-stu-id="33176-181">x</span></span>|||
|[<span data-ttu-id="33176-182">Permissions</span><span class="sxs-lookup"><span data-stu-id="33176-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="33176-183">x</span><span class="sxs-lookup"><span data-stu-id="33176-183">x</span></span>||
|[<span data-ttu-id="33176-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="33176-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="33176-185">x</span><span class="sxs-lookup"><span data-stu-id="33176-185">x</span></span>||
|[<span data-ttu-id="33176-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="33176-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="33176-187">x</span><span class="sxs-lookup"><span data-stu-id="33176-187">x</span></span>|
|[<span data-ttu-id="33176-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="33176-188">VersionOverrides</span></span>](versionoverrides.md)||<span data-ttu-id="33176-189">x</span><span class="sxs-lookup"><span data-stu-id="33176-189">x</span></span>||

## <a name="attributes"></a><span data-ttu-id="33176-190">属性</span><span class="sxs-lookup"><span data-stu-id="33176-190">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="33176-191">xmlns</span><span class="sxs-lookup"><span data-stu-id="33176-191">xmlns</span></span>|<span data-ttu-id="33176-p101">定义的 Office 外接程序清单命名空间和架构版本。应始终将此属性设置为 `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span><span class="sxs-lookup"><span data-stu-id="33176-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="33176-194">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="33176-194">xmlns:xsi</span></span>|<span data-ttu-id="33176-p102">定义 XMLSchema 实例。应始终将此属性设置为 `"http://www.w3.org/2001/XMLSchema-instance"`</span><span class="sxs-lookup"><span data-stu-id="33176-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="33176-197">xsi:type</span><span class="sxs-lookup"><span data-stu-id="33176-197">xsi:type</span></span>|<span data-ttu-id="33176-p103">定义的 Office 外接程序的类型。应始终将此属性设置为下列值之一：`"ContentApp"`、`"MailApp"` 或 `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="33176-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
