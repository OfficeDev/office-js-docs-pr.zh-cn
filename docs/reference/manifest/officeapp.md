---
title: 清单文件中的 OfficeApp 元素
description: ''
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 080025e62a56421dff942792f99ee672ce1db69a
ms.sourcegitcommit: c1dbea577ae6183523fb663d364422d2adbc8bcf
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/05/2020
ms.locfileid: "41773577"
---
# <a name="officeapp-element"></a><span data-ttu-id="fa9f9-102">OfficeApp 元素</span><span class="sxs-lookup"><span data-stu-id="fa9f9-102">OfficeApp element</span></span>

<span data-ttu-id="fa9f9-103">Office 外接程序清单中的根元素。</span><span class="sxs-lookup"><span data-stu-id="fa9f9-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="fa9f9-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="fa9f9-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="fa9f9-105">语法</span><span class="sxs-lookup"><span data-stu-id="fa9f9-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="fa9f9-106">包含于</span><span class="sxs-lookup"><span data-stu-id="fa9f9-106">Contained in</span></span>

 <span data-ttu-id="fa9f9-107">_none_</span><span class="sxs-lookup"><span data-stu-id="fa9f9-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="fa9f9-108">必须包含</span><span class="sxs-lookup"><span data-stu-id="fa9f9-108">Must contain</span></span>

|<span data-ttu-id="fa9f9-109">**元素**</span><span class="sxs-lookup"><span data-stu-id="fa9f9-109">**Element**</span></span>|<span data-ttu-id="fa9f9-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="fa9f9-110">**Content**</span></span>|<span data-ttu-id="fa9f9-111">**Mail**</span><span class="sxs-lookup"><span data-stu-id="fa9f9-111">**Mail**</span></span>|<span data-ttu-id="fa9f9-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="fa9f9-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="fa9f9-113">Id</span><span class="sxs-lookup"><span data-stu-id="fa9f9-113">Id</span></span>](id.md)|<span data-ttu-id="fa9f9-114">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-114">x</span></span>|<span data-ttu-id="fa9f9-115">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-115">x</span></span>|<span data-ttu-id="fa9f9-116">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-116">x</span></span>|
|[<span data-ttu-id="fa9f9-117">版本</span><span class="sxs-lookup"><span data-stu-id="fa9f9-117">Version</span></span>](version.md)|<span data-ttu-id="fa9f9-118">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-118">x</span></span>|<span data-ttu-id="fa9f9-119">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-119">x</span></span>|<span data-ttu-id="fa9f9-120">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-120">x</span></span>|
|[<span data-ttu-id="fa9f9-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="fa9f9-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="fa9f9-122">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-122">x</span></span>|<span data-ttu-id="fa9f9-123">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-123">x</span></span>|<span data-ttu-id="fa9f9-124">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-124">x</span></span>|
|[<span data-ttu-id="fa9f9-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="fa9f9-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="fa9f9-126">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-126">x</span></span>|<span data-ttu-id="fa9f9-127">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-127">x</span></span>|<span data-ttu-id="fa9f9-128">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-128">x</span></span>|
|[<span data-ttu-id="fa9f9-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="fa9f9-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="fa9f9-130">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-130">x</span></span>||<span data-ttu-id="fa9f9-131">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-131">x</span></span>|
|[<span data-ttu-id="fa9f9-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="fa9f9-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="fa9f9-133">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-133">x</span></span>|<span data-ttu-id="fa9f9-134">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-134">x</span></span>|<span data-ttu-id="fa9f9-135">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-135">x</span></span>|
|[<span data-ttu-id="fa9f9-136">说明</span><span class="sxs-lookup"><span data-stu-id="fa9f9-136">Description</span></span>](description.md)|<span data-ttu-id="fa9f9-137">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-137">x</span></span>|<span data-ttu-id="fa9f9-138">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-138">x</span></span>|<span data-ttu-id="fa9f9-139">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-139">x</span></span>|
|[<span data-ttu-id="fa9f9-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="fa9f9-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="fa9f9-141">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-141">x</span></span>||
|[<span data-ttu-id="fa9f9-142">Permissions</span><span class="sxs-lookup"><span data-stu-id="fa9f9-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="fa9f9-143">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-143">x</span></span>||<span data-ttu-id="fa9f9-144">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-144">x</span></span>|
|[<span data-ttu-id="fa9f9-145">Rule</span><span class="sxs-lookup"><span data-stu-id="fa9f9-145">Rule</span></span>](rule.md)||<span data-ttu-id="fa9f9-146">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="fa9f9-147">可以包含</span><span class="sxs-lookup"><span data-stu-id="fa9f9-147">Can contain</span></span>

|<span data-ttu-id="fa9f9-148">**Element**</span><span class="sxs-lookup"><span data-stu-id="fa9f9-148">**Element**</span></span>|<span data-ttu-id="fa9f9-149">**Content**</span><span class="sxs-lookup"><span data-stu-id="fa9f9-149">**Content**</span></span>|<span data-ttu-id="fa9f9-150">**Mail**</span><span class="sxs-lookup"><span data-stu-id="fa9f9-150">**Mail**</span></span>|<span data-ttu-id="fa9f9-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="fa9f9-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="fa9f9-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="fa9f9-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="fa9f9-153">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-153">x</span></span>|<span data-ttu-id="fa9f9-154">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-154">x</span></span>|<span data-ttu-id="fa9f9-155">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-155">x</span></span>|
|[<span data-ttu-id="fa9f9-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="fa9f9-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="fa9f9-157">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-157">x</span></span>|<span data-ttu-id="fa9f9-158">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-158">x</span></span>|<span data-ttu-id="fa9f9-159">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-159">x</span></span>|
|[<span data-ttu-id="fa9f9-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="fa9f9-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="fa9f9-161">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-161">x</span></span>|<span data-ttu-id="fa9f9-162">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-162">x</span></span>|<span data-ttu-id="fa9f9-163">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-163">x</span></span>|
|[<span data-ttu-id="fa9f9-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="fa9f9-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="fa9f9-165">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-165">x</span></span>|<span data-ttu-id="fa9f9-166">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-166">x</span></span>|<span data-ttu-id="fa9f9-167">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-167">x</span></span>|
|[<span data-ttu-id="fa9f9-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="fa9f9-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="fa9f9-169">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-169">x</span></span>|<span data-ttu-id="fa9f9-170">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-170">x</span></span>|<span data-ttu-id="fa9f9-171">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-171">x</span></span>|
|[<span data-ttu-id="fa9f9-172">Hosts</span><span class="sxs-lookup"><span data-stu-id="fa9f9-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="fa9f9-173">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-173">x</span></span>|<span data-ttu-id="fa9f9-174">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-174">x</span></span>|<span data-ttu-id="fa9f9-175">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-175">x</span></span>|
|[<span data-ttu-id="fa9f9-176">Requirements</span><span class="sxs-lookup"><span data-stu-id="fa9f9-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="fa9f9-177">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-177">x</span></span>|<span data-ttu-id="fa9f9-178">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-178">x</span></span>|<span data-ttu-id="fa9f9-179">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-179">x</span></span>|
|[<span data-ttu-id="fa9f9-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="fa9f9-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="fa9f9-181">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-181">x</span></span>|||
|[<span data-ttu-id="fa9f9-182">Permissions</span><span class="sxs-lookup"><span data-stu-id="fa9f9-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="fa9f9-183">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-183">x</span></span>||
|[<span data-ttu-id="fa9f9-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="fa9f9-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="fa9f9-185">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-185">x</span></span>||
|[<span data-ttu-id="fa9f9-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="fa9f9-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="fa9f9-187">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-187">x</span></span>|
|[<span data-ttu-id="fa9f9-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="fa9f9-188">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="fa9f9-189">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-189">x</span></span>|<span data-ttu-id="fa9f9-190">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-190">x</span></span>|<span data-ttu-id="fa9f9-191">x</span><span class="sxs-lookup"><span data-stu-id="fa9f9-191">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="fa9f9-192">属性</span><span class="sxs-lookup"><span data-stu-id="fa9f9-192">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="fa9f9-193">xmlns</span><span class="sxs-lookup"><span data-stu-id="fa9f9-193">xmlns</span></span>|<span data-ttu-id="fa9f9-p101">定义的 Office 外接程序清单命名空间和架构版本。应始终将此属性设置为 `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span><span class="sxs-lookup"><span data-stu-id="fa9f9-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="fa9f9-196">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="fa9f9-196">xmlns:xsi</span></span>|<span data-ttu-id="fa9f9-p102">定义 XMLSchema 实例。应始终将此属性设置为 `"http://www.w3.org/2001/XMLSchema-instance"`</span><span class="sxs-lookup"><span data-stu-id="fa9f9-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="fa9f9-199">xsi:type</span><span class="sxs-lookup"><span data-stu-id="fa9f9-199">xsi:type</span></span>|<span data-ttu-id="fa9f9-p103">定义的 Office 外接程序的类型。应始终将此属性设置为下列值之一：`"ContentApp"`、`"MailApp"` 或 `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="fa9f9-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
