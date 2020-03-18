---
title: 清单文件中的 OfficeApp 元素
description: OfficeApp 元素是 Office 外接程序清单的根元素。
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 038933f2d06ee5f485dbdb7dd7abdbd95fb97c7d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720594"
---
# <a name="officeapp-element"></a><span data-ttu-id="905e1-103">OfficeApp 元素</span><span class="sxs-lookup"><span data-stu-id="905e1-103">OfficeApp element</span></span>

<span data-ttu-id="905e1-104">Office 外接程序清单中的根元素。</span><span class="sxs-lookup"><span data-stu-id="905e1-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="905e1-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="905e1-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="905e1-106">语法</span><span class="sxs-lookup"><span data-stu-id="905e1-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="905e1-107">包含于</span><span class="sxs-lookup"><span data-stu-id="905e1-107">Contained in</span></span>

 <span data-ttu-id="905e1-108">_none_</span><span class="sxs-lookup"><span data-stu-id="905e1-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="905e1-109">必须包含</span><span class="sxs-lookup"><span data-stu-id="905e1-109">Must contain</span></span>

|<span data-ttu-id="905e1-110">**元素**</span><span class="sxs-lookup"><span data-stu-id="905e1-110">**Element**</span></span>|<span data-ttu-id="905e1-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="905e1-111">**Content**</span></span>|<span data-ttu-id="905e1-112">**Mail**</span><span class="sxs-lookup"><span data-stu-id="905e1-112">**Mail**</span></span>|<span data-ttu-id="905e1-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="905e1-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="905e1-114">Id</span><span class="sxs-lookup"><span data-stu-id="905e1-114">Id</span></span>](id.md)|<span data-ttu-id="905e1-115">x</span><span class="sxs-lookup"><span data-stu-id="905e1-115">x</span></span>|<span data-ttu-id="905e1-116">x</span><span class="sxs-lookup"><span data-stu-id="905e1-116">x</span></span>|<span data-ttu-id="905e1-117">x</span><span class="sxs-lookup"><span data-stu-id="905e1-117">x</span></span>|
|[<span data-ttu-id="905e1-118">版本</span><span class="sxs-lookup"><span data-stu-id="905e1-118">Version</span></span>](version.md)|<span data-ttu-id="905e1-119">x</span><span class="sxs-lookup"><span data-stu-id="905e1-119">x</span></span>|<span data-ttu-id="905e1-120">x</span><span class="sxs-lookup"><span data-stu-id="905e1-120">x</span></span>|<span data-ttu-id="905e1-121">x</span><span class="sxs-lookup"><span data-stu-id="905e1-121">x</span></span>|
|[<span data-ttu-id="905e1-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="905e1-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="905e1-123">x</span><span class="sxs-lookup"><span data-stu-id="905e1-123">x</span></span>|<span data-ttu-id="905e1-124">x</span><span class="sxs-lookup"><span data-stu-id="905e1-124">x</span></span>|<span data-ttu-id="905e1-125">x</span><span class="sxs-lookup"><span data-stu-id="905e1-125">x</span></span>|
|[<span data-ttu-id="905e1-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="905e1-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="905e1-127">x</span><span class="sxs-lookup"><span data-stu-id="905e1-127">x</span></span>|<span data-ttu-id="905e1-128">x</span><span class="sxs-lookup"><span data-stu-id="905e1-128">x</span></span>|<span data-ttu-id="905e1-129">x</span><span class="sxs-lookup"><span data-stu-id="905e1-129">x</span></span>|
|[<span data-ttu-id="905e1-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="905e1-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="905e1-131">x</span><span class="sxs-lookup"><span data-stu-id="905e1-131">x</span></span>||<span data-ttu-id="905e1-132">x</span><span class="sxs-lookup"><span data-stu-id="905e1-132">x</span></span>|
|[<span data-ttu-id="905e1-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="905e1-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="905e1-134">x</span><span class="sxs-lookup"><span data-stu-id="905e1-134">x</span></span>|<span data-ttu-id="905e1-135">x</span><span class="sxs-lookup"><span data-stu-id="905e1-135">x</span></span>|<span data-ttu-id="905e1-136">x</span><span class="sxs-lookup"><span data-stu-id="905e1-136">x</span></span>|
|[<span data-ttu-id="905e1-137">说明</span><span class="sxs-lookup"><span data-stu-id="905e1-137">Description</span></span>](description.md)|<span data-ttu-id="905e1-138">x</span><span class="sxs-lookup"><span data-stu-id="905e1-138">x</span></span>|<span data-ttu-id="905e1-139">x</span><span class="sxs-lookup"><span data-stu-id="905e1-139">x</span></span>|<span data-ttu-id="905e1-140">x</span><span class="sxs-lookup"><span data-stu-id="905e1-140">x</span></span>|
|[<span data-ttu-id="905e1-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="905e1-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="905e1-142">x</span><span class="sxs-lookup"><span data-stu-id="905e1-142">x</span></span>||
|[<span data-ttu-id="905e1-143">Permissions</span><span class="sxs-lookup"><span data-stu-id="905e1-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="905e1-144">x</span><span class="sxs-lookup"><span data-stu-id="905e1-144">x</span></span>||<span data-ttu-id="905e1-145">x</span><span class="sxs-lookup"><span data-stu-id="905e1-145">x</span></span>|
|[<span data-ttu-id="905e1-146">Rule</span><span class="sxs-lookup"><span data-stu-id="905e1-146">Rule</span></span>](rule.md)||<span data-ttu-id="905e1-147">x</span><span class="sxs-lookup"><span data-stu-id="905e1-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="905e1-148">可以包含</span><span class="sxs-lookup"><span data-stu-id="905e1-148">Can contain</span></span>

|<span data-ttu-id="905e1-149">**Element**</span><span class="sxs-lookup"><span data-stu-id="905e1-149">**Element**</span></span>|<span data-ttu-id="905e1-150">**Content**</span><span class="sxs-lookup"><span data-stu-id="905e1-150">**Content**</span></span>|<span data-ttu-id="905e1-151">**Mail**</span><span class="sxs-lookup"><span data-stu-id="905e1-151">**Mail**</span></span>|<span data-ttu-id="905e1-152">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="905e1-152">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="905e1-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="905e1-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="905e1-154">x</span><span class="sxs-lookup"><span data-stu-id="905e1-154">x</span></span>|<span data-ttu-id="905e1-155">x</span><span class="sxs-lookup"><span data-stu-id="905e1-155">x</span></span>|<span data-ttu-id="905e1-156">x</span><span class="sxs-lookup"><span data-stu-id="905e1-156">x</span></span>|
|[<span data-ttu-id="905e1-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="905e1-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="905e1-158">x</span><span class="sxs-lookup"><span data-stu-id="905e1-158">x</span></span>|<span data-ttu-id="905e1-159">x</span><span class="sxs-lookup"><span data-stu-id="905e1-159">x</span></span>|<span data-ttu-id="905e1-160">x</span><span class="sxs-lookup"><span data-stu-id="905e1-160">x</span></span>|
|[<span data-ttu-id="905e1-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="905e1-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="905e1-162">x</span><span class="sxs-lookup"><span data-stu-id="905e1-162">x</span></span>|<span data-ttu-id="905e1-163">x</span><span class="sxs-lookup"><span data-stu-id="905e1-163">x</span></span>|<span data-ttu-id="905e1-164">x</span><span class="sxs-lookup"><span data-stu-id="905e1-164">x</span></span>|
|[<span data-ttu-id="905e1-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="905e1-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="905e1-166">x</span><span class="sxs-lookup"><span data-stu-id="905e1-166">x</span></span>|<span data-ttu-id="905e1-167">x</span><span class="sxs-lookup"><span data-stu-id="905e1-167">x</span></span>|<span data-ttu-id="905e1-168">x</span><span class="sxs-lookup"><span data-stu-id="905e1-168">x</span></span>|
|[<span data-ttu-id="905e1-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="905e1-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="905e1-170">x</span><span class="sxs-lookup"><span data-stu-id="905e1-170">x</span></span>|<span data-ttu-id="905e1-171">x</span><span class="sxs-lookup"><span data-stu-id="905e1-171">x</span></span>|<span data-ttu-id="905e1-172">x</span><span class="sxs-lookup"><span data-stu-id="905e1-172">x</span></span>|
|[<span data-ttu-id="905e1-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="905e1-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="905e1-174">x</span><span class="sxs-lookup"><span data-stu-id="905e1-174">x</span></span>|<span data-ttu-id="905e1-175">x</span><span class="sxs-lookup"><span data-stu-id="905e1-175">x</span></span>|<span data-ttu-id="905e1-176">x</span><span class="sxs-lookup"><span data-stu-id="905e1-176">x</span></span>|
|[<span data-ttu-id="905e1-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="905e1-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="905e1-178">x</span><span class="sxs-lookup"><span data-stu-id="905e1-178">x</span></span>|<span data-ttu-id="905e1-179">x</span><span class="sxs-lookup"><span data-stu-id="905e1-179">x</span></span>|<span data-ttu-id="905e1-180">x</span><span class="sxs-lookup"><span data-stu-id="905e1-180">x</span></span>|
|[<span data-ttu-id="905e1-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="905e1-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="905e1-182">x</span><span class="sxs-lookup"><span data-stu-id="905e1-182">x</span></span>|||
|[<span data-ttu-id="905e1-183">Permissions</span><span class="sxs-lookup"><span data-stu-id="905e1-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="905e1-184">x</span><span class="sxs-lookup"><span data-stu-id="905e1-184">x</span></span>||
|[<span data-ttu-id="905e1-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="905e1-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="905e1-186">x</span><span class="sxs-lookup"><span data-stu-id="905e1-186">x</span></span>||
|[<span data-ttu-id="905e1-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="905e1-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="905e1-188">x</span><span class="sxs-lookup"><span data-stu-id="905e1-188">x</span></span>|
|[<span data-ttu-id="905e1-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="905e1-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="905e1-190">x</span><span class="sxs-lookup"><span data-stu-id="905e1-190">x</span></span>|<span data-ttu-id="905e1-191">x</span><span class="sxs-lookup"><span data-stu-id="905e1-191">x</span></span>|<span data-ttu-id="905e1-192">x</span><span class="sxs-lookup"><span data-stu-id="905e1-192">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="905e1-193">属性</span><span class="sxs-lookup"><span data-stu-id="905e1-193">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="905e1-194">xmlns</span><span class="sxs-lookup"><span data-stu-id="905e1-194">xmlns</span></span>|<span data-ttu-id="905e1-p101">定义的 Office 外接程序清单命名空间和架构版本。应始终将此属性设置为 `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span><span class="sxs-lookup"><span data-stu-id="905e1-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="905e1-197">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="905e1-197">xmlns:xsi</span></span>|<span data-ttu-id="905e1-p102">定义 XMLSchema 实例。应始终将此属性设置为 `"http://www.w3.org/2001/XMLSchema-instance"`</span><span class="sxs-lookup"><span data-stu-id="905e1-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="905e1-200">xsi:type</span><span class="sxs-lookup"><span data-stu-id="905e1-200">xsi:type</span></span>|<span data-ttu-id="905e1-p103">定义的 Office 外接程序的类型。应始终将此属性设置为下列值之一：`"ContentApp"`、`"MailApp"` 或 `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="905e1-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
