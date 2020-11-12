---
title: 清单文件中的 OfficeApp 元素
description: OfficeApp 元素是 Office 外接程序清单的根元素。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: c5786343173d0e130df4b786f28a8689d573b6ca
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996317"
---
# <a name="officeapp-element"></a><span data-ttu-id="4df96-103">OfficeApp 元素</span><span class="sxs-lookup"><span data-stu-id="4df96-103">OfficeApp element</span></span>

<span data-ttu-id="4df96-104">Office 外接程序清单中的根元素。</span><span class="sxs-lookup"><span data-stu-id="4df96-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="4df96-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="4df96-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4df96-106">语法</span><span class="sxs-lookup"><span data-stu-id="4df96-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="4df96-107">包含于</span><span class="sxs-lookup"><span data-stu-id="4df96-107">Contained in</span></span>

 <span data-ttu-id="4df96-108">_none_</span><span class="sxs-lookup"><span data-stu-id="4df96-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="4df96-109">必须包含</span><span class="sxs-lookup"><span data-stu-id="4df96-109">Must contain</span></span>

|<span data-ttu-id="4df96-110">元素</span><span class="sxs-lookup"><span data-stu-id="4df96-110">Element</span></span>|<span data-ttu-id="4df96-111">内容</span><span class="sxs-lookup"><span data-stu-id="4df96-111">Content</span></span>|<span data-ttu-id="4df96-112">邮件</span><span class="sxs-lookup"><span data-stu-id="4df96-112">Mail</span></span>|<span data-ttu-id="4df96-113">任务窗格</span><span class="sxs-lookup"><span data-stu-id="4df96-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="4df96-114">Id</span><span class="sxs-lookup"><span data-stu-id="4df96-114">Id</span></span>](id.md)|<span data-ttu-id="4df96-115">x</span><span class="sxs-lookup"><span data-stu-id="4df96-115">x</span></span>|<span data-ttu-id="4df96-116">x</span><span class="sxs-lookup"><span data-stu-id="4df96-116">x</span></span>|<span data-ttu-id="4df96-117">x</span><span class="sxs-lookup"><span data-stu-id="4df96-117">x</span></span>|
|[<span data-ttu-id="4df96-118">版本</span><span class="sxs-lookup"><span data-stu-id="4df96-118">Version</span></span>](version.md)|<span data-ttu-id="4df96-119">x</span><span class="sxs-lookup"><span data-stu-id="4df96-119">x</span></span>|<span data-ttu-id="4df96-120">x</span><span class="sxs-lookup"><span data-stu-id="4df96-120">x</span></span>|<span data-ttu-id="4df96-121">x</span><span class="sxs-lookup"><span data-stu-id="4df96-121">x</span></span>|
|[<span data-ttu-id="4df96-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="4df96-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="4df96-123">x</span><span class="sxs-lookup"><span data-stu-id="4df96-123">x</span></span>|<span data-ttu-id="4df96-124">x</span><span class="sxs-lookup"><span data-stu-id="4df96-124">x</span></span>|<span data-ttu-id="4df96-125">x</span><span class="sxs-lookup"><span data-stu-id="4df96-125">x</span></span>|
|[<span data-ttu-id="4df96-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="4df96-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="4df96-127">x</span><span class="sxs-lookup"><span data-stu-id="4df96-127">x</span></span>|<span data-ttu-id="4df96-128">x</span><span class="sxs-lookup"><span data-stu-id="4df96-128">x</span></span>|<span data-ttu-id="4df96-129">x</span><span class="sxs-lookup"><span data-stu-id="4df96-129">x</span></span>|
|[<span data-ttu-id="4df96-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="4df96-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="4df96-131">x</span><span class="sxs-lookup"><span data-stu-id="4df96-131">x</span></span>||<span data-ttu-id="4df96-132">x</span><span class="sxs-lookup"><span data-stu-id="4df96-132">x</span></span>|
|[<span data-ttu-id="4df96-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="4df96-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="4df96-134">x</span><span class="sxs-lookup"><span data-stu-id="4df96-134">x</span></span>|<span data-ttu-id="4df96-135">x</span><span class="sxs-lookup"><span data-stu-id="4df96-135">x</span></span>|<span data-ttu-id="4df96-136">x</span><span class="sxs-lookup"><span data-stu-id="4df96-136">x</span></span>|
|[<span data-ttu-id="4df96-137">说明</span><span class="sxs-lookup"><span data-stu-id="4df96-137">Description</span></span>](description.md)|<span data-ttu-id="4df96-138">x</span><span class="sxs-lookup"><span data-stu-id="4df96-138">x</span></span>|<span data-ttu-id="4df96-139">x</span><span class="sxs-lookup"><span data-stu-id="4df96-139">x</span></span>|<span data-ttu-id="4df96-140">x</span><span class="sxs-lookup"><span data-stu-id="4df96-140">x</span></span>|
|[<span data-ttu-id="4df96-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="4df96-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="4df96-142">x</span><span class="sxs-lookup"><span data-stu-id="4df96-142">x</span></span>||
|[<span data-ttu-id="4df96-143">Permissions</span><span class="sxs-lookup"><span data-stu-id="4df96-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="4df96-144">x</span><span class="sxs-lookup"><span data-stu-id="4df96-144">x</span></span>||<span data-ttu-id="4df96-145">x</span><span class="sxs-lookup"><span data-stu-id="4df96-145">x</span></span>|
|[<span data-ttu-id="4df96-146">Rule</span><span class="sxs-lookup"><span data-stu-id="4df96-146">Rule</span></span>](rule.md)||<span data-ttu-id="4df96-147">x</span><span class="sxs-lookup"><span data-stu-id="4df96-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="4df96-148">可以包含</span><span class="sxs-lookup"><span data-stu-id="4df96-148">Can contain</span></span>

|<span data-ttu-id="4df96-149">元素</span><span class="sxs-lookup"><span data-stu-id="4df96-149">Element</span></span>|<span data-ttu-id="4df96-150">内容</span><span class="sxs-lookup"><span data-stu-id="4df96-150">Content</span></span>|<span data-ttu-id="4df96-151">邮件</span><span class="sxs-lookup"><span data-stu-id="4df96-151">Mail</span></span>|<span data-ttu-id="4df96-152">任务窗格</span><span class="sxs-lookup"><span data-stu-id="4df96-152">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="4df96-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="4df96-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="4df96-154">x</span><span class="sxs-lookup"><span data-stu-id="4df96-154">x</span></span>|<span data-ttu-id="4df96-155">x</span><span class="sxs-lookup"><span data-stu-id="4df96-155">x</span></span>|<span data-ttu-id="4df96-156">x</span><span class="sxs-lookup"><span data-stu-id="4df96-156">x</span></span>|
|[<span data-ttu-id="4df96-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="4df96-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="4df96-158">x</span><span class="sxs-lookup"><span data-stu-id="4df96-158">x</span></span>|<span data-ttu-id="4df96-159">x</span><span class="sxs-lookup"><span data-stu-id="4df96-159">x</span></span>|<span data-ttu-id="4df96-160">x</span><span class="sxs-lookup"><span data-stu-id="4df96-160">x</span></span>|
|[<span data-ttu-id="4df96-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="4df96-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="4df96-162">x</span><span class="sxs-lookup"><span data-stu-id="4df96-162">x</span></span>|<span data-ttu-id="4df96-163">x</span><span class="sxs-lookup"><span data-stu-id="4df96-163">x</span></span>|<span data-ttu-id="4df96-164">x</span><span class="sxs-lookup"><span data-stu-id="4df96-164">x</span></span>|
|[<span data-ttu-id="4df96-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="4df96-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="4df96-166">x</span><span class="sxs-lookup"><span data-stu-id="4df96-166">x</span></span>|<span data-ttu-id="4df96-167">x</span><span class="sxs-lookup"><span data-stu-id="4df96-167">x</span></span>|<span data-ttu-id="4df96-168">x</span><span class="sxs-lookup"><span data-stu-id="4df96-168">x</span></span>|
|[<span data-ttu-id="4df96-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="4df96-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="4df96-170">x</span><span class="sxs-lookup"><span data-stu-id="4df96-170">x</span></span>|<span data-ttu-id="4df96-171">x</span><span class="sxs-lookup"><span data-stu-id="4df96-171">x</span></span>|<span data-ttu-id="4df96-172">x</span><span class="sxs-lookup"><span data-stu-id="4df96-172">x</span></span>|
|[<span data-ttu-id="4df96-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="4df96-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="4df96-174">x</span><span class="sxs-lookup"><span data-stu-id="4df96-174">x</span></span>|<span data-ttu-id="4df96-175">x</span><span class="sxs-lookup"><span data-stu-id="4df96-175">x</span></span>|<span data-ttu-id="4df96-176">x</span><span class="sxs-lookup"><span data-stu-id="4df96-176">x</span></span>|
|[<span data-ttu-id="4df96-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="4df96-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="4df96-178">x</span><span class="sxs-lookup"><span data-stu-id="4df96-178">x</span></span>|<span data-ttu-id="4df96-179">x</span><span class="sxs-lookup"><span data-stu-id="4df96-179">x</span></span>|<span data-ttu-id="4df96-180">x</span><span class="sxs-lookup"><span data-stu-id="4df96-180">x</span></span>|
|[<span data-ttu-id="4df96-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="4df96-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="4df96-182">x</span><span class="sxs-lookup"><span data-stu-id="4df96-182">x</span></span>|||
|[<span data-ttu-id="4df96-183">Permissions</span><span class="sxs-lookup"><span data-stu-id="4df96-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="4df96-184">x</span><span class="sxs-lookup"><span data-stu-id="4df96-184">x</span></span>||
|[<span data-ttu-id="4df96-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="4df96-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="4df96-186">x</span><span class="sxs-lookup"><span data-stu-id="4df96-186">x</span></span>||
|[<span data-ttu-id="4df96-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="4df96-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="4df96-188">x</span><span class="sxs-lookup"><span data-stu-id="4df96-188">x</span></span>|
|[<span data-ttu-id="4df96-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="4df96-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="4df96-190">x</span><span class="sxs-lookup"><span data-stu-id="4df96-190">x</span></span>|<span data-ttu-id="4df96-191">x</span><span class="sxs-lookup"><span data-stu-id="4df96-191">x</span></span>|<span data-ttu-id="4df96-192">x</span><span class="sxs-lookup"><span data-stu-id="4df96-192">x</span></span>|
|[<span data-ttu-id="4df96-193">ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="4df96-193">ExtendedOverrides</span></span>](extendedoverrides.md)|||<span data-ttu-id="4df96-194">x</span><span class="sxs-lookup"><span data-stu-id="4df96-194">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="4df96-195">属性</span><span class="sxs-lookup"><span data-stu-id="4df96-195">Attributes</span></span>

|<span data-ttu-id="4df96-196">属性</span><span class="sxs-lookup"><span data-stu-id="4df96-196">Attribute</span></span>|<span data-ttu-id="4df96-197">说明</span><span class="sxs-lookup"><span data-stu-id="4df96-197">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="4df96-198">xmlns</span><span class="sxs-lookup"><span data-stu-id="4df96-198">xmlns</span></span>|<span data-ttu-id="4df96-p101">定义的 Office 外接程序清单命名空间和架构版本。应始终将此属性设置为 `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span><span class="sxs-lookup"><span data-stu-id="4df96-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="4df96-201">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="4df96-201">xmlns:xsi</span></span>|<span data-ttu-id="4df96-p102">定义 XMLSchema 实例。应始终将此属性设置为 `"http://www.w3.org/2001/XMLSchema-instance"`</span><span class="sxs-lookup"><span data-stu-id="4df96-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="4df96-204">xsi:type</span><span class="sxs-lookup"><span data-stu-id="4df96-204">xsi:type</span></span>|<span data-ttu-id="4df96-p103">定义的 Office 外接程序的类型。应始终将此属性设置为下列值之一：`"ContentApp"`、`"MailApp"` 或 `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="4df96-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
