---
title: 清单文件中的 OfficeApp 元素
description: OfficeApp 元素是 Office 外接程序清单的根元素。
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: b6f3102a97794a19366b06734789e01fc4bc4f9d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611524"
---
# <a name="officeapp-element"></a><span data-ttu-id="29b46-103">OfficeApp 元素</span><span class="sxs-lookup"><span data-stu-id="29b46-103">OfficeApp element</span></span>

<span data-ttu-id="29b46-104">Office 外接程序清单中的根元素。</span><span class="sxs-lookup"><span data-stu-id="29b46-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="29b46-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="29b46-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="29b46-106">语法</span><span class="sxs-lookup"><span data-stu-id="29b46-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="29b46-107">包含于</span><span class="sxs-lookup"><span data-stu-id="29b46-107">Contained in</span></span>

 <span data-ttu-id="29b46-108">_none_</span><span class="sxs-lookup"><span data-stu-id="29b46-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="29b46-109">必须包含</span><span class="sxs-lookup"><span data-stu-id="29b46-109">Must contain</span></span>

|<span data-ttu-id="29b46-110">**元素**</span><span class="sxs-lookup"><span data-stu-id="29b46-110">**Element**</span></span>|<span data-ttu-id="29b46-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="29b46-111">**Content**</span></span>|<span data-ttu-id="29b46-112">**Mail**</span><span class="sxs-lookup"><span data-stu-id="29b46-112">**Mail**</span></span>|<span data-ttu-id="29b46-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="29b46-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="29b46-114">Id</span><span class="sxs-lookup"><span data-stu-id="29b46-114">Id</span></span>](id.md)|<span data-ttu-id="29b46-115">x</span><span class="sxs-lookup"><span data-stu-id="29b46-115">x</span></span>|<span data-ttu-id="29b46-116">x</span><span class="sxs-lookup"><span data-stu-id="29b46-116">x</span></span>|<span data-ttu-id="29b46-117">x</span><span class="sxs-lookup"><span data-stu-id="29b46-117">x</span></span>|
|[<span data-ttu-id="29b46-118">版本</span><span class="sxs-lookup"><span data-stu-id="29b46-118">Version</span></span>](version.md)|<span data-ttu-id="29b46-119">x</span><span class="sxs-lookup"><span data-stu-id="29b46-119">x</span></span>|<span data-ttu-id="29b46-120">x</span><span class="sxs-lookup"><span data-stu-id="29b46-120">x</span></span>|<span data-ttu-id="29b46-121">x</span><span class="sxs-lookup"><span data-stu-id="29b46-121">x</span></span>|
|[<span data-ttu-id="29b46-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="29b46-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="29b46-123">x</span><span class="sxs-lookup"><span data-stu-id="29b46-123">x</span></span>|<span data-ttu-id="29b46-124">x</span><span class="sxs-lookup"><span data-stu-id="29b46-124">x</span></span>|<span data-ttu-id="29b46-125">x</span><span class="sxs-lookup"><span data-stu-id="29b46-125">x</span></span>|
|[<span data-ttu-id="29b46-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="29b46-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="29b46-127">x</span><span class="sxs-lookup"><span data-stu-id="29b46-127">x</span></span>|<span data-ttu-id="29b46-128">x</span><span class="sxs-lookup"><span data-stu-id="29b46-128">x</span></span>|<span data-ttu-id="29b46-129">x</span><span class="sxs-lookup"><span data-stu-id="29b46-129">x</span></span>|
|[<span data-ttu-id="29b46-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="29b46-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="29b46-131">x</span><span class="sxs-lookup"><span data-stu-id="29b46-131">x</span></span>||<span data-ttu-id="29b46-132">x</span><span class="sxs-lookup"><span data-stu-id="29b46-132">x</span></span>|
|[<span data-ttu-id="29b46-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="29b46-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="29b46-134">x</span><span class="sxs-lookup"><span data-stu-id="29b46-134">x</span></span>|<span data-ttu-id="29b46-135">x</span><span class="sxs-lookup"><span data-stu-id="29b46-135">x</span></span>|<span data-ttu-id="29b46-136">x</span><span class="sxs-lookup"><span data-stu-id="29b46-136">x</span></span>|
|[<span data-ttu-id="29b46-137">说明</span><span class="sxs-lookup"><span data-stu-id="29b46-137">Description</span></span>](description.md)|<span data-ttu-id="29b46-138">x</span><span class="sxs-lookup"><span data-stu-id="29b46-138">x</span></span>|<span data-ttu-id="29b46-139">x</span><span class="sxs-lookup"><span data-stu-id="29b46-139">x</span></span>|<span data-ttu-id="29b46-140">x</span><span class="sxs-lookup"><span data-stu-id="29b46-140">x</span></span>|
|[<span data-ttu-id="29b46-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="29b46-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="29b46-142">x</span><span class="sxs-lookup"><span data-stu-id="29b46-142">x</span></span>||
|[<span data-ttu-id="29b46-143">Permissions</span><span class="sxs-lookup"><span data-stu-id="29b46-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="29b46-144">x</span><span class="sxs-lookup"><span data-stu-id="29b46-144">x</span></span>||<span data-ttu-id="29b46-145">x</span><span class="sxs-lookup"><span data-stu-id="29b46-145">x</span></span>|
|[<span data-ttu-id="29b46-146">Rule</span><span class="sxs-lookup"><span data-stu-id="29b46-146">Rule</span></span>](rule.md)||<span data-ttu-id="29b46-147">x</span><span class="sxs-lookup"><span data-stu-id="29b46-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="29b46-148">可以包含</span><span class="sxs-lookup"><span data-stu-id="29b46-148">Can contain</span></span>

|<span data-ttu-id="29b46-149">**Element**</span><span class="sxs-lookup"><span data-stu-id="29b46-149">**Element**</span></span>|<span data-ttu-id="29b46-150">**Content**</span><span class="sxs-lookup"><span data-stu-id="29b46-150">**Content**</span></span>|<span data-ttu-id="29b46-151">**Mail**</span><span class="sxs-lookup"><span data-stu-id="29b46-151">**Mail**</span></span>|<span data-ttu-id="29b46-152">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="29b46-152">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="29b46-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="29b46-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="29b46-154">x</span><span class="sxs-lookup"><span data-stu-id="29b46-154">x</span></span>|<span data-ttu-id="29b46-155">x</span><span class="sxs-lookup"><span data-stu-id="29b46-155">x</span></span>|<span data-ttu-id="29b46-156">x</span><span class="sxs-lookup"><span data-stu-id="29b46-156">x</span></span>|
|[<span data-ttu-id="29b46-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="29b46-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="29b46-158">x</span><span class="sxs-lookup"><span data-stu-id="29b46-158">x</span></span>|<span data-ttu-id="29b46-159">x</span><span class="sxs-lookup"><span data-stu-id="29b46-159">x</span></span>|<span data-ttu-id="29b46-160">x</span><span class="sxs-lookup"><span data-stu-id="29b46-160">x</span></span>|
|[<span data-ttu-id="29b46-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="29b46-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="29b46-162">x</span><span class="sxs-lookup"><span data-stu-id="29b46-162">x</span></span>|<span data-ttu-id="29b46-163">x</span><span class="sxs-lookup"><span data-stu-id="29b46-163">x</span></span>|<span data-ttu-id="29b46-164">x</span><span class="sxs-lookup"><span data-stu-id="29b46-164">x</span></span>|
|[<span data-ttu-id="29b46-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="29b46-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="29b46-166">x</span><span class="sxs-lookup"><span data-stu-id="29b46-166">x</span></span>|<span data-ttu-id="29b46-167">x</span><span class="sxs-lookup"><span data-stu-id="29b46-167">x</span></span>|<span data-ttu-id="29b46-168">x</span><span class="sxs-lookup"><span data-stu-id="29b46-168">x</span></span>|
|[<span data-ttu-id="29b46-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="29b46-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="29b46-170">x</span><span class="sxs-lookup"><span data-stu-id="29b46-170">x</span></span>|<span data-ttu-id="29b46-171">x</span><span class="sxs-lookup"><span data-stu-id="29b46-171">x</span></span>|<span data-ttu-id="29b46-172">x</span><span class="sxs-lookup"><span data-stu-id="29b46-172">x</span></span>|
|[<span data-ttu-id="29b46-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="29b46-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="29b46-174">x</span><span class="sxs-lookup"><span data-stu-id="29b46-174">x</span></span>|<span data-ttu-id="29b46-175">x</span><span class="sxs-lookup"><span data-stu-id="29b46-175">x</span></span>|<span data-ttu-id="29b46-176">x</span><span class="sxs-lookup"><span data-stu-id="29b46-176">x</span></span>|
|[<span data-ttu-id="29b46-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="29b46-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="29b46-178">x</span><span class="sxs-lookup"><span data-stu-id="29b46-178">x</span></span>|<span data-ttu-id="29b46-179">x</span><span class="sxs-lookup"><span data-stu-id="29b46-179">x</span></span>|<span data-ttu-id="29b46-180">x</span><span class="sxs-lookup"><span data-stu-id="29b46-180">x</span></span>|
|[<span data-ttu-id="29b46-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="29b46-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="29b46-182">x</span><span class="sxs-lookup"><span data-stu-id="29b46-182">x</span></span>|||
|[<span data-ttu-id="29b46-183">Permissions</span><span class="sxs-lookup"><span data-stu-id="29b46-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="29b46-184">x</span><span class="sxs-lookup"><span data-stu-id="29b46-184">x</span></span>||
|[<span data-ttu-id="29b46-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="29b46-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="29b46-186">x</span><span class="sxs-lookup"><span data-stu-id="29b46-186">x</span></span>||
|[<span data-ttu-id="29b46-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="29b46-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="29b46-188">x</span><span class="sxs-lookup"><span data-stu-id="29b46-188">x</span></span>|
|[<span data-ttu-id="29b46-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="29b46-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="29b46-190">x</span><span class="sxs-lookup"><span data-stu-id="29b46-190">x</span></span>|<span data-ttu-id="29b46-191">x</span><span class="sxs-lookup"><span data-stu-id="29b46-191">x</span></span>|<span data-ttu-id="29b46-192">x</span><span class="sxs-lookup"><span data-stu-id="29b46-192">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="29b46-193">属性</span><span class="sxs-lookup"><span data-stu-id="29b46-193">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="29b46-194">xmlns</span><span class="sxs-lookup"><span data-stu-id="29b46-194">xmlns</span></span>|<span data-ttu-id="29b46-p101">定义的 Office 外接程序清单命名空间和架构版本。应始终将此属性设置为 `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span><span class="sxs-lookup"><span data-stu-id="29b46-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="29b46-197">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="29b46-197">xmlns:xsi</span></span>|<span data-ttu-id="29b46-p102">定义 XMLSchema 实例。应始终将此属性设置为 `"http://www.w3.org/2001/XMLSchema-instance"`</span><span class="sxs-lookup"><span data-stu-id="29b46-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="29b46-200">xsi:type</span><span class="sxs-lookup"><span data-stu-id="29b46-200">xsi:type</span></span>|<span data-ttu-id="29b46-p103">定义的 Office 外接程序的类型。应始终将此属性设置为下列值之一：`"ContentApp"`、`"MailApp"` 或 `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="29b46-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
