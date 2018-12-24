---
title: 清单文件中的 OfficeApp 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 42b6fe2e1c33322b90016d5e7ceec7b1bfe5b72d
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433164"
---
# <a name="officeapp-element"></a><span data-ttu-id="296f4-102">OfficeApp 元素</span><span class="sxs-lookup"><span data-stu-id="296f4-102">OfficeApp element</span></span>

<span data-ttu-id="296f4-103">Office 外接程序清单中的根元素。</span><span class="sxs-lookup"><span data-stu-id="296f4-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="296f4-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="296f4-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="296f4-105">语法</span><span class="sxs-lookup"><span data-stu-id="296f4-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="296f4-106">包含于</span><span class="sxs-lookup"><span data-stu-id="296f4-106">Contained in</span></span>

 <span data-ttu-id="296f4-107">_none_</span><span class="sxs-lookup"><span data-stu-id="296f4-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="296f4-108">必须包含</span><span class="sxs-lookup"><span data-stu-id="296f4-108">Must contain</span></span>

|<span data-ttu-id="296f4-109">**元素**</span><span class="sxs-lookup"><span data-stu-id="296f4-109">**Element**</span></span>|<span data-ttu-id="296f4-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="296f4-110">**Content**</span></span>|<span data-ttu-id="296f4-111">**Mail**</span><span class="sxs-lookup"><span data-stu-id="296f4-111">**Mail**</span></span>|<span data-ttu-id="296f4-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="296f4-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="296f4-113">Id</span><span class="sxs-lookup"><span data-stu-id="296f4-113">Id</span></span>](id.md)|<span data-ttu-id="296f4-114">x</span><span class="sxs-lookup"><span data-stu-id="296f4-114">x</span></span>|<span data-ttu-id="296f4-115">x</span><span class="sxs-lookup"><span data-stu-id="296f4-115">x</span></span>|<span data-ttu-id="296f4-116">x</span><span class="sxs-lookup"><span data-stu-id="296f4-116">x</span></span>|
|[<span data-ttu-id="296f4-117">Version</span><span class="sxs-lookup"><span data-stu-id="296f4-117">Version</span></span>](version.md)|<span data-ttu-id="296f4-118">x</span><span class="sxs-lookup"><span data-stu-id="296f4-118">x</span></span>|<span data-ttu-id="296f4-119">x</span><span class="sxs-lookup"><span data-stu-id="296f4-119">x</span></span>|<span data-ttu-id="296f4-120">x</span><span class="sxs-lookup"><span data-stu-id="296f4-120">x</span></span>|
|[<span data-ttu-id="296f4-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="296f4-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="296f4-122">x</span><span class="sxs-lookup"><span data-stu-id="296f4-122">x</span></span>|<span data-ttu-id="296f4-123">x</span><span class="sxs-lookup"><span data-stu-id="296f4-123">x</span></span>|<span data-ttu-id="296f4-124">x</span><span class="sxs-lookup"><span data-stu-id="296f4-124">x</span></span>|
|[<span data-ttu-id="296f4-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="296f4-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="296f4-126">x</span><span class="sxs-lookup"><span data-stu-id="296f4-126">x</span></span>|<span data-ttu-id="296f4-127">x</span><span class="sxs-lookup"><span data-stu-id="296f4-127">x</span></span>|<span data-ttu-id="296f4-128">x</span><span class="sxs-lookup"><span data-stu-id="296f4-128">x</span></span>|
|[<span data-ttu-id="296f4-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="296f4-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="296f4-130">x</span><span class="sxs-lookup"><span data-stu-id="296f4-130">x</span></span>||<span data-ttu-id="296f4-131">x</span><span class="sxs-lookup"><span data-stu-id="296f4-131">x</span></span>|
|[<span data-ttu-id="296f4-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="296f4-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="296f4-133">x</span><span class="sxs-lookup"><span data-stu-id="296f4-133">x</span></span>|<span data-ttu-id="296f4-134">x</span><span class="sxs-lookup"><span data-stu-id="296f4-134">x</span></span>|<span data-ttu-id="296f4-135">x</span><span class="sxs-lookup"><span data-stu-id="296f4-135">x</span></span>|
|[<span data-ttu-id="296f4-136">说明</span><span class="sxs-lookup"><span data-stu-id="296f4-136">Description</span></span>](description.md)|<span data-ttu-id="296f4-137">x</span><span class="sxs-lookup"><span data-stu-id="296f4-137">x</span></span>|<span data-ttu-id="296f4-138">x</span><span class="sxs-lookup"><span data-stu-id="296f4-138">x</span></span>|<span data-ttu-id="296f4-139">x</span><span class="sxs-lookup"><span data-stu-id="296f4-139">x</span></span>|
|[<span data-ttu-id="296f4-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="296f4-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="296f4-141">x</span><span class="sxs-lookup"><span data-stu-id="296f4-141">x</span></span>||
|[<span data-ttu-id="296f4-142">Permissions</span><span class="sxs-lookup"><span data-stu-id="296f4-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="296f4-143">x</span><span class="sxs-lookup"><span data-stu-id="296f4-143">x</span></span>||<span data-ttu-id="296f4-144">x</span><span class="sxs-lookup"><span data-stu-id="296f4-144">x</span></span>|
|[<span data-ttu-id="296f4-145">Rule</span><span class="sxs-lookup"><span data-stu-id="296f4-145">Rule</span></span>](rule.md)||<span data-ttu-id="296f4-146">x</span><span class="sxs-lookup"><span data-stu-id="296f4-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="296f4-147">可以包含</span><span class="sxs-lookup"><span data-stu-id="296f4-147">Can contain</span></span>

|<span data-ttu-id="296f4-148">**元素**</span><span class="sxs-lookup"><span data-stu-id="296f4-148">**Element**</span></span>|<span data-ttu-id="296f4-149">**Content**</span><span class="sxs-lookup"><span data-stu-id="296f4-149">**Content**</span></span>|<span data-ttu-id="296f4-150">**Mail**</span><span class="sxs-lookup"><span data-stu-id="296f4-150">**Mail**</span></span>|<span data-ttu-id="296f4-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="296f4-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="296f4-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="296f4-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="296f4-153">x</span><span class="sxs-lookup"><span data-stu-id="296f4-153">x</span></span>|<span data-ttu-id="296f4-154">x</span><span class="sxs-lookup"><span data-stu-id="296f4-154">x</span></span>|<span data-ttu-id="296f4-155">x</span><span class="sxs-lookup"><span data-stu-id="296f4-155">x</span></span>|
|[<span data-ttu-id="296f4-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="296f4-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="296f4-157">x</span><span class="sxs-lookup"><span data-stu-id="296f4-157">x</span></span>|<span data-ttu-id="296f4-158">x</span><span class="sxs-lookup"><span data-stu-id="296f4-158">x</span></span>|<span data-ttu-id="296f4-159">x</span><span class="sxs-lookup"><span data-stu-id="296f4-159">x</span></span>|
|[<span data-ttu-id="296f4-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="296f4-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="296f4-161">x</span><span class="sxs-lookup"><span data-stu-id="296f4-161">x</span></span>|<span data-ttu-id="296f4-162">x</span><span class="sxs-lookup"><span data-stu-id="296f4-162">x</span></span>|<span data-ttu-id="296f4-163">x</span><span class="sxs-lookup"><span data-stu-id="296f4-163">x</span></span>|
|[<span data-ttu-id="296f4-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="296f4-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="296f4-165">x</span><span class="sxs-lookup"><span data-stu-id="296f4-165">x</span></span>|<span data-ttu-id="296f4-166">x</span><span class="sxs-lookup"><span data-stu-id="296f4-166">x</span></span>|<span data-ttu-id="296f4-167">x</span><span class="sxs-lookup"><span data-stu-id="296f4-167">x</span></span>|
|[<span data-ttu-id="296f4-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="296f4-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="296f4-169">x</span><span class="sxs-lookup"><span data-stu-id="296f4-169">x</span></span>|<span data-ttu-id="296f4-170">x</span><span class="sxs-lookup"><span data-stu-id="296f4-170">x</span></span>|<span data-ttu-id="296f4-171">x</span><span class="sxs-lookup"><span data-stu-id="296f4-171">x</span></span>|
|[<span data-ttu-id="296f4-172">Hosts</span><span class="sxs-lookup"><span data-stu-id="296f4-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="296f4-173">x</span><span class="sxs-lookup"><span data-stu-id="296f4-173">x</span></span>|<span data-ttu-id="296f4-174">x</span><span class="sxs-lookup"><span data-stu-id="296f4-174">x</span></span>|<span data-ttu-id="296f4-175">x</span><span class="sxs-lookup"><span data-stu-id="296f4-175">x</span></span>|
|[<span data-ttu-id="296f4-176">Requirements</span><span class="sxs-lookup"><span data-stu-id="296f4-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="296f4-177">x</span><span class="sxs-lookup"><span data-stu-id="296f4-177">x</span></span>|<span data-ttu-id="296f4-178">x</span><span class="sxs-lookup"><span data-stu-id="296f4-178">x</span></span>|<span data-ttu-id="296f4-179">x</span><span class="sxs-lookup"><span data-stu-id="296f4-179">x</span></span>|
|[<span data-ttu-id="296f4-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="296f4-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="296f4-181">x</span><span class="sxs-lookup"><span data-stu-id="296f4-181">x</span></span>|||
|[<span data-ttu-id="296f4-182">Permissions</span><span class="sxs-lookup"><span data-stu-id="296f4-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="296f4-183">x</span><span class="sxs-lookup"><span data-stu-id="296f4-183">x</span></span>||
|[<span data-ttu-id="296f4-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="296f4-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="296f4-185">x</span><span class="sxs-lookup"><span data-stu-id="296f4-185">x</span></span>||
|[<span data-ttu-id="296f4-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="296f4-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="296f4-187">x</span><span class="sxs-lookup"><span data-stu-id="296f4-187">x</span></span>|
|[<span data-ttu-id="296f4-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="296f4-188">VersionOverrides</span></span>](versionoverrides.md)||<span data-ttu-id="296f4-189">x</span><span class="sxs-lookup"><span data-stu-id="296f4-189">x</span></span>||

## <a name="attributes"></a><span data-ttu-id="296f4-190">属性</span><span class="sxs-lookup"><span data-stu-id="296f4-190">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="296f4-191">xmlns</span><span class="sxs-lookup"><span data-stu-id="296f4-191">xmlns</span></span>|<span data-ttu-id="296f4-p101">定义的 Office 外接程序清单命名空间和架构版本。应始终将此属性设置为 `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span><span class="sxs-lookup"><span data-stu-id="296f4-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="296f4-194">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="296f4-194">xmlns:xsi</span></span>|<span data-ttu-id="296f4-p102">定义 XMLSchema 实例。应始终将此属性设置为 `"http://www.w3.org/2001/XMLSchema-instance"`</span><span class="sxs-lookup"><span data-stu-id="296f4-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="296f4-197">xsi:type</span><span class="sxs-lookup"><span data-stu-id="296f4-197">xsi:type</span></span>|<span data-ttu-id="296f4-p103">定义的 Office 外接程序的类型。应始终将此属性设置为下列值之一：`"ContentApp"`、`"MailApp"` 或 `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="296f4-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
