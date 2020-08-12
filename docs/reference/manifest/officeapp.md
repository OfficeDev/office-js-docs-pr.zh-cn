---
title: 清单文件中的 OfficeApp 元素
description: OfficeApp 元素是 Office 外接程序清单的根元素。
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 770c764db6d8d7d1d2e870e48437de7c8f887101
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641457"
---
# <a name="officeapp-element"></a><span data-ttu-id="033a6-103">OfficeApp 元素</span><span class="sxs-lookup"><span data-stu-id="033a6-103">OfficeApp element</span></span>

<span data-ttu-id="033a6-104">Office 外接程序清单中的根元素。</span><span class="sxs-lookup"><span data-stu-id="033a6-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="033a6-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="033a6-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="033a6-106">语法</span><span class="sxs-lookup"><span data-stu-id="033a6-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="033a6-107">包含于</span><span class="sxs-lookup"><span data-stu-id="033a6-107">Contained in</span></span>

 <span data-ttu-id="033a6-108">_none_</span><span class="sxs-lookup"><span data-stu-id="033a6-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="033a6-109">必须包含</span><span class="sxs-lookup"><span data-stu-id="033a6-109">Must contain</span></span>

|<span data-ttu-id="033a6-110">元素</span><span class="sxs-lookup"><span data-stu-id="033a6-110">Element</span></span>|<span data-ttu-id="033a6-111">内容</span><span class="sxs-lookup"><span data-stu-id="033a6-111">Content</span></span>|<span data-ttu-id="033a6-112">邮件</span><span class="sxs-lookup"><span data-stu-id="033a6-112">Mail</span></span>|<span data-ttu-id="033a6-113">任务窗格</span><span class="sxs-lookup"><span data-stu-id="033a6-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="033a6-114">Id</span><span class="sxs-lookup"><span data-stu-id="033a6-114">Id</span></span>](id.md)|<span data-ttu-id="033a6-115">x</span><span class="sxs-lookup"><span data-stu-id="033a6-115">x</span></span>|<span data-ttu-id="033a6-116">x</span><span class="sxs-lookup"><span data-stu-id="033a6-116">x</span></span>|<span data-ttu-id="033a6-117">x</span><span class="sxs-lookup"><span data-stu-id="033a6-117">x</span></span>|
|[<span data-ttu-id="033a6-118">版本</span><span class="sxs-lookup"><span data-stu-id="033a6-118">Version</span></span>](version.md)|<span data-ttu-id="033a6-119">x</span><span class="sxs-lookup"><span data-stu-id="033a6-119">x</span></span>|<span data-ttu-id="033a6-120">x</span><span class="sxs-lookup"><span data-stu-id="033a6-120">x</span></span>|<span data-ttu-id="033a6-121">x</span><span class="sxs-lookup"><span data-stu-id="033a6-121">x</span></span>|
|[<span data-ttu-id="033a6-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="033a6-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="033a6-123">x</span><span class="sxs-lookup"><span data-stu-id="033a6-123">x</span></span>|<span data-ttu-id="033a6-124">x</span><span class="sxs-lookup"><span data-stu-id="033a6-124">x</span></span>|<span data-ttu-id="033a6-125">x</span><span class="sxs-lookup"><span data-stu-id="033a6-125">x</span></span>|
|[<span data-ttu-id="033a6-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="033a6-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="033a6-127">x</span><span class="sxs-lookup"><span data-stu-id="033a6-127">x</span></span>|<span data-ttu-id="033a6-128">x</span><span class="sxs-lookup"><span data-stu-id="033a6-128">x</span></span>|<span data-ttu-id="033a6-129">x</span><span class="sxs-lookup"><span data-stu-id="033a6-129">x</span></span>|
|[<span data-ttu-id="033a6-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="033a6-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="033a6-131">x</span><span class="sxs-lookup"><span data-stu-id="033a6-131">x</span></span>||<span data-ttu-id="033a6-132">x</span><span class="sxs-lookup"><span data-stu-id="033a6-132">x</span></span>|
|[<span data-ttu-id="033a6-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="033a6-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="033a6-134">x</span><span class="sxs-lookup"><span data-stu-id="033a6-134">x</span></span>|<span data-ttu-id="033a6-135">x</span><span class="sxs-lookup"><span data-stu-id="033a6-135">x</span></span>|<span data-ttu-id="033a6-136">x</span><span class="sxs-lookup"><span data-stu-id="033a6-136">x</span></span>|
|[<span data-ttu-id="033a6-137">说明</span><span class="sxs-lookup"><span data-stu-id="033a6-137">Description</span></span>](description.md)|<span data-ttu-id="033a6-138">x</span><span class="sxs-lookup"><span data-stu-id="033a6-138">x</span></span>|<span data-ttu-id="033a6-139">x</span><span class="sxs-lookup"><span data-stu-id="033a6-139">x</span></span>|<span data-ttu-id="033a6-140">x</span><span class="sxs-lookup"><span data-stu-id="033a6-140">x</span></span>|
|[<span data-ttu-id="033a6-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="033a6-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="033a6-142">x</span><span class="sxs-lookup"><span data-stu-id="033a6-142">x</span></span>||
|[<span data-ttu-id="033a6-143">Permissions</span><span class="sxs-lookup"><span data-stu-id="033a6-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="033a6-144">x</span><span class="sxs-lookup"><span data-stu-id="033a6-144">x</span></span>||<span data-ttu-id="033a6-145">x</span><span class="sxs-lookup"><span data-stu-id="033a6-145">x</span></span>|
|[<span data-ttu-id="033a6-146">Rule</span><span class="sxs-lookup"><span data-stu-id="033a6-146">Rule</span></span>](rule.md)||<span data-ttu-id="033a6-147">x</span><span class="sxs-lookup"><span data-stu-id="033a6-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="033a6-148">可以包含</span><span class="sxs-lookup"><span data-stu-id="033a6-148">Can contain</span></span>

|<span data-ttu-id="033a6-149">元素</span><span class="sxs-lookup"><span data-stu-id="033a6-149">Element</span></span>|<span data-ttu-id="033a6-150">内容</span><span class="sxs-lookup"><span data-stu-id="033a6-150">Content</span></span>|<span data-ttu-id="033a6-151">邮件</span><span class="sxs-lookup"><span data-stu-id="033a6-151">Mail</span></span>|<span data-ttu-id="033a6-152">任务窗格</span><span class="sxs-lookup"><span data-stu-id="033a6-152">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="033a6-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="033a6-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="033a6-154">x</span><span class="sxs-lookup"><span data-stu-id="033a6-154">x</span></span>|<span data-ttu-id="033a6-155">x</span><span class="sxs-lookup"><span data-stu-id="033a6-155">x</span></span>|<span data-ttu-id="033a6-156">x</span><span class="sxs-lookup"><span data-stu-id="033a6-156">x</span></span>|
|[<span data-ttu-id="033a6-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="033a6-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="033a6-158">x</span><span class="sxs-lookup"><span data-stu-id="033a6-158">x</span></span>|<span data-ttu-id="033a6-159">x</span><span class="sxs-lookup"><span data-stu-id="033a6-159">x</span></span>|<span data-ttu-id="033a6-160">x</span><span class="sxs-lookup"><span data-stu-id="033a6-160">x</span></span>|
|[<span data-ttu-id="033a6-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="033a6-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="033a6-162">x</span><span class="sxs-lookup"><span data-stu-id="033a6-162">x</span></span>|<span data-ttu-id="033a6-163">x</span><span class="sxs-lookup"><span data-stu-id="033a6-163">x</span></span>|<span data-ttu-id="033a6-164">x</span><span class="sxs-lookup"><span data-stu-id="033a6-164">x</span></span>|
|[<span data-ttu-id="033a6-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="033a6-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="033a6-166">x</span><span class="sxs-lookup"><span data-stu-id="033a6-166">x</span></span>|<span data-ttu-id="033a6-167">x</span><span class="sxs-lookup"><span data-stu-id="033a6-167">x</span></span>|<span data-ttu-id="033a6-168">x</span><span class="sxs-lookup"><span data-stu-id="033a6-168">x</span></span>|
|[<span data-ttu-id="033a6-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="033a6-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="033a6-170">x</span><span class="sxs-lookup"><span data-stu-id="033a6-170">x</span></span>|<span data-ttu-id="033a6-171">x</span><span class="sxs-lookup"><span data-stu-id="033a6-171">x</span></span>|<span data-ttu-id="033a6-172">x</span><span class="sxs-lookup"><span data-stu-id="033a6-172">x</span></span>|
|[<span data-ttu-id="033a6-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="033a6-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="033a6-174">x</span><span class="sxs-lookup"><span data-stu-id="033a6-174">x</span></span>|<span data-ttu-id="033a6-175">x</span><span class="sxs-lookup"><span data-stu-id="033a6-175">x</span></span>|<span data-ttu-id="033a6-176">x</span><span class="sxs-lookup"><span data-stu-id="033a6-176">x</span></span>|
|[<span data-ttu-id="033a6-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="033a6-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="033a6-178">x</span><span class="sxs-lookup"><span data-stu-id="033a6-178">x</span></span>|<span data-ttu-id="033a6-179">x</span><span class="sxs-lookup"><span data-stu-id="033a6-179">x</span></span>|<span data-ttu-id="033a6-180">x</span><span class="sxs-lookup"><span data-stu-id="033a6-180">x</span></span>|
|[<span data-ttu-id="033a6-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="033a6-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="033a6-182">x</span><span class="sxs-lookup"><span data-stu-id="033a6-182">x</span></span>|||
|[<span data-ttu-id="033a6-183">Permissions</span><span class="sxs-lookup"><span data-stu-id="033a6-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="033a6-184">x</span><span class="sxs-lookup"><span data-stu-id="033a6-184">x</span></span>||
|[<span data-ttu-id="033a6-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="033a6-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="033a6-186">x</span><span class="sxs-lookup"><span data-stu-id="033a6-186">x</span></span>||
|[<span data-ttu-id="033a6-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="033a6-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="033a6-188">x</span><span class="sxs-lookup"><span data-stu-id="033a6-188">x</span></span>|
|[<span data-ttu-id="033a6-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="033a6-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="033a6-190">x</span><span class="sxs-lookup"><span data-stu-id="033a6-190">x</span></span>|<span data-ttu-id="033a6-191">x</span><span class="sxs-lookup"><span data-stu-id="033a6-191">x</span></span>|<span data-ttu-id="033a6-192">x</span><span class="sxs-lookup"><span data-stu-id="033a6-192">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="033a6-193">属性</span><span class="sxs-lookup"><span data-stu-id="033a6-193">Attributes</span></span>

|<span data-ttu-id="033a6-194">属性</span><span class="sxs-lookup"><span data-stu-id="033a6-194">Attribute</span></span>|<span data-ttu-id="033a6-195">说明</span><span class="sxs-lookup"><span data-stu-id="033a6-195">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="033a6-196">xmlns</span><span class="sxs-lookup"><span data-stu-id="033a6-196">xmlns</span></span>|<span data-ttu-id="033a6-p101">定义的 Office 外接程序清单命名空间和架构版本。应始终将此属性设置为 `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span><span class="sxs-lookup"><span data-stu-id="033a6-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="033a6-199">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="033a6-199">xmlns:xsi</span></span>|<span data-ttu-id="033a6-p102">定义 XMLSchema 实例。应始终将此属性设置为 `"http://www.w3.org/2001/XMLSchema-instance"`</span><span class="sxs-lookup"><span data-stu-id="033a6-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="033a6-202">xsi:type</span><span class="sxs-lookup"><span data-stu-id="033a6-202">xsi:type</span></span>|<span data-ttu-id="033a6-p103">定义的 Office 外接程序的类型。应始终将此属性设置为下列值之一：`"ContentApp"`、`"MailApp"` 或 `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="033a6-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
