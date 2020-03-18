---
title: Office。上下文要求集1。3
description: Outlook 外接程序 API 中的 Outlook 上下文对象的对象模型（邮箱 API 1.3 版本）。
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: dc524c5657380845744e241a600fbf66f1eed5ce
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720118"
---
# <a name="context"></a><span data-ttu-id="3fe17-103">context</span><span class="sxs-lookup"><span data-stu-id="3fe17-103">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="3fe17-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="3fe17-104">[Office](office.md).context</span></span>

<span data-ttu-id="3fe17-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="3fe17-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="3fe17-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅[通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.3)"。</span><span class="sxs-lookup"><span data-stu-id="3fe17-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.3).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3fe17-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="3fe17-107">Requirements</span></span>

|<span data-ttu-id="3fe17-108">要求</span><span class="sxs-lookup"><span data-stu-id="3fe17-108">Requirement</span></span>| <span data-ttu-id="3fe17-109">值</span><span class="sxs-lookup"><span data-stu-id="3fe17-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fe17-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3fe17-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fe17-111">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-111">1.1</span></span>|
|[<span data-ttu-id="3fe17-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3fe17-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fe17-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3fe17-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="3fe17-114">属性</span><span class="sxs-lookup"><span data-stu-id="3fe17-114">Properties</span></span>

| <span data-ttu-id="3fe17-115">属性</span><span class="sxs-lookup"><span data-stu-id="3fe17-115">Property</span></span> | <span data-ttu-id="3fe17-116">型号</span><span class="sxs-lookup"><span data-stu-id="3fe17-116">Modes</span></span> | <span data-ttu-id="3fe17-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="3fe17-117">Return type</span></span> | <span data-ttu-id="3fe17-118">最低</span><span class="sxs-lookup"><span data-stu-id="3fe17-118">Minimum</span></span><br><span data-ttu-id="3fe17-119">要求集</span><span class="sxs-lookup"><span data-stu-id="3fe17-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3fe17-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="3fe17-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="3fe17-121">撰写</span><span class="sxs-lookup"><span data-stu-id="3fe17-121">Compose</span></span><br><span data-ttu-id="3fe17-122">读取</span><span class="sxs-lookup"><span data-stu-id="3fe17-122">Read</span></span> | <span data-ttu-id="3fe17-123">String</span><span class="sxs-lookup"><span data-stu-id="3fe17-123">String</span></span> | [<span data-ttu-id="3fe17-124">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fe17-125">过程</span><span class="sxs-lookup"><span data-stu-id="3fe17-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="3fe17-126">撰写</span><span class="sxs-lookup"><span data-stu-id="3fe17-126">Compose</span></span><br><span data-ttu-id="3fe17-127">读取</span><span class="sxs-lookup"><span data-stu-id="3fe17-127">Read</span></span> | [<span data-ttu-id="3fe17-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="3fe17-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.3) | [<span data-ttu-id="3fe17-129">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fe17-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="3fe17-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="3fe17-131">撰写</span><span class="sxs-lookup"><span data-stu-id="3fe17-131">Compose</span></span><br><span data-ttu-id="3fe17-132">读取</span><span class="sxs-lookup"><span data-stu-id="3fe17-132">Read</span></span> | <span data-ttu-id="3fe17-133">String</span><span class="sxs-lookup"><span data-stu-id="3fe17-133">String</span></span> | [<span data-ttu-id="3fe17-134">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fe17-135">host</span><span class="sxs-lookup"><span data-stu-id="3fe17-135">host</span></span>](#host-hosttype) | <span data-ttu-id="3fe17-136">撰写</span><span class="sxs-lookup"><span data-stu-id="3fe17-136">Compose</span></span><br><span data-ttu-id="3fe17-137">读取</span><span class="sxs-lookup"><span data-stu-id="3fe17-137">Read</span></span> | [<span data-ttu-id="3fe17-138">HostType</span><span class="sxs-lookup"><span data-stu-id="3fe17-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.3) | [<span data-ttu-id="3fe17-139">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fe17-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="3fe17-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="3fe17-141">撰写</span><span class="sxs-lookup"><span data-stu-id="3fe17-141">Compose</span></span><br><span data-ttu-id="3fe17-142">读取</span><span class="sxs-lookup"><span data-stu-id="3fe17-142">Read</span></span> | [<span data-ttu-id="3fe17-143">邮箱</span><span class="sxs-lookup"><span data-stu-id="3fe17-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3) | [<span data-ttu-id="3fe17-144">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fe17-145">平台</span><span class="sxs-lookup"><span data-stu-id="3fe17-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="3fe17-146">撰写</span><span class="sxs-lookup"><span data-stu-id="3fe17-146">Compose</span></span><br><span data-ttu-id="3fe17-147">读取</span><span class="sxs-lookup"><span data-stu-id="3fe17-147">Read</span></span> | [<span data-ttu-id="3fe17-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="3fe17-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.3) | [<span data-ttu-id="3fe17-149">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fe17-150">满足</span><span class="sxs-lookup"><span data-stu-id="3fe17-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="3fe17-151">撰写</span><span class="sxs-lookup"><span data-stu-id="3fe17-151">Compose</span></span><br><span data-ttu-id="3fe17-152">读取</span><span class="sxs-lookup"><span data-stu-id="3fe17-152">Read</span></span> | [<span data-ttu-id="3fe17-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="3fe17-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.3) | [<span data-ttu-id="3fe17-154">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fe17-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="3fe17-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="3fe17-156">撰写</span><span class="sxs-lookup"><span data-stu-id="3fe17-156">Compose</span></span><br><span data-ttu-id="3fe17-157">读取</span><span class="sxs-lookup"><span data-stu-id="3fe17-157">Read</span></span> | [<span data-ttu-id="3fe17-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="3fe17-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.3) | [<span data-ttu-id="3fe17-159">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fe17-160">ui</span><span class="sxs-lookup"><span data-stu-id="3fe17-160">ui</span></span>](#ui-ui) | <span data-ttu-id="3fe17-161">撰写</span><span class="sxs-lookup"><span data-stu-id="3fe17-161">Compose</span></span><br><span data-ttu-id="3fe17-162">读取</span><span class="sxs-lookup"><span data-stu-id="3fe17-162">Read</span></span> | [<span data-ttu-id="3fe17-163">UI</span><span class="sxs-lookup"><span data-stu-id="3fe17-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.3) | [<span data-ttu-id="3fe17-164">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="3fe17-165">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="3fe17-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="3fe17-166">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="3fe17-166">contentLanguage: String</span></span>

<span data-ttu-id="3fe17-167">获取用户指定的用于编辑项的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="3fe17-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="3fe17-168">此`contentLanguage`值反映了在 Office 主机应用程序中使用**File > Options > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="3fe17-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="3fe17-169">类型</span><span class="sxs-lookup"><span data-stu-id="3fe17-169">Type</span></span>

*   <span data-ttu-id="3fe17-170">String</span><span class="sxs-lookup"><span data-stu-id="3fe17-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3fe17-171">要求</span><span class="sxs-lookup"><span data-stu-id="3fe17-171">Requirements</span></span>

|<span data-ttu-id="3fe17-172">要求</span><span class="sxs-lookup"><span data-stu-id="3fe17-172">Requirement</span></span>| <span data-ttu-id="3fe17-173">值</span><span class="sxs-lookup"><span data-stu-id="3fe17-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fe17-174">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3fe17-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fe17-175">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-175">1.1</span></span>|
|[<span data-ttu-id="3fe17-176">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3fe17-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fe17-177">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3fe17-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3fe17-178">示例</span><span class="sxs-lookup"><span data-stu-id="3fe17-178">Example</span></span>

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="3fe17-179">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="3fe17-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="3fe17-180">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="3fe17-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="3fe17-181">类型</span><span class="sxs-lookup"><span data-stu-id="3fe17-181">Type</span></span>

*   [<span data-ttu-id="3fe17-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="3fe17-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="3fe17-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="3fe17-183">Requirements</span></span>

|<span data-ttu-id="3fe17-184">要求</span><span class="sxs-lookup"><span data-stu-id="3fe17-184">Requirement</span></span>| <span data-ttu-id="3fe17-185">值</span><span class="sxs-lookup"><span data-stu-id="3fe17-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fe17-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3fe17-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fe17-187">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-187">1.1</span></span>|
|[<span data-ttu-id="3fe17-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3fe17-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fe17-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3fe17-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3fe17-190">示例</span><span class="sxs-lookup"><span data-stu-id="3fe17-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="3fe17-191">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="3fe17-191">displayLanguage: String</span></span>

<span data-ttu-id="3fe17-192">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="3fe17-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="3fe17-193">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="3fe17-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="3fe17-194">类型</span><span class="sxs-lookup"><span data-stu-id="3fe17-194">Type</span></span>

*   <span data-ttu-id="3fe17-195">String</span><span class="sxs-lookup"><span data-stu-id="3fe17-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3fe17-196">要求</span><span class="sxs-lookup"><span data-stu-id="3fe17-196">Requirements</span></span>

|<span data-ttu-id="3fe17-197">要求</span><span class="sxs-lookup"><span data-stu-id="3fe17-197">Requirement</span></span>| <span data-ttu-id="3fe17-198">值</span><span class="sxs-lookup"><span data-stu-id="3fe17-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fe17-199">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3fe17-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fe17-200">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-200">1.1</span></span>|
|[<span data-ttu-id="3fe17-201">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3fe17-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fe17-202">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3fe17-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3fe17-203">示例</span><span class="sxs-lookup"><span data-stu-id="3fe17-203">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="host-hosttype"></a><span data-ttu-id="3fe17-204">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="3fe17-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="3fe17-205">获取运行外接程序的 Office 应用程序主机。</span><span class="sxs-lookup"><span data-stu-id="3fe17-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="3fe17-206">类型</span><span class="sxs-lookup"><span data-stu-id="3fe17-206">Type</span></span>

*   [<span data-ttu-id="3fe17-207">HostType</span><span class="sxs-lookup"><span data-stu-id="3fe17-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="3fe17-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="3fe17-208">Requirements</span></span>

|<span data-ttu-id="3fe17-209">要求</span><span class="sxs-lookup"><span data-stu-id="3fe17-209">Requirement</span></span>| <span data-ttu-id="3fe17-210">值</span><span class="sxs-lookup"><span data-stu-id="3fe17-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fe17-211">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3fe17-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fe17-212">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-212">1.1</span></span>|
|[<span data-ttu-id="3fe17-213">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3fe17-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fe17-214">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3fe17-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3fe17-215">示例</span><span class="sxs-lookup"><span data-stu-id="3fe17-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="3fe17-216">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="3fe17-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="3fe17-217">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="3fe17-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="3fe17-218">类型</span><span class="sxs-lookup"><span data-stu-id="3fe17-218">Type</span></span>

*   [<span data-ttu-id="3fe17-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="3fe17-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="3fe17-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="3fe17-220">Requirements</span></span>

|<span data-ttu-id="3fe17-221">要求</span><span class="sxs-lookup"><span data-stu-id="3fe17-221">Requirement</span></span>| <span data-ttu-id="3fe17-222">值</span><span class="sxs-lookup"><span data-stu-id="3fe17-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fe17-223">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3fe17-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fe17-224">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-224">1.1</span></span>|
|[<span data-ttu-id="3fe17-225">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3fe17-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fe17-226">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3fe17-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3fe17-227">示例</span><span class="sxs-lookup"><span data-stu-id="3fe17-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="3fe17-228">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="3fe17-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="3fe17-229">提供用于确定当前主机和平台上支持的要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="3fe17-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="3fe17-230">类型</span><span class="sxs-lookup"><span data-stu-id="3fe17-230">Type</span></span>

*   [<span data-ttu-id="3fe17-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="3fe17-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="3fe17-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="3fe17-232">Requirements</span></span>

|<span data-ttu-id="3fe17-233">要求</span><span class="sxs-lookup"><span data-stu-id="3fe17-233">Requirement</span></span>| <span data-ttu-id="3fe17-234">值</span><span class="sxs-lookup"><span data-stu-id="3fe17-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fe17-235">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3fe17-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fe17-236">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-236">1.1</span></span>|
|[<span data-ttu-id="3fe17-237">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3fe17-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fe17-238">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3fe17-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3fe17-239">示例</span><span class="sxs-lookup"><span data-stu-id="3fe17-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="3fe17-240">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="3fe17-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="3fe17-241">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="3fe17-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="3fe17-242">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="3fe17-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="3fe17-243">类型</span><span class="sxs-lookup"><span data-stu-id="3fe17-243">Type</span></span>

*   [<span data-ttu-id="3fe17-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="3fe17-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="3fe17-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="3fe17-245">Requirements</span></span>

|<span data-ttu-id="3fe17-246">要求</span><span class="sxs-lookup"><span data-stu-id="3fe17-246">Requirement</span></span>| <span data-ttu-id="3fe17-247">值</span><span class="sxs-lookup"><span data-stu-id="3fe17-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fe17-248">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3fe17-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fe17-249">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-249">1.1</span></span>|
|[<span data-ttu-id="3fe17-250">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3fe17-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="3fe17-251">受限</span><span class="sxs-lookup"><span data-stu-id="3fe17-251">Restricted</span></span>|
|[<span data-ttu-id="3fe17-252">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3fe17-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fe17-253">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3fe17-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="3fe17-254">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="3fe17-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="3fe17-255">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="3fe17-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="3fe17-256">类型</span><span class="sxs-lookup"><span data-stu-id="3fe17-256">Type</span></span>

*   [<span data-ttu-id="3fe17-257">UI</span><span class="sxs-lookup"><span data-stu-id="3fe17-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="3fe17-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="3fe17-258">Requirements</span></span>

|<span data-ttu-id="3fe17-259">要求</span><span class="sxs-lookup"><span data-stu-id="3fe17-259">Requirement</span></span>| <span data-ttu-id="3fe17-260">值</span><span class="sxs-lookup"><span data-stu-id="3fe17-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fe17-261">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3fe17-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fe17-262">1.1</span><span class="sxs-lookup"><span data-stu-id="3fe17-262">1.1</span></span>|
|[<span data-ttu-id="3fe17-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3fe17-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fe17-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3fe17-264">Compose or Read</span></span>|
