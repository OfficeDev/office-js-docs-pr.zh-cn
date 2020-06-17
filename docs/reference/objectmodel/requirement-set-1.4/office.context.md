---
title: Office。上下文要求集1。4
description: 使用邮箱 API 要求集1.4 的 Outlook 外接程序可用的 Office 对象成员。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 5d15e3bd70d95693eecab2d124ef43e6bf492446
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609920"
---
# <a name="context-mailbox-requirement-set-14"></a><span data-ttu-id="ae465-103">context （邮箱要求集1.4）</span><span class="sxs-lookup"><span data-stu-id="ae465-103">context (Mailbox requirement set 1.4)</span></span>

### <a name="officecontext"></a><span data-ttu-id="ae465-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="ae465-104">[Office](office.md).context</span></span>

<span data-ttu-id="ae465-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="ae465-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="ae465-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅[通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.4)"。</span><span class="sxs-lookup"><span data-stu-id="ae465-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.4).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae465-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae465-107">Requirements</span></span>

|<span data-ttu-id="ae465-108">要求</span><span class="sxs-lookup"><span data-stu-id="ae465-108">Requirement</span></span>| <span data-ttu-id="ae465-109">值</span><span class="sxs-lookup"><span data-stu-id="ae465-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae465-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae465-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae465-111">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-111">1.1</span></span>|
|[<span data-ttu-id="ae465-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae465-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae465-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae465-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ae465-114">属性</span><span class="sxs-lookup"><span data-stu-id="ae465-114">Properties</span></span>

| <span data-ttu-id="ae465-115">属性</span><span class="sxs-lookup"><span data-stu-id="ae465-115">Property</span></span> | <span data-ttu-id="ae465-116">型号</span><span class="sxs-lookup"><span data-stu-id="ae465-116">Modes</span></span> | <span data-ttu-id="ae465-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="ae465-117">Return type</span></span> | <span data-ttu-id="ae465-118">最低</span><span class="sxs-lookup"><span data-stu-id="ae465-118">Minimum</span></span><br><span data-ttu-id="ae465-119">要求集</span><span class="sxs-lookup"><span data-stu-id="ae465-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ae465-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="ae465-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="ae465-121">撰写</span><span class="sxs-lookup"><span data-stu-id="ae465-121">Compose</span></span><br><span data-ttu-id="ae465-122">Read</span><span class="sxs-lookup"><span data-stu-id="ae465-122">Read</span></span> | <span data-ttu-id="ae465-123">String</span><span class="sxs-lookup"><span data-stu-id="ae465-123">String</span></span> | [<span data-ttu-id="ae465-124">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ae465-125">过程</span><span class="sxs-lookup"><span data-stu-id="ae465-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="ae465-126">撰写</span><span class="sxs-lookup"><span data-stu-id="ae465-126">Compose</span></span><br><span data-ttu-id="ae465-127">Read</span><span class="sxs-lookup"><span data-stu-id="ae465-127">Read</span></span> | [<span data-ttu-id="ae465-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ae465-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.4) | [<span data-ttu-id="ae465-129">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ae465-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="ae465-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="ae465-131">撰写</span><span class="sxs-lookup"><span data-stu-id="ae465-131">Compose</span></span><br><span data-ttu-id="ae465-132">Read</span><span class="sxs-lookup"><span data-stu-id="ae465-132">Read</span></span> | <span data-ttu-id="ae465-133">String</span><span class="sxs-lookup"><span data-stu-id="ae465-133">String</span></span> | [<span data-ttu-id="ae465-134">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ae465-135">host</span><span class="sxs-lookup"><span data-stu-id="ae465-135">host</span></span>](#host-hosttype) | <span data-ttu-id="ae465-136">撰写</span><span class="sxs-lookup"><span data-stu-id="ae465-136">Compose</span></span><br><span data-ttu-id="ae465-137">Read</span><span class="sxs-lookup"><span data-stu-id="ae465-137">Read</span></span> | [<span data-ttu-id="ae465-138">HostType</span><span class="sxs-lookup"><span data-stu-id="ae465-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.4) | [<span data-ttu-id="ae465-139">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ae465-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="ae465-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="ae465-141">撰写</span><span class="sxs-lookup"><span data-stu-id="ae465-141">Compose</span></span><br><span data-ttu-id="ae465-142">Read</span><span class="sxs-lookup"><span data-stu-id="ae465-142">Read</span></span> | [<span data-ttu-id="ae465-143">邮箱</span><span class="sxs-lookup"><span data-stu-id="ae465-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4) | [<span data-ttu-id="ae465-144">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ae465-145">平台</span><span class="sxs-lookup"><span data-stu-id="ae465-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="ae465-146">撰写</span><span class="sxs-lookup"><span data-stu-id="ae465-146">Compose</span></span><br><span data-ttu-id="ae465-147">Read</span><span class="sxs-lookup"><span data-stu-id="ae465-147">Read</span></span> | [<span data-ttu-id="ae465-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ae465-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.4) | [<span data-ttu-id="ae465-149">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ae465-150">满足</span><span class="sxs-lookup"><span data-stu-id="ae465-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="ae465-151">撰写</span><span class="sxs-lookup"><span data-stu-id="ae465-151">Compose</span></span><br><span data-ttu-id="ae465-152">Read</span><span class="sxs-lookup"><span data-stu-id="ae465-152">Read</span></span> | [<span data-ttu-id="ae465-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ae465-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4) | [<span data-ttu-id="ae465-154">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ae465-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="ae465-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="ae465-156">撰写</span><span class="sxs-lookup"><span data-stu-id="ae465-156">Compose</span></span><br><span data-ttu-id="ae465-157">Read</span><span class="sxs-lookup"><span data-stu-id="ae465-157">Read</span></span> | [<span data-ttu-id="ae465-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ae465-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4) | [<span data-ttu-id="ae465-159">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ae465-160">ui</span><span class="sxs-lookup"><span data-stu-id="ae465-160">ui</span></span>](#ui-ui) | <span data-ttu-id="ae465-161">撰写</span><span class="sxs-lookup"><span data-stu-id="ae465-161">Compose</span></span><br><span data-ttu-id="ae465-162">Read</span><span class="sxs-lookup"><span data-stu-id="ae465-162">Read</span></span> | [<span data-ttu-id="ae465-163">UI</span><span class="sxs-lookup"><span data-stu-id="ae465-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.4) | [<span data-ttu-id="ae465-164">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="ae465-165">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="ae465-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="ae465-166">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="ae465-166">contentLanguage: String</span></span>

<span data-ttu-id="ae465-167">获取用户指定的用于编辑项的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="ae465-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="ae465-168">此 `contentLanguage` 值反映了在 Office 主机应用程序中使用**File > Options > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="ae465-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="ae465-169">类型</span><span class="sxs-lookup"><span data-stu-id="ae465-169">Type</span></span>

*   <span data-ttu-id="ae465-170">String</span><span class="sxs-lookup"><span data-stu-id="ae465-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae465-171">要求</span><span class="sxs-lookup"><span data-stu-id="ae465-171">Requirements</span></span>

|<span data-ttu-id="ae465-172">要求</span><span class="sxs-lookup"><span data-stu-id="ae465-172">Requirement</span></span>| <span data-ttu-id="ae465-173">值</span><span class="sxs-lookup"><span data-stu-id="ae465-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae465-174">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae465-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae465-175">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-175">1.1</span></span>|
|[<span data-ttu-id="ae465-176">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae465-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae465-177">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae465-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae465-178">示例</span><span class="sxs-lookup"><span data-stu-id="ae465-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="ae465-179">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="ae465-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="ae465-180">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="ae465-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ae465-181">类型</span><span class="sxs-lookup"><span data-stu-id="ae465-181">Type</span></span>

*   [<span data-ttu-id="ae465-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ae465-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="ae465-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae465-183">Requirements</span></span>

|<span data-ttu-id="ae465-184">要求</span><span class="sxs-lookup"><span data-stu-id="ae465-184">Requirement</span></span>| <span data-ttu-id="ae465-185">值</span><span class="sxs-lookup"><span data-stu-id="ae465-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae465-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae465-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae465-187">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-187">1.1</span></span>|
|[<span data-ttu-id="ae465-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae465-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae465-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae465-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae465-190">示例</span><span class="sxs-lookup"><span data-stu-id="ae465-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="ae465-191">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="ae465-191">displayLanguage: String</span></span>

<span data-ttu-id="ae465-192">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="ae465-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="ae465-193">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="ae465-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="ae465-194">类型</span><span class="sxs-lookup"><span data-stu-id="ae465-194">Type</span></span>

*   <span data-ttu-id="ae465-195">String</span><span class="sxs-lookup"><span data-stu-id="ae465-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae465-196">要求</span><span class="sxs-lookup"><span data-stu-id="ae465-196">Requirements</span></span>

|<span data-ttu-id="ae465-197">要求</span><span class="sxs-lookup"><span data-stu-id="ae465-197">Requirement</span></span>| <span data-ttu-id="ae465-198">值</span><span class="sxs-lookup"><span data-stu-id="ae465-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae465-199">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae465-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae465-200">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-200">1.1</span></span>|
|[<span data-ttu-id="ae465-201">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae465-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae465-202">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae465-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae465-203">示例</span><span class="sxs-lookup"><span data-stu-id="ae465-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="ae465-204">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="ae465-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="ae465-205">获取运行外接程序的 Office 应用程序主机。</span><span class="sxs-lookup"><span data-stu-id="ae465-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ae465-206">类型</span><span class="sxs-lookup"><span data-stu-id="ae465-206">Type</span></span>

*   [<span data-ttu-id="ae465-207">HostType</span><span class="sxs-lookup"><span data-stu-id="ae465-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="ae465-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae465-208">Requirements</span></span>

|<span data-ttu-id="ae465-209">要求</span><span class="sxs-lookup"><span data-stu-id="ae465-209">Requirement</span></span>| <span data-ttu-id="ae465-210">值</span><span class="sxs-lookup"><span data-stu-id="ae465-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae465-211">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae465-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae465-212">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-212">1.1</span></span>|
|[<span data-ttu-id="ae465-213">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae465-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae465-214">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae465-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae465-215">示例</span><span class="sxs-lookup"><span data-stu-id="ae465-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="ae465-216">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="ae465-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="ae465-217">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="ae465-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ae465-218">类型</span><span class="sxs-lookup"><span data-stu-id="ae465-218">Type</span></span>

*   [<span data-ttu-id="ae465-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ae465-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="ae465-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae465-220">Requirements</span></span>

|<span data-ttu-id="ae465-221">要求</span><span class="sxs-lookup"><span data-stu-id="ae465-221">Requirement</span></span>| <span data-ttu-id="ae465-222">值</span><span class="sxs-lookup"><span data-stu-id="ae465-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae465-223">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae465-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae465-224">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-224">1.1</span></span>|
|[<span data-ttu-id="ae465-225">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae465-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae465-226">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae465-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae465-227">示例</span><span class="sxs-lookup"><span data-stu-id="ae465-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="ae465-228">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="ae465-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="ae465-229">提供用于确定当前主机和平台上支持的要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="ae465-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="ae465-230">类型</span><span class="sxs-lookup"><span data-stu-id="ae465-230">Type</span></span>

*   [<span data-ttu-id="ae465-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ae465-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="ae465-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae465-232">Requirements</span></span>

|<span data-ttu-id="ae465-233">要求</span><span class="sxs-lookup"><span data-stu-id="ae465-233">Requirement</span></span>| <span data-ttu-id="ae465-234">值</span><span class="sxs-lookup"><span data-stu-id="ae465-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae465-235">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae465-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae465-236">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-236">1.1</span></span>|
|[<span data-ttu-id="ae465-237">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae465-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae465-238">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae465-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae465-239">示例</span><span class="sxs-lookup"><span data-stu-id="ae465-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="ae465-240">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="ae465-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="ae465-241">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="ae465-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="ae465-242">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="ae465-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="ae465-243">类型</span><span class="sxs-lookup"><span data-stu-id="ae465-243">Type</span></span>

*   [<span data-ttu-id="ae465-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ae465-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="ae465-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae465-245">Requirements</span></span>

|<span data-ttu-id="ae465-246">要求</span><span class="sxs-lookup"><span data-stu-id="ae465-246">Requirement</span></span>| <span data-ttu-id="ae465-247">值</span><span class="sxs-lookup"><span data-stu-id="ae465-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae465-248">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae465-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae465-249">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-249">1.1</span></span>|
|[<span data-ttu-id="ae465-250">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae465-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="ae465-251">受限</span><span class="sxs-lookup"><span data-stu-id="ae465-251">Restricted</span></span>|
|[<span data-ttu-id="ae465-252">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae465-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae465-253">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae465-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="ae465-254">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="ae465-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="ae465-255">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="ae465-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ae465-256">类型</span><span class="sxs-lookup"><span data-stu-id="ae465-256">Type</span></span>

*   [<span data-ttu-id="ae465-257">UI</span><span class="sxs-lookup"><span data-stu-id="ae465-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="ae465-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae465-258">Requirements</span></span>

|<span data-ttu-id="ae465-259">要求</span><span class="sxs-lookup"><span data-stu-id="ae465-259">Requirement</span></span>| <span data-ttu-id="ae465-260">值</span><span class="sxs-lookup"><span data-stu-id="ae465-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae465-261">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae465-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae465-262">1.1</span><span class="sxs-lookup"><span data-stu-id="ae465-262">1.1</span></span>|
|[<span data-ttu-id="ae465-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae465-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae465-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae465-264">Compose or Read</span></span>|
