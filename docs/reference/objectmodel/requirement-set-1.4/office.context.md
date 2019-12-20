---
title: Office。上下文要求集1。4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 160a10cbcfa4ce1e2bac6c37d961a17ef14024f4
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814869"
---
# <a name="context"></a><span data-ttu-id="867f3-102">context</span><span class="sxs-lookup"><span data-stu-id="867f3-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="867f3-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="867f3-103">[Office](office.md).context</span></span>

<span data-ttu-id="867f3-104">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="867f3-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="867f3-105">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅[通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.4)"。</span><span class="sxs-lookup"><span data-stu-id="867f3-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.4).</span></span>

##### <a name="requirements"></a><span data-ttu-id="867f3-106">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-106">Requirements</span></span>

|<span data-ttu-id="867f3-107">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-107">Requirement</span></span>| <span data-ttu-id="867f3-108">值</span><span class="sxs-lookup"><span data-stu-id="867f3-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="867f3-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="867f3-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="867f3-110">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-110">1.1</span></span>|
|[<span data-ttu-id="867f3-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="867f3-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="867f3-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="867f3-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="867f3-113">属性</span><span class="sxs-lookup"><span data-stu-id="867f3-113">Properties</span></span>

| <span data-ttu-id="867f3-114">属性</span><span class="sxs-lookup"><span data-stu-id="867f3-114">Property</span></span> | <span data-ttu-id="867f3-115">型号</span><span class="sxs-lookup"><span data-stu-id="867f3-115">Modes</span></span> | <span data-ttu-id="867f3-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="867f3-116">Return type</span></span> | <span data-ttu-id="867f3-117">最低</span><span class="sxs-lookup"><span data-stu-id="867f3-117">Minimum</span></span><br><span data-ttu-id="867f3-118">要求集</span><span class="sxs-lookup"><span data-stu-id="867f3-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="867f3-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="867f3-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="867f3-120">撰写</span><span class="sxs-lookup"><span data-stu-id="867f3-120">Compose</span></span><br><span data-ttu-id="867f3-121">读取</span><span class="sxs-lookup"><span data-stu-id="867f3-121">Read</span></span> | <span data-ttu-id="867f3-122">String</span><span class="sxs-lookup"><span data-stu-id="867f3-122">String</span></span> | [<span data-ttu-id="867f3-123">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="867f3-124">过程</span><span class="sxs-lookup"><span data-stu-id="867f3-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="867f3-125">撰写</span><span class="sxs-lookup"><span data-stu-id="867f3-125">Compose</span></span><br><span data-ttu-id="867f3-126">读取</span><span class="sxs-lookup"><span data-stu-id="867f3-126">Read</span></span> | [<span data-ttu-id="867f3-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="867f3-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.4) | [<span data-ttu-id="867f3-128">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="867f3-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="867f3-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="867f3-130">撰写</span><span class="sxs-lookup"><span data-stu-id="867f3-130">Compose</span></span><br><span data-ttu-id="867f3-131">读取</span><span class="sxs-lookup"><span data-stu-id="867f3-131">Read</span></span> | <span data-ttu-id="867f3-132">String</span><span class="sxs-lookup"><span data-stu-id="867f3-132">String</span></span> | [<span data-ttu-id="867f3-133">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="867f3-134">host</span><span class="sxs-lookup"><span data-stu-id="867f3-134">host</span></span>](#host-hosttype) | <span data-ttu-id="867f3-135">撰写</span><span class="sxs-lookup"><span data-stu-id="867f3-135">Compose</span></span><br><span data-ttu-id="867f3-136">读取</span><span class="sxs-lookup"><span data-stu-id="867f3-136">Read</span></span> | [<span data-ttu-id="867f3-137">HostType</span><span class="sxs-lookup"><span data-stu-id="867f3-137">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.4) | [<span data-ttu-id="867f3-138">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="867f3-139">mailbox</span><span class="sxs-lookup"><span data-stu-id="867f3-139">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="867f3-140">撰写</span><span class="sxs-lookup"><span data-stu-id="867f3-140">Compose</span></span><br><span data-ttu-id="867f3-141">读取</span><span class="sxs-lookup"><span data-stu-id="867f3-141">Read</span></span> | [<span data-ttu-id="867f3-142">邮箱</span><span class="sxs-lookup"><span data-stu-id="867f3-142">Mailbox</span></span>](/javascript/api/office/office.mailbox?view=outlook-js-1.4) | [<span data-ttu-id="867f3-143">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="867f3-144">平台</span><span class="sxs-lookup"><span data-stu-id="867f3-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="867f3-145">撰写</span><span class="sxs-lookup"><span data-stu-id="867f3-145">Compose</span></span><br><span data-ttu-id="867f3-146">读取</span><span class="sxs-lookup"><span data-stu-id="867f3-146">Read</span></span> | [<span data-ttu-id="867f3-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="867f3-147">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.4) | [<span data-ttu-id="867f3-148">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="867f3-149">满足</span><span class="sxs-lookup"><span data-stu-id="867f3-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="867f3-150">撰写</span><span class="sxs-lookup"><span data-stu-id="867f3-150">Compose</span></span><br><span data-ttu-id="867f3-151">读取</span><span class="sxs-lookup"><span data-stu-id="867f3-151">Read</span></span> | [<span data-ttu-id="867f3-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="867f3-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4) | [<span data-ttu-id="867f3-153">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-153">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="867f3-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="867f3-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="867f3-155">撰写</span><span class="sxs-lookup"><span data-stu-id="867f3-155">Compose</span></span><br><span data-ttu-id="867f3-156">读取</span><span class="sxs-lookup"><span data-stu-id="867f3-156">Read</span></span> | [<span data-ttu-id="867f3-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="867f3-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4) | [<span data-ttu-id="867f3-158">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="867f3-159">ui</span><span class="sxs-lookup"><span data-stu-id="867f3-159">ui</span></span>](#ui-ui) | <span data-ttu-id="867f3-160">撰写</span><span class="sxs-lookup"><span data-stu-id="867f3-160">Compose</span></span><br><span data-ttu-id="867f3-161">读取</span><span class="sxs-lookup"><span data-stu-id="867f3-161">Read</span></span> | [<span data-ttu-id="867f3-162">UI</span><span class="sxs-lookup"><span data-stu-id="867f3-162">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.4) | [<span data-ttu-id="867f3-163">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="867f3-164">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="867f3-164">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="867f3-165">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="867f3-165">contentLanguage: String</span></span>

<span data-ttu-id="867f3-166">获取用户指定的用于编辑项的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="867f3-166">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="867f3-167">此`contentLanguage`值反映了在 Office 主机应用程序中使用**File > Options > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="867f3-167">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="867f3-168">类型</span><span class="sxs-lookup"><span data-stu-id="867f3-168">Type</span></span>

*   <span data-ttu-id="867f3-169">String</span><span class="sxs-lookup"><span data-stu-id="867f3-169">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="867f3-170">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-170">Requirements</span></span>

|<span data-ttu-id="867f3-171">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-171">Requirement</span></span>| <span data-ttu-id="867f3-172">值</span><span class="sxs-lookup"><span data-stu-id="867f3-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="867f3-173">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="867f3-173">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="867f3-174">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-174">1.1</span></span>|
|[<span data-ttu-id="867f3-175">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="867f3-175">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="867f3-176">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="867f3-176">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="867f3-177">示例</span><span class="sxs-lookup"><span data-stu-id="867f3-177">Example</span></span>

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="867f3-178">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="867f3-178">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="867f3-179">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="867f3-179">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="867f3-180">类型</span><span class="sxs-lookup"><span data-stu-id="867f3-180">Type</span></span>

*   [<span data-ttu-id="867f3-181">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="867f3-181">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="867f3-182">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-182">Requirements</span></span>

|<span data-ttu-id="867f3-183">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-183">Requirement</span></span>| <span data-ttu-id="867f3-184">值</span><span class="sxs-lookup"><span data-stu-id="867f3-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="867f3-185">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="867f3-185">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="867f3-186">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-186">1.1</span></span>|
|[<span data-ttu-id="867f3-187">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="867f3-187">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="867f3-188">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="867f3-188">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="867f3-189">示例</span><span class="sxs-lookup"><span data-stu-id="867f3-189">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="867f3-190">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="867f3-190">displayLanguage: String</span></span>

<span data-ttu-id="867f3-191">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="867f3-191">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="867f3-192">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="867f3-192">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="867f3-193">类型</span><span class="sxs-lookup"><span data-stu-id="867f3-193">Type</span></span>

*   <span data-ttu-id="867f3-194">String</span><span class="sxs-lookup"><span data-stu-id="867f3-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="867f3-195">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-195">Requirements</span></span>

|<span data-ttu-id="867f3-196">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-196">Requirement</span></span>| <span data-ttu-id="867f3-197">值</span><span class="sxs-lookup"><span data-stu-id="867f3-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="867f3-198">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="867f3-198">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="867f3-199">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-199">1.1</span></span>|
|[<span data-ttu-id="867f3-200">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="867f3-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="867f3-201">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="867f3-201">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="867f3-202">示例</span><span class="sxs-lookup"><span data-stu-id="867f3-202">Example</span></span>

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="867f3-203">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="867f3-203">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="867f3-204">获取运行外接程序的 Office 应用程序主机。</span><span class="sxs-lookup"><span data-stu-id="867f3-204">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="867f3-205">类型</span><span class="sxs-lookup"><span data-stu-id="867f3-205">Type</span></span>

*   [<span data-ttu-id="867f3-206">HostType</span><span class="sxs-lookup"><span data-stu-id="867f3-206">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="867f3-207">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-207">Requirements</span></span>

|<span data-ttu-id="867f3-208">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-208">Requirement</span></span>| <span data-ttu-id="867f3-209">值</span><span class="sxs-lookup"><span data-stu-id="867f3-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="867f3-210">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="867f3-210">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="867f3-211">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-211">1.1</span></span>|
|[<span data-ttu-id="867f3-212">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="867f3-212">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="867f3-213">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="867f3-213">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="867f3-214">示例</span><span class="sxs-lookup"><span data-stu-id="867f3-214">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="867f3-215">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="867f3-215">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="867f3-216">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="867f3-216">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="867f3-217">类型</span><span class="sxs-lookup"><span data-stu-id="867f3-217">Type</span></span>

*   [<span data-ttu-id="867f3-218">PlatformType</span><span class="sxs-lookup"><span data-stu-id="867f3-218">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="867f3-219">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-219">Requirements</span></span>

|<span data-ttu-id="867f3-220">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-220">Requirement</span></span>| <span data-ttu-id="867f3-221">值</span><span class="sxs-lookup"><span data-stu-id="867f3-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="867f3-222">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="867f3-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="867f3-223">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-223">1.1</span></span>|
|[<span data-ttu-id="867f3-224">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="867f3-224">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="867f3-225">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="867f3-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="867f3-226">示例</span><span class="sxs-lookup"><span data-stu-id="867f3-226">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="867f3-227">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="867f3-227">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="867f3-228">提供用于确定当前主机和平台上支持的要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="867f3-228">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="867f3-229">类型</span><span class="sxs-lookup"><span data-stu-id="867f3-229">Type</span></span>

*   [<span data-ttu-id="867f3-230">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="867f3-230">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="867f3-231">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-231">Requirements</span></span>

|<span data-ttu-id="867f3-232">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-232">Requirement</span></span>| <span data-ttu-id="867f3-233">值</span><span class="sxs-lookup"><span data-stu-id="867f3-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="867f3-234">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="867f3-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="867f3-235">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-235">1.1</span></span>|
|[<span data-ttu-id="867f3-236">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="867f3-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="867f3-237">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="867f3-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="867f3-238">示例</span><span class="sxs-lookup"><span data-stu-id="867f3-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="867f3-239">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="867f3-239">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="867f3-240">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="867f3-240">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="867f3-241">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="867f3-241">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="867f3-242">类型</span><span class="sxs-lookup"><span data-stu-id="867f3-242">Type</span></span>

*   [<span data-ttu-id="867f3-243">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="867f3-243">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="867f3-244">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-244">Requirements</span></span>

|<span data-ttu-id="867f3-245">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-245">Requirement</span></span>| <span data-ttu-id="867f3-246">值</span><span class="sxs-lookup"><span data-stu-id="867f3-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="867f3-247">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="867f3-247">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="867f3-248">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-248">1.1</span></span>|
|[<span data-ttu-id="867f3-249">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="867f3-249">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="867f3-250">受限</span><span class="sxs-lookup"><span data-stu-id="867f3-250">Restricted</span></span>|
|[<span data-ttu-id="867f3-251">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="867f3-251">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="867f3-252">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="867f3-252">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="867f3-253">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="867f3-253">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="867f3-254">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="867f3-254">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="867f3-255">类型</span><span class="sxs-lookup"><span data-stu-id="867f3-255">Type</span></span>

*   [<span data-ttu-id="867f3-256">UI</span><span class="sxs-lookup"><span data-stu-id="867f3-256">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="867f3-257">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-257">Requirements</span></span>

|<span data-ttu-id="867f3-258">要求</span><span class="sxs-lookup"><span data-stu-id="867f3-258">Requirement</span></span>| <span data-ttu-id="867f3-259">值</span><span class="sxs-lookup"><span data-stu-id="867f3-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="867f3-260">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="867f3-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="867f3-261">1.1</span><span class="sxs-lookup"><span data-stu-id="867f3-261">1.1</span></span>|
|[<span data-ttu-id="867f3-262">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="867f3-262">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="867f3-263">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="867f3-263">Compose or Read</span></span>|
