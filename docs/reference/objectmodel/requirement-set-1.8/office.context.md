---
title: Office。上下文要求集1。8
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: fdbd559ffc672b5ddf25fed069da8ef4c05af286
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814178"
---
# <a name="context"></a><span data-ttu-id="259f5-102">context</span><span class="sxs-lookup"><span data-stu-id="259f5-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="259f5-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="259f5-103">[Office](office.md).context</span></span>

<span data-ttu-id="259f5-104">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="259f5-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="259f5-105">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅[通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.8)"。</span><span class="sxs-lookup"><span data-stu-id="259f5-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.8).</span></span>

##### <a name="requirements"></a><span data-ttu-id="259f5-106">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-106">Requirements</span></span>

|<span data-ttu-id="259f5-107">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-107">Requirement</span></span>| <span data-ttu-id="259f5-108">值</span><span class="sxs-lookup"><span data-stu-id="259f5-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="259f5-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="259f5-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="259f5-110">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-110">1.1</span></span>|
|[<span data-ttu-id="259f5-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="259f5-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="259f5-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="259f5-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="259f5-113">属性</span><span class="sxs-lookup"><span data-stu-id="259f5-113">Properties</span></span>

| <span data-ttu-id="259f5-114">属性</span><span class="sxs-lookup"><span data-stu-id="259f5-114">Property</span></span> | <span data-ttu-id="259f5-115">型号</span><span class="sxs-lookup"><span data-stu-id="259f5-115">Modes</span></span> | <span data-ttu-id="259f5-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="259f5-116">Return type</span></span> | <span data-ttu-id="259f5-117">最低</span><span class="sxs-lookup"><span data-stu-id="259f5-117">Minimum</span></span><br><span data-ttu-id="259f5-118">要求集</span><span class="sxs-lookup"><span data-stu-id="259f5-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="259f5-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="259f5-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="259f5-120">撰写</span><span class="sxs-lookup"><span data-stu-id="259f5-120">Compose</span></span><br><span data-ttu-id="259f5-121">读取</span><span class="sxs-lookup"><span data-stu-id="259f5-121">Read</span></span> | <span data-ttu-id="259f5-122">String</span><span class="sxs-lookup"><span data-stu-id="259f5-122">String</span></span> | [<span data-ttu-id="259f5-123">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="259f5-124">过程</span><span class="sxs-lookup"><span data-stu-id="259f5-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="259f5-125">撰写</span><span class="sxs-lookup"><span data-stu-id="259f5-125">Compose</span></span><br><span data-ttu-id="259f5-126">读取</span><span class="sxs-lookup"><span data-stu-id="259f5-126">Read</span></span> | [<span data-ttu-id="259f5-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="259f5-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.8) | [<span data-ttu-id="259f5-128">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="259f5-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="259f5-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="259f5-130">撰写</span><span class="sxs-lookup"><span data-stu-id="259f5-130">Compose</span></span><br><span data-ttu-id="259f5-131">读取</span><span class="sxs-lookup"><span data-stu-id="259f5-131">Read</span></span> | <span data-ttu-id="259f5-132">String</span><span class="sxs-lookup"><span data-stu-id="259f5-132">String</span></span> | [<span data-ttu-id="259f5-133">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="259f5-134">host</span><span class="sxs-lookup"><span data-stu-id="259f5-134">host</span></span>](#host-hosttype) | <span data-ttu-id="259f5-135">撰写</span><span class="sxs-lookup"><span data-stu-id="259f5-135">Compose</span></span><br><span data-ttu-id="259f5-136">读取</span><span class="sxs-lookup"><span data-stu-id="259f5-136">Read</span></span> | [<span data-ttu-id="259f5-137">HostType</span><span class="sxs-lookup"><span data-stu-id="259f5-137">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.8) | [<span data-ttu-id="259f5-138">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="259f5-139">mailbox</span><span class="sxs-lookup"><span data-stu-id="259f5-139">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="259f5-140">撰写</span><span class="sxs-lookup"><span data-stu-id="259f5-140">Compose</span></span><br><span data-ttu-id="259f5-141">读取</span><span class="sxs-lookup"><span data-stu-id="259f5-141">Read</span></span> | [<span data-ttu-id="259f5-142">邮箱</span><span class="sxs-lookup"><span data-stu-id="259f5-142">Mailbox</span></span>](/javascript/api/office/office.mailbox?view=outlook-js-1.8) | [<span data-ttu-id="259f5-143">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="259f5-144">平台</span><span class="sxs-lookup"><span data-stu-id="259f5-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="259f5-145">撰写</span><span class="sxs-lookup"><span data-stu-id="259f5-145">Compose</span></span><br><span data-ttu-id="259f5-146">读取</span><span class="sxs-lookup"><span data-stu-id="259f5-146">Read</span></span> | [<span data-ttu-id="259f5-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="259f5-147">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.8) | [<span data-ttu-id="259f5-148">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="259f5-149">满足</span><span class="sxs-lookup"><span data-stu-id="259f5-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="259f5-150">撰写</span><span class="sxs-lookup"><span data-stu-id="259f5-150">Compose</span></span><br><span data-ttu-id="259f5-151">读取</span><span class="sxs-lookup"><span data-stu-id="259f5-151">Read</span></span> | [<span data-ttu-id="259f5-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="259f5-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.8) | [<span data-ttu-id="259f5-153">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-153">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="259f5-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="259f5-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="259f5-155">撰写</span><span class="sxs-lookup"><span data-stu-id="259f5-155">Compose</span></span><br><span data-ttu-id="259f5-156">读取</span><span class="sxs-lookup"><span data-stu-id="259f5-156">Read</span></span> | [<span data-ttu-id="259f5-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="259f5-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.8) | [<span data-ttu-id="259f5-158">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="259f5-159">ui</span><span class="sxs-lookup"><span data-stu-id="259f5-159">ui</span></span>](#ui-ui) | <span data-ttu-id="259f5-160">撰写</span><span class="sxs-lookup"><span data-stu-id="259f5-160">Compose</span></span><br><span data-ttu-id="259f5-161">读取</span><span class="sxs-lookup"><span data-stu-id="259f5-161">Read</span></span> | [<span data-ttu-id="259f5-162">UI</span><span class="sxs-lookup"><span data-stu-id="259f5-162">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.8) | [<span data-ttu-id="259f5-163">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="259f5-164">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="259f5-164">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="259f5-165">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="259f5-165">contentLanguage: String</span></span>

<span data-ttu-id="259f5-166">获取用户指定的用于编辑项的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="259f5-166">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="259f5-167">此`contentLanguage`值反映了在 Office 主机应用程序中使用**File > Options > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="259f5-167">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="259f5-168">类型</span><span class="sxs-lookup"><span data-stu-id="259f5-168">Type</span></span>

*   <span data-ttu-id="259f5-169">String</span><span class="sxs-lookup"><span data-stu-id="259f5-169">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="259f5-170">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-170">Requirements</span></span>

|<span data-ttu-id="259f5-171">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-171">Requirement</span></span>| <span data-ttu-id="259f5-172">值</span><span class="sxs-lookup"><span data-stu-id="259f5-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="259f5-173">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="259f5-173">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="259f5-174">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-174">1.1</span></span>|
|[<span data-ttu-id="259f5-175">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="259f5-175">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="259f5-176">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="259f5-176">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="259f5-177">示例</span><span class="sxs-lookup"><span data-stu-id="259f5-177">Example</span></span>

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="259f5-178">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="259f5-178">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="259f5-179">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="259f5-179">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="259f5-180">类型</span><span class="sxs-lookup"><span data-stu-id="259f5-180">Type</span></span>

*   [<span data-ttu-id="259f5-181">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="259f5-181">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="259f5-182">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-182">Requirements</span></span>

|<span data-ttu-id="259f5-183">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-183">Requirement</span></span>| <span data-ttu-id="259f5-184">值</span><span class="sxs-lookup"><span data-stu-id="259f5-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="259f5-185">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="259f5-185">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="259f5-186">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-186">1.1</span></span>|
|[<span data-ttu-id="259f5-187">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="259f5-187">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="259f5-188">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="259f5-188">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="259f5-189">示例</span><span class="sxs-lookup"><span data-stu-id="259f5-189">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="259f5-190">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="259f5-190">displayLanguage: String</span></span>

<span data-ttu-id="259f5-191">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="259f5-191">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="259f5-192">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="259f5-192">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="259f5-193">类型</span><span class="sxs-lookup"><span data-stu-id="259f5-193">Type</span></span>

*   <span data-ttu-id="259f5-194">String</span><span class="sxs-lookup"><span data-stu-id="259f5-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="259f5-195">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-195">Requirements</span></span>

|<span data-ttu-id="259f5-196">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-196">Requirement</span></span>| <span data-ttu-id="259f5-197">值</span><span class="sxs-lookup"><span data-stu-id="259f5-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="259f5-198">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="259f5-198">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="259f5-199">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-199">1.1</span></span>|
|[<span data-ttu-id="259f5-200">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="259f5-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="259f5-201">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="259f5-201">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="259f5-202">示例</span><span class="sxs-lookup"><span data-stu-id="259f5-202">Example</span></span>

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="259f5-203">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="259f5-203">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="259f5-204">获取运行外接程序的 Office 应用程序主机。</span><span class="sxs-lookup"><span data-stu-id="259f5-204">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="259f5-205">类型</span><span class="sxs-lookup"><span data-stu-id="259f5-205">Type</span></span>

*   [<span data-ttu-id="259f5-206">HostType</span><span class="sxs-lookup"><span data-stu-id="259f5-206">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="259f5-207">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-207">Requirements</span></span>

|<span data-ttu-id="259f5-208">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-208">Requirement</span></span>| <span data-ttu-id="259f5-209">值</span><span class="sxs-lookup"><span data-stu-id="259f5-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="259f5-210">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="259f5-210">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="259f5-211">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-211">1.1</span></span>|
|[<span data-ttu-id="259f5-212">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="259f5-212">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="259f5-213">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="259f5-213">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="259f5-214">示例</span><span class="sxs-lookup"><span data-stu-id="259f5-214">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="259f5-215">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="259f5-215">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="259f5-216">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="259f5-216">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="259f5-217">类型</span><span class="sxs-lookup"><span data-stu-id="259f5-217">Type</span></span>

*   [<span data-ttu-id="259f5-218">PlatformType</span><span class="sxs-lookup"><span data-stu-id="259f5-218">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="259f5-219">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-219">Requirements</span></span>

|<span data-ttu-id="259f5-220">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-220">Requirement</span></span>| <span data-ttu-id="259f5-221">值</span><span class="sxs-lookup"><span data-stu-id="259f5-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="259f5-222">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="259f5-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="259f5-223">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-223">1.1</span></span>|
|[<span data-ttu-id="259f5-224">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="259f5-224">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="259f5-225">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="259f5-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="259f5-226">示例</span><span class="sxs-lookup"><span data-stu-id="259f5-226">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="259f5-227">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="259f5-227">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="259f5-228">提供用于确定当前主机和平台上支持的要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="259f5-228">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="259f5-229">类型</span><span class="sxs-lookup"><span data-stu-id="259f5-229">Type</span></span>

*   [<span data-ttu-id="259f5-230">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="259f5-230">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="259f5-231">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-231">Requirements</span></span>

|<span data-ttu-id="259f5-232">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-232">Requirement</span></span>| <span data-ttu-id="259f5-233">值</span><span class="sxs-lookup"><span data-stu-id="259f5-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="259f5-234">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="259f5-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="259f5-235">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-235">1.1</span></span>|
|[<span data-ttu-id="259f5-236">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="259f5-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="259f5-237">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="259f5-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="259f5-238">示例</span><span class="sxs-lookup"><span data-stu-id="259f5-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="259f5-239">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="259f5-239">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="259f5-240">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="259f5-240">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="259f5-241">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="259f5-241">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="259f5-242">类型</span><span class="sxs-lookup"><span data-stu-id="259f5-242">Type</span></span>

*   [<span data-ttu-id="259f5-243">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="259f5-243">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="259f5-244">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-244">Requirements</span></span>

|<span data-ttu-id="259f5-245">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-245">Requirement</span></span>| <span data-ttu-id="259f5-246">值</span><span class="sxs-lookup"><span data-stu-id="259f5-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="259f5-247">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="259f5-247">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="259f5-248">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-248">1.1</span></span>|
|[<span data-ttu-id="259f5-249">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="259f5-249">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="259f5-250">受限</span><span class="sxs-lookup"><span data-stu-id="259f5-250">Restricted</span></span>|
|[<span data-ttu-id="259f5-251">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="259f5-251">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="259f5-252">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="259f5-252">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="259f5-253">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="259f5-253">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="259f5-254">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="259f5-254">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="259f5-255">类型</span><span class="sxs-lookup"><span data-stu-id="259f5-255">Type</span></span>

*   [<span data-ttu-id="259f5-256">UI</span><span class="sxs-lookup"><span data-stu-id="259f5-256">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="259f5-257">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-257">Requirements</span></span>

|<span data-ttu-id="259f5-258">要求</span><span class="sxs-lookup"><span data-stu-id="259f5-258">Requirement</span></span>| <span data-ttu-id="259f5-259">值</span><span class="sxs-lookup"><span data-stu-id="259f5-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="259f5-260">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="259f5-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="259f5-261">1.1</span><span class="sxs-lookup"><span data-stu-id="259f5-261">1.1</span></span>|
|[<span data-ttu-id="259f5-262">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="259f5-262">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="259f5-263">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="259f5-263">Compose or Read</span></span>|
