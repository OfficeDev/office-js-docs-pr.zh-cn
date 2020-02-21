---
title: Office。上下文要求集1。4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 9103f6666b9a7903393729bceb157dfae615da06
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163750"
---
# <a name="context"></a><span data-ttu-id="b0e81-102">context</span><span class="sxs-lookup"><span data-stu-id="b0e81-102">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="b0e81-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="b0e81-103">[Office](office.md).context</span></span>

<span data-ttu-id="b0e81-104">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="b0e81-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="b0e81-105">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅[通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.4)"。</span><span class="sxs-lookup"><span data-stu-id="b0e81-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.4).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b0e81-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="b0e81-106">Requirements</span></span>

|<span data-ttu-id="b0e81-107">要求</span><span class="sxs-lookup"><span data-stu-id="b0e81-107">Requirement</span></span>| <span data-ttu-id="b0e81-108">值</span><span class="sxs-lookup"><span data-stu-id="b0e81-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="b0e81-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b0e81-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b0e81-110">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-110">1.1</span></span>|
|[<span data-ttu-id="b0e81-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b0e81-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b0e81-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b0e81-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="b0e81-113">属性</span><span class="sxs-lookup"><span data-stu-id="b0e81-113">Properties</span></span>

| <span data-ttu-id="b0e81-114">属性</span><span class="sxs-lookup"><span data-stu-id="b0e81-114">Property</span></span> | <span data-ttu-id="b0e81-115">型号</span><span class="sxs-lookup"><span data-stu-id="b0e81-115">Modes</span></span> | <span data-ttu-id="b0e81-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="b0e81-116">Return type</span></span> | <span data-ttu-id="b0e81-117">最低</span><span class="sxs-lookup"><span data-stu-id="b0e81-117">Minimum</span></span><br><span data-ttu-id="b0e81-118">要求集</span><span class="sxs-lookup"><span data-stu-id="b0e81-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b0e81-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="b0e81-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="b0e81-120">撰写</span><span class="sxs-lookup"><span data-stu-id="b0e81-120">Compose</span></span><br><span data-ttu-id="b0e81-121">读取</span><span class="sxs-lookup"><span data-stu-id="b0e81-121">Read</span></span> | <span data-ttu-id="b0e81-122">String</span><span class="sxs-lookup"><span data-stu-id="b0e81-122">String</span></span> | [<span data-ttu-id="b0e81-123">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b0e81-124">过程</span><span class="sxs-lookup"><span data-stu-id="b0e81-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="b0e81-125">撰写</span><span class="sxs-lookup"><span data-stu-id="b0e81-125">Compose</span></span><br><span data-ttu-id="b0e81-126">读取</span><span class="sxs-lookup"><span data-stu-id="b0e81-126">Read</span></span> | [<span data-ttu-id="b0e81-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b0e81-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.4) | [<span data-ttu-id="b0e81-128">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b0e81-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="b0e81-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="b0e81-130">撰写</span><span class="sxs-lookup"><span data-stu-id="b0e81-130">Compose</span></span><br><span data-ttu-id="b0e81-131">读取</span><span class="sxs-lookup"><span data-stu-id="b0e81-131">Read</span></span> | <span data-ttu-id="b0e81-132">String</span><span class="sxs-lookup"><span data-stu-id="b0e81-132">String</span></span> | [<span data-ttu-id="b0e81-133">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b0e81-134">host</span><span class="sxs-lookup"><span data-stu-id="b0e81-134">host</span></span>](#host-hosttype) | <span data-ttu-id="b0e81-135">撰写</span><span class="sxs-lookup"><span data-stu-id="b0e81-135">Compose</span></span><br><span data-ttu-id="b0e81-136">读取</span><span class="sxs-lookup"><span data-stu-id="b0e81-136">Read</span></span> | [<span data-ttu-id="b0e81-137">HostType</span><span class="sxs-lookup"><span data-stu-id="b0e81-137">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.4) | [<span data-ttu-id="b0e81-138">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b0e81-139">mailbox</span><span class="sxs-lookup"><span data-stu-id="b0e81-139">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="b0e81-140">撰写</span><span class="sxs-lookup"><span data-stu-id="b0e81-140">Compose</span></span><br><span data-ttu-id="b0e81-141">读取</span><span class="sxs-lookup"><span data-stu-id="b0e81-141">Read</span></span> | [<span data-ttu-id="b0e81-142">邮箱</span><span class="sxs-lookup"><span data-stu-id="b0e81-142">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4) | [<span data-ttu-id="b0e81-143">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b0e81-144">平台</span><span class="sxs-lookup"><span data-stu-id="b0e81-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="b0e81-145">撰写</span><span class="sxs-lookup"><span data-stu-id="b0e81-145">Compose</span></span><br><span data-ttu-id="b0e81-146">读取</span><span class="sxs-lookup"><span data-stu-id="b0e81-146">Read</span></span> | [<span data-ttu-id="b0e81-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b0e81-147">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.4) | [<span data-ttu-id="b0e81-148">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b0e81-149">满足</span><span class="sxs-lookup"><span data-stu-id="b0e81-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="b0e81-150">撰写</span><span class="sxs-lookup"><span data-stu-id="b0e81-150">Compose</span></span><br><span data-ttu-id="b0e81-151">读取</span><span class="sxs-lookup"><span data-stu-id="b0e81-151">Read</span></span> | [<span data-ttu-id="b0e81-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b0e81-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4) | [<span data-ttu-id="b0e81-153">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-153">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b0e81-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="b0e81-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="b0e81-155">撰写</span><span class="sxs-lookup"><span data-stu-id="b0e81-155">Compose</span></span><br><span data-ttu-id="b0e81-156">读取</span><span class="sxs-lookup"><span data-stu-id="b0e81-156">Read</span></span> | [<span data-ttu-id="b0e81-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b0e81-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4) | [<span data-ttu-id="b0e81-158">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b0e81-159">ui</span><span class="sxs-lookup"><span data-stu-id="b0e81-159">ui</span></span>](#ui-ui) | <span data-ttu-id="b0e81-160">撰写</span><span class="sxs-lookup"><span data-stu-id="b0e81-160">Compose</span></span><br><span data-ttu-id="b0e81-161">读取</span><span class="sxs-lookup"><span data-stu-id="b0e81-161">Read</span></span> | [<span data-ttu-id="b0e81-162">UI</span><span class="sxs-lookup"><span data-stu-id="b0e81-162">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.4) | [<span data-ttu-id="b0e81-163">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="b0e81-164">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="b0e81-164">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="b0e81-165">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="b0e81-165">contentLanguage: String</span></span>

<span data-ttu-id="b0e81-166">获取用户指定的用于编辑项的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="b0e81-166">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="b0e81-167">此`contentLanguage`值反映了在 Office 主机应用程序中使用**File > Options > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="b0e81-167">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="b0e81-168">类型</span><span class="sxs-lookup"><span data-stu-id="b0e81-168">Type</span></span>

*   <span data-ttu-id="b0e81-169">String</span><span class="sxs-lookup"><span data-stu-id="b0e81-169">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b0e81-170">要求</span><span class="sxs-lookup"><span data-stu-id="b0e81-170">Requirements</span></span>

|<span data-ttu-id="b0e81-171">要求</span><span class="sxs-lookup"><span data-stu-id="b0e81-171">Requirement</span></span>| <span data-ttu-id="b0e81-172">值</span><span class="sxs-lookup"><span data-stu-id="b0e81-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="b0e81-173">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b0e81-173">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b0e81-174">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-174">1.1</span></span>|
|[<span data-ttu-id="b0e81-175">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b0e81-175">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b0e81-176">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b0e81-176">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b0e81-177">示例</span><span class="sxs-lookup"><span data-stu-id="b0e81-177">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="b0e81-178">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="b0e81-178">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="b0e81-179">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="b0e81-179">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b0e81-180">类型</span><span class="sxs-lookup"><span data-stu-id="b0e81-180">Type</span></span>

*   [<span data-ttu-id="b0e81-181">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b0e81-181">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="b0e81-182">Requirements</span><span class="sxs-lookup"><span data-stu-id="b0e81-182">Requirements</span></span>

|<span data-ttu-id="b0e81-183">要求</span><span class="sxs-lookup"><span data-stu-id="b0e81-183">Requirement</span></span>| <span data-ttu-id="b0e81-184">值</span><span class="sxs-lookup"><span data-stu-id="b0e81-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="b0e81-185">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b0e81-185">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b0e81-186">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-186">1.1</span></span>|
|[<span data-ttu-id="b0e81-187">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b0e81-187">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b0e81-188">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b0e81-188">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b0e81-189">示例</span><span class="sxs-lookup"><span data-stu-id="b0e81-189">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="b0e81-190">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="b0e81-190">displayLanguage: String</span></span>

<span data-ttu-id="b0e81-191">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="b0e81-191">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="b0e81-192">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="b0e81-192">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="b0e81-193">类型</span><span class="sxs-lookup"><span data-stu-id="b0e81-193">Type</span></span>

*   <span data-ttu-id="b0e81-194">String</span><span class="sxs-lookup"><span data-stu-id="b0e81-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b0e81-195">要求</span><span class="sxs-lookup"><span data-stu-id="b0e81-195">Requirements</span></span>

|<span data-ttu-id="b0e81-196">要求</span><span class="sxs-lookup"><span data-stu-id="b0e81-196">Requirement</span></span>| <span data-ttu-id="b0e81-197">值</span><span class="sxs-lookup"><span data-stu-id="b0e81-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="b0e81-198">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b0e81-198">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b0e81-199">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-199">1.1</span></span>|
|[<span data-ttu-id="b0e81-200">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b0e81-200">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b0e81-201">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b0e81-201">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b0e81-202">示例</span><span class="sxs-lookup"><span data-stu-id="b0e81-202">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="b0e81-203">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="b0e81-203">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="b0e81-204">获取运行外接程序的 Office 应用程序主机。</span><span class="sxs-lookup"><span data-stu-id="b0e81-204">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b0e81-205">类型</span><span class="sxs-lookup"><span data-stu-id="b0e81-205">Type</span></span>

*   [<span data-ttu-id="b0e81-206">HostType</span><span class="sxs-lookup"><span data-stu-id="b0e81-206">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="b0e81-207">Requirements</span><span class="sxs-lookup"><span data-stu-id="b0e81-207">Requirements</span></span>

|<span data-ttu-id="b0e81-208">要求</span><span class="sxs-lookup"><span data-stu-id="b0e81-208">Requirement</span></span>| <span data-ttu-id="b0e81-209">值</span><span class="sxs-lookup"><span data-stu-id="b0e81-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="b0e81-210">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b0e81-210">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b0e81-211">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-211">1.1</span></span>|
|[<span data-ttu-id="b0e81-212">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b0e81-212">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b0e81-213">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b0e81-213">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b0e81-214">示例</span><span class="sxs-lookup"><span data-stu-id="b0e81-214">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="b0e81-215">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="b0e81-215">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="b0e81-216">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="b0e81-216">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b0e81-217">类型</span><span class="sxs-lookup"><span data-stu-id="b0e81-217">Type</span></span>

*   [<span data-ttu-id="b0e81-218">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b0e81-218">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="b0e81-219">Requirements</span><span class="sxs-lookup"><span data-stu-id="b0e81-219">Requirements</span></span>

|<span data-ttu-id="b0e81-220">要求</span><span class="sxs-lookup"><span data-stu-id="b0e81-220">Requirement</span></span>| <span data-ttu-id="b0e81-221">值</span><span class="sxs-lookup"><span data-stu-id="b0e81-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="b0e81-222">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b0e81-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b0e81-223">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-223">1.1</span></span>|
|[<span data-ttu-id="b0e81-224">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b0e81-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b0e81-225">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b0e81-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b0e81-226">示例</span><span class="sxs-lookup"><span data-stu-id="b0e81-226">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="b0e81-227">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="b0e81-227">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="b0e81-228">提供用于确定当前主机和平台上支持的要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="b0e81-228">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b0e81-229">类型</span><span class="sxs-lookup"><span data-stu-id="b0e81-229">Type</span></span>

*   [<span data-ttu-id="b0e81-230">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b0e81-230">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="b0e81-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="b0e81-231">Requirements</span></span>

|<span data-ttu-id="b0e81-232">要求</span><span class="sxs-lookup"><span data-stu-id="b0e81-232">Requirement</span></span>| <span data-ttu-id="b0e81-233">值</span><span class="sxs-lookup"><span data-stu-id="b0e81-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="b0e81-234">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b0e81-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b0e81-235">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-235">1.1</span></span>|
|[<span data-ttu-id="b0e81-236">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b0e81-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b0e81-237">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b0e81-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b0e81-238">示例</span><span class="sxs-lookup"><span data-stu-id="b0e81-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="b0e81-239">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="b0e81-239">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="b0e81-240">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="b0e81-240">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="b0e81-241">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="b0e81-241">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="b0e81-242">类型</span><span class="sxs-lookup"><span data-stu-id="b0e81-242">Type</span></span>

*   [<span data-ttu-id="b0e81-243">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b0e81-243">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="b0e81-244">Requirements</span><span class="sxs-lookup"><span data-stu-id="b0e81-244">Requirements</span></span>

|<span data-ttu-id="b0e81-245">要求</span><span class="sxs-lookup"><span data-stu-id="b0e81-245">Requirement</span></span>| <span data-ttu-id="b0e81-246">值</span><span class="sxs-lookup"><span data-stu-id="b0e81-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="b0e81-247">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b0e81-247">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b0e81-248">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-248">1.1</span></span>|
|[<span data-ttu-id="b0e81-249">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b0e81-249">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="b0e81-250">受限</span><span class="sxs-lookup"><span data-stu-id="b0e81-250">Restricted</span></span>|
|[<span data-ttu-id="b0e81-251">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b0e81-251">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b0e81-252">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b0e81-252">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="b0e81-253">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="b0e81-253">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="b0e81-254">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="b0e81-254">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="b0e81-255">类型</span><span class="sxs-lookup"><span data-stu-id="b0e81-255">Type</span></span>

*   [<span data-ttu-id="b0e81-256">UI</span><span class="sxs-lookup"><span data-stu-id="b0e81-256">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="b0e81-257">Requirements</span><span class="sxs-lookup"><span data-stu-id="b0e81-257">Requirements</span></span>

|<span data-ttu-id="b0e81-258">要求</span><span class="sxs-lookup"><span data-stu-id="b0e81-258">Requirement</span></span>| <span data-ttu-id="b0e81-259">值</span><span class="sxs-lookup"><span data-stu-id="b0e81-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="b0e81-260">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b0e81-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b0e81-261">1.1</span><span class="sxs-lookup"><span data-stu-id="b0e81-261">1.1</span></span>|
|[<span data-ttu-id="b0e81-262">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b0e81-262">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b0e81-263">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b0e81-263">Compose or Read</span></span>|
