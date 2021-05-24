---
title: Office.context - 要求集 1.4
description: Office。适用于使用邮箱 API 要求Outlook集 1.4 的外接程序的上下文对象成员。
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: 6183715090cbbca008b0a750012c65da0ac21d7c
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591028"
---
# <a name="context-mailbox-requirement-set-14"></a><span data-ttu-id="ca781-103">context (Mailbox requirement set 1.4) </span><span class="sxs-lookup"><span data-stu-id="ca781-103">context (Mailbox requirement set 1.4)</span></span>

### <a name="officecontext"></a><span data-ttu-id="ca781-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="ca781-104">[Office](office.md).context</span></span>

<span data-ttu-id="ca781-105">Office.context 提供了外接程序在所有应用程序中使用的共享Office接口。</span><span class="sxs-lookup"><span data-stu-id="ca781-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="ca781-106">此列表仅记录外接程序Outlook接口。有关 Office.context 命名空间的完整列表，请参阅通用 API 中的[Office.context 引用](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true)。</span><span class="sxs-lookup"><span data-stu-id="ca781-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca781-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca781-107">Requirements</span></span>

|<span data-ttu-id="ca781-108">要求</span><span class="sxs-lookup"><span data-stu-id="ca781-108">Requirement</span></span>| <span data-ttu-id="ca781-109">值</span><span class="sxs-lookup"><span data-stu-id="ca781-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca781-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca781-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ca781-111">1.1</span><span class="sxs-lookup"><span data-stu-id="ca781-111">1.1</span></span>|
|[<span data-ttu-id="ca781-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca781-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ca781-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca781-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="ca781-114">属性</span><span class="sxs-lookup"><span data-stu-id="ca781-114">Properties</span></span>

| <span data-ttu-id="ca781-115">属性</span><span class="sxs-lookup"><span data-stu-id="ca781-115">Property</span></span> | <span data-ttu-id="ca781-116">模式</span><span class="sxs-lookup"><span data-stu-id="ca781-116">Modes</span></span> | <span data-ttu-id="ca781-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="ca781-117">Return type</span></span> | <span data-ttu-id="ca781-118">最小值</span><span class="sxs-lookup"><span data-stu-id="ca781-118">Minimum</span></span><br><span data-ttu-id="ca781-119">要求集</span><span class="sxs-lookup"><span data-stu-id="ca781-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ca781-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="ca781-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="ca781-121">撰写</span><span class="sxs-lookup"><span data-stu-id="ca781-121">Compose</span></span><br><span data-ttu-id="ca781-122">阅读</span><span class="sxs-lookup"><span data-stu-id="ca781-122">Read</span></span> | <span data-ttu-id="ca781-123">字符串</span><span class="sxs-lookup"><span data-stu-id="ca781-123">String</span></span> | [<span data-ttu-id="ca781-124">1.1</span><span class="sxs-lookup"><span data-stu-id="ca781-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ca781-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="ca781-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="ca781-126">撰写</span><span class="sxs-lookup"><span data-stu-id="ca781-126">Compose</span></span><br><span data-ttu-id="ca781-127">阅读</span><span class="sxs-lookup"><span data-stu-id="ca781-127">Read</span></span> | [<span data-ttu-id="ca781-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ca781-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="ca781-129">1.1</span><span class="sxs-lookup"><span data-stu-id="ca781-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ca781-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="ca781-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="ca781-131">撰写</span><span class="sxs-lookup"><span data-stu-id="ca781-131">Compose</span></span><br><span data-ttu-id="ca781-132">阅读</span><span class="sxs-lookup"><span data-stu-id="ca781-132">Read</span></span> | <span data-ttu-id="ca781-133">字符串</span><span class="sxs-lookup"><span data-stu-id="ca781-133">String</span></span> | [<span data-ttu-id="ca781-134">1.1</span><span class="sxs-lookup"><span data-stu-id="ca781-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ca781-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="ca781-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="ca781-136">撰写</span><span class="sxs-lookup"><span data-stu-id="ca781-136">Compose</span></span><br><span data-ttu-id="ca781-137">阅读</span><span class="sxs-lookup"><span data-stu-id="ca781-137">Read</span></span> | [<span data-ttu-id="ca781-138">邮箱</span><span class="sxs-lookup"><span data-stu-id="ca781-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="ca781-139">1.1</span><span class="sxs-lookup"><span data-stu-id="ca781-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ca781-140">requirements</span><span class="sxs-lookup"><span data-stu-id="ca781-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="ca781-141">撰写</span><span class="sxs-lookup"><span data-stu-id="ca781-141">Compose</span></span><br><span data-ttu-id="ca781-142">阅读</span><span class="sxs-lookup"><span data-stu-id="ca781-142">Read</span></span> | [<span data-ttu-id="ca781-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ca781-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="ca781-144">1.1</span><span class="sxs-lookup"><span data-stu-id="ca781-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ca781-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="ca781-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="ca781-146">撰写</span><span class="sxs-lookup"><span data-stu-id="ca781-146">Compose</span></span><br><span data-ttu-id="ca781-147">阅读</span><span class="sxs-lookup"><span data-stu-id="ca781-147">Read</span></span> | [<span data-ttu-id="ca781-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ca781-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="ca781-149">1.1</span><span class="sxs-lookup"><span data-stu-id="ca781-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ca781-150">ui</span><span class="sxs-lookup"><span data-stu-id="ca781-150">ui</span></span>](#ui-ui) | <span data-ttu-id="ca781-151">撰写</span><span class="sxs-lookup"><span data-stu-id="ca781-151">Compose</span></span><br><span data-ttu-id="ca781-152">阅读</span><span class="sxs-lookup"><span data-stu-id="ca781-152">Read</span></span> | [<span data-ttu-id="ca781-153">UI</span><span class="sxs-lookup"><span data-stu-id="ca781-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="ca781-154">1.1</span><span class="sxs-lookup"><span data-stu-id="ca781-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="ca781-155">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="ca781-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="ca781-156">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="ca781-156">contentLanguage: String</span></span>

<span data-ttu-id="ca781-157">获取用户 (编辑) 的语言区域设置。</span><span class="sxs-lookup"><span data-stu-id="ca781-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="ca781-158">该值 `contentLanguage` 反映当前在客户端 **应用程序中** 由 File **> Options > Language** 指定的Office设置。</span><span class="sxs-lookup"><span data-stu-id="ca781-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="ca781-159">类型</span><span class="sxs-lookup"><span data-stu-id="ca781-159">Type</span></span>

*   <span data-ttu-id="ca781-160">String</span><span class="sxs-lookup"><span data-stu-id="ca781-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca781-161">要求</span><span class="sxs-lookup"><span data-stu-id="ca781-161">Requirements</span></span>

|<span data-ttu-id="ca781-162">要求</span><span class="sxs-lookup"><span data-stu-id="ca781-162">Requirement</span></span>| <span data-ttu-id="ca781-163">值</span><span class="sxs-lookup"><span data-stu-id="ca781-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca781-164">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca781-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ca781-165">1.1</span><span class="sxs-lookup"><span data-stu-id="ca781-165">1.1</span></span>|
|[<span data-ttu-id="ca781-166">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca781-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ca781-167">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca781-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca781-168">示例</span><span class="sxs-lookup"><span data-stu-id="ca781-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="ca781-169">diagnostics： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="ca781-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="ca781-170">获取加载项运行环境的信息。</span><span class="sxs-lookup"><span data-stu-id="ca781-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ca781-171">类型</span><span class="sxs-lookup"><span data-stu-id="ca781-171">Type</span></span>

*   [<span data-ttu-id="ca781-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ca781-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="ca781-173">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca781-173">Requirements</span></span>

|<span data-ttu-id="ca781-174">要求</span><span class="sxs-lookup"><span data-stu-id="ca781-174">Requirement</span></span>| <span data-ttu-id="ca781-175">值</span><span class="sxs-lookup"><span data-stu-id="ca781-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca781-176">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca781-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ca781-177">1.1</span><span class="sxs-lookup"><span data-stu-id="ca781-177">1.1</span></span>|
|[<span data-ttu-id="ca781-178">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca781-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ca781-179">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca781-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca781-180">示例</span><span class="sxs-lookup"><span data-stu-id="ca781-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="ca781-181">displayLanguage：String</span><span class="sxs-lookup"><span data-stu-id="ca781-181">displayLanguage: String</span></span>

<span data-ttu-id="ca781-182">获取区域设置 (语言) RFC 1766 语言标记格式，该标记格式由用户为 Office 客户端应用程序的 UI 指定。</span><span class="sxs-lookup"><span data-stu-id="ca781-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="ca781-183">该值反映当前显示语言设置，该设置由 > `displayLanguage` **客户端** 应用程序中>选项Office语言。 </span><span class="sxs-lookup"><span data-stu-id="ca781-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="ca781-184">类型</span><span class="sxs-lookup"><span data-stu-id="ca781-184">Type</span></span>

*   <span data-ttu-id="ca781-185">String</span><span class="sxs-lookup"><span data-stu-id="ca781-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ca781-186">要求</span><span class="sxs-lookup"><span data-stu-id="ca781-186">Requirements</span></span>

|<span data-ttu-id="ca781-187">要求</span><span class="sxs-lookup"><span data-stu-id="ca781-187">Requirement</span></span>| <span data-ttu-id="ca781-188">值</span><span class="sxs-lookup"><span data-stu-id="ca781-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca781-189">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca781-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ca781-190">1.1</span><span class="sxs-lookup"><span data-stu-id="ca781-190">1.1</span></span>|
|[<span data-ttu-id="ca781-191">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca781-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ca781-192">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca781-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca781-193">示例</span><span class="sxs-lookup"><span data-stu-id="ca781-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="ca781-194">requirements： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="ca781-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="ca781-195">提供用于确定当前应用程序和平台上支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="ca781-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="ca781-196">类型</span><span class="sxs-lookup"><span data-stu-id="ca781-196">Type</span></span>

*   [<span data-ttu-id="ca781-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ca781-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="ca781-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca781-198">Requirements</span></span>

|<span data-ttu-id="ca781-199">要求</span><span class="sxs-lookup"><span data-stu-id="ca781-199">Requirement</span></span>| <span data-ttu-id="ca781-200">值</span><span class="sxs-lookup"><span data-stu-id="ca781-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca781-201">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca781-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ca781-202">1.1</span><span class="sxs-lookup"><span data-stu-id="ca781-202">1.1</span></span>|
|[<span data-ttu-id="ca781-203">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca781-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ca781-204">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca781-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ca781-205">示例</span><span class="sxs-lookup"><span data-stu-id="ca781-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="ca781-206">[roamingSettings：RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="ca781-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="ca781-207">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="ca781-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="ca781-208">该对象允许您存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时可供该外接程序使用 `RoamingSettings` 。</span><span class="sxs-lookup"><span data-stu-id="ca781-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="ca781-209">类型</span><span class="sxs-lookup"><span data-stu-id="ca781-209">Type</span></span>

*   [<span data-ttu-id="ca781-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ca781-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="ca781-211">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca781-211">Requirements</span></span>

|<span data-ttu-id="ca781-212">要求</span><span class="sxs-lookup"><span data-stu-id="ca781-212">Requirement</span></span>| <span data-ttu-id="ca781-213">值</span><span class="sxs-lookup"><span data-stu-id="ca781-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca781-214">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca781-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ca781-215">1.1</span><span class="sxs-lookup"><span data-stu-id="ca781-215">1.1</span></span>|
|[<span data-ttu-id="ca781-216">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ca781-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="ca781-217">受限</span><span class="sxs-lookup"><span data-stu-id="ca781-217">Restricted</span></span>|
|[<span data-ttu-id="ca781-218">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca781-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ca781-219">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca781-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="ca781-220">[ui：UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="ca781-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="ca781-221">提供可用于在加载项中创建和操作 UI 组件（如对话框）Office方法。</span><span class="sxs-lookup"><span data-stu-id="ca781-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ca781-222">类型</span><span class="sxs-lookup"><span data-stu-id="ca781-222">Type</span></span>

*   [<span data-ttu-id="ca781-223">UI</span><span class="sxs-lookup"><span data-stu-id="ca781-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="ca781-224">Requirements</span><span class="sxs-lookup"><span data-stu-id="ca781-224">Requirements</span></span>

|<span data-ttu-id="ca781-225">要求</span><span class="sxs-lookup"><span data-stu-id="ca781-225">Requirement</span></span>| <span data-ttu-id="ca781-226">值</span><span class="sxs-lookup"><span data-stu-id="ca781-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="ca781-227">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ca781-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ca781-228">1.1</span><span class="sxs-lookup"><span data-stu-id="ca781-228">1.1</span></span>|
|[<span data-ttu-id="ca781-229">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ca781-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ca781-230">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ca781-230">Compose or Read</span></span>|
