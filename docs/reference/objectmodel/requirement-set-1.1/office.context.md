---
title: Office.context - 要求集 1.1
description: Office。使用邮箱 API 要求集 1.1 Outlook外接程序可用的上下文对象成员。
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: 41273bfc5362a9d5572e38b8e80b81041f5aa312
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590874"
---
# <a name="context-mailbox-requirement-set-11"></a><span data-ttu-id="a9a27-103">context (Mailbox requirement set 1.1) </span><span class="sxs-lookup"><span data-stu-id="a9a27-103">context (Mailbox requirement set 1.1)</span></span>

### <a name="officecontext"></a><span data-ttu-id="a9a27-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="a9a27-104">[Office](office.md).context</span></span>

<span data-ttu-id="a9a27-105">Office.context 提供了外接程序在所有应用程序中使用的共享Office接口。</span><span class="sxs-lookup"><span data-stu-id="a9a27-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="a9a27-106">此列表仅记录外接程序Outlook接口。有关 Office.context 命名空间的完整列表，请参阅通用 API 中的[Office.context 引用](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true)。</span><span class="sxs-lookup"><span data-stu-id="a9a27-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9a27-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="a9a27-107">Requirements</span></span>

|<span data-ttu-id="a9a27-108">要求</span><span class="sxs-lookup"><span data-stu-id="a9a27-108">Requirement</span></span>| <span data-ttu-id="a9a27-109">值</span><span class="sxs-lookup"><span data-stu-id="a9a27-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9a27-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a9a27-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a9a27-111">1.1</span><span class="sxs-lookup"><span data-stu-id="a9a27-111">1.1</span></span>|
|[<span data-ttu-id="a9a27-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a9a27-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a9a27-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a9a27-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="a9a27-114">属性</span><span class="sxs-lookup"><span data-stu-id="a9a27-114">Properties</span></span>

| <span data-ttu-id="a9a27-115">属性</span><span class="sxs-lookup"><span data-stu-id="a9a27-115">Property</span></span> | <span data-ttu-id="a9a27-116">模式</span><span class="sxs-lookup"><span data-stu-id="a9a27-116">Modes</span></span> | <span data-ttu-id="a9a27-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="a9a27-117">Return type</span></span> | <span data-ttu-id="a9a27-118">最小值</span><span class="sxs-lookup"><span data-stu-id="a9a27-118">Minimum</span></span><br><span data-ttu-id="a9a27-119">要求集</span><span class="sxs-lookup"><span data-stu-id="a9a27-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a9a27-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="a9a27-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="a9a27-121">撰写</span><span class="sxs-lookup"><span data-stu-id="a9a27-121">Compose</span></span><br><span data-ttu-id="a9a27-122">阅读</span><span class="sxs-lookup"><span data-stu-id="a9a27-122">Read</span></span> | <span data-ttu-id="a9a27-123">字符串</span><span class="sxs-lookup"><span data-stu-id="a9a27-123">String</span></span> | [<span data-ttu-id="a9a27-124">1.1</span><span class="sxs-lookup"><span data-stu-id="a9a27-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a9a27-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="a9a27-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="a9a27-126">撰写</span><span class="sxs-lookup"><span data-stu-id="a9a27-126">Compose</span></span><br><span data-ttu-id="a9a27-127">阅读</span><span class="sxs-lookup"><span data-stu-id="a9a27-127">Read</span></span> | [<span data-ttu-id="a9a27-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="a9a27-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="a9a27-129">1.1</span><span class="sxs-lookup"><span data-stu-id="a9a27-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a9a27-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="a9a27-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="a9a27-131">撰写</span><span class="sxs-lookup"><span data-stu-id="a9a27-131">Compose</span></span><br><span data-ttu-id="a9a27-132">阅读</span><span class="sxs-lookup"><span data-stu-id="a9a27-132">Read</span></span> | <span data-ttu-id="a9a27-133">字符串</span><span class="sxs-lookup"><span data-stu-id="a9a27-133">String</span></span> | [<span data-ttu-id="a9a27-134">1.1</span><span class="sxs-lookup"><span data-stu-id="a9a27-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a9a27-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="a9a27-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="a9a27-136">撰写</span><span class="sxs-lookup"><span data-stu-id="a9a27-136">Compose</span></span><br><span data-ttu-id="a9a27-137">阅读</span><span class="sxs-lookup"><span data-stu-id="a9a27-137">Read</span></span> | [<span data-ttu-id="a9a27-138">邮箱</span><span class="sxs-lookup"><span data-stu-id="a9a27-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="a9a27-139">1.1</span><span class="sxs-lookup"><span data-stu-id="a9a27-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a9a27-140">requirements</span><span class="sxs-lookup"><span data-stu-id="a9a27-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="a9a27-141">撰写</span><span class="sxs-lookup"><span data-stu-id="a9a27-141">Compose</span></span><br><span data-ttu-id="a9a27-142">阅读</span><span class="sxs-lookup"><span data-stu-id="a9a27-142">Read</span></span> | [<span data-ttu-id="a9a27-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="a9a27-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="a9a27-144">1.1</span><span class="sxs-lookup"><span data-stu-id="a9a27-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a9a27-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="a9a27-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="a9a27-146">撰写</span><span class="sxs-lookup"><span data-stu-id="a9a27-146">Compose</span></span><br><span data-ttu-id="a9a27-147">阅读</span><span class="sxs-lookup"><span data-stu-id="a9a27-147">Read</span></span> | [<span data-ttu-id="a9a27-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="a9a27-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="a9a27-149">1.1</span><span class="sxs-lookup"><span data-stu-id="a9a27-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a9a27-150">ui</span><span class="sxs-lookup"><span data-stu-id="a9a27-150">ui</span></span>](#ui-ui) | <span data-ttu-id="a9a27-151">撰写</span><span class="sxs-lookup"><span data-stu-id="a9a27-151">Compose</span></span><br><span data-ttu-id="a9a27-152">阅读</span><span class="sxs-lookup"><span data-stu-id="a9a27-152">Read</span></span> | [<span data-ttu-id="a9a27-153">UI</span><span class="sxs-lookup"><span data-stu-id="a9a27-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="a9a27-154">1.1</span><span class="sxs-lookup"><span data-stu-id="a9a27-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="a9a27-155">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="a9a27-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="a9a27-156">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="a9a27-156">contentLanguage: String</span></span>

<span data-ttu-id="a9a27-157">获取用户 (编辑) 的语言区域设置。</span><span class="sxs-lookup"><span data-stu-id="a9a27-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="a9a27-158">该值 `contentLanguage` 反映当前在客户端 **应用程序中** 由 File **> Options > Language** 指定的Office设置。</span><span class="sxs-lookup"><span data-stu-id="a9a27-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="a9a27-159">类型</span><span class="sxs-lookup"><span data-stu-id="a9a27-159">Type</span></span>

*   <span data-ttu-id="a9a27-160">String</span><span class="sxs-lookup"><span data-stu-id="a9a27-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9a27-161">要求</span><span class="sxs-lookup"><span data-stu-id="a9a27-161">Requirements</span></span>

|<span data-ttu-id="a9a27-162">要求</span><span class="sxs-lookup"><span data-stu-id="a9a27-162">Requirement</span></span>| <span data-ttu-id="a9a27-163">值</span><span class="sxs-lookup"><span data-stu-id="a9a27-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9a27-164">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a9a27-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a9a27-165">1.1</span><span class="sxs-lookup"><span data-stu-id="a9a27-165">1.1</span></span>|
|[<span data-ttu-id="a9a27-166">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a9a27-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a9a27-167">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a9a27-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9a27-168">示例</span><span class="sxs-lookup"><span data-stu-id="a9a27-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="a9a27-169">diagnostics： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="a9a27-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="a9a27-170">获取加载项运行环境的信息。</span><span class="sxs-lookup"><span data-stu-id="a9a27-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="a9a27-171">类型</span><span class="sxs-lookup"><span data-stu-id="a9a27-171">Type</span></span>

*   [<span data-ttu-id="a9a27-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="a9a27-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="a9a27-173">Requirements</span><span class="sxs-lookup"><span data-stu-id="a9a27-173">Requirements</span></span>

|<span data-ttu-id="a9a27-174">要求</span><span class="sxs-lookup"><span data-stu-id="a9a27-174">Requirement</span></span>| <span data-ttu-id="a9a27-175">值</span><span class="sxs-lookup"><span data-stu-id="a9a27-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9a27-176">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a9a27-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a9a27-177">1.1</span><span class="sxs-lookup"><span data-stu-id="a9a27-177">1.1</span></span>|
|[<span data-ttu-id="a9a27-178">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a9a27-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a9a27-179">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a9a27-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9a27-180">示例</span><span class="sxs-lookup"><span data-stu-id="a9a27-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="a9a27-181">displayLanguage：String</span><span class="sxs-lookup"><span data-stu-id="a9a27-181">displayLanguage: String</span></span>

<span data-ttu-id="a9a27-182">获取区域设置 (语言) RFC 1766 语言标记格式，该标记格式由用户为 Office 客户端应用程序的 UI 指定。</span><span class="sxs-lookup"><span data-stu-id="a9a27-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="a9a27-183">该值反映当前显示语言设置，该设置由 > `displayLanguage` **客户端** 应用程序中>选项Office语言。 </span><span class="sxs-lookup"><span data-stu-id="a9a27-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="a9a27-184">类型</span><span class="sxs-lookup"><span data-stu-id="a9a27-184">Type</span></span>

*   <span data-ttu-id="a9a27-185">String</span><span class="sxs-lookup"><span data-stu-id="a9a27-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9a27-186">要求</span><span class="sxs-lookup"><span data-stu-id="a9a27-186">Requirements</span></span>

|<span data-ttu-id="a9a27-187">要求</span><span class="sxs-lookup"><span data-stu-id="a9a27-187">Requirement</span></span>| <span data-ttu-id="a9a27-188">值</span><span class="sxs-lookup"><span data-stu-id="a9a27-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9a27-189">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a9a27-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a9a27-190">1.1</span><span class="sxs-lookup"><span data-stu-id="a9a27-190">1.1</span></span>|
|[<span data-ttu-id="a9a27-191">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a9a27-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a9a27-192">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a9a27-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9a27-193">示例</span><span class="sxs-lookup"><span data-stu-id="a9a27-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="a9a27-194">requirements： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="a9a27-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="a9a27-195">提供用于确定当前应用程序和平台上支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="a9a27-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="a9a27-196">类型</span><span class="sxs-lookup"><span data-stu-id="a9a27-196">Type</span></span>

*   [<span data-ttu-id="a9a27-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="a9a27-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="a9a27-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="a9a27-198">Requirements</span></span>

|<span data-ttu-id="a9a27-199">要求</span><span class="sxs-lookup"><span data-stu-id="a9a27-199">Requirement</span></span>| <span data-ttu-id="a9a27-200">值</span><span class="sxs-lookup"><span data-stu-id="a9a27-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9a27-201">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a9a27-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a9a27-202">1.1</span><span class="sxs-lookup"><span data-stu-id="a9a27-202">1.1</span></span>|
|[<span data-ttu-id="a9a27-203">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a9a27-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a9a27-204">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a9a27-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a9a27-205">示例</span><span class="sxs-lookup"><span data-stu-id="a9a27-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="a9a27-206">[roamingSettings：RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="a9a27-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="a9a27-207">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="a9a27-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="a9a27-208">该对象允许您存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时可供该外接程序使用 `RoamingSettings` 。</span><span class="sxs-lookup"><span data-stu-id="a9a27-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="a9a27-209">类型</span><span class="sxs-lookup"><span data-stu-id="a9a27-209">Type</span></span>

*   [<span data-ttu-id="a9a27-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="a9a27-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="a9a27-211">Requirements</span><span class="sxs-lookup"><span data-stu-id="a9a27-211">Requirements</span></span>

|<span data-ttu-id="a9a27-212">要求</span><span class="sxs-lookup"><span data-stu-id="a9a27-212">Requirement</span></span>| <span data-ttu-id="a9a27-213">值</span><span class="sxs-lookup"><span data-stu-id="a9a27-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9a27-214">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a9a27-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a9a27-215">1.1</span><span class="sxs-lookup"><span data-stu-id="a9a27-215">1.1</span></span>|
|[<span data-ttu-id="a9a27-216">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a9a27-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="a9a27-217">受限</span><span class="sxs-lookup"><span data-stu-id="a9a27-217">Restricted</span></span>|
|[<span data-ttu-id="a9a27-218">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a9a27-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a9a27-219">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a9a27-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="a9a27-220">[ui：UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="a9a27-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="a9a27-221">提供可用于在加载项中创建和操作 UI 组件（如对话框）Office方法。</span><span class="sxs-lookup"><span data-stu-id="a9a27-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="a9a27-222">类型</span><span class="sxs-lookup"><span data-stu-id="a9a27-222">Type</span></span>

*   [<span data-ttu-id="a9a27-223">UI</span><span class="sxs-lookup"><span data-stu-id="a9a27-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="a9a27-224">Requirements</span><span class="sxs-lookup"><span data-stu-id="a9a27-224">Requirements</span></span>

|<span data-ttu-id="a9a27-225">要求</span><span class="sxs-lookup"><span data-stu-id="a9a27-225">Requirement</span></span>| <span data-ttu-id="a9a27-226">值</span><span class="sxs-lookup"><span data-stu-id="a9a27-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9a27-227">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a9a27-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a9a27-228">1.1</span><span class="sxs-lookup"><span data-stu-id="a9a27-228">1.1</span></span>|
|[<span data-ttu-id="a9a27-229">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a9a27-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a9a27-230">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a9a27-230">Compose or Read</span></span>|
