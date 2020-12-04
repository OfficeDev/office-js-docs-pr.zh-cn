---
title: Office。上下文要求集1。3
description: 使用邮箱 API 要求集1.3 的 Outlook 外接程序可用的 Office 对象成员。
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: b497cdf3f878df7efd816f236bd565c8fad7d922
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570742"
---
# <a name="context-mailbox-requirement-set-13"></a><span data-ttu-id="fc9d0-103"> (邮箱要求集1.3 的上下文) </span><span class="sxs-lookup"><span data-stu-id="fc9d0-103">context (Mailbox requirement set 1.3)</span></span>

### <a name="officecontext"></a><span data-ttu-id="fc9d0-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="fc9d0-104">[Office](office.md).context</span></span>

<span data-ttu-id="fc9d0-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="fc9d0-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="fc9d0-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅 [通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.3&preserve-view=true)"。</span><span class="sxs-lookup"><span data-stu-id="fc9d0-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.3&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="fc9d0-107">要求</span><span class="sxs-lookup"><span data-stu-id="fc9d0-107">Requirements</span></span>

|<span data-ttu-id="fc9d0-108">要求</span><span class="sxs-lookup"><span data-stu-id="fc9d0-108">Requirement</span></span>| <span data-ttu-id="fc9d0-109">值</span><span class="sxs-lookup"><span data-stu-id="fc9d0-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc9d0-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fc9d0-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fc9d0-111">1.1</span><span class="sxs-lookup"><span data-stu-id="fc9d0-111">1.1</span></span>|
|[<span data-ttu-id="fc9d0-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fc9d0-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="fc9d0-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fc9d0-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="fc9d0-114">属性</span><span class="sxs-lookup"><span data-stu-id="fc9d0-114">Properties</span></span>

| <span data-ttu-id="fc9d0-115">属性</span><span class="sxs-lookup"><span data-stu-id="fc9d0-115">Property</span></span> | <span data-ttu-id="fc9d0-116">型号</span><span class="sxs-lookup"><span data-stu-id="fc9d0-116">Modes</span></span> | <span data-ttu-id="fc9d0-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="fc9d0-117">Return type</span></span> | <span data-ttu-id="fc9d0-118">最小值</span><span class="sxs-lookup"><span data-stu-id="fc9d0-118">Minimum</span></span><br><span data-ttu-id="fc9d0-119">要求集</span><span class="sxs-lookup"><span data-stu-id="fc9d0-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="fc9d0-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="fc9d0-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="fc9d0-121">撰写</span><span class="sxs-lookup"><span data-stu-id="fc9d0-121">Compose</span></span><br><span data-ttu-id="fc9d0-122">阅读</span><span class="sxs-lookup"><span data-stu-id="fc9d0-122">Read</span></span> | <span data-ttu-id="fc9d0-123">String</span><span class="sxs-lookup"><span data-stu-id="fc9d0-123">String</span></span> | [<span data-ttu-id="fc9d0-124">1.1</span><span class="sxs-lookup"><span data-stu-id="fc9d0-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="fc9d0-125">过程</span><span class="sxs-lookup"><span data-stu-id="fc9d0-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="fc9d0-126">撰写</span><span class="sxs-lookup"><span data-stu-id="fc9d0-126">Compose</span></span><br><span data-ttu-id="fc9d0-127">阅读</span><span class="sxs-lookup"><span data-stu-id="fc9d0-127">Read</span></span> | [<span data-ttu-id="fc9d0-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="fc9d0-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="fc9d0-129">1.1</span><span class="sxs-lookup"><span data-stu-id="fc9d0-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="fc9d0-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="fc9d0-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="fc9d0-131">撰写</span><span class="sxs-lookup"><span data-stu-id="fc9d0-131">Compose</span></span><br><span data-ttu-id="fc9d0-132">阅读</span><span class="sxs-lookup"><span data-stu-id="fc9d0-132">Read</span></span> | <span data-ttu-id="fc9d0-133">String</span><span class="sxs-lookup"><span data-stu-id="fc9d0-133">String</span></span> | [<span data-ttu-id="fc9d0-134">1.1</span><span class="sxs-lookup"><span data-stu-id="fc9d0-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="fc9d0-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="fc9d0-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="fc9d0-136">撰写</span><span class="sxs-lookup"><span data-stu-id="fc9d0-136">Compose</span></span><br><span data-ttu-id="fc9d0-137">阅读</span><span class="sxs-lookup"><span data-stu-id="fc9d0-137">Read</span></span> | [<span data-ttu-id="fc9d0-138">邮箱</span><span class="sxs-lookup"><span data-stu-id="fc9d0-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="fc9d0-139">1.1</span><span class="sxs-lookup"><span data-stu-id="fc9d0-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="fc9d0-140">满足</span><span class="sxs-lookup"><span data-stu-id="fc9d0-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="fc9d0-141">撰写</span><span class="sxs-lookup"><span data-stu-id="fc9d0-141">Compose</span></span><br><span data-ttu-id="fc9d0-142">阅读</span><span class="sxs-lookup"><span data-stu-id="fc9d0-142">Read</span></span> | [<span data-ttu-id="fc9d0-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="fc9d0-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="fc9d0-144">1.1</span><span class="sxs-lookup"><span data-stu-id="fc9d0-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="fc9d0-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="fc9d0-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="fc9d0-146">撰写</span><span class="sxs-lookup"><span data-stu-id="fc9d0-146">Compose</span></span><br><span data-ttu-id="fc9d0-147">阅读</span><span class="sxs-lookup"><span data-stu-id="fc9d0-147">Read</span></span> | [<span data-ttu-id="fc9d0-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="fc9d0-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="fc9d0-149">1.1</span><span class="sxs-lookup"><span data-stu-id="fc9d0-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="fc9d0-150">ui</span><span class="sxs-lookup"><span data-stu-id="fc9d0-150">ui</span></span>](#ui-ui) | <span data-ttu-id="fc9d0-151">撰写</span><span class="sxs-lookup"><span data-stu-id="fc9d0-151">Compose</span></span><br><span data-ttu-id="fc9d0-152">阅读</span><span class="sxs-lookup"><span data-stu-id="fc9d0-152">Read</span></span> | [<span data-ttu-id="fc9d0-153">UI</span><span class="sxs-lookup"><span data-stu-id="fc9d0-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="fc9d0-154">1.1</span><span class="sxs-lookup"><span data-stu-id="fc9d0-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="fc9d0-155">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="fc9d0-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="fc9d0-156">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="fc9d0-156">contentLanguage: String</span></span>

<span data-ttu-id="fc9d0-157">获取用户指定的用于编辑项目的区域设置 (语言) 。</span><span class="sxs-lookup"><span data-stu-id="fc9d0-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="fc9d0-158">此 `contentLanguage` 值反映了使用 Office 客户端应用程序中的 "**文件 > 选项" > 语言** 指定的当前 **编辑语言** 设置。</span><span class="sxs-lookup"><span data-stu-id="fc9d0-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="fc9d0-159">类型</span><span class="sxs-lookup"><span data-stu-id="fc9d0-159">Type</span></span>

*   <span data-ttu-id="fc9d0-160">String</span><span class="sxs-lookup"><span data-stu-id="fc9d0-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fc9d0-161">要求</span><span class="sxs-lookup"><span data-stu-id="fc9d0-161">Requirements</span></span>

|<span data-ttu-id="fc9d0-162">要求</span><span class="sxs-lookup"><span data-stu-id="fc9d0-162">Requirement</span></span>| <span data-ttu-id="fc9d0-163">值</span><span class="sxs-lookup"><span data-stu-id="fc9d0-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc9d0-164">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fc9d0-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fc9d0-165">1.1</span><span class="sxs-lookup"><span data-stu-id="fc9d0-165">1.1</span></span>|
|[<span data-ttu-id="fc9d0-166">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fc9d0-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="fc9d0-167">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fc9d0-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fc9d0-168">示例</span><span class="sxs-lookup"><span data-stu-id="fc9d0-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="fc9d0-169">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="fc9d0-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="fc9d0-170">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="fc9d0-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="fc9d0-171">类型</span><span class="sxs-lookup"><span data-stu-id="fc9d0-171">Type</span></span>

*   [<span data-ttu-id="fc9d0-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="fc9d0-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="fc9d0-173">要求</span><span class="sxs-lookup"><span data-stu-id="fc9d0-173">Requirements</span></span>

|<span data-ttu-id="fc9d0-174">要求</span><span class="sxs-lookup"><span data-stu-id="fc9d0-174">Requirement</span></span>| <span data-ttu-id="fc9d0-175">值</span><span class="sxs-lookup"><span data-stu-id="fc9d0-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc9d0-176">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fc9d0-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fc9d0-177">1.1</span><span class="sxs-lookup"><span data-stu-id="fc9d0-177">1.1</span></span>|
|[<span data-ttu-id="fc9d0-178">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fc9d0-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="fc9d0-179">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fc9d0-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fc9d0-180">示例</span><span class="sxs-lookup"><span data-stu-id="fc9d0-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="fc9d0-181">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="fc9d0-181">displayLanguage: String</span></span>

<span data-ttu-id="fc9d0-182">获取用户为 Office 客户端应用程序的 UI 指定的 RFC 1766 语言标记格式中 (语言) 的区域设置。</span><span class="sxs-lookup"><span data-stu-id="fc9d0-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="fc9d0-183">此 `displayLanguage` 值反映了使用 Office 客户端应用程序中的 **File > Options > 语言** 指定的当前 **显示语言** 设置。</span><span class="sxs-lookup"><span data-stu-id="fc9d0-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="fc9d0-184">类型</span><span class="sxs-lookup"><span data-stu-id="fc9d0-184">Type</span></span>

*   <span data-ttu-id="fc9d0-185">String</span><span class="sxs-lookup"><span data-stu-id="fc9d0-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fc9d0-186">要求</span><span class="sxs-lookup"><span data-stu-id="fc9d0-186">Requirements</span></span>

|<span data-ttu-id="fc9d0-187">要求</span><span class="sxs-lookup"><span data-stu-id="fc9d0-187">Requirement</span></span>| <span data-ttu-id="fc9d0-188">值</span><span class="sxs-lookup"><span data-stu-id="fc9d0-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc9d0-189">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fc9d0-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fc9d0-190">1.1</span><span class="sxs-lookup"><span data-stu-id="fc9d0-190">1.1</span></span>|
|[<span data-ttu-id="fc9d0-191">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fc9d0-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="fc9d0-192">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fc9d0-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fc9d0-193">示例</span><span class="sxs-lookup"><span data-stu-id="fc9d0-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="fc9d0-194">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="fc9d0-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="fc9d0-195">提供用于确定当前应用程序和平台支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="fc9d0-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="fc9d0-196">类型</span><span class="sxs-lookup"><span data-stu-id="fc9d0-196">Type</span></span>

*   [<span data-ttu-id="fc9d0-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="fc9d0-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="fc9d0-198">要求</span><span class="sxs-lookup"><span data-stu-id="fc9d0-198">Requirements</span></span>

|<span data-ttu-id="fc9d0-199">要求</span><span class="sxs-lookup"><span data-stu-id="fc9d0-199">Requirement</span></span>| <span data-ttu-id="fc9d0-200">值</span><span class="sxs-lookup"><span data-stu-id="fc9d0-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc9d0-201">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fc9d0-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fc9d0-202">1.1</span><span class="sxs-lookup"><span data-stu-id="fc9d0-202">1.1</span></span>|
|[<span data-ttu-id="fc9d0-203">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fc9d0-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="fc9d0-204">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fc9d0-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fc9d0-205">示例</span><span class="sxs-lookup"><span data-stu-id="fc9d0-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="fc9d0-206">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="fc9d0-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="fc9d0-207">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="fc9d0-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="fc9d0-208">该 `RoamingSettings` 对象使您可以存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时都可使用该加载项。</span><span class="sxs-lookup"><span data-stu-id="fc9d0-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="fc9d0-209">类型</span><span class="sxs-lookup"><span data-stu-id="fc9d0-209">Type</span></span>

*   [<span data-ttu-id="fc9d0-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="fc9d0-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="fc9d0-211">要求</span><span class="sxs-lookup"><span data-stu-id="fc9d0-211">Requirements</span></span>

|<span data-ttu-id="fc9d0-212">要求</span><span class="sxs-lookup"><span data-stu-id="fc9d0-212">Requirement</span></span>| <span data-ttu-id="fc9d0-213">值</span><span class="sxs-lookup"><span data-stu-id="fc9d0-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc9d0-214">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fc9d0-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fc9d0-215">1.1</span><span class="sxs-lookup"><span data-stu-id="fc9d0-215">1.1</span></span>|
|[<span data-ttu-id="fc9d0-216">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fc9d0-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="fc9d0-217">受限</span><span class="sxs-lookup"><span data-stu-id="fc9d0-217">Restricted</span></span>|
|[<span data-ttu-id="fc9d0-218">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fc9d0-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="fc9d0-219">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fc9d0-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="fc9d0-220">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="fc9d0-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="fc9d0-221">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="fc9d0-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="fc9d0-222">类型</span><span class="sxs-lookup"><span data-stu-id="fc9d0-222">Type</span></span>

*   [<span data-ttu-id="fc9d0-223">UI</span><span class="sxs-lookup"><span data-stu-id="fc9d0-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="fc9d0-224">要求</span><span class="sxs-lookup"><span data-stu-id="fc9d0-224">Requirements</span></span>

|<span data-ttu-id="fc9d0-225">要求</span><span class="sxs-lookup"><span data-stu-id="fc9d0-225">Requirement</span></span>| <span data-ttu-id="fc9d0-226">值</span><span class="sxs-lookup"><span data-stu-id="fc9d0-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc9d0-227">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fc9d0-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fc9d0-228">1.1</span><span class="sxs-lookup"><span data-stu-id="fc9d0-228">1.1</span></span>|
|[<span data-ttu-id="fc9d0-229">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fc9d0-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="fc9d0-230">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fc9d0-230">Compose or Read</span></span>|
