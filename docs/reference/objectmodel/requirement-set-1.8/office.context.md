---
title: Office.context - 要求集 1.8
description: Office。使用邮箱 API 要求集 1.8 Outlook外接程序可用的上下文对象成员。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 99573d9984c571c99461e90e8bdccdca35fe30b7
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590965"
---
# <a name="context-mailbox-requirement-set-18"></a><span data-ttu-id="c5bb0-103">context (Mailbox requirement set 1.8) </span><span class="sxs-lookup"><span data-stu-id="c5bb0-103">context (Mailbox requirement set 1.8)</span></span>

### <a name="officecontext"></a><span data-ttu-id="c5bb0-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="c5bb0-104">[Office](office.md).context</span></span>

<span data-ttu-id="c5bb0-105">Office.context 提供了外接程序在所有应用程序中使用的共享Office接口。</span><span class="sxs-lookup"><span data-stu-id="c5bb0-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="c5bb0-106">此列表仅记录外接程序Outlook接口。有关 Office.context 命名空间的完整列表，请参阅通用 API 中的[Office.context 引用](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true)。</span><span class="sxs-lookup"><span data-stu-id="c5bb0-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5bb0-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5bb0-107">Requirements</span></span>

|<span data-ttu-id="c5bb0-108">要求</span><span class="sxs-lookup"><span data-stu-id="c5bb0-108">Requirement</span></span>| <span data-ttu-id="c5bb0-109">值</span><span class="sxs-lookup"><span data-stu-id="c5bb0-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5bb0-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c5bb0-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5bb0-111">1.1</span><span class="sxs-lookup"><span data-stu-id="c5bb0-111">1.1</span></span>|
|[<span data-ttu-id="c5bb0-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c5bb0-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5bb0-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="c5bb0-114">属性</span><span class="sxs-lookup"><span data-stu-id="c5bb0-114">Properties</span></span>

| <span data-ttu-id="c5bb0-115">属性</span><span class="sxs-lookup"><span data-stu-id="c5bb0-115">Property</span></span> | <span data-ttu-id="c5bb0-116">模式</span><span class="sxs-lookup"><span data-stu-id="c5bb0-116">Modes</span></span> | <span data-ttu-id="c5bb0-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="c5bb0-117">Return type</span></span> | <span data-ttu-id="c5bb0-118">最小值</span><span class="sxs-lookup"><span data-stu-id="c5bb0-118">Minimum</span></span><br><span data-ttu-id="c5bb0-119">要求集</span><span class="sxs-lookup"><span data-stu-id="c5bb0-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="c5bb0-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="c5bb0-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="c5bb0-121">撰写</span><span class="sxs-lookup"><span data-stu-id="c5bb0-121">Compose</span></span><br><span data-ttu-id="c5bb0-122">阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-122">Read</span></span> | <span data-ttu-id="c5bb0-123">字符串</span><span class="sxs-lookup"><span data-stu-id="c5bb0-123">String</span></span> | [<span data-ttu-id="c5bb0-124">1.1</span><span class="sxs-lookup"><span data-stu-id="c5bb0-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5bb0-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="c5bb0-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="c5bb0-126">撰写</span><span class="sxs-lookup"><span data-stu-id="c5bb0-126">Compose</span></span><br><span data-ttu-id="c5bb0-127">阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-127">Read</span></span> | [<span data-ttu-id="c5bb0-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="c5bb0-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="c5bb0-129">1.1</span><span class="sxs-lookup"><span data-stu-id="c5bb0-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5bb0-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="c5bb0-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="c5bb0-131">撰写</span><span class="sxs-lookup"><span data-stu-id="c5bb0-131">Compose</span></span><br><span data-ttu-id="c5bb0-132">阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-132">Read</span></span> | <span data-ttu-id="c5bb0-133">字符串</span><span class="sxs-lookup"><span data-stu-id="c5bb0-133">String</span></span> | [<span data-ttu-id="c5bb0-134">1.1</span><span class="sxs-lookup"><span data-stu-id="c5bb0-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5bb0-135">host</span><span class="sxs-lookup"><span data-stu-id="c5bb0-135">host</span></span>](#host-hosttype) | <span data-ttu-id="c5bb0-136">撰写</span><span class="sxs-lookup"><span data-stu-id="c5bb0-136">Compose</span></span><br><span data-ttu-id="c5bb0-137">阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-137">Read</span></span> | [<span data-ttu-id="c5bb0-138">HostType</span><span class="sxs-lookup"><span data-stu-id="c5bb0-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="c5bb0-139">1.5</span><span class="sxs-lookup"><span data-stu-id="c5bb0-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="c5bb0-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="c5bb0-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="c5bb0-141">撰写</span><span class="sxs-lookup"><span data-stu-id="c5bb0-141">Compose</span></span><br><span data-ttu-id="c5bb0-142">阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-142">Read</span></span> | [<span data-ttu-id="c5bb0-143">邮箱</span><span class="sxs-lookup"><span data-stu-id="c5bb0-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="c5bb0-144">1.1</span><span class="sxs-lookup"><span data-stu-id="c5bb0-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5bb0-145">平台</span><span class="sxs-lookup"><span data-stu-id="c5bb0-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="c5bb0-146">撰写</span><span class="sxs-lookup"><span data-stu-id="c5bb0-146">Compose</span></span><br><span data-ttu-id="c5bb0-147">阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-147">Read</span></span> | [<span data-ttu-id="c5bb0-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="c5bb0-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="c5bb0-149">1.5</span><span class="sxs-lookup"><span data-stu-id="c5bb0-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="c5bb0-150">requirements</span><span class="sxs-lookup"><span data-stu-id="c5bb0-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="c5bb0-151">撰写</span><span class="sxs-lookup"><span data-stu-id="c5bb0-151">Compose</span></span><br><span data-ttu-id="c5bb0-152">阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-152">Read</span></span> | [<span data-ttu-id="c5bb0-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="c5bb0-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="c5bb0-154">1.1</span><span class="sxs-lookup"><span data-stu-id="c5bb0-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5bb0-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="c5bb0-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="c5bb0-156">撰写</span><span class="sxs-lookup"><span data-stu-id="c5bb0-156">Compose</span></span><br><span data-ttu-id="c5bb0-157">阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-157">Read</span></span> | [<span data-ttu-id="c5bb0-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c5bb0-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="c5bb0-159">1.1</span><span class="sxs-lookup"><span data-stu-id="c5bb0-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5bb0-160">ui</span><span class="sxs-lookup"><span data-stu-id="c5bb0-160">ui</span></span>](#ui-ui) | <span data-ttu-id="c5bb0-161">撰写</span><span class="sxs-lookup"><span data-stu-id="c5bb0-161">Compose</span></span><br><span data-ttu-id="c5bb0-162">阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-162">Read</span></span> | [<span data-ttu-id="c5bb0-163">UI</span><span class="sxs-lookup"><span data-stu-id="c5bb0-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="c5bb0-164">1.1</span><span class="sxs-lookup"><span data-stu-id="c5bb0-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="c5bb0-165">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="c5bb0-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="c5bb0-166">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="c5bb0-166">contentLanguage: String</span></span>

<span data-ttu-id="c5bb0-167">获取用户 (编辑) 的语言区域设置。</span><span class="sxs-lookup"><span data-stu-id="c5bb0-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="c5bb0-168">该值 `contentLanguage` 反映当前在客户端 **应用程序中** 由 File **> Options > Language** 指定的Office设置。</span><span class="sxs-lookup"><span data-stu-id="c5bb0-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="c5bb0-169">类型</span><span class="sxs-lookup"><span data-stu-id="c5bb0-169">Type</span></span>

*   <span data-ttu-id="c5bb0-170">String</span><span class="sxs-lookup"><span data-stu-id="c5bb0-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5bb0-171">要求</span><span class="sxs-lookup"><span data-stu-id="c5bb0-171">Requirements</span></span>

|<span data-ttu-id="c5bb0-172">要求</span><span class="sxs-lookup"><span data-stu-id="c5bb0-172">Requirement</span></span>| <span data-ttu-id="c5bb0-173">值</span><span class="sxs-lookup"><span data-stu-id="c5bb0-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5bb0-174">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c5bb0-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5bb0-175">1.1</span><span class="sxs-lookup"><span data-stu-id="c5bb0-175">1.1</span></span>|
|[<span data-ttu-id="c5bb0-176">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c5bb0-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5bb0-177">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5bb0-178">示例</span><span class="sxs-lookup"><span data-stu-id="c5bb0-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="c5bb0-179">diagnostics： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="c5bb0-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="c5bb0-180">获取加载项运行环境的信息。</span><span class="sxs-lookup"><span data-stu-id="c5bb0-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="c5bb0-181">类型</span><span class="sxs-lookup"><span data-stu-id="c5bb0-181">Type</span></span>

*   [<span data-ttu-id="c5bb0-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="c5bb0-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="c5bb0-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5bb0-183">Requirements</span></span>

|<span data-ttu-id="c5bb0-184">要求</span><span class="sxs-lookup"><span data-stu-id="c5bb0-184">Requirement</span></span>| <span data-ttu-id="c5bb0-185">值</span><span class="sxs-lookup"><span data-stu-id="c5bb0-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5bb0-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c5bb0-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5bb0-187">1.1</span><span class="sxs-lookup"><span data-stu-id="c5bb0-187">1.1</span></span>|
|[<span data-ttu-id="c5bb0-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c5bb0-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5bb0-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5bb0-190">示例</span><span class="sxs-lookup"><span data-stu-id="c5bb0-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="c5bb0-191">displayLanguage：String</span><span class="sxs-lookup"><span data-stu-id="c5bb0-191">displayLanguage: String</span></span>

<span data-ttu-id="c5bb0-192">获取区域设置 (语言) RFC 1766 语言标记格式，该标记格式由用户为 Office 客户端应用程序的 UI 指定。</span><span class="sxs-lookup"><span data-stu-id="c5bb0-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="c5bb0-193">该值反映当前显示语言设置，该设置由 > `displayLanguage` **客户端** 应用程序中>选项Office语言。 </span><span class="sxs-lookup"><span data-stu-id="c5bb0-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="c5bb0-194">类型</span><span class="sxs-lookup"><span data-stu-id="c5bb0-194">Type</span></span>

*   <span data-ttu-id="c5bb0-195">String</span><span class="sxs-lookup"><span data-stu-id="c5bb0-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5bb0-196">要求</span><span class="sxs-lookup"><span data-stu-id="c5bb0-196">Requirements</span></span>

|<span data-ttu-id="c5bb0-197">要求</span><span class="sxs-lookup"><span data-stu-id="c5bb0-197">Requirement</span></span>| <span data-ttu-id="c5bb0-198">值</span><span class="sxs-lookup"><span data-stu-id="c5bb0-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5bb0-199">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c5bb0-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5bb0-200">1.1</span><span class="sxs-lookup"><span data-stu-id="c5bb0-200">1.1</span></span>|
|[<span data-ttu-id="c5bb0-201">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c5bb0-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5bb0-202">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5bb0-203">示例</span><span class="sxs-lookup"><span data-stu-id="c5bb0-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="c5bb0-204">host： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="c5bb0-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="c5bb0-205">获取Office加载项的加载项应用程序。</span><span class="sxs-lookup"><span data-stu-id="c5bb0-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="c5bb0-206">或者，您可以使用[Office.context.diagnostics](#diagnostics-contextinformation)属性获取主机。</span><span class="sxs-lookup"><span data-stu-id="c5bb0-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="c5bb0-207">类型</span><span class="sxs-lookup"><span data-stu-id="c5bb0-207">Type</span></span>

*   [<span data-ttu-id="c5bb0-208">HostType</span><span class="sxs-lookup"><span data-stu-id="c5bb0-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="c5bb0-209">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5bb0-209">Requirements</span></span>

|<span data-ttu-id="c5bb0-210">要求</span><span class="sxs-lookup"><span data-stu-id="c5bb0-210">Requirement</span></span>| <span data-ttu-id="c5bb0-211">值</span><span class="sxs-lookup"><span data-stu-id="c5bb0-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5bb0-212">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c5bb0-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5bb0-213">1.5</span><span class="sxs-lookup"><span data-stu-id="c5bb0-213">1.5</span></span>|
|[<span data-ttu-id="c5bb0-214">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c5bb0-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5bb0-215">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5bb0-216">示例</span><span class="sxs-lookup"><span data-stu-id="c5bb0-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="c5bb0-217">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="c5bb0-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="c5bb0-218">提供运行加载项的平台。</span><span class="sxs-lookup"><span data-stu-id="c5bb0-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="c5bb0-219">或者，您可以使用[Office.context.diagnostics](#diagnostics-contextinformation)属性获取平台。</span><span class="sxs-lookup"><span data-stu-id="c5bb0-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="c5bb0-220">类型</span><span class="sxs-lookup"><span data-stu-id="c5bb0-220">Type</span></span>

*   [<span data-ttu-id="c5bb0-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="c5bb0-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="c5bb0-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5bb0-222">Requirements</span></span>

|<span data-ttu-id="c5bb0-223">要求</span><span class="sxs-lookup"><span data-stu-id="c5bb0-223">Requirement</span></span>| <span data-ttu-id="c5bb0-224">值</span><span class="sxs-lookup"><span data-stu-id="c5bb0-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5bb0-225">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c5bb0-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5bb0-226">1.5</span><span class="sxs-lookup"><span data-stu-id="c5bb0-226">1.5</span></span>|
|[<span data-ttu-id="c5bb0-227">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c5bb0-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5bb0-228">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5bb0-229">示例</span><span class="sxs-lookup"><span data-stu-id="c5bb0-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="c5bb0-230">requirements： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="c5bb0-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="c5bb0-231">提供用于确定当前应用程序和平台上支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="c5bb0-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="c5bb0-232">类型</span><span class="sxs-lookup"><span data-stu-id="c5bb0-232">Type</span></span>

*   [<span data-ttu-id="c5bb0-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="c5bb0-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="c5bb0-234">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5bb0-234">Requirements</span></span>

|<span data-ttu-id="c5bb0-235">要求</span><span class="sxs-lookup"><span data-stu-id="c5bb0-235">Requirement</span></span>| <span data-ttu-id="c5bb0-236">值</span><span class="sxs-lookup"><span data-stu-id="c5bb0-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5bb0-237">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c5bb0-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5bb0-238">1.1</span><span class="sxs-lookup"><span data-stu-id="c5bb0-238">1.1</span></span>|
|[<span data-ttu-id="c5bb0-239">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c5bb0-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5bb0-240">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5bb0-241">示例</span><span class="sxs-lookup"><span data-stu-id="c5bb0-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="c5bb0-242">[roamingSettings：RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="c5bb0-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="c5bb0-243">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="c5bb0-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="c5bb0-244">该对象允许您存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时可供该外接程序使用 `RoamingSettings` 。</span><span class="sxs-lookup"><span data-stu-id="c5bb0-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="c5bb0-245">类型</span><span class="sxs-lookup"><span data-stu-id="c5bb0-245">Type</span></span>

*   [<span data-ttu-id="c5bb0-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c5bb0-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="c5bb0-247">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5bb0-247">Requirements</span></span>

|<span data-ttu-id="c5bb0-248">要求</span><span class="sxs-lookup"><span data-stu-id="c5bb0-248">Requirement</span></span>| <span data-ttu-id="c5bb0-249">值</span><span class="sxs-lookup"><span data-stu-id="c5bb0-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5bb0-250">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c5bb0-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5bb0-251">1.1</span><span class="sxs-lookup"><span data-stu-id="c5bb0-251">1.1</span></span>|
|[<span data-ttu-id="c5bb0-252">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c5bb0-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="c5bb0-253">受限</span><span class="sxs-lookup"><span data-stu-id="c5bb0-253">Restricted</span></span>|
|[<span data-ttu-id="c5bb0-254">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c5bb0-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5bb0-255">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="c5bb0-256">[ui：UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="c5bb0-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="c5bb0-257">提供可用于在加载项中创建和操作 UI 组件（如对话框）Office方法。</span><span class="sxs-lookup"><span data-stu-id="c5bb0-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="c5bb0-258">类型</span><span class="sxs-lookup"><span data-stu-id="c5bb0-258">Type</span></span>

*   [<span data-ttu-id="c5bb0-259">UI</span><span class="sxs-lookup"><span data-stu-id="c5bb0-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="c5bb0-260">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5bb0-260">Requirements</span></span>

|<span data-ttu-id="c5bb0-261">要求</span><span class="sxs-lookup"><span data-stu-id="c5bb0-261">Requirement</span></span>| <span data-ttu-id="c5bb0-262">值</span><span class="sxs-lookup"><span data-stu-id="c5bb0-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5bb0-263">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c5bb0-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5bb0-264">1.1</span><span class="sxs-lookup"><span data-stu-id="c5bb0-264">1.1</span></span>|
|[<span data-ttu-id="c5bb0-265">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c5bb0-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5bb0-266">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c5bb0-266">Compose or Read</span></span>|
