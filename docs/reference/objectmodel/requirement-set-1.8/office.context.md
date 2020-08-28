---
title: Office。上下文要求集1。8
description: 使用邮箱 API 要求集1.8 的 Outlook 外接程序可用的 Office 对象成员。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 09c298f6c4e793bc52e87e4892143d174bb2656b
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293665"
---
# <a name="context-mailbox-requirement-set-18"></a><span data-ttu-id="381eb-103"> (邮箱要求集1.8 的上下文) </span><span class="sxs-lookup"><span data-stu-id="381eb-103">context (Mailbox requirement set 1.8)</span></span>

### <a name="officecontext"></a><span data-ttu-id="381eb-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="381eb-104">[Office](office.md).context</span></span>

<span data-ttu-id="381eb-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="381eb-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="381eb-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅 [通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.8)"。</span><span class="sxs-lookup"><span data-stu-id="381eb-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.8).</span></span>

##### <a name="requirements"></a><span data-ttu-id="381eb-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="381eb-107">Requirements</span></span>

|<span data-ttu-id="381eb-108">要求</span><span class="sxs-lookup"><span data-stu-id="381eb-108">Requirement</span></span>| <span data-ttu-id="381eb-109">值</span><span class="sxs-lookup"><span data-stu-id="381eb-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="381eb-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="381eb-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="381eb-111">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-111">1.1</span></span>|
|[<span data-ttu-id="381eb-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="381eb-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="381eb-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="381eb-114">属性</span><span class="sxs-lookup"><span data-stu-id="381eb-114">Properties</span></span>

| <span data-ttu-id="381eb-115">属性</span><span class="sxs-lookup"><span data-stu-id="381eb-115">Property</span></span> | <span data-ttu-id="381eb-116">型号</span><span class="sxs-lookup"><span data-stu-id="381eb-116">Modes</span></span> | <span data-ttu-id="381eb-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="381eb-117">Return type</span></span> | <span data-ttu-id="381eb-118">最小值</span><span class="sxs-lookup"><span data-stu-id="381eb-118">Minimum</span></span><br><span data-ttu-id="381eb-119">要求集</span><span class="sxs-lookup"><span data-stu-id="381eb-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="381eb-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="381eb-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="381eb-121">撰写</span><span class="sxs-lookup"><span data-stu-id="381eb-121">Compose</span></span><br><span data-ttu-id="381eb-122">阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-122">Read</span></span> | <span data-ttu-id="381eb-123">String</span><span class="sxs-lookup"><span data-stu-id="381eb-123">String</span></span> | [<span data-ttu-id="381eb-124">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="381eb-125">过程</span><span class="sxs-lookup"><span data-stu-id="381eb-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="381eb-126">撰写</span><span class="sxs-lookup"><span data-stu-id="381eb-126">Compose</span></span><br><span data-ttu-id="381eb-127">阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-127">Read</span></span> | [<span data-ttu-id="381eb-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="381eb-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.8) | [<span data-ttu-id="381eb-129">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="381eb-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="381eb-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="381eb-131">撰写</span><span class="sxs-lookup"><span data-stu-id="381eb-131">Compose</span></span><br><span data-ttu-id="381eb-132">阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-132">Read</span></span> | <span data-ttu-id="381eb-133">String</span><span class="sxs-lookup"><span data-stu-id="381eb-133">String</span></span> | [<span data-ttu-id="381eb-134">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="381eb-135">host</span><span class="sxs-lookup"><span data-stu-id="381eb-135">host</span></span>](#host-hosttype) | <span data-ttu-id="381eb-136">撰写</span><span class="sxs-lookup"><span data-stu-id="381eb-136">Compose</span></span><br><span data-ttu-id="381eb-137">阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-137">Read</span></span> | [<span data-ttu-id="381eb-138">HostType</span><span class="sxs-lookup"><span data-stu-id="381eb-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.8) | [<span data-ttu-id="381eb-139">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="381eb-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="381eb-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="381eb-141">撰写</span><span class="sxs-lookup"><span data-stu-id="381eb-141">Compose</span></span><br><span data-ttu-id="381eb-142">阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-142">Read</span></span> | [<span data-ttu-id="381eb-143">邮箱</span><span class="sxs-lookup"><span data-stu-id="381eb-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8) | [<span data-ttu-id="381eb-144">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="381eb-145">平台</span><span class="sxs-lookup"><span data-stu-id="381eb-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="381eb-146">撰写</span><span class="sxs-lookup"><span data-stu-id="381eb-146">Compose</span></span><br><span data-ttu-id="381eb-147">阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-147">Read</span></span> | [<span data-ttu-id="381eb-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="381eb-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.8) | [<span data-ttu-id="381eb-149">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="381eb-150">满足</span><span class="sxs-lookup"><span data-stu-id="381eb-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="381eb-151">撰写</span><span class="sxs-lookup"><span data-stu-id="381eb-151">Compose</span></span><br><span data-ttu-id="381eb-152">阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-152">Read</span></span> | [<span data-ttu-id="381eb-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="381eb-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.8) | [<span data-ttu-id="381eb-154">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="381eb-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="381eb-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="381eb-156">撰写</span><span class="sxs-lookup"><span data-stu-id="381eb-156">Compose</span></span><br><span data-ttu-id="381eb-157">阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-157">Read</span></span> | [<span data-ttu-id="381eb-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="381eb-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.8) | [<span data-ttu-id="381eb-159">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="381eb-160">ui</span><span class="sxs-lookup"><span data-stu-id="381eb-160">ui</span></span>](#ui-ui) | <span data-ttu-id="381eb-161">撰写</span><span class="sxs-lookup"><span data-stu-id="381eb-161">Compose</span></span><br><span data-ttu-id="381eb-162">阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-162">Read</span></span> | [<span data-ttu-id="381eb-163">UI</span><span class="sxs-lookup"><span data-stu-id="381eb-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.8) | [<span data-ttu-id="381eb-164">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="381eb-165">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="381eb-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="381eb-166">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="381eb-166">contentLanguage: String</span></span>

<span data-ttu-id="381eb-167">获取用户指定的用于编辑项目的区域设置 (语言) 。</span><span class="sxs-lookup"><span data-stu-id="381eb-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="381eb-168">此 `contentLanguage` 值反映了使用 Office 客户端应用程序中的 "**文件 > 选项" > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="381eb-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="381eb-169">类型</span><span class="sxs-lookup"><span data-stu-id="381eb-169">Type</span></span>

*   <span data-ttu-id="381eb-170">String</span><span class="sxs-lookup"><span data-stu-id="381eb-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="381eb-171">要求</span><span class="sxs-lookup"><span data-stu-id="381eb-171">Requirements</span></span>

|<span data-ttu-id="381eb-172">要求</span><span class="sxs-lookup"><span data-stu-id="381eb-172">Requirement</span></span>| <span data-ttu-id="381eb-173">值</span><span class="sxs-lookup"><span data-stu-id="381eb-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="381eb-174">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="381eb-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="381eb-175">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-175">1.1</span></span>|
|[<span data-ttu-id="381eb-176">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="381eb-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="381eb-177">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="381eb-178">示例</span><span class="sxs-lookup"><span data-stu-id="381eb-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="381eb-179">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="381eb-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="381eb-180">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="381eb-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="381eb-181">类型</span><span class="sxs-lookup"><span data-stu-id="381eb-181">Type</span></span>

*   [<span data-ttu-id="381eb-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="381eb-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="381eb-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="381eb-183">Requirements</span></span>

|<span data-ttu-id="381eb-184">要求</span><span class="sxs-lookup"><span data-stu-id="381eb-184">Requirement</span></span>| <span data-ttu-id="381eb-185">值</span><span class="sxs-lookup"><span data-stu-id="381eb-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="381eb-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="381eb-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="381eb-187">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-187">1.1</span></span>|
|[<span data-ttu-id="381eb-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="381eb-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="381eb-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="381eb-190">示例</span><span class="sxs-lookup"><span data-stu-id="381eb-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="381eb-191">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="381eb-191">displayLanguage: String</span></span>

<span data-ttu-id="381eb-192">获取用户为 Office 客户端应用程序的 UI 指定的 RFC 1766 语言标记格式中 (语言) 的区域设置。</span><span class="sxs-lookup"><span data-stu-id="381eb-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="381eb-193">此 `displayLanguage` 值反映了使用 Office 客户端应用程序中的**File > Options > 语言**指定的当前**显示语言**设置。</span><span class="sxs-lookup"><span data-stu-id="381eb-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="381eb-194">类型</span><span class="sxs-lookup"><span data-stu-id="381eb-194">Type</span></span>

*   <span data-ttu-id="381eb-195">String</span><span class="sxs-lookup"><span data-stu-id="381eb-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="381eb-196">要求</span><span class="sxs-lookup"><span data-stu-id="381eb-196">Requirements</span></span>

|<span data-ttu-id="381eb-197">要求</span><span class="sxs-lookup"><span data-stu-id="381eb-197">Requirement</span></span>| <span data-ttu-id="381eb-198">值</span><span class="sxs-lookup"><span data-stu-id="381eb-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="381eb-199">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="381eb-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="381eb-200">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-200">1.1</span></span>|
|[<span data-ttu-id="381eb-201">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="381eb-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="381eb-202">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="381eb-203">示例</span><span class="sxs-lookup"><span data-stu-id="381eb-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="381eb-204">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="381eb-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="381eb-205">获取承载外接程序的 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="381eb-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="381eb-206">类型</span><span class="sxs-lookup"><span data-stu-id="381eb-206">Type</span></span>

*   [<span data-ttu-id="381eb-207">HostType</span><span class="sxs-lookup"><span data-stu-id="381eb-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="381eb-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="381eb-208">Requirements</span></span>

|<span data-ttu-id="381eb-209">要求</span><span class="sxs-lookup"><span data-stu-id="381eb-209">Requirement</span></span>| <span data-ttu-id="381eb-210">值</span><span class="sxs-lookup"><span data-stu-id="381eb-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="381eb-211">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="381eb-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="381eb-212">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-212">1.1</span></span>|
|[<span data-ttu-id="381eb-213">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="381eb-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="381eb-214">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="381eb-215">示例</span><span class="sxs-lookup"><span data-stu-id="381eb-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="381eb-216">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="381eb-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="381eb-217">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="381eb-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="381eb-218">类型</span><span class="sxs-lookup"><span data-stu-id="381eb-218">Type</span></span>

*   [<span data-ttu-id="381eb-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="381eb-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="381eb-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="381eb-220">Requirements</span></span>

|<span data-ttu-id="381eb-221">要求</span><span class="sxs-lookup"><span data-stu-id="381eb-221">Requirement</span></span>| <span data-ttu-id="381eb-222">值</span><span class="sxs-lookup"><span data-stu-id="381eb-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="381eb-223">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="381eb-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="381eb-224">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-224">1.1</span></span>|
|[<span data-ttu-id="381eb-225">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="381eb-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="381eb-226">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="381eb-227">示例</span><span class="sxs-lookup"><span data-stu-id="381eb-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="381eb-228">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="381eb-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="381eb-229">提供用于确定当前应用程序和平台支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="381eb-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="381eb-230">类型</span><span class="sxs-lookup"><span data-stu-id="381eb-230">Type</span></span>

*   [<span data-ttu-id="381eb-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="381eb-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="381eb-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="381eb-232">Requirements</span></span>

|<span data-ttu-id="381eb-233">要求</span><span class="sxs-lookup"><span data-stu-id="381eb-233">Requirement</span></span>| <span data-ttu-id="381eb-234">值</span><span class="sxs-lookup"><span data-stu-id="381eb-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="381eb-235">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="381eb-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="381eb-236">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-236">1.1</span></span>|
|[<span data-ttu-id="381eb-237">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="381eb-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="381eb-238">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="381eb-239">示例</span><span class="sxs-lookup"><span data-stu-id="381eb-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="381eb-240">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="381eb-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="381eb-241">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="381eb-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="381eb-242">该 `RoamingSettings` 对象使您可以存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时都可使用该加载项。</span><span class="sxs-lookup"><span data-stu-id="381eb-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="381eb-243">类型</span><span class="sxs-lookup"><span data-stu-id="381eb-243">Type</span></span>

*   [<span data-ttu-id="381eb-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="381eb-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="381eb-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="381eb-245">Requirements</span></span>

|<span data-ttu-id="381eb-246">要求</span><span class="sxs-lookup"><span data-stu-id="381eb-246">Requirement</span></span>| <span data-ttu-id="381eb-247">值</span><span class="sxs-lookup"><span data-stu-id="381eb-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="381eb-248">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="381eb-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="381eb-249">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-249">1.1</span></span>|
|[<span data-ttu-id="381eb-250">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="381eb-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="381eb-251">受限</span><span class="sxs-lookup"><span data-stu-id="381eb-251">Restricted</span></span>|
|[<span data-ttu-id="381eb-252">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="381eb-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="381eb-253">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="381eb-254">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="381eb-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="381eb-255">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="381eb-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="381eb-256">类型</span><span class="sxs-lookup"><span data-stu-id="381eb-256">Type</span></span>

*   [<span data-ttu-id="381eb-257">UI</span><span class="sxs-lookup"><span data-stu-id="381eb-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="381eb-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="381eb-258">Requirements</span></span>

|<span data-ttu-id="381eb-259">要求</span><span class="sxs-lookup"><span data-stu-id="381eb-259">Requirement</span></span>| <span data-ttu-id="381eb-260">值</span><span class="sxs-lookup"><span data-stu-id="381eb-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="381eb-261">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="381eb-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="381eb-262">1.1</span><span class="sxs-lookup"><span data-stu-id="381eb-262">1.1</span></span>|
|[<span data-ttu-id="381eb-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="381eb-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="381eb-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="381eb-264">Compose or Read</span></span>|
