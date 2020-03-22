---
title: Office。上下文要求集1。8
description: 使用邮箱 API 要求集1.8 的 Outlook 外接程序可用的 Office 对象成员。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 46e54d2eece113681cec7a2f86dfa0c1bc1b3993
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891171"
---
# <a name="context-mailbox-requirement-set-18"></a><span data-ttu-id="b26df-103">context （邮箱要求集1.8）</span><span class="sxs-lookup"><span data-stu-id="b26df-103">context (Mailbox requirement set 1.8)</span></span>

### <a name="officecontext"></a><span data-ttu-id="b26df-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="b26df-104">[Office](office.md).context</span></span>

<span data-ttu-id="b26df-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="b26df-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="b26df-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅[通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.8)"。</span><span class="sxs-lookup"><span data-stu-id="b26df-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.8).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b26df-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="b26df-107">Requirements</span></span>

|<span data-ttu-id="b26df-108">要求</span><span class="sxs-lookup"><span data-stu-id="b26df-108">Requirement</span></span>| <span data-ttu-id="b26df-109">值</span><span class="sxs-lookup"><span data-stu-id="b26df-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="b26df-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b26df-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b26df-111">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-111">1.1</span></span>|
|[<span data-ttu-id="b26df-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b26df-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b26df-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b26df-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="b26df-114">属性</span><span class="sxs-lookup"><span data-stu-id="b26df-114">Properties</span></span>

| <span data-ttu-id="b26df-115">属性</span><span class="sxs-lookup"><span data-stu-id="b26df-115">Property</span></span> | <span data-ttu-id="b26df-116">型号</span><span class="sxs-lookup"><span data-stu-id="b26df-116">Modes</span></span> | <span data-ttu-id="b26df-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="b26df-117">Return type</span></span> | <span data-ttu-id="b26df-118">最低</span><span class="sxs-lookup"><span data-stu-id="b26df-118">Minimum</span></span><br><span data-ttu-id="b26df-119">要求集</span><span class="sxs-lookup"><span data-stu-id="b26df-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b26df-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="b26df-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="b26df-121">撰写</span><span class="sxs-lookup"><span data-stu-id="b26df-121">Compose</span></span><br><span data-ttu-id="b26df-122">读取</span><span class="sxs-lookup"><span data-stu-id="b26df-122">Read</span></span> | <span data-ttu-id="b26df-123">String</span><span class="sxs-lookup"><span data-stu-id="b26df-123">String</span></span> | [<span data-ttu-id="b26df-124">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b26df-125">过程</span><span class="sxs-lookup"><span data-stu-id="b26df-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="b26df-126">撰写</span><span class="sxs-lookup"><span data-stu-id="b26df-126">Compose</span></span><br><span data-ttu-id="b26df-127">读取</span><span class="sxs-lookup"><span data-stu-id="b26df-127">Read</span></span> | [<span data-ttu-id="b26df-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b26df-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.8) | [<span data-ttu-id="b26df-129">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b26df-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="b26df-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="b26df-131">撰写</span><span class="sxs-lookup"><span data-stu-id="b26df-131">Compose</span></span><br><span data-ttu-id="b26df-132">读取</span><span class="sxs-lookup"><span data-stu-id="b26df-132">Read</span></span> | <span data-ttu-id="b26df-133">String</span><span class="sxs-lookup"><span data-stu-id="b26df-133">String</span></span> | [<span data-ttu-id="b26df-134">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b26df-135">host</span><span class="sxs-lookup"><span data-stu-id="b26df-135">host</span></span>](#host-hosttype) | <span data-ttu-id="b26df-136">撰写</span><span class="sxs-lookup"><span data-stu-id="b26df-136">Compose</span></span><br><span data-ttu-id="b26df-137">读取</span><span class="sxs-lookup"><span data-stu-id="b26df-137">Read</span></span> | [<span data-ttu-id="b26df-138">HostType</span><span class="sxs-lookup"><span data-stu-id="b26df-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.8) | [<span data-ttu-id="b26df-139">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b26df-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="b26df-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="b26df-141">撰写</span><span class="sxs-lookup"><span data-stu-id="b26df-141">Compose</span></span><br><span data-ttu-id="b26df-142">读取</span><span class="sxs-lookup"><span data-stu-id="b26df-142">Read</span></span> | [<span data-ttu-id="b26df-143">邮箱</span><span class="sxs-lookup"><span data-stu-id="b26df-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8) | [<span data-ttu-id="b26df-144">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b26df-145">平台</span><span class="sxs-lookup"><span data-stu-id="b26df-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="b26df-146">撰写</span><span class="sxs-lookup"><span data-stu-id="b26df-146">Compose</span></span><br><span data-ttu-id="b26df-147">读取</span><span class="sxs-lookup"><span data-stu-id="b26df-147">Read</span></span> | [<span data-ttu-id="b26df-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b26df-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.8) | [<span data-ttu-id="b26df-149">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b26df-150">满足</span><span class="sxs-lookup"><span data-stu-id="b26df-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="b26df-151">撰写</span><span class="sxs-lookup"><span data-stu-id="b26df-151">Compose</span></span><br><span data-ttu-id="b26df-152">读取</span><span class="sxs-lookup"><span data-stu-id="b26df-152">Read</span></span> | [<span data-ttu-id="b26df-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b26df-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.8) | [<span data-ttu-id="b26df-154">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b26df-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="b26df-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="b26df-156">撰写</span><span class="sxs-lookup"><span data-stu-id="b26df-156">Compose</span></span><br><span data-ttu-id="b26df-157">读取</span><span class="sxs-lookup"><span data-stu-id="b26df-157">Read</span></span> | [<span data-ttu-id="b26df-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b26df-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.8) | [<span data-ttu-id="b26df-159">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b26df-160">ui</span><span class="sxs-lookup"><span data-stu-id="b26df-160">ui</span></span>](#ui-ui) | <span data-ttu-id="b26df-161">撰写</span><span class="sxs-lookup"><span data-stu-id="b26df-161">Compose</span></span><br><span data-ttu-id="b26df-162">读取</span><span class="sxs-lookup"><span data-stu-id="b26df-162">Read</span></span> | [<span data-ttu-id="b26df-163">UI</span><span class="sxs-lookup"><span data-stu-id="b26df-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.8) | [<span data-ttu-id="b26df-164">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="b26df-165">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="b26df-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="b26df-166">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="b26df-166">contentLanguage: String</span></span>

<span data-ttu-id="b26df-167">获取用户指定的用于编辑项的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="b26df-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="b26df-168">此`contentLanguage`值反映了在 Office 主机应用程序中使用**File > Options > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="b26df-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="b26df-169">类型</span><span class="sxs-lookup"><span data-stu-id="b26df-169">Type</span></span>

*   <span data-ttu-id="b26df-170">String</span><span class="sxs-lookup"><span data-stu-id="b26df-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b26df-171">要求</span><span class="sxs-lookup"><span data-stu-id="b26df-171">Requirements</span></span>

|<span data-ttu-id="b26df-172">要求</span><span class="sxs-lookup"><span data-stu-id="b26df-172">Requirement</span></span>| <span data-ttu-id="b26df-173">值</span><span class="sxs-lookup"><span data-stu-id="b26df-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="b26df-174">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b26df-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b26df-175">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-175">1.1</span></span>|
|[<span data-ttu-id="b26df-176">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b26df-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b26df-177">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b26df-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b26df-178">示例</span><span class="sxs-lookup"><span data-stu-id="b26df-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="b26df-179">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="b26df-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="b26df-180">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="b26df-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b26df-181">类型</span><span class="sxs-lookup"><span data-stu-id="b26df-181">Type</span></span>

*   [<span data-ttu-id="b26df-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b26df-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="b26df-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="b26df-183">Requirements</span></span>

|<span data-ttu-id="b26df-184">要求</span><span class="sxs-lookup"><span data-stu-id="b26df-184">Requirement</span></span>| <span data-ttu-id="b26df-185">值</span><span class="sxs-lookup"><span data-stu-id="b26df-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="b26df-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b26df-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b26df-187">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-187">1.1</span></span>|
|[<span data-ttu-id="b26df-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b26df-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b26df-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b26df-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b26df-190">示例</span><span class="sxs-lookup"><span data-stu-id="b26df-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="b26df-191">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="b26df-191">displayLanguage: String</span></span>

<span data-ttu-id="b26df-192">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="b26df-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="b26df-193">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="b26df-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="b26df-194">类型</span><span class="sxs-lookup"><span data-stu-id="b26df-194">Type</span></span>

*   <span data-ttu-id="b26df-195">String</span><span class="sxs-lookup"><span data-stu-id="b26df-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b26df-196">要求</span><span class="sxs-lookup"><span data-stu-id="b26df-196">Requirements</span></span>

|<span data-ttu-id="b26df-197">要求</span><span class="sxs-lookup"><span data-stu-id="b26df-197">Requirement</span></span>| <span data-ttu-id="b26df-198">值</span><span class="sxs-lookup"><span data-stu-id="b26df-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="b26df-199">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b26df-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b26df-200">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-200">1.1</span></span>|
|[<span data-ttu-id="b26df-201">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b26df-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b26df-202">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b26df-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b26df-203">示例</span><span class="sxs-lookup"><span data-stu-id="b26df-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="b26df-204">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="b26df-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="b26df-205">获取运行外接程序的 Office 应用程序主机。</span><span class="sxs-lookup"><span data-stu-id="b26df-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b26df-206">类型</span><span class="sxs-lookup"><span data-stu-id="b26df-206">Type</span></span>

*   [<span data-ttu-id="b26df-207">HostType</span><span class="sxs-lookup"><span data-stu-id="b26df-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="b26df-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="b26df-208">Requirements</span></span>

|<span data-ttu-id="b26df-209">要求</span><span class="sxs-lookup"><span data-stu-id="b26df-209">Requirement</span></span>| <span data-ttu-id="b26df-210">值</span><span class="sxs-lookup"><span data-stu-id="b26df-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="b26df-211">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b26df-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b26df-212">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-212">1.1</span></span>|
|[<span data-ttu-id="b26df-213">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b26df-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b26df-214">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b26df-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b26df-215">示例</span><span class="sxs-lookup"><span data-stu-id="b26df-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="b26df-216">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="b26df-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="b26df-217">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="b26df-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b26df-218">类型</span><span class="sxs-lookup"><span data-stu-id="b26df-218">Type</span></span>

*   [<span data-ttu-id="b26df-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b26df-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="b26df-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="b26df-220">Requirements</span></span>

|<span data-ttu-id="b26df-221">要求</span><span class="sxs-lookup"><span data-stu-id="b26df-221">Requirement</span></span>| <span data-ttu-id="b26df-222">值</span><span class="sxs-lookup"><span data-stu-id="b26df-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="b26df-223">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b26df-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b26df-224">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-224">1.1</span></span>|
|[<span data-ttu-id="b26df-225">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b26df-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b26df-226">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b26df-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b26df-227">示例</span><span class="sxs-lookup"><span data-stu-id="b26df-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="b26df-228">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="b26df-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="b26df-229">提供用于确定当前主机和平台上支持的要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="b26df-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b26df-230">类型</span><span class="sxs-lookup"><span data-stu-id="b26df-230">Type</span></span>

*   [<span data-ttu-id="b26df-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b26df-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="b26df-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="b26df-232">Requirements</span></span>

|<span data-ttu-id="b26df-233">要求</span><span class="sxs-lookup"><span data-stu-id="b26df-233">Requirement</span></span>| <span data-ttu-id="b26df-234">值</span><span class="sxs-lookup"><span data-stu-id="b26df-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="b26df-235">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b26df-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b26df-236">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-236">1.1</span></span>|
|[<span data-ttu-id="b26df-237">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b26df-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b26df-238">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b26df-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b26df-239">示例</span><span class="sxs-lookup"><span data-stu-id="b26df-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="b26df-240">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="b26df-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="b26df-241">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="b26df-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="b26df-242">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="b26df-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="b26df-243">类型</span><span class="sxs-lookup"><span data-stu-id="b26df-243">Type</span></span>

*   [<span data-ttu-id="b26df-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b26df-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="b26df-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="b26df-245">Requirements</span></span>

|<span data-ttu-id="b26df-246">要求</span><span class="sxs-lookup"><span data-stu-id="b26df-246">Requirement</span></span>| <span data-ttu-id="b26df-247">值</span><span class="sxs-lookup"><span data-stu-id="b26df-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="b26df-248">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b26df-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b26df-249">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-249">1.1</span></span>|
|[<span data-ttu-id="b26df-250">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b26df-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="b26df-251">受限</span><span class="sxs-lookup"><span data-stu-id="b26df-251">Restricted</span></span>|
|[<span data-ttu-id="b26df-252">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b26df-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b26df-253">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b26df-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="b26df-254">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="b26df-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="b26df-255">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="b26df-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="b26df-256">类型</span><span class="sxs-lookup"><span data-stu-id="b26df-256">Type</span></span>

*   [<span data-ttu-id="b26df-257">UI</span><span class="sxs-lookup"><span data-stu-id="b26df-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="b26df-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="b26df-258">Requirements</span></span>

|<span data-ttu-id="b26df-259">要求</span><span class="sxs-lookup"><span data-stu-id="b26df-259">Requirement</span></span>| <span data-ttu-id="b26df-260">值</span><span class="sxs-lookup"><span data-stu-id="b26df-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="b26df-261">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b26df-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b26df-262">1.1</span><span class="sxs-lookup"><span data-stu-id="b26df-262">1.1</span></span>|
|[<span data-ttu-id="b26df-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b26df-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b26df-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b26df-264">Compose or Read</span></span>|
