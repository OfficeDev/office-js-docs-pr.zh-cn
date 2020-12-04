---
title: Office。上下文要求集1。6
description: 使用邮箱 API 要求集1.6 的 Outlook 外接程序可用的 Office 对象成员。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 55e3761aea94d902903c53a9b3be687d94b42e12
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570756"
---
# <a name="context-mailbox-requirement-set-16"></a><span data-ttu-id="f6710-103"> (邮箱要求集1.6 的上下文) </span><span class="sxs-lookup"><span data-stu-id="f6710-103">context (Mailbox requirement set 1.6)</span></span>

### <a name="officecontext"></a><span data-ttu-id="f6710-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="f6710-104">[Office](office.md).context</span></span>

<span data-ttu-id="f6710-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="f6710-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="f6710-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅 [通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true)"。</span><span class="sxs-lookup"><span data-stu-id="f6710-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6710-107">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-107">Requirements</span></span>

|<span data-ttu-id="f6710-108">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-108">Requirement</span></span>| <span data-ttu-id="f6710-109">值</span><span class="sxs-lookup"><span data-stu-id="f6710-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6710-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f6710-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f6710-111">1.1</span><span class="sxs-lookup"><span data-stu-id="f6710-111">1.1</span></span>|
|[<span data-ttu-id="f6710-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f6710-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f6710-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="f6710-114">属性</span><span class="sxs-lookup"><span data-stu-id="f6710-114">Properties</span></span>

| <span data-ttu-id="f6710-115">属性</span><span class="sxs-lookup"><span data-stu-id="f6710-115">Property</span></span> | <span data-ttu-id="f6710-116">型号</span><span class="sxs-lookup"><span data-stu-id="f6710-116">Modes</span></span> | <span data-ttu-id="f6710-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="f6710-117">Return type</span></span> | <span data-ttu-id="f6710-118">最小值</span><span class="sxs-lookup"><span data-stu-id="f6710-118">Minimum</span></span><br><span data-ttu-id="f6710-119">要求集</span><span class="sxs-lookup"><span data-stu-id="f6710-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f6710-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="f6710-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="f6710-121">撰写</span><span class="sxs-lookup"><span data-stu-id="f6710-121">Compose</span></span><br><span data-ttu-id="f6710-122">阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-122">Read</span></span> | <span data-ttu-id="f6710-123">String</span><span class="sxs-lookup"><span data-stu-id="f6710-123">String</span></span> | [<span data-ttu-id="f6710-124">1.1</span><span class="sxs-lookup"><span data-stu-id="f6710-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f6710-125">过程</span><span class="sxs-lookup"><span data-stu-id="f6710-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="f6710-126">撰写</span><span class="sxs-lookup"><span data-stu-id="f6710-126">Compose</span></span><br><span data-ttu-id="f6710-127">阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-127">Read</span></span> | [<span data-ttu-id="f6710-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="f6710-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="f6710-129">1.1</span><span class="sxs-lookup"><span data-stu-id="f6710-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f6710-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="f6710-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="f6710-131">撰写</span><span class="sxs-lookup"><span data-stu-id="f6710-131">Compose</span></span><br><span data-ttu-id="f6710-132">阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-132">Read</span></span> | <span data-ttu-id="f6710-133">String</span><span class="sxs-lookup"><span data-stu-id="f6710-133">String</span></span> | [<span data-ttu-id="f6710-134">1.1</span><span class="sxs-lookup"><span data-stu-id="f6710-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f6710-135">host</span><span class="sxs-lookup"><span data-stu-id="f6710-135">host</span></span>](#host-hosttype) | <span data-ttu-id="f6710-136">撰写</span><span class="sxs-lookup"><span data-stu-id="f6710-136">Compose</span></span><br><span data-ttu-id="f6710-137">阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-137">Read</span></span> | [<span data-ttu-id="f6710-138">HostType</span><span class="sxs-lookup"><span data-stu-id="f6710-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="f6710-139">1.5</span><span class="sxs-lookup"><span data-stu-id="f6710-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="f6710-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="f6710-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="f6710-141">撰写</span><span class="sxs-lookup"><span data-stu-id="f6710-141">Compose</span></span><br><span data-ttu-id="f6710-142">阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-142">Read</span></span> | [<span data-ttu-id="f6710-143">邮箱</span><span class="sxs-lookup"><span data-stu-id="f6710-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="f6710-144">1.1</span><span class="sxs-lookup"><span data-stu-id="f6710-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f6710-145">平台</span><span class="sxs-lookup"><span data-stu-id="f6710-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="f6710-146">撰写</span><span class="sxs-lookup"><span data-stu-id="f6710-146">Compose</span></span><br><span data-ttu-id="f6710-147">阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-147">Read</span></span> | [<span data-ttu-id="f6710-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="f6710-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="f6710-149">1.5</span><span class="sxs-lookup"><span data-stu-id="f6710-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="f6710-150">满足</span><span class="sxs-lookup"><span data-stu-id="f6710-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="f6710-151">撰写</span><span class="sxs-lookup"><span data-stu-id="f6710-151">Compose</span></span><br><span data-ttu-id="f6710-152">阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-152">Read</span></span> | [<span data-ttu-id="f6710-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="f6710-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="f6710-154">1.1</span><span class="sxs-lookup"><span data-stu-id="f6710-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f6710-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="f6710-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="f6710-156">撰写</span><span class="sxs-lookup"><span data-stu-id="f6710-156">Compose</span></span><br><span data-ttu-id="f6710-157">阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-157">Read</span></span> | [<span data-ttu-id="f6710-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="f6710-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="f6710-159">1.1</span><span class="sxs-lookup"><span data-stu-id="f6710-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f6710-160">ui</span><span class="sxs-lookup"><span data-stu-id="f6710-160">ui</span></span>](#ui-ui) | <span data-ttu-id="f6710-161">撰写</span><span class="sxs-lookup"><span data-stu-id="f6710-161">Compose</span></span><br><span data-ttu-id="f6710-162">阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-162">Read</span></span> | [<span data-ttu-id="f6710-163">UI</span><span class="sxs-lookup"><span data-stu-id="f6710-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="f6710-164">1.1</span><span class="sxs-lookup"><span data-stu-id="f6710-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="f6710-165">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="f6710-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="f6710-166">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="f6710-166">contentLanguage: String</span></span>

<span data-ttu-id="f6710-167">获取用户指定的用于编辑项目的区域设置 (语言) 。</span><span class="sxs-lookup"><span data-stu-id="f6710-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="f6710-168">此 `contentLanguage` 值反映了使用 Office 客户端应用程序中的 "**文件 > 选项" > 语言** 指定的当前 **编辑语言** 设置。</span><span class="sxs-lookup"><span data-stu-id="f6710-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="f6710-169">类型</span><span class="sxs-lookup"><span data-stu-id="f6710-169">Type</span></span>

*   <span data-ttu-id="f6710-170">String</span><span class="sxs-lookup"><span data-stu-id="f6710-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6710-171">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-171">Requirements</span></span>

|<span data-ttu-id="f6710-172">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-172">Requirement</span></span>| <span data-ttu-id="f6710-173">值</span><span class="sxs-lookup"><span data-stu-id="f6710-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6710-174">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f6710-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f6710-175">1.1</span><span class="sxs-lookup"><span data-stu-id="f6710-175">1.1</span></span>|
|[<span data-ttu-id="f6710-176">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f6710-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f6710-177">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6710-178">示例</span><span class="sxs-lookup"><span data-stu-id="f6710-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="f6710-179">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="f6710-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="f6710-180">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="f6710-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="f6710-181">类型</span><span class="sxs-lookup"><span data-stu-id="f6710-181">Type</span></span>

*   [<span data-ttu-id="f6710-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="f6710-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="f6710-183">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-183">Requirements</span></span>

|<span data-ttu-id="f6710-184">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-184">Requirement</span></span>| <span data-ttu-id="f6710-185">值</span><span class="sxs-lookup"><span data-stu-id="f6710-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6710-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f6710-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f6710-187">1.1</span><span class="sxs-lookup"><span data-stu-id="f6710-187">1.1</span></span>|
|[<span data-ttu-id="f6710-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f6710-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f6710-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6710-190">示例</span><span class="sxs-lookup"><span data-stu-id="f6710-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="f6710-191">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="f6710-191">displayLanguage: String</span></span>

<span data-ttu-id="f6710-192">获取用户为 Office 客户端应用程序的 UI 指定的 RFC 1766 语言标记格式中 (语言) 的区域设置。</span><span class="sxs-lookup"><span data-stu-id="f6710-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="f6710-193">此 `displayLanguage` 值反映了使用 Office 客户端应用程序中的 **File > Options > 语言** 指定的当前 **显示语言** 设置。</span><span class="sxs-lookup"><span data-stu-id="f6710-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="f6710-194">类型</span><span class="sxs-lookup"><span data-stu-id="f6710-194">Type</span></span>

*   <span data-ttu-id="f6710-195">String</span><span class="sxs-lookup"><span data-stu-id="f6710-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6710-196">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-196">Requirements</span></span>

|<span data-ttu-id="f6710-197">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-197">Requirement</span></span>| <span data-ttu-id="f6710-198">值</span><span class="sxs-lookup"><span data-stu-id="f6710-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6710-199">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f6710-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f6710-200">1.1</span><span class="sxs-lookup"><span data-stu-id="f6710-200">1.1</span></span>|
|[<span data-ttu-id="f6710-201">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f6710-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f6710-202">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6710-203">示例</span><span class="sxs-lookup"><span data-stu-id="f6710-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="f6710-204">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="f6710-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="f6710-205">获取承载外接程序的 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="f6710-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="f6710-206">或者，也可以使用 " [context.subname](#diagnostics-contextinformation) " 属性获取主机。</span><span class="sxs-lookup"><span data-stu-id="f6710-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="f6710-207">类型</span><span class="sxs-lookup"><span data-stu-id="f6710-207">Type</span></span>

*   [<span data-ttu-id="f6710-208">HostType</span><span class="sxs-lookup"><span data-stu-id="f6710-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="f6710-209">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-209">Requirements</span></span>

|<span data-ttu-id="f6710-210">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-210">Requirement</span></span>| <span data-ttu-id="f6710-211">值</span><span class="sxs-lookup"><span data-stu-id="f6710-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6710-212">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f6710-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f6710-213">1.5</span><span class="sxs-lookup"><span data-stu-id="f6710-213">1.5</span></span>|
|[<span data-ttu-id="f6710-214">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f6710-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f6710-215">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6710-216">示例</span><span class="sxs-lookup"><span data-stu-id="f6710-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="f6710-217">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="f6710-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="f6710-218">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="f6710-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="f6710-219">或者，也可以使用 " [context.subname](#diagnostics-contextinformation) " 属性获取平台。</span><span class="sxs-lookup"><span data-stu-id="f6710-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="f6710-220">类型</span><span class="sxs-lookup"><span data-stu-id="f6710-220">Type</span></span>

*   [<span data-ttu-id="f6710-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="f6710-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="f6710-222">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-222">Requirements</span></span>

|<span data-ttu-id="f6710-223">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-223">Requirement</span></span>| <span data-ttu-id="f6710-224">值</span><span class="sxs-lookup"><span data-stu-id="f6710-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6710-225">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f6710-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f6710-226">1.5</span><span class="sxs-lookup"><span data-stu-id="f6710-226">1.5</span></span>|
|[<span data-ttu-id="f6710-227">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f6710-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f6710-228">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6710-229">示例</span><span class="sxs-lookup"><span data-stu-id="f6710-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="f6710-230">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="f6710-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="f6710-231">提供用于确定当前应用程序和平台支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="f6710-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="f6710-232">类型</span><span class="sxs-lookup"><span data-stu-id="f6710-232">Type</span></span>

*   [<span data-ttu-id="f6710-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="f6710-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="f6710-234">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-234">Requirements</span></span>

|<span data-ttu-id="f6710-235">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-235">Requirement</span></span>| <span data-ttu-id="f6710-236">值</span><span class="sxs-lookup"><span data-stu-id="f6710-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6710-237">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f6710-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f6710-238">1.1</span><span class="sxs-lookup"><span data-stu-id="f6710-238">1.1</span></span>|
|[<span data-ttu-id="f6710-239">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f6710-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f6710-240">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6710-241">示例</span><span class="sxs-lookup"><span data-stu-id="f6710-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="f6710-242">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="f6710-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="f6710-243">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="f6710-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="f6710-244">该 `RoamingSettings` 对象使您可以存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时都可使用该加载项。</span><span class="sxs-lookup"><span data-stu-id="f6710-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="f6710-245">类型</span><span class="sxs-lookup"><span data-stu-id="f6710-245">Type</span></span>

*   [<span data-ttu-id="f6710-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="f6710-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="f6710-247">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-247">Requirements</span></span>

|<span data-ttu-id="f6710-248">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-248">Requirement</span></span>| <span data-ttu-id="f6710-249">值</span><span class="sxs-lookup"><span data-stu-id="f6710-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6710-250">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f6710-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f6710-251">1.1</span><span class="sxs-lookup"><span data-stu-id="f6710-251">1.1</span></span>|
|[<span data-ttu-id="f6710-252">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f6710-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="f6710-253">受限</span><span class="sxs-lookup"><span data-stu-id="f6710-253">Restricted</span></span>|
|[<span data-ttu-id="f6710-254">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f6710-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f6710-255">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="f6710-256">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="f6710-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="f6710-257">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="f6710-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="f6710-258">类型</span><span class="sxs-lookup"><span data-stu-id="f6710-258">Type</span></span>

*   [<span data-ttu-id="f6710-259">UI</span><span class="sxs-lookup"><span data-stu-id="f6710-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="f6710-260">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-260">Requirements</span></span>

|<span data-ttu-id="f6710-261">要求</span><span class="sxs-lookup"><span data-stu-id="f6710-261">Requirement</span></span>| <span data-ttu-id="f6710-262">值</span><span class="sxs-lookup"><span data-stu-id="f6710-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6710-263">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f6710-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f6710-264">1.1</span><span class="sxs-lookup"><span data-stu-id="f6710-264">1.1</span></span>|
|[<span data-ttu-id="f6710-265">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f6710-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f6710-266">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f6710-266">Compose or Read</span></span>|
