---
title: Office. context-预览要求集
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 5c34a7a0db5880a94ba5519059a93010a5243978
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629186"
---
# <a name="context"></a><span data-ttu-id="4a009-102">context</span><span class="sxs-lookup"><span data-stu-id="4a009-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="4a009-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="4a009-103">[Office](Office.md).context</span></span>

<span data-ttu-id="4a009-p101">Office.context 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office.context 命名空间的完整列表，请参阅[通用 API 中的 Office.context 引用](/javascript/api/office/office.context)。</span><span class="sxs-lookup"><span data-stu-id="4a009-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a009-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a009-106">Requirements</span></span>

|<span data-ttu-id="4a009-107">要求</span><span class="sxs-lookup"><span data-stu-id="4a009-107">Requirement</span></span>| <span data-ttu-id="4a009-108">值</span><span class="sxs-lookup"><span data-stu-id="4a009-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a009-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a009-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a009-110">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-110">1.0</span></span>|
|[<span data-ttu-id="4a009-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a009-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a009-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a009-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="4a009-113">属性</span><span class="sxs-lookup"><span data-stu-id="4a009-113">Properties</span></span>

| <span data-ttu-id="4a009-114">属性</span><span class="sxs-lookup"><span data-stu-id="4a009-114">Property</span></span> | <span data-ttu-id="4a009-115">型号</span><span class="sxs-lookup"><span data-stu-id="4a009-115">Modes</span></span> | <span data-ttu-id="4a009-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="4a009-116">Return type</span></span> | <span data-ttu-id="4a009-117">最低</span><span class="sxs-lookup"><span data-stu-id="4a009-117">Minimum</span></span><br><span data-ttu-id="4a009-118">要求集</span><span class="sxs-lookup"><span data-stu-id="4a009-118">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="4a009-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="4a009-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="4a009-120">撰写</span><span class="sxs-lookup"><span data-stu-id="4a009-120">Compose</span></span><br><span data-ttu-id="4a009-121">读取</span><span class="sxs-lookup"><span data-stu-id="4a009-121">Read</span></span> | <span data-ttu-id="4a009-122">String</span><span class="sxs-lookup"><span data-stu-id="4a009-122">String</span></span> | <span data-ttu-id="4a009-123">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-123">1.0</span></span> |
| [<span data-ttu-id="4a009-124">过程</span><span class="sxs-lookup"><span data-stu-id="4a009-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="4a009-125">撰写</span><span class="sxs-lookup"><span data-stu-id="4a009-125">Compose</span></span><br><span data-ttu-id="4a009-126">读取</span><span class="sxs-lookup"><span data-stu-id="4a009-126">Read</span></span> | [<span data-ttu-id="4a009-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="4a009-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation) | <span data-ttu-id="4a009-128">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-128">1.0</span></span> |
| [<span data-ttu-id="4a009-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="4a009-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="4a009-130">撰写</span><span class="sxs-lookup"><span data-stu-id="4a009-130">Compose</span></span><br><span data-ttu-id="4a009-131">读取</span><span class="sxs-lookup"><span data-stu-id="4a009-131">Read</span></span> | <span data-ttu-id="4a009-132">String</span><span class="sxs-lookup"><span data-stu-id="4a009-132">String</span></span> | <span data-ttu-id="4a009-133">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-133">1.0</span></span> |
| [<span data-ttu-id="4a009-134">host</span><span class="sxs-lookup"><span data-stu-id="4a009-134">host</span></span>](#host-hosttype) | <span data-ttu-id="4a009-135">撰写</span><span class="sxs-lookup"><span data-stu-id="4a009-135">Compose</span></span><br><span data-ttu-id="4a009-136">读取</span><span class="sxs-lookup"><span data-stu-id="4a009-136">Read</span></span> | [<span data-ttu-id="4a009-137">HostType</span><span class="sxs-lookup"><span data-stu-id="4a009-137">HostType</span></span>](/javascript/api/office/office.hosttype) | <span data-ttu-id="4a009-138">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-138">1.0</span></span> |
| [<span data-ttu-id="4a009-139">officeTheme</span><span class="sxs-lookup"><span data-stu-id="4a009-139">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="4a009-140">撰写</span><span class="sxs-lookup"><span data-stu-id="4a009-140">Compose</span></span><br><span data-ttu-id="4a009-141">读取</span><span class="sxs-lookup"><span data-stu-id="4a009-141">Read</span></span> | [<span data-ttu-id="4a009-142">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="4a009-142">OfficeTheme</span></span>](/javascript/api/office/office.officetheme) | <span data-ttu-id="4a009-143">预览</span><span class="sxs-lookup"><span data-stu-id="4a009-143">Preview</span></span> |
| [<span data-ttu-id="4a009-144">平台</span><span class="sxs-lookup"><span data-stu-id="4a009-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="4a009-145">撰写</span><span class="sxs-lookup"><span data-stu-id="4a009-145">Compose</span></span><br><span data-ttu-id="4a009-146">读取</span><span class="sxs-lookup"><span data-stu-id="4a009-146">Read</span></span> | [<span data-ttu-id="4a009-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="4a009-147">PlatformType</span></span>](/javascript/api/office/office.platformtype) | <span data-ttu-id="4a009-148">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-148">1.0</span></span> |
| [<span data-ttu-id="4a009-149">满足</span><span class="sxs-lookup"><span data-stu-id="4a009-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="4a009-150">撰写</span><span class="sxs-lookup"><span data-stu-id="4a009-150">Compose</span></span><br><span data-ttu-id="4a009-151">读取</span><span class="sxs-lookup"><span data-stu-id="4a009-151">Read</span></span> | [<span data-ttu-id="4a009-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="4a009-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport) | <span data-ttu-id="4a009-153">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-153">1.0</span></span> |
| [<span data-ttu-id="4a009-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="4a009-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="4a009-155">撰写</span><span class="sxs-lookup"><span data-stu-id="4a009-155">Compose</span></span><br><span data-ttu-id="4a009-156">读取</span><span class="sxs-lookup"><span data-stu-id="4a009-156">Read</span></span> | [<span data-ttu-id="4a009-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="4a009-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings) | <span data-ttu-id="4a009-158">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-158">1.0</span></span> |
| [<span data-ttu-id="4a009-159">ui</span><span class="sxs-lookup"><span data-stu-id="4a009-159">ui</span></span>](#ui-ui) | <span data-ttu-id="4a009-160">撰写</span><span class="sxs-lookup"><span data-stu-id="4a009-160">Compose</span></span><br><span data-ttu-id="4a009-161">读取</span><span class="sxs-lookup"><span data-stu-id="4a009-161">Read</span></span> | [<span data-ttu-id="4a009-162">UI</span><span class="sxs-lookup"><span data-stu-id="4a009-162">UI</span></span>](/javascript/api/office/office.ui) | <span data-ttu-id="4a009-163">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-163">1.0</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4a009-164">命名空间</span><span class="sxs-lookup"><span data-stu-id="4a009-164">Namespaces</span></span>

<span data-ttu-id="4a009-165">[auth](/javascript/api/office/office.auth)：提供对[单一登录（SSO）](/outlook/add-ins/authenticate-a-user-with-an-sso-token)的支持。</span><span class="sxs-lookup"><span data-stu-id="4a009-165">[auth](/javascript/api/office/office.auth): Provides support for [single sign-on (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token).</span></span>

<span data-ttu-id="4a009-166">[邮箱](office.context.mailbox.md)：提供对 Microsoft Outlook 的 outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="4a009-166">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

## <a name="property-details"></a><span data-ttu-id="4a009-167">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="4a009-167">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="4a009-168">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="4a009-168">contentLanguage: String</span></span>

<span data-ttu-id="4a009-169">获取用户指定的用于编辑项的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="4a009-169">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="4a009-170">此`contentLanguage`值反映了在 Office 主机应用程序中使用**File > Options > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="4a009-170">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="4a009-171">类型</span><span class="sxs-lookup"><span data-stu-id="4a009-171">Type</span></span>

*   <span data-ttu-id="4a009-172">String</span><span class="sxs-lookup"><span data-stu-id="4a009-172">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a009-173">要求</span><span class="sxs-lookup"><span data-stu-id="4a009-173">Requirements</span></span>

|<span data-ttu-id="4a009-174">要求</span><span class="sxs-lookup"><span data-stu-id="4a009-174">Requirement</span></span>| <span data-ttu-id="4a009-175">值</span><span class="sxs-lookup"><span data-stu-id="4a009-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a009-176">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a009-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a009-177">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-177">1.0</span></span>|
|[<span data-ttu-id="4a009-178">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a009-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a009-179">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a009-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a009-180">示例</span><span class="sxs-lookup"><span data-stu-id="4a009-180">Example</span></span>

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="4a009-181">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="4a009-181">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="4a009-182">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="4a009-182">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="4a009-183">类型</span><span class="sxs-lookup"><span data-stu-id="4a009-183">Type</span></span>

*   [<span data-ttu-id="4a009-184">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="4a009-184">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="4a009-185">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a009-185">Requirements</span></span>

|<span data-ttu-id="4a009-186">要求</span><span class="sxs-lookup"><span data-stu-id="4a009-186">Requirement</span></span>| <span data-ttu-id="4a009-187">值</span><span class="sxs-lookup"><span data-stu-id="4a009-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a009-188">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a009-188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a009-189">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-189">1.0</span></span>|
|[<span data-ttu-id="4a009-190">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a009-190">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a009-191">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a009-191">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a009-192">示例</span><span class="sxs-lookup"><span data-stu-id="4a009-192">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="4a009-193">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="4a009-193">displayLanguage: String</span></span>

<span data-ttu-id="4a009-194">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="4a009-194">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="4a009-195">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="4a009-195">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="4a009-196">类型</span><span class="sxs-lookup"><span data-stu-id="4a009-196">Type</span></span>

*   <span data-ttu-id="4a009-197">String</span><span class="sxs-lookup"><span data-stu-id="4a009-197">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a009-198">要求</span><span class="sxs-lookup"><span data-stu-id="4a009-198">Requirements</span></span>

|<span data-ttu-id="4a009-199">要求</span><span class="sxs-lookup"><span data-stu-id="4a009-199">Requirement</span></span>| <span data-ttu-id="4a009-200">值</span><span class="sxs-lookup"><span data-stu-id="4a009-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a009-201">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a009-201">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a009-202">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-202">1.0</span></span>|
|[<span data-ttu-id="4a009-203">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a009-203">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a009-204">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a009-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a009-205">示例</span><span class="sxs-lookup"><span data-stu-id="4a009-205">Example</span></span>

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="4a009-206">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="4a009-206">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="4a009-207">获取运行外接程序的 Office 应用程序主机。</span><span class="sxs-lookup"><span data-stu-id="4a009-207">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="4a009-208">类型</span><span class="sxs-lookup"><span data-stu-id="4a009-208">Type</span></span>

*   [<span data-ttu-id="4a009-209">HostType</span><span class="sxs-lookup"><span data-stu-id="4a009-209">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="4a009-210">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a009-210">Requirements</span></span>

|<span data-ttu-id="4a009-211">要求</span><span class="sxs-lookup"><span data-stu-id="4a009-211">Requirement</span></span>| <span data-ttu-id="4a009-212">值</span><span class="sxs-lookup"><span data-stu-id="4a009-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a009-213">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a009-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a009-214">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-214">1.0</span></span>|
|[<span data-ttu-id="4a009-215">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a009-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a009-216">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a009-216">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a009-217">示例</span><span class="sxs-lookup"><span data-stu-id="4a009-217">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officethemejavascriptapiofficeofficeofficetheme"></a><span data-ttu-id="4a009-218">officeTheme： [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="4a009-218">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="4a009-219">提供了访问 Office 主题颜色的属性。</span><span class="sxs-lookup"><span data-stu-id="4a009-219">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="4a009-220">此成员仅在 Windows 中的 Outlook 中受支持。</span><span class="sxs-lookup"><span data-stu-id="4a009-220">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="4a009-221">使用 Office 主题颜色，您可以将加载项的配色方案与用户选择的当前 Office 主题进行协调，以供用户使用**office > Office 帐户 > Office 主题 UI**，该用户在所有 Office 主机应用程序中应用。</span><span class="sxs-lookup"><span data-stu-id="4a009-221">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="4a009-222">使用 Office 主题颜色适用于邮件和任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="4a009-222">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="4a009-223">类型</span><span class="sxs-lookup"><span data-stu-id="4a009-223">Type</span></span>

*   [<span data-ttu-id="4a009-224">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="4a009-224">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="4a009-225">属性：</span><span class="sxs-lookup"><span data-stu-id="4a009-225">Properties:</span></span>

|<span data-ttu-id="4a009-226">名称</span><span class="sxs-lookup"><span data-stu-id="4a009-226">Name</span></span>| <span data-ttu-id="4a009-227">类型</span><span class="sxs-lookup"><span data-stu-id="4a009-227">Type</span></span>| <span data-ttu-id="4a009-228">说明</span><span class="sxs-lookup"><span data-stu-id="4a009-228">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="4a009-229">String</span><span class="sxs-lookup"><span data-stu-id="4a009-229">String</span></span>|<span data-ttu-id="4a009-230">获取十六进制三原色形式的 Office 主题正文背景色。</span><span class="sxs-lookup"><span data-stu-id="4a009-230">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="4a009-231">String</span><span class="sxs-lookup"><span data-stu-id="4a009-231">String</span></span>|<span data-ttu-id="4a009-232">获取十六进制三原色形式的 Office 主题正文前景色。</span><span class="sxs-lookup"><span data-stu-id="4a009-232">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="4a009-233">String</span><span class="sxs-lookup"><span data-stu-id="4a009-233">String</span></span>|<span data-ttu-id="4a009-234">获取十六进制三原色形式的 Office 主题控制背景色。</span><span class="sxs-lookup"><span data-stu-id="4a009-234">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="4a009-235">字符串</span><span class="sxs-lookup"><span data-stu-id="4a009-235">String</span></span>|<span data-ttu-id="4a009-236">获取十六进制三原色形式的 Office 主题正文控制颜色。</span><span class="sxs-lookup"><span data-stu-id="4a009-236">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4a009-237">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a009-237">Requirements</span></span>

|<span data-ttu-id="4a009-238">要求</span><span class="sxs-lookup"><span data-stu-id="4a009-238">Requirement</span></span>| <span data-ttu-id="4a009-239">值</span><span class="sxs-lookup"><span data-stu-id="4a009-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a009-240">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a009-240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a009-241">预览</span><span class="sxs-lookup"><span data-stu-id="4a009-241">Preview</span></span>|
|[<span data-ttu-id="4a009-242">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a009-242">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a009-243">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a009-243">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a009-244">示例</span><span class="sxs-lookup"><span data-stu-id="4a009-244">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

<br>

---
---

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="4a009-245">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="4a009-245">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="4a009-246">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="4a009-246">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="4a009-247">类型</span><span class="sxs-lookup"><span data-stu-id="4a009-247">Type</span></span>

*   [<span data-ttu-id="4a009-248">PlatformType</span><span class="sxs-lookup"><span data-stu-id="4a009-248">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="4a009-249">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a009-249">Requirements</span></span>

|<span data-ttu-id="4a009-250">要求</span><span class="sxs-lookup"><span data-stu-id="4a009-250">Requirement</span></span>| <span data-ttu-id="4a009-251">值</span><span class="sxs-lookup"><span data-stu-id="4a009-251">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a009-252">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a009-252">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a009-253">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-253">1.0</span></span>|
|[<span data-ttu-id="4a009-254">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a009-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a009-255">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a009-255">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a009-256">示例</span><span class="sxs-lookup"><span data-stu-id="4a009-256">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="4a009-257">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="4a009-257">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="4a009-258">提供用于确定当前主机和平台上支持的要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="4a009-258">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="4a009-259">类型</span><span class="sxs-lookup"><span data-stu-id="4a009-259">Type</span></span>

*   [<span data-ttu-id="4a009-260">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="4a009-260">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="4a009-261">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a009-261">Requirements</span></span>

|<span data-ttu-id="4a009-262">要求</span><span class="sxs-lookup"><span data-stu-id="4a009-262">Requirement</span></span>| <span data-ttu-id="4a009-263">值</span><span class="sxs-lookup"><span data-stu-id="4a009-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a009-264">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a009-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a009-265">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-265">1.0</span></span>|
|[<span data-ttu-id="4a009-266">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a009-266">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a009-267">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a009-267">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a009-268">示例</span><span class="sxs-lookup"><span data-stu-id="4a009-268">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.8")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="4a009-269">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="4a009-269">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="4a009-270">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="4a009-270">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="4a009-271">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="4a009-271">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="4a009-272">类型</span><span class="sxs-lookup"><span data-stu-id="4a009-272">Type</span></span>

*   [<span data-ttu-id="4a009-273">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="4a009-273">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="4a009-274">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a009-274">Requirements</span></span>

|<span data-ttu-id="4a009-275">要求</span><span class="sxs-lookup"><span data-stu-id="4a009-275">Requirement</span></span>| <span data-ttu-id="4a009-276">值</span><span class="sxs-lookup"><span data-stu-id="4a009-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a009-277">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a009-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a009-278">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-278">1.0</span></span>|
|[<span data-ttu-id="4a009-279">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a009-279">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a009-280">受限</span><span class="sxs-lookup"><span data-stu-id="4a009-280">Restricted</span></span>|
|[<span data-ttu-id="4a009-281">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a009-281">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a009-282">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a009-282">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="4a009-283">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="4a009-283">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="4a009-284">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="4a009-284">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="4a009-285">类型</span><span class="sxs-lookup"><span data-stu-id="4a009-285">Type</span></span>

*   [<span data-ttu-id="4a009-286">UI</span><span class="sxs-lookup"><span data-stu-id="4a009-286">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="4a009-287">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a009-287">Requirements</span></span>

|<span data-ttu-id="4a009-288">要求</span><span class="sxs-lookup"><span data-stu-id="4a009-288">Requirement</span></span>| <span data-ttu-id="4a009-289">值</span><span class="sxs-lookup"><span data-stu-id="4a009-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a009-290">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a009-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a009-291">1.0</span><span class="sxs-lookup"><span data-stu-id="4a009-291">1.0</span></span>|
|[<span data-ttu-id="4a009-292">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a009-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a009-293">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a009-293">Compose or Read</span></span>|
