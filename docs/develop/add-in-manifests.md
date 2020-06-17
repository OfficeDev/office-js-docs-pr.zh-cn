---
title: Office 加载项 XML 清单
description: 获取 Office 加载项清单及其用途概述。
ms.date: 03/18/2020
localization_priority: Priority
ms.openlocfilehash: 0df47ac67a924ab9fd2b3064e0a1ff1b4aa63360
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608992"
---
# <a name="office-add-ins-xml-manifest"></a><span data-ttu-id="62be2-103">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="62be2-103">Office Add-ins XML manifest</span></span>

<span data-ttu-id="62be2-104">Office 外接程序的 XML 清单文件描述，当最终用户安装外接程序并将其与 Office 文档和应用程序配合使用时，应如何激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="62be2-104">The XML manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications.</span></span>

<span data-ttu-id="62be2-105">基于此架构的 XML 清单文件允许 Office 外接程序执行以下内容：</span><span class="sxs-lookup"><span data-stu-id="62be2-105">An XML manifest file based on this schema enables an Office Add-in to do the following:</span></span>

* <span data-ttu-id="62be2-106">通过提供 ID、版本、说明、显示名称和默认区域设置进行自我描述。</span><span class="sxs-lookup"><span data-stu-id="62be2-106">Describe itself by providing an ID, version, description, display name, and default locale.</span></span>

* <span data-ttu-id="62be2-107">指定用于为外接程序塑造品牌的图像，以及用于 Office 功能区中[外接程序命令][]的图标。</span><span class="sxs-lookup"><span data-stu-id="62be2-107">Specify the images used for branding the add-in and iconography used for [add-in commands][] in the Office Ribbon.</span></span>

* <span data-ttu-id="62be2-108">指定外接程序如何与 Office 集成，包括任何自定义 UI，如外接程序创建的功能区按钮。</span><span class="sxs-lookup"><span data-stu-id="62be2-108">Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.</span></span>

* <span data-ttu-id="62be2-109">指定内容外接程序请求的默认尺寸和 Outlook 外接程序请求的高度。</span><span class="sxs-lookup"><span data-stu-id="62be2-109">Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.</span></span>

* <span data-ttu-id="62be2-110">声明 Office 外接程序所需的权限，例如读取或写入文档。</span><span class="sxs-lookup"><span data-stu-id="62be2-110">Declare permissions that the Office Add-in requires, such as reading or writing to the document.</span></span>

* <span data-ttu-id="62be2-111">对于 Outlook 外接程序，定义一个或多个规则，以指定将在其中激活规则并与邮件、约会或会议请求项目交互的上下文。</span><span class="sxs-lookup"><span data-stu-id="62be2-111">For Outlook add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="required-elements"></a><span data-ttu-id="62be2-112">必需元素</span><span class="sxs-lookup"><span data-stu-id="62be2-112">Required elements</span></span>

<span data-ttu-id="62be2-113">下表指定了三种类型 Office 加载项的必需元素。</span><span class="sxs-lookup"><span data-stu-id="62be2-113">The following table specifies the elements that are required for the three types of Office Add-ins.</span></span>

> [!NOTE]
> <span data-ttu-id="62be2-114">还存在强制性命令，其中元素必须出现在其父元素中。</span><span class="sxs-lookup"><span data-stu-id="62be2-114">There is also a mandatory order in which elements must appear within their parent element.</span></span> <span data-ttu-id="62be2-115">有关详细信息，请参阅[如何查找清单元素的正确顺序](manifest-element-ordering.md)。</span><span class="sxs-lookup"><span data-stu-id="62be2-115">For more information see [How to find the proper order of manifest elements](manifest-element-ordering.md).</span></span>


### <a name="required-elements-by-office-add-in-type"></a><span data-ttu-id="62be2-116">Office 加载项类型的必需元素</span><span class="sxs-lookup"><span data-stu-id="62be2-116">Required elements by Office Add-in type</span></span>

| <span data-ttu-id="62be2-117">元素</span><span class="sxs-lookup"><span data-stu-id="62be2-117">Element</span></span>                                                                                      | <span data-ttu-id="62be2-118">内容</span><span class="sxs-lookup"><span data-stu-id="62be2-118">Content</span></span> | <span data-ttu-id="62be2-119">任务窗格</span><span class="sxs-lookup"><span data-stu-id="62be2-119">Task pane</span></span> | <span data-ttu-id="62be2-120">Outlook</span><span class="sxs-lookup"><span data-stu-id="62be2-120">Outlook</span></span> |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| <span data-ttu-id="62be2-121">[OfficeApp][]</span><span class="sxs-lookup"><span data-stu-id="62be2-121">[OfficeApp][]</span></span>                                                                                |    <span data-ttu-id="62be2-122">X</span><span class="sxs-lookup"><span data-stu-id="62be2-122">X</span></span>    |     <span data-ttu-id="62be2-123">X</span><span class="sxs-lookup"><span data-stu-id="62be2-123">X</span></span>     |    <span data-ttu-id="62be2-124">X</span><span class="sxs-lookup"><span data-stu-id="62be2-124">X</span></span>    |
| <span data-ttu-id="62be2-125">[Id][]</span><span class="sxs-lookup"><span data-stu-id="62be2-125">[Id][]</span></span>                                                                                       |    <span data-ttu-id="62be2-126">X</span><span class="sxs-lookup"><span data-stu-id="62be2-126">X</span></span>    |     <span data-ttu-id="62be2-127">X</span><span class="sxs-lookup"><span data-stu-id="62be2-127">X</span></span>     |    <span data-ttu-id="62be2-128">X</span><span class="sxs-lookup"><span data-stu-id="62be2-128">X</span></span>    |
| <span data-ttu-id="62be2-129">[Version][]</span><span class="sxs-lookup"><span data-stu-id="62be2-129">[Version][]</span></span>                                                                                  |    <span data-ttu-id="62be2-130">X</span><span class="sxs-lookup"><span data-stu-id="62be2-130">X</span></span>    |     <span data-ttu-id="62be2-131">X</span><span class="sxs-lookup"><span data-stu-id="62be2-131">X</span></span>     |    <span data-ttu-id="62be2-132">X</span><span class="sxs-lookup"><span data-stu-id="62be2-132">X</span></span>    |
| <span data-ttu-id="62be2-133">[ProviderName][]</span><span class="sxs-lookup"><span data-stu-id="62be2-133">[ProviderName][]</span></span>                                                                             |    <span data-ttu-id="62be2-134">X</span><span class="sxs-lookup"><span data-stu-id="62be2-134">X</span></span>    |     <span data-ttu-id="62be2-135">X</span><span class="sxs-lookup"><span data-stu-id="62be2-135">X</span></span>     |    <span data-ttu-id="62be2-136">X</span><span class="sxs-lookup"><span data-stu-id="62be2-136">X</span></span>    |
| <span data-ttu-id="62be2-137">[DefaultLocale][]</span><span class="sxs-lookup"><span data-stu-id="62be2-137">[DefaultLocale][]</span></span>                                                                            |    <span data-ttu-id="62be2-138">X</span><span class="sxs-lookup"><span data-stu-id="62be2-138">X</span></span>    |     <span data-ttu-id="62be2-139">X</span><span class="sxs-lookup"><span data-stu-id="62be2-139">X</span></span>     |    <span data-ttu-id="62be2-140">X</span><span class="sxs-lookup"><span data-stu-id="62be2-140">X</span></span>    |
| <span data-ttu-id="62be2-141">[DisplayName][]</span><span class="sxs-lookup"><span data-stu-id="62be2-141">[DisplayName][]</span></span>                                                                              |    <span data-ttu-id="62be2-142">X</span><span class="sxs-lookup"><span data-stu-id="62be2-142">X</span></span>    |     <span data-ttu-id="62be2-143">X</span><span class="sxs-lookup"><span data-stu-id="62be2-143">X</span></span>     |    <span data-ttu-id="62be2-144">X</span><span class="sxs-lookup"><span data-stu-id="62be2-144">X</span></span>    |
| <span data-ttu-id="62be2-145">[Description][]</span><span class="sxs-lookup"><span data-stu-id="62be2-145">[Description][]</span></span>                                                                              |    <span data-ttu-id="62be2-146">X</span><span class="sxs-lookup"><span data-stu-id="62be2-146">X</span></span>    |     <span data-ttu-id="62be2-147">X</span><span class="sxs-lookup"><span data-stu-id="62be2-147">X</span></span>     |    <span data-ttu-id="62be2-148">X</span><span class="sxs-lookup"><span data-stu-id="62be2-148">X</span></span>    |
| <span data-ttu-id="62be2-149">[IconUrl][]</span><span class="sxs-lookup"><span data-stu-id="62be2-149">[IconUrl][]</span></span>                                                                                  |    <span data-ttu-id="62be2-150">X</span><span class="sxs-lookup"><span data-stu-id="62be2-150">X</span></span>    |     <span data-ttu-id="62be2-151">X</span><span class="sxs-lookup"><span data-stu-id="62be2-151">X</span></span>     |    <span data-ttu-id="62be2-152">X</span><span class="sxs-lookup"><span data-stu-id="62be2-152">X</span></span>    |
| <span data-ttu-id="62be2-153">[SupportUrl][]\*\*</span><span class="sxs-lookup"><span data-stu-id="62be2-153">[SupportUrl][]\*\*</span></span>                                                                           |    <span data-ttu-id="62be2-154">X</span><span class="sxs-lookup"><span data-stu-id="62be2-154">X</span></span>    |     <span data-ttu-id="62be2-155">X</span><span class="sxs-lookup"><span data-stu-id="62be2-155">X</span></span>     |    <span data-ttu-id="62be2-156">X</span><span class="sxs-lookup"><span data-stu-id="62be2-156">X</span></span>    |
| <span data-ttu-id="62be2-157">[DefaultSettings (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="62be2-157">[DefaultSettings (ContentApp)][]</span></span><br/><span data-ttu-id="62be2-158">[DefaultSettings (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="62be2-158">[DefaultSettings (TaskPaneApp)][]</span></span>                       |    <span data-ttu-id="62be2-159">X</span><span class="sxs-lookup"><span data-stu-id="62be2-159">X</span></span>    |     <span data-ttu-id="62be2-160">X</span><span class="sxs-lookup"><span data-stu-id="62be2-160">X</span></span>     |         |
| <span data-ttu-id="62be2-161">[SourceLocation (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="62be2-161">[SourceLocation (ContentApp)][]</span></span><br/><span data-ttu-id="62be2-162">[SourceLocation (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="62be2-162">[SourceLocation (TaskPaneApp)][]</span></span>                         |    <span data-ttu-id="62be2-163">X</span><span class="sxs-lookup"><span data-stu-id="62be2-163">X</span></span>    |     <span data-ttu-id="62be2-164">X</span><span class="sxs-lookup"><span data-stu-id="62be2-164">X</span></span>     |         |
| <span data-ttu-id="62be2-165">[DesktopSettings][]</span><span class="sxs-lookup"><span data-stu-id="62be2-165">[DesktopSettings][]</span></span>                                                                          |         |           |    <span data-ttu-id="62be2-166">X</span><span class="sxs-lookup"><span data-stu-id="62be2-166">X</span></span>    |
| <span data-ttu-id="62be2-167">[SourceLocation (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="62be2-167">[SourceLocation (MailApp)][]</span></span>                                                                 |         |           |    <span data-ttu-id="62be2-168">X</span><span class="sxs-lookup"><span data-stu-id="62be2-168">X</span></span>    |
| <span data-ttu-id="62be2-169">[Permissions (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="62be2-169">[Permissions (ContentApp)][]</span></span><br/><span data-ttu-id="62be2-170">[Permissions (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="62be2-170">[Permissions (TaskPaneApp)][]</span></span><br/><span data-ttu-id="62be2-171">[Permissions (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="62be2-171">[Permissions (MailApp)][]</span></span> |    <span data-ttu-id="62be2-172">X</span><span class="sxs-lookup"><span data-stu-id="62be2-172">X</span></span>    |     <span data-ttu-id="62be2-173">X</span><span class="sxs-lookup"><span data-stu-id="62be2-173">X</span></span>     |    <span data-ttu-id="62be2-174">X</span><span class="sxs-lookup"><span data-stu-id="62be2-174">X</span></span>    |
| <span data-ttu-id="62be2-175">[Rule (RuleCollection)][]</span><span class="sxs-lookup"><span data-stu-id="62be2-175">[Rule (RuleCollection)][]</span></span><br/><span data-ttu-id="62be2-176">[Rule (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="62be2-176">[Rule (MailApp)][]</span></span>                                             |         |           |    <span data-ttu-id="62be2-177">X</span><span class="sxs-lookup"><span data-stu-id="62be2-177">X</span></span>    |
| <span data-ttu-id="62be2-178">[Requirements (MailApp)\*][]</span><span class="sxs-lookup"><span data-stu-id="62be2-178">[Requirements (MailApp)\*][]</span></span>                                                                  |         |           |    <span data-ttu-id="62be2-179">X</span><span class="sxs-lookup"><span data-stu-id="62be2-179">X</span></span>    |
| <span data-ttu-id="62be2-180">[Set\*][]</span><span class="sxs-lookup"><span data-stu-id="62be2-180">[Set\*][]</span></span><br/><span data-ttu-id="62be2-181">[Sets (MailAppRequirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="62be2-181">[Sets (MailAppRequirements)\*][]</span></span>                                                 |         |           |    <span data-ttu-id="62be2-182">X</span><span class="sxs-lookup"><span data-stu-id="62be2-182">X</span></span>    |
| <span data-ttu-id="62be2-183">[Form\*][]</span><span class="sxs-lookup"><span data-stu-id="62be2-183">[Form\*][]</span></span><br/><span data-ttu-id="62be2-184">[FormSettings\*][]</span><span class="sxs-lookup"><span data-stu-id="62be2-184">[FormSettings\*][]</span></span>                                                              |         |           |    <span data-ttu-id="62be2-185">X</span><span class="sxs-lookup"><span data-stu-id="62be2-185">X</span></span>    |
| <span data-ttu-id="62be2-186">[Sets (Requirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="62be2-186">[Sets (Requirements)\*][]</span></span>                                                                     |    <span data-ttu-id="62be2-187">X</span><span class="sxs-lookup"><span data-stu-id="62be2-187">X</span></span>    |     <span data-ttu-id="62be2-188">X</span><span class="sxs-lookup"><span data-stu-id="62be2-188">X</span></span>     |         |
| <span data-ttu-id="62be2-189">[Hosts\*][]</span><span class="sxs-lookup"><span data-stu-id="62be2-189">[Hosts\*][]</span></span>                                                                                   |    <span data-ttu-id="62be2-190">X</span><span class="sxs-lookup"><span data-stu-id="62be2-190">X</span></span>    |     <span data-ttu-id="62be2-191">X</span><span class="sxs-lookup"><span data-stu-id="62be2-191">X</span></span>     |         |

<span data-ttu-id="62be2-192">_\*Office 加载项清单架构版本 1.1 中新增_</span><span class="sxs-lookup"><span data-stu-id="62be2-192">_\*Added in the Office Add-in Manifest Schema version 1.1._</span></span>

<span data-ttu-id="62be2-193">_\*\* 仅通过 AppSource 分发的加载项才需要 SupportUrl。_</span><span class="sxs-lookup"><span data-stu-id="62be2-193">_\*\* SupportUrl is only required for add-ins that are distributed through AppSource._</span></span>

<!-- Links for above table -->

[officeapp]: ../reference/manifest/officeapp.md
[id]: ../reference/manifest/id.md
[version]: ../reference/manifest/version.md
[providername]: ../reference/manifest/providername.md
[defaultlocale]: ../reference/manifest/defaultlocale.md
[displayname]: ../reference/manifest/displayname.md
[description]: ../reference/manifest/description.md
[iconurl]: ../reference/manifest/iconurl.md
[supporturl]: ../reference/manifest/supporturl.md
[defaultsettings (contentapp)]: ../reference/manifest/defaultsettings.md
[defaultsettings (taskpaneapp)]: ../reference/manifest/defaultsettings.md
[sourcelocation (contentapp)]: ../reference/manifest/sourcelocation.md
[sourcelocation (taskpaneapp)]: ../reference/manifest/sourcelocation.md
[desktopsettings]: /previous-versions/office/fp179684%28v=office.15%29
[sourcelocation (mailapp)]: /previous-versions/office/fp123668%28v=office.15%29
[permissions (contentapp)]: ../reference/manifest/permissions.md
[permissions (taskpaneapp)]: ../reference/manifest/permissions.md
[permissions (mailapp)]: ../reference/manifest/permissions.md
[rule (rulecollection)]: ../reference/manifest/rule.md
[rule (mailapp)]: ../reference/manifest/rule.md
[requirements (mailapp)*]: ../reference/manifest/requirements.md
[set*]: ../reference/manifest/set.md
[sets (mailapprequirements)*]: ../reference/manifest/sets.md
[form*]: ../reference/manifest/form.md
[formsettings*]: ../reference/manifest/formsettings.md
[sets (requirements)*]: ../reference/manifest/sets.md
[hosts*]: ../reference/manifest/hosts.md

## <a name="hosting-requirements"></a><span data-ttu-id="62be2-221">托管要求</span><span class="sxs-lookup"><span data-stu-id="62be2-221">Hosting requirements</span></span>

<span data-ttu-id="62be2-222">所有图像 URI（如用于[外接程序命令][]的 URI）都必须支持缓存。</span><span class="sxs-lookup"><span data-stu-id="62be2-222">All image URIs, such as those used for [add-in commands][], must support caching.</span></span> <span data-ttu-id="62be2-223">托管图像的服务器不得在 HTTP 响应中返回指定 `no-cache`、`no-store` 或类似选项的 `Cache-Control` 标头。</span><span class="sxs-lookup"><span data-stu-id="62be2-223">The server hosting the image should not return a `Cache-Control` header specifying `no-cache`, `no-store`, or similar options in the HTTP response.</span></span>

<span data-ttu-id="62be2-224">所有 URL（如 [SourceLocation](../reference/manifest/sourcelocation.md) 元素中指定的源文件位置）都应**受 SSL 保护 (HTTPS)**。</span><span class="sxs-lookup"><span data-stu-id="62be2-224">All URLs, such as the source file locations specified in the [SourceLocation](../reference/manifest/sourcelocation.md) element, should be **SSL-secured (HTTPS)**.</span></span> [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a><span data-ttu-id="62be2-225">关于提交到 AppSource 的最佳做法</span><span class="sxs-lookup"><span data-stu-id="62be2-225">Best practices for submitting to AppSource</span></span>

<span data-ttu-id="62be2-p103">确保外接程序 ID 有效且具有唯一 GUID。Web 上提供可用于创建唯一 GUID 的各种 GUID 生成器工具。</span><span class="sxs-lookup"><span data-stu-id="62be2-p103">Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.</span></span>

<span data-ttu-id="62be2-228">提交到 AppSource 的加载项还必须包括 [SupportUrl](../reference/manifest/supporturl.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="62be2-228">Add-ins submitted to AppSource must also include the [SupportUrl](../reference/manifest/supporturl.md) element.</span></span> <span data-ttu-id="62be2-229">有关详细信息，请参阅[提交到 AppSource 的应用和加载项的验证策略](/legal/marketplace/certification-policies)。</span><span class="sxs-lookup"><span data-stu-id="62be2-229">For more information, see [Validation policies for apps and add-ins submitted to AppSource](/legal/marketplace/certification-policies).</span></span>

<span data-ttu-id="62be2-230">仅使用 [AppDomain](../reference/manifest/appdomains.md) 元素指定域（除了在 [SourceLocation](../reference/manifest/sourcelocation.md) 元素中指定的用于身份验证方案的域）。</span><span class="sxs-lookup"><span data-stu-id="62be2-230">Only use the [AppDomains](../reference/manifest/appdomains.md) element to specify domains other than the one specified in the [SourceLocation](../reference/manifest/sourcelocation.md) element for authentication scenarios.</span></span>

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a><span data-ttu-id="62be2-231">指定要在外接程序窗口中打开的域</span><span class="sxs-lookup"><span data-stu-id="62be2-231">Specify domains you want to open in the add-in window</span></span>

<span data-ttu-id="62be2-232">在 Office 网页版中运行时，可以将任务窗格导航到任何 URL。</span><span class="sxs-lookup"><span data-stu-id="62be2-232">When running in Office on the web, your task pane can be navigated to any URL.</span></span> <span data-ttu-id="62be2-233">但在桌面平台中，如果外接程序尝试转到托管起始页（如清单文件的 [SourceLocation](../reference/manifest/sourcelocation.md) 元素中所指定的）的域之外的域中的 URL，则该 URL 将在 Office 主机应用程序的外接程序窗格外的新浏览器窗口中打开。</span><span class="sxs-lookup"><span data-stu-id="62be2-233">However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](../reference/manifest/sourcelocation.md) element of the manifest file), that URL opens in a new browser window outside the add-in pane of the Office host application.</span></span>

<span data-ttu-id="62be2-234">若要重写此（桌面版 Office）操作，请在清单文件的 [AppDomains](../reference/manifest/appdomains.md) 元素中指定的域列表中指定要在外接程序窗口中打开的每个域。</span><span class="sxs-lookup"><span data-stu-id="62be2-234">To override this (desktop Office) behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](../reference/manifest/appdomains.md) element of the manifest file.</span></span> <span data-ttu-id="62be2-235">如果加载项尝试转至该列表的域中的 URL，则它将在 Office 网页版和桌面版中的任务窗口中打开。</span><span class="sxs-lookup"><span data-stu-id="62be2-235">If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both Office on the web and desktop.</span></span> <span data-ttu-id="62be2-236">如果它尝试转至列表之外的域中的 URL，则在桌面版 Office 中，该 URL 将在新的浏览器窗口中（外接程序窗格之外）打开。</span><span class="sxs-lookup"><span data-stu-id="62be2-236">If it tries to go to a URL that isn't in the list, then, in desktop Office, that URL opens in a new browser window (outside the add-in pane).</span></span>

> [!NOTE]
> <span data-ttu-id="62be2-237">该行为有两个例外情况：</span><span class="sxs-lookup"><span data-stu-id="62be2-237">There are two exceptions to this behavior:</span></span>
>
> - <span data-ttu-id="62be2-238">它仅适用于外接程序的根窗格。</span><span class="sxs-lookup"><span data-stu-id="62be2-238">It applies only to the root pane of the add-in.</span></span> <span data-ttu-id="62be2-239">如果外接程序页面中嵌入有 iframe，则可以将该 iframe 定向到任何 URL，不论它是否列在 **AppDomains** 中，即使在桌面版 Office 中也是如此。</span><span class="sxs-lookup"><span data-stu-id="62be2-239">If there is an iframe embedded in the add-in page, the iframe can be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>
> - <span data-ttu-id="62be2-240">使用 [displayDialogAsync](/javascript/api/office/office.ui?view=common-js#displaydialogasync-startaddress--options--callback-) API 打开对话框时，传递到方法的 URL 必须与外接程序位于相同的域，但是之后对话框可以定向到任意 URL，无论其是否列入 **AppDomains** 甚至桌面 Office 中。</span><span class="sxs-lookup"><span data-stu-id="62be2-240">When a dialog is opened with the [displayDialogAsync](/javascript/api/office/office.ui?view=common-js#displaydialogasync-startaddress--options--callback-) API, the URL that is passed to the method must be in the same domain as the add-in, but the dialog can then be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>

<span data-ttu-id="62be2-241">以下 XML 清单示例在 **SourceLocation** 元素中指定的 `https://www.contoso.com` 域中托管其外接程序页面。</span><span class="sxs-lookup"><span data-stu-id="62be2-241">The following XML manifest example hosts its main add-in page in the `https://www.contoso.com` domain as specified in the **SourceLocation** element.</span></span> <span data-ttu-id="62be2-242">它还指定 **AppDomains** 元素列表内 [AppDomain](../reference/manifest/appdomain.md) 元素中的 `https://www.northwindtraders.com` 域。</span><span class="sxs-lookup"><span data-stu-id="62be2-242">It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](../reference/manifest/appdomain.md) element within the **AppDomains** element list.</span></span> <span data-ttu-id="62be2-243">如果加载项转到 `www.northwindtraders.com` 域中的页面，此页面会在加载项窗格中打开，即使是在 Office 桌面版中，也不例外。</span><span class="sxs-lookup"><span data-stu-id="62be2-243">If the add-in goes to a page in the `www.northwindtraders.com` domain, that page opens in the add-in pane, even in Office desktop.</span></span>

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="specify-domains-from-which-officejs-api-calls-are-made"></a><span data-ttu-id="62be2-244">指定从中执行 Office .js API 调用的域</span><span class="sxs-lookup"><span data-stu-id="62be2-244">Specify domains from which Office.js API calls are made</span></span>

<span data-ttu-id="62be2-245">你的加载项可以从清单文件的 [SourceLocation](../reference/manifest/sourcelocation.md) 元素中引用的域执行 Office.js API 调用。</span><span class="sxs-lookup"><span data-stu-id="62be2-245">Your add-in can make Office.js API calls from the domain referenced in the [SourceLocation](../reference/manifest/sourcelocation.md) element of the manifest file.</span></span> <span data-ttu-id="62be2-246">如果加载项中有需要访问 Office.js API 的其他 IFrame，请将该源 URL 的域添加到在清单文件的 [AppDomains](../reference/manifest/appdomains.md) 元素中指定的列表。</span><span class="sxs-lookup"><span data-stu-id="62be2-246">If you have other IFrames within your add-in that need to access Office.js APIs, add the domain of that source URL to the list specified in the [AppDomains](../reference/manifest/appdomains.md) element of the manifest file.</span></span> <span data-ttu-id="62be2-247">如果有一个未包含在 `AppDomains` 列表中且具有源的 IFrame 尝试执行 Office.js API 调用，则加载项将收到[“权限被拒绝”错误](../reference/javascript-api-for-office-error-codes.md)。</span><span class="sxs-lookup"><span data-stu-id="62be2-247">If an IFrame with a source not contained in the `AppDomains` list attempts to make an Office.js API call, then the add-in will receive a [permission denied error](../reference/javascript-api-for-office-error-codes.md).</span></span>

## <a name="manifest-v11-xml-file-examples-and-schemas"></a><span data-ttu-id="62be2-248">清单 v1.1 XML 文件示例和架构</span><span class="sxs-lookup"><span data-stu-id="62be2-248">Manifest v1.1 XML file examples and schemas</span></span>

<span data-ttu-id="62be2-249">下面各部分展示了内容加载项、任务窗格加载项和 Outlook 加载项的清单 v1.1 XML 文件示例。</span><span class="sxs-lookup"><span data-stu-id="62be2-249">The following sections show examples of manifest v1.1 XML files for content, task pane, and Outlook add-ins.</span></span>

# <a name="task-pane"></a>[<span data-ttu-id="62be2-250">任务窗格</span><span class="sxs-lookup"><span data-stu-id="62be2-250">Task pane</span></span>](#tab/tabid-1)

[<span data-ttu-id="62be2-251">加载项清单架构</span><span class="sxs-lookup"><span data-stu-id="62be2-251">Add-in manifest schemas</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

  <!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

  <!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

  <!--IMPORTANT! Id must be unique for your add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
  <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <!--BeginTaskpaneMode integration. Office 2013 and any client that doesn't understand commands will use this section.
    This section will also be used if there are no VersionOverrides -->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
  </DefaultSettings>
  <!--EndTaskpaneMode integration -->

  <Permissions>ReadWriteDocument</Permissions>

  <!--BeginAddinCommandsMode integration-->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <!--Each host can have a different set of commands. Cool huh!? -->
      <!-- Workbook=Excel Document=Word Presentation=PowerPoint -->
      <!-- Make sure the hosts you override match the hosts declared in the top section of the manifest -->
      <Host xsi:type="Document">
        <!-- Form factor. Currently only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
          <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called.
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office Ribbon-->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!--Use OfficeTab to extend an existing Tab. Use CustomTab to create a new tab -->
            <!-- Documentation includes all the IDs currently tested to work -->
            <CustomTab id="Contoso.Tab1">
              <!--Group ID-->
              <Group id="Contoso.Tab1.Group1">
                <!--Label for your group. resid must point to a ShortString resource -->
                <Label resid="Contoso.Tab1.GroupLabel" />
                <Icon>
                  <!-- Sample Todo: Each size needs its own icon resource or it will look distorted when resized -->
                  <!--Icons. Required sizes: 16, 32, 80; optional: 20, 24, 40, 48, 64. You should provide as many sizes as possible for a great user experience. -->
                  <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                </Icon>

                <!--Control. It can be of type "Button" or "Menu" -->
                <Control xsi:type="Button" id="Contoso.FunctionButton">
                  <!--Label for your button. resid must point to a ShortString resource -->
                  <Label resid="Contoso.FunctionButton.Label" />
                  <Supertip>
                    <!--ToolTip title. resid must point to a ShortString resource -->
                    <Title resid="Contoso.FunctionButton.Label" />
                    <!--ToolTip description. resid must point to a LongString resource -->
                    <Description resid="Contoso.FunctionButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon80" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFunction or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  <Action xsi:type="ExecuteFunction">
                    <!--Name of the function to call. This function needs to exist in the global DOM namespace of the function file-->
                    <FunctionName>writeText</FunctionName>
                  </Action>
                </Control>

                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Button2Id1</TaskpaneId>
                    <!--Provide a url resource id for the location that will be displayed on the task pane -->
                    <SourceLocation resid="Contoso.Taskpane1.Url" />
                  </Action>
                </Control>
                <!-- Menu example -->
                <Control xsi:type="Menu" id="Contoso.Menu">
                  <Label resid="Contoso.Dropdown.Label" />
                  <Supertip>
                    <Title resid="Contoso.Dropdown.Label" />
                    <Description resid="Contoso.Dropdown.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane1.Url" />
                      </Action>
                    </Item>

                    <Item id="Contoso.Menu.Item2">
                      <Label resid="Contoso.Item2.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item2.Label" />
                        <Description resid="Contoso.Item2.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID2</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane2.Url" />
                      </Action>
                    </Item>

                  </Items>
                </Control>

              </Group>

              <!-- Label of your tab -->
              <!-- If validating with XSD it needs to be at the end -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon16" DefaultValue="https://myCDN/Images/Button16x16.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon32" DefaultValue="https://myCDN/Images/Button32x32.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon80" DefaultValue="https://myCDN/Images/Button80x80.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://i.imgur.com/qDujiX0.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Contoso.FunctionFile.Url" DefaultValue="https://commandsimple.azurewebsites.net/FunctionFile.html" />
        <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
        <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Contoso.FunctionButton.Label" DefaultValue="Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Label" DefaultValue="Show Taskpane" />
        <bt:String id="Contoso.Dropdown.Label" DefaultValue="Dropdown" />
        <bt:String id="Contoso.Item1.Label" DefaultValue="Show Taskpane 1" />
        <bt:String id="Contoso.Item2.Label" DefaultValue="Show Taskpane 2" />
        <bt:String id="Contoso.Tab1.GroupLabel" DefaultValue="Test Group" />
         <bt:String id="Contoso.Tab1.TabLabel" DefaultValue="Test Tab" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Contoso.FunctionButton.Tooltip" DefaultValue="Click to Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane" />
        <bt:String id="Contoso.Dropdown.Tooltip" DefaultValue="Click to Show Options on this Menu" />
        <bt:String id="Contoso.Item1.Tooltip" DefaultValue="Click to Show Taskpane1" />
        <bt:String id="Contoso.Item2.Tooltip" DefaultValue="Click to Show Taskpane2" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

# <a name="content"></a>[<span data-ttu-id="62be2-252">内容</span><span class="sxs-lookup"><span data-stu-id="62be2-252">Content</span></span>](#tab/tabid-2)

[<span data-ttu-id="62be2-253">加载项清单架构</span><span class="sxs-lookup"><span data-stu-id="62be2-253">Add-in manifest schemas</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Database" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" />
    </Sets>
  </Requirements>  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com/apps/content.html" />
    <RequestedWidth>400</RequestedWidth>
    <RequestedHeight>400</RequestedHeight>
  </DefaultSettings>
  <Permissions>Restricted</Permissions>
  <AllowSnapshot>true</AllowSnapshot>
</OfficeApp>
```

# <a name="mail"></a>[<span data-ttu-id="62be2-254">邮件</span><span class="sxs-lookup"><span data-stu-id="62be2-254">Mail</span></span>](#tab/tabid-3)

[<span data-ttu-id="62be2-255">加载项清单架构</span><span class="sxs-lookup"><span data-stu-id="62be2-255">Add-in manifest schemas</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>971E76EF-D73E-567F-ADAE-5A76B39052CF</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="YouTube"/>
  <Description DefaultValue=
    "Watch YouTube videos referenced in the e-mails you  
    receive without leaving your email client.">
    <Override Locale="fr-fr" Value="Visualisez les vidéos
      YouTube références dans vos courriers électronique
      directement depuis Outlook."/>
  </Description>
  <!-- Change the following lines to specify    -->
  <!-- the web server that hosts the icon files. -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-64.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_desktop.htm" />
        <RequestedHeight>216</RequestedHeight>
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_tablet.htm" />
        <RequestedHeight>216</RequestedHeight>
      </TabletSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_desktop.htm" />
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_tablet.htm" />
      </TabletSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="RuleCollection" Mode="And">
      <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
      </Rule>
      <Rule xsi:type="ItemHasRegularExpressionMatch"
        PropertyName="BodyAsPlaintext" RegExName="VideoURL"
        RegExValue=
        "http://(((www\.)?youtube\.com/watch\?v=)|
        (youtu\.be/))[a-zA-Z0-9_-]{11}" />
    </Rule>
    <Rule xsi:type="RuleCollection" Mode="Or">
      <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
      <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    </Rule>
  </Rule>
</OfficeApp>
```

---

## <a name="validate-an-office-add-ins-manifest"></a><span data-ttu-id="62be2-256">验证 Office 加载项的清单</span><span class="sxs-lookup"><span data-stu-id="62be2-256">Validate an Office Add-in's manifest</span></span>

<span data-ttu-id="62be2-257">有关根据 [XML 架构定义 (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) 验证清单的信息，请参阅[验证 Office 加载项的清单](../testing/troubleshoot-manifest.md)。</span><span class="sxs-lookup"><span data-stu-id="62be2-257">For information about validating a manifest against the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8), see [Validate an Office Add-in's manifest](../testing/troubleshoot-manifest.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="62be2-258">另请参阅</span><span class="sxs-lookup"><span data-stu-id="62be2-258">See also</span></span>

* [<span data-ttu-id="62be2-259">如何查找清单元素的正确顺序</span><span class="sxs-lookup"><span data-stu-id="62be2-259">How to find the proper order of manifest elements</span></span>](manifest-element-ordering.md)
* <span data-ttu-id="62be2-260">[在清单中创建加载项命令][加载项命令]</span><span class="sxs-lookup"><span data-stu-id="62be2-260">[Create add-in commands in your manifest][add-in commands]</span></span>
* [<span data-ttu-id="62be2-261">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="62be2-261">Specify Office hosts and API requirements</span></span>](specify-office-hosts-and-api-requirements.md)
* [<span data-ttu-id="62be2-262">Office 外接程序的本地化</span><span class="sxs-lookup"><span data-stu-id="62be2-262">Localization for Office Add-ins</span></span>](localization.md)
* [<span data-ttu-id="62be2-263">Office 外接程序清单的架构参考</span><span class="sxs-lookup"><span data-stu-id="62be2-263">Schema reference for Office Add-ins manifests</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
* [<span data-ttu-id="62be2-264">更新 API 和清单版本</span><span class="sxs-lookup"><span data-stu-id="62be2-264">Update API and manifest version</span></span>](update-your-javascript-api-for-office-and-manifest-schema-version.md)
* [<span data-ttu-id="62be2-265">标识等效的 COM 加载项</span><span class="sxs-lookup"><span data-stu-id="62be2-265">Identify an equivalent COM add-in</span></span>](make-office-add-in-compatible-with-existing-com-add-in.md)
* [<span data-ttu-id="62be2-266">在加载项中请求获取 API 使用权限</span><span class="sxs-lookup"><span data-stu-id="62be2-266">Requesting permissions for API use in add-ins</span></span>](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
* [<span data-ttu-id="62be2-267">验证 Office 加载项的清单</span><span class="sxs-lookup"><span data-stu-id="62be2-267">Validate an Office Add-in's manifest</span></span>](../testing/troubleshoot-manifest.md)

[加载项命令]: create-addin-commands.md
[add-in commands]: create-addin-commands.md
