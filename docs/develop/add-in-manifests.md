---
title: Office 加载项 XML 清单
description: ''
ms.date: 02/09/2018
ms.openlocfilehash: e25d465b39cea0a13a890fec95fafdbeafff0ca5
ms.sourcegitcommit: 9b021af6cb23a58486d6c5c7492be425e309bea1
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/15/2018
ms.locfileid: "26533705"
---
# <a name="office-add-ins-xml-manifest"></a><span data-ttu-id="de648-102">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="de648-102">Office Add-ins XML manifest</span></span>

<span data-ttu-id="de648-103">Office 外接程序的 XML 清单文件描述，当最终用户安装外接程序并将其与 Office 文档和应用程序配合使用时，应如何激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="de648-103">The XML manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications.</span></span>

<span data-ttu-id="de648-104">基于此架构的 XML 清单文件允许 Office 外接程序执行以下内容：</span><span class="sxs-lookup"><span data-stu-id="de648-104">An XML manifest file based on this schema enables an Office Add-in to do the following:</span></span>

* <span data-ttu-id="de648-105">通过提供 ID、版本、说明、显示名称和默认区域设置进行自我描述。</span><span class="sxs-lookup"><span data-stu-id="de648-105">Describe itself by providing an ID, version, description, display name, and default locale.</span></span>

* <span data-ttu-id="de648-106">指定用于为外接程序塑造品牌的图像，以及用于 Office 功能区中[外接程序命令][]的图标。</span><span class="sxs-lookup"><span data-stu-id="de648-106">Specify the images used for branding the Add-in and iconography used for [Add-in Commands][] in the Office Ribbon.</span></span>

* <span data-ttu-id="de648-107">指定外接程序如何与 Office 集成，包括任何自定义 UI，如外接程序创建的功能区按钮。</span><span class="sxs-lookup"><span data-stu-id="de648-107">Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.</span></span>

* <span data-ttu-id="de648-108">指定内容外接程序请求的默认尺寸和 Outlook 外接程序请求的高度。</span><span class="sxs-lookup"><span data-stu-id="de648-108">Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.</span></span>

* <span data-ttu-id="de648-109">声明 Office 外接程序所需的权限，例如读取或写入文档。</span><span class="sxs-lookup"><span data-stu-id="de648-109">Declare permissions that the Office Add-in requires, such as reading or writing to the document.</span></span>

* <span data-ttu-id="de648-110">对于 Outlook 外接程序，定义一个或多个规则，以指定将在其中激活规则并与邮件、约会或会议请求项目交互的上下文。</span><span class="sxs-lookup"><span data-stu-id="de648-110">For Outlook add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.</span></span>

> [!NOTE]
> <span data-ttu-id="de648-p101">如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](https://docs.microsoft.com/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性](../overview/office-add-in-availability.md)页面）。</span><span class="sxs-lookup"><span data-stu-id="de648-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="required-elements"></a><span data-ttu-id="de648-113">必需元素</span><span class="sxs-lookup"><span data-stu-id="de648-113">Required elements</span></span>

<span data-ttu-id="de648-114">下表指定了三种类型 Office 加载项的必需元素。</span><span class="sxs-lookup"><span data-stu-id="de648-114">The following table specifies the elements that are required for the three types of Office Add-ins.</span></span>

### <a name="required-elements-by-office-add-in-type"></a><span data-ttu-id="de648-115">Office 加载项类型的必需元素</span><span class="sxs-lookup"><span data-stu-id="de648-115">Required elements by Office Add-in type</span></span>

| <span data-ttu-id="de648-116">元素</span><span class="sxs-lookup"><span data-stu-id="de648-116">Element</span></span>                                                                                      | <span data-ttu-id="de648-117">内容</span><span class="sxs-lookup"><span data-stu-id="de648-117">Content</span></span> | <span data-ttu-id="de648-118">任务窗格</span><span class="sxs-lookup"><span data-stu-id="de648-118">Task pane</span></span> | <span data-ttu-id="de648-119">Outlook</span><span class="sxs-lookup"><span data-stu-id="de648-119">Outlook</span></span> |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| <span data-ttu-id="de648-120">[OfficeApp][]</span><span class="sxs-lookup"><span data-stu-id="de648-120">[OfficeApp][]</span></span>                                                                                |    <span data-ttu-id="de648-121">X</span><span class="sxs-lookup"><span data-stu-id="de648-121">X</span></span>    |     <span data-ttu-id="de648-122">X</span><span class="sxs-lookup"><span data-stu-id="de648-122">X</span></span>     |    <span data-ttu-id="de648-123">X</span><span class="sxs-lookup"><span data-stu-id="de648-123">X</span></span>    |
| <span data-ttu-id="de648-124">
  [Id][]</span><span class="sxs-lookup"><span data-stu-id="de648-124">[Id][]</span></span>                                                                                       |    <span data-ttu-id="de648-125">X</span><span class="sxs-lookup"><span data-stu-id="de648-125">X</span></span>    |     <span data-ttu-id="de648-126">X</span><span class="sxs-lookup"><span data-stu-id="de648-126">X</span></span>     |    <span data-ttu-id="de648-127">X</span><span class="sxs-lookup"><span data-stu-id="de648-127">X</span></span>    |
| <span data-ttu-id="de648-128">[版本][]</span><span class="sxs-lookup"><span data-stu-id="de648-128">[Version][]</span></span>                                                                                  |    <span data-ttu-id="de648-129">X</span><span class="sxs-lookup"><span data-stu-id="de648-129">X</span></span>    |     <span data-ttu-id="de648-130">X</span><span class="sxs-lookup"><span data-stu-id="de648-130">X</span></span>     |    <span data-ttu-id="de648-131">X</span><span class="sxs-lookup"><span data-stu-id="de648-131">X</span></span>    |
| <span data-ttu-id="de648-132">[ProviderName][]</span><span class="sxs-lookup"><span data-stu-id="de648-132">[ProviderName][]</span></span>                                                                             |    <span data-ttu-id="de648-133">X</span><span class="sxs-lookup"><span data-stu-id="de648-133">X</span></span>    |     <span data-ttu-id="de648-134">X</span><span class="sxs-lookup"><span data-stu-id="de648-134">X</span></span>     |    <span data-ttu-id="de648-135">X</span><span class="sxs-lookup"><span data-stu-id="de648-135">X</span></span>    |
| <span data-ttu-id="de648-136">[DefaultLocale][]</span><span class="sxs-lookup"><span data-stu-id="de648-136">[DefaultLocale][]</span></span>                                                                            |    <span data-ttu-id="de648-137">X</span><span class="sxs-lookup"><span data-stu-id="de648-137">X</span></span>    |     <span data-ttu-id="de648-138">X</span><span class="sxs-lookup"><span data-stu-id="de648-138">X</span></span>     |    <span data-ttu-id="de648-139">X</span><span class="sxs-lookup"><span data-stu-id="de648-139">X</span></span>    |
| <span data-ttu-id="de648-140">[DisplayName][]</span><span class="sxs-lookup"><span data-stu-id="de648-140">[DisplayName][]</span></span>                                                                              |    <span data-ttu-id="de648-141">X</span><span class="sxs-lookup"><span data-stu-id="de648-141">X</span></span>    |     <span data-ttu-id="de648-142">X</span><span class="sxs-lookup"><span data-stu-id="de648-142">X</span></span>     |    <span data-ttu-id="de648-143">X</span><span class="sxs-lookup"><span data-stu-id="de648-143">X</span></span>    |
| <span data-ttu-id="de648-144">[说明][]</span><span class="sxs-lookup"><span data-stu-id="de648-144">[Description][]</span></span>                                                                              |    <span data-ttu-id="de648-145">X</span><span class="sxs-lookup"><span data-stu-id="de648-145">X</span></span>    |     <span data-ttu-id="de648-146">X</span><span class="sxs-lookup"><span data-stu-id="de648-146">X</span></span>     |    <span data-ttu-id="de648-147">X</span><span class="sxs-lookup"><span data-stu-id="de648-147">X</span></span>    |
| <span data-ttu-id="de648-148">[IconUrl][]</span><span class="sxs-lookup"><span data-stu-id="de648-148">[IconUrl][]</span></span>                                                                                  |    <span data-ttu-id="de648-149">X</span><span class="sxs-lookup"><span data-stu-id="de648-149">X</span></span>    |     <span data-ttu-id="de648-150">X</span><span class="sxs-lookup"><span data-stu-id="de648-150">X</span></span>     |    <span data-ttu-id="de648-151">X</span><span class="sxs-lookup"><span data-stu-id="de648-151">X</span></span>    |
| <span data-ttu-id="de648-152">[DefaultSettings (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="de648-152">[DefaultSettings (ContentApp)][]</span></span><br/><span data-ttu-id="de648-153">[DefaultSettings (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="de648-153">[DefaultSettings (TaskPaneApp)][]</span></span>                       |    <span data-ttu-id="de648-154">X</span><span class="sxs-lookup"><span data-stu-id="de648-154">X</span></span>    |     <span data-ttu-id="de648-155">X</span><span class="sxs-lookup"><span data-stu-id="de648-155">X</span></span>     |         |
| <span data-ttu-id="de648-156">[SourceLocation (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="de648-156">[SourceLocation (ContentApp)][]</span></span><br/><span data-ttu-id="de648-157">[SourceLocation (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="de648-157">[SourceLocation (TaskPaneApp)][]</span></span>                         |    <span data-ttu-id="de648-158">X</span><span class="sxs-lookup"><span data-stu-id="de648-158">X</span></span>    |     <span data-ttu-id="de648-159">X</span><span class="sxs-lookup"><span data-stu-id="de648-159">X</span></span>     |         |
| <span data-ttu-id="de648-160">[DesktopSettings][]</span><span class="sxs-lookup"><span data-stu-id="de648-160">[DesktopSettings][]</span></span>                                                                          |         |           |    <span data-ttu-id="de648-161">X</span><span class="sxs-lookup"><span data-stu-id="de648-161">X</span></span>    |
| <span data-ttu-id="de648-162">[SourceLocation (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="de648-162">[SourceLocation (MailApp)][]</span></span>                                                                 |         |           |    <span data-ttu-id="de648-163">X</span><span class="sxs-lookup"><span data-stu-id="de648-163">X</span></span>    |
| <span data-ttu-id="de648-164">
  [Permissions (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="de648-164">[Permissions (ContentApp)][]</span></span><br/><span data-ttu-id="de648-165">
  [Permissions (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="de648-165">[Permissions (TaskPaneApp)][]</span></span><br/><span data-ttu-id="de648-166">
  [Permissions (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="de648-166">[Permissions (MailApp)][]</span></span> |    <span data-ttu-id="de648-167">X</span><span class="sxs-lookup"><span data-stu-id="de648-167">X</span></span>    |     <span data-ttu-id="de648-168">X</span><span class="sxs-lookup"><span data-stu-id="de648-168">X</span></span>     |    <span data-ttu-id="de648-169">X</span><span class="sxs-lookup"><span data-stu-id="de648-169">X</span></span>    |
| <span data-ttu-id="de648-170">
  [Rule (RuleCollection)][]</span><span class="sxs-lookup"><span data-stu-id="de648-170">[Rule (RuleCollection)][]</span></span><br/><span data-ttu-id="de648-171">
  [Rule (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="de648-171">[Rule (MailApp)][]</span></span>                                             |         |           |    <span data-ttu-id="de648-172">X</span><span class="sxs-lookup"><span data-stu-id="de648-172">X</span></span>    |
| <span data-ttu-id="de648-173">[Requirements (MailApp)\*][]</span><span class="sxs-lookup"><span data-stu-id="de648-173">[\*Requirements (MailApp)][]</span></span>                                                                  |         |           |    <span data-ttu-id="de648-174">X</span><span class="sxs-lookup"><span data-stu-id="de648-174">X</span></span>    |
| <span data-ttu-id="de648-175">[Set\*][]</span><span class="sxs-lookup"><span data-stu-id="de648-175">[\*Set][]</span></span><br/><span data-ttu-id="de648-176">[Sets (MailAppRequirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="de648-176">[\*\*Sets (MailAppRequirements)][]</span></span>                                                 |         |           |    <span data-ttu-id="de648-177">X</span><span class="sxs-lookup"><span data-stu-id="de648-177">X</span></span>    |
| <span data-ttu-id="de648-178">[Form\*][]</span><span class="sxs-lookup"><span data-stu-id="de648-178">[\*Form][]</span></span><br/><span data-ttu-id="de648-179">[FormSettings\*][]</span><span class="sxs-lookup"><span data-stu-id="de648-179">[\*\*FormSettings][]</span></span>                                                              |         |           |    <span data-ttu-id="de648-180">X</span><span class="sxs-lookup"><span data-stu-id="de648-180">X</span></span>    |
| <span data-ttu-id="de648-181">[Sets (Requirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="de648-181">[\*Sets (Requirements)][]</span></span>                                                                     |    <span data-ttu-id="de648-182">X</span><span class="sxs-lookup"><span data-stu-id="de648-182">X</span></span>    |     <span data-ttu-id="de648-183">X</span><span class="sxs-lookup"><span data-stu-id="de648-183">X</span></span>     |         |
| <span data-ttu-id="de648-184">[Hosts\*][]</span><span class="sxs-lookup"><span data-stu-id="de648-184">[\*Hosts][]</span></span>                                                                                   |    <span data-ttu-id="de648-185">X</span><span class="sxs-lookup"><span data-stu-id="de648-185">X</span></span>    |     <span data-ttu-id="de648-186">X</span><span class="sxs-lookup"><span data-stu-id="de648-186">X</span></span>     |         |

<span data-ttu-id="de648-187">_\*Office 加载项清单架构版本 1.1 中新增_</span><span class="sxs-lookup"><span data-stu-id="de648-187">_\*Added in the Office Add-in Manifest Schema version 1.1._</span></span>

<!-- Links for above table -->

[officeapp]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/officeapp?view=office-js
[id]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/id
[version]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/version
[providername]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/providername
[defaultlocale]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultlocale
[displayname]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/displayname
[description]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/description
[iconurl]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/iconurl
[IconUrl]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/iconurl
[defaultsettings (contentapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultsettings
[defaultsettings (taskpaneapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultsettings
[sourcelocation (contentapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation
[sourcelocation (taskpaneapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation
[desktopsettings]: https://msdn.microsoft.com/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c(Office.15).aspx
[sourcelocation (mailapp)]: http://msdn.microsoft.com/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx
[permissions (contentapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/permissions
[permissions (taskpaneapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/permissions
[permissions (mailapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/permissions
[rule (rulecollection)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/rule
[rule (mailapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/rule
[requirements (mailapp)*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/requirements
[set*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/set
[sets (mailapprequirements)*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sets
[form*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/form
[formsettings*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/formsettings
[sets (requirements)*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sets
[hosts\*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/hosts
[Hosts]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/hosts

## <a name="hosting-requirements"></a><span data-ttu-id="de648-214">托管要求</span><span class="sxs-lookup"><span data-stu-id="de648-214">Hosting requirements</span></span>

<span data-ttu-id="de648-215">所有图像 URI（如用于[外接程序命令][]的 URI）都必须支持缓存。</span><span class="sxs-lookup"><span data-stu-id="de648-215">All image URIs, such as those used for [Add-in Commands][], must support caching.</span></span> <span data-ttu-id="de648-216">托管图像的服务器不得在 HTTP 响应中返回指定 `no-cache`、`no-store` 或类似选项的 `Cache-Control` 标头。</span><span class="sxs-lookup"><span data-stu-id="de648-216">The server hosting the image should not return a `Cache-Control` header specifying `no-cache`, `no-store`, or similar options in the HTTP response.</span></span>

<span data-ttu-id="de648-217">所有 URL（如 [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) 元素中指定的源文件位置）都应**受 SSL 保护 (HTTPS)**。</span><span class="sxs-lookup"><span data-stu-id="de648-217">All URLs, such as the source file locations specified in the [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) element, should be **SSL-secured (HTTPS)**.</span></span> [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a><span data-ttu-id="de648-218">关于提交到 AppSource 的最佳做法</span><span class="sxs-lookup"><span data-stu-id="de648-218">Best practices for submitting to AppSource</span></span>

<span data-ttu-id="de648-p103">确保外接程序 ID 有效且具有唯一 GUID。Web 上提供可用于创建唯一 GUID 的各种 GUID 生成器工具。</span><span class="sxs-lookup"><span data-stu-id="de648-p103">Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.</span></span>

<span data-ttu-id="de648-221">提交到 AppSource 的加载项还必须包括 [SupportUrl](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/supporturl) 元素。</span><span class="sxs-lookup"><span data-stu-id="de648-221">Add-ins submitted to AppSource must also include the [SupportUrl](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/supporturl) element.</span></span> <span data-ttu-id="de648-222">有关详细信息，请参阅[提交到 AppSource 的应用和加载项的验证策略](https://docs.microsoft.com/office/dev/store/validation-policies)。</span><span class="sxs-lookup"><span data-stu-id="de648-222">For more information, see [Validation policies for apps and add-ins submitted to AppSource](https://docs.microsoft.com/office/dev/store/validation-policies).</span></span>

<span data-ttu-id="de648-223">仅使用 [AppDomain](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomains) 元素指定域（除了在 [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) 元素中指定的用于身份验证方案的域）。</span><span class="sxs-lookup"><span data-stu-id="de648-223">Only use the [AppDomains](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomains) element to specify domains other than the one specified in the [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) element for authentication scenarios.</span></span>

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a><span data-ttu-id="de648-224">指定要在外接程序窗口中打开的域</span><span class="sxs-lookup"><span data-stu-id="de648-224">Specify domains you want to open in the add-in window</span></span>

<span data-ttu-id="de648-225">在 Office Online 中运行时，可以将任务窗格导航到任何 URL。</span><span class="sxs-lookup"><span data-stu-id="de648-225">When running in Office Online, your task pane can be navigated to any URL.</span></span> <span data-ttu-id="de648-226">但在桌面平台中，如果外接程序尝试转到托管起始页（如清单文件的 [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) 元素中所指定的）的域之外的域中的 URL，则该 URL 将在 Office 主机应用程序的外接程序窗格外的新浏览器窗口中打开。</span><span class="sxs-lookup"><span data-stu-id="de648-226">By default, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) element of the manifest file), that URL will open in a new browser window outside the add-in pane of the Office host application.</span></span>

<span data-ttu-id="de648-227">若要重写此（桌面版 Office）操作，请在清单文件的 [AppDomains](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomains) 元素中指定的域列表中指定要在外接程序窗口中打开的每个域。</span><span class="sxs-lookup"><span data-stu-id="de648-227">To override this behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomains) element of the manifest file.</span></span> <span data-ttu-id="de648-228">如果外接程序尝试转至该列表的域中的 URL，则它将在桌面版 Office 和 Office Online 中的任务窗口中打开。</span><span class="sxs-lookup"><span data-stu-id="de648-228">If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both desktop Office and Office Online.</span></span> <span data-ttu-id="de648-229">如果它尝试转至列表之外的域中的 URL，则在桌面版 Office 中，该 URL 将在新的浏览器窗口中（外接程序窗格之外）打开。</span><span class="sxs-lookup"><span data-stu-id="de648-229">If the add-in tries to go to a URL in a domain that isn't in the list, that URL will open in a new browser window (outside the add-in pane).</span></span>

> [!NOTE]
> <span data-ttu-id="de648-230">此操作仅适用于外接程序的根窗格。</span><span class="sxs-lookup"><span data-stu-id="de648-230">This behavior applies only to the root pane of the add-in.</span></span> <span data-ttu-id="de648-231">如果外接程序页面中嵌入有 iframe，则可以将该 iframe 定向到任何 URL，不论它是否列在 **AppDomains** 中，即使在桌面版 Office 中也是如此。</span><span class="sxs-lookup"><span data-stu-id="de648-231">If there is an iframe embedded in the add-in page, the iframe can be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>

<span data-ttu-id="de648-232">以下 XML 清单示例在 **SourceLocation** 元素中指定的 `https://www.contoso.com` 域中托管其外接程序页面。</span><span class="sxs-lookup"><span data-stu-id="de648-232">The following XML manifest example hosts its main add-in page in the  `https://www.contoso.com` domain as specified in the **SourceLocation** element.</span></span> <span data-ttu-id="de648-233">它还指定 **AppDomains** 元素列表内 [AppDomain](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomain) 元素中的 `https://www.northwindtraders.com` 域。</span><span class="sxs-lookup"><span data-stu-id="de648-233">It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomain) element within the **AppDomains** element list.</span></span> <span data-ttu-id="de648-234">如果外接程序转至 www.northwindtraders.com 域中的页面，则该页面将在外接程序窗格中打开，即使在 Office 桌面版中也是如此。</span><span class="sxs-lookup"><span data-stu-id="de648-234">If the add-in goes to a page in the www.northwindtraders.com domain, that page will open in the add-in pane.</span></span>

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="manifest-v11-xml-file-examples-and-schemas"></a><span data-ttu-id="de648-235">清单 v1.1 XML 文件示例和架构</span><span class="sxs-lookup"><span data-stu-id="de648-235">Manifest v1.1 XML file examples and schemas</span></span>

<span data-ttu-id="de648-236">下面各部分展示了内容加载项、任务窗格加载项和 Outlook 加载项的清单 v1.1 XML 文件示例。</span><span class="sxs-lookup"><span data-stu-id="de648-236">The following sections show examples of manifest v1.1 XML files for content, task pane, and Outlook add-ins.</span></span>

# <a name="task-panetabtabid-1"></a>[<span data-ttu-id="de648-237">任务窗格</span><span class="sxs-lookup"><span data-stu-id="de648-237">Task pane</span></span>](#tab/tabid-1)

[<span data-ttu-id="de648-238">任务窗格应用程序清单架构</span><span class="sxs-lookup"><span data-stu-id="de648-238">Task pane app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/taskpane)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

  <!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

  <!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

  <!--IMPORTANT! Id must be unique for your add-in. If you clone this manifest ensure that you change this id to your own GUID -->
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
                  <!--Icons. Required sizes 16,31,80, optional 20, 24, 40, 48, 64. Strongly recommended to provide all sizes for great UX -->
                  <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon" />
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
              <!-- If validating with XSD it needs to be at the end, we might change this before release -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon" DefaultValue="https://i.imgur.com/FkSShX9.png" />
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

# <a name="contenttabtabid-2"></a>[<span data-ttu-id="de648-239">内容</span><span class="sxs-lookup"><span data-stu-id="de648-239">Content</span></span>](#tab/tabid-2)

[<span data-ttu-id="de648-240">内容应用程序清单架构</span><span class="sxs-lookup"><span data-stu-id="de648-240">Content app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/content)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
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

# <a name="mailtabtabid-3"></a>[<span data-ttu-id="de648-241">邮件</span><span class="sxs-lookup"><span data-stu-id="de648-241">Mail</span></span>](#tab/tabid-3)

[<span data-ttu-id="de648-242">邮件应用程序清单架构</span><span class="sxs-lookup"><span data-stu-id="de648-242">Mail app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/mail)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">

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
      directement depuis Outlook et Outlook Web App."/>
  </Description>
  <!-- Change the following lines to specify    -->
  <!-- the web server that hosts the icon files. -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />

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

## <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="de648-243">验证并排查清单问题</span><span class="sxs-lookup"><span data-stu-id="de648-243">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="de648-p109">如需排查清单问题，请参阅[验证并排查清单问题](../testing/troubleshoot-manifest.md)。其中介绍了如何针对 [XML 架构定义 (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) 验证清单，以及如何使用运行时日志记录功能调试清单。</span><span class="sxs-lookup"><span data-stu-id="de648-p109">For troubleshooting issues with your manifest, see [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md). There, you will find information on how to validate the manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), and also how to use runtime logging to debug the manifest.</span></span>

## <a name="see-also"></a><span data-ttu-id="de648-246">另请参阅</span><span class="sxs-lookup"><span data-stu-id="de648-246">See also</span></span>

* <span data-ttu-id="de648-247">[在清单中创建加载项命令][加载项命令]</span><span class="sxs-lookup"><span data-stu-id="de648-247">[Create add-in commands in your manifest][add-in commands]</span></span>
* [<span data-ttu-id="de648-248">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="de648-248">Specify Office hosts and API requirements</span></span>](specify-office-hosts-and-api-requirements.md)
* [<span data-ttu-id="de648-249">Office 外接程序的本地化</span><span class="sxs-lookup"><span data-stu-id="de648-249">Localization for Office Add-ins</span></span>](localization.md)
* [<span data-ttu-id="de648-250">Office 外接程序清单的架构参考</span><span class="sxs-lookup"><span data-stu-id="de648-250">Schema reference for Office Add-ins manifests</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)
* [<span data-ttu-id="de648-251">验证并排查清单问题</span><span class="sxs-lookup"><span data-stu-id="de648-251">Validate and troubleshoot issues with your manifest</span></span>](../testing/troubleshoot-manifest.md)

[加载项命令]: create-addin-commands.md
[add-in commands]: create-addin-commands.md