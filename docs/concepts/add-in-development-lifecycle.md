---
title: Office 加载项开发生命周期
description: ''
ms.date: 07/01/2019
localization_priority: Priority
ms.openlocfilehash: 44e2792f030662bd89b272998ad47fd0a645d785
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454571"
---
# <a name="office-add-ins-development-lifecycle"></a><span data-ttu-id="d5600-102">Office 加载项开发生命周期</span><span class="sxs-lookup"><span data-stu-id="d5600-102">Office Add-ins development lifecycle</span></span>

> [!NOTE]
> <span data-ttu-id="d5600-p101">如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性](../overview/office-add-in-availability.md)页面）。</span><span class="sxs-lookup"><span data-stu-id="d5600-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

<span data-ttu-id="d5600-105">Office 加载项的典型开发生命周期包括下列步骤：</span><span class="sxs-lookup"><span data-stu-id="d5600-105">The typical development lifecycle of an Office Add-in includes the following steps:</span></span>


## <a name="1-decide-on-the-purpose-of-the-add-in"></a><span data-ttu-id="d5600-106">1. 确定加载项的用途</span><span class="sxs-lookup"><span data-stu-id="d5600-106">1. Decide on the purpose of the add-in</span></span>

<span data-ttu-id="d5600-107">提出以下问题：</span><span class="sxs-lookup"><span data-stu-id="d5600-107">Ask the following questions:</span></span>

- <span data-ttu-id="d5600-108">加载项有何作用？</span><span class="sxs-lookup"><span data-stu-id="d5600-108">How is the add-in useful?</span></span>

- <span data-ttu-id="d5600-109">它如何帮助您的客户提高工作效率？</span><span class="sxs-lookup"><span data-stu-id="d5600-109">How does it help your customers be more productive?</span></span>

- <span data-ttu-id="d5600-110">您的加载项功能支持哪些方案？</span><span class="sxs-lookup"><span data-stu-id="d5600-110">What scenarios does your add-in's features support?</span></span>

<span data-ttu-id="d5600-111">确定最重要的功能和方案，并围绕它们进行集中设计。</span><span class="sxs-lookup"><span data-stu-id="d5600-111">Decide the most important features and scenarios and focus your design around them.</span></span>


## <a name="2-identify-the-data-and-data-source-for-the-add-in"></a><span data-ttu-id="d5600-112">2. 确定加载项的数据和数据源</span><span class="sxs-lookup"><span data-stu-id="d5600-112">2. Identify the data and data source for the add-in</span></span>

- <span data-ttu-id="d5600-113">数据是在文档、工作簿、演示文稿还是项目中？</span><span class="sxs-lookup"><span data-stu-id="d5600-113">Is the data in a document, workbook, presentation, project, or an Access browser-based database?</span></span>

- <span data-ttu-id="d5600-114">数据是否关于 Exchange Server 或 Exchange Online 邮箱中的一个或多个项？</span><span class="sxs-lookup"><span data-stu-id="d5600-114">Is the data about an item or items in an Exchange Server or Exchange Online mailbox?</span></span>

- <span data-ttu-id="d5600-115">数据是否来自外部源（如 Web 服务）？</span><span class="sxs-lookup"><span data-stu-id="d5600-115">Is the data from an external source such as a web service?</span></span>


## <a name="3-identify-the-type-of-add-in-and-office-host-applications-that-best-support-the-purpose-of-the-add-in"></a><span data-ttu-id="d5600-116">3. 确定加载项类型和最能支持其用途的 Office 主机应用</span><span class="sxs-lookup"><span data-stu-id="d5600-116">3. Identify the type of add-in and Office host applications that best support the purpose of the add-in</span></span>

<span data-ttu-id="d5600-117">为确定方案，请考虑以下几点：</span><span class="sxs-lookup"><span data-stu-id="d5600-117">Consider the following to identify the scenarios:</span></span>

- <span data-ttu-id="d5600-p102">客户是否要使用加载项来丰富文档的内容？如果是，建议考虑创建**内容加载项**。</span><span class="sxs-lookup"><span data-stu-id="d5600-p102">Will customers use the add-in to enrich the content of a document or Access browser-based database? If so, you may want to consider creating a **content add-in**.</span></span>

- <span data-ttu-id="d5600-p103">客户是否要在查看或撰写电子邮件或约会时使用该外接程序？能够根据当前上下文公开外接程序是否很重要？是否优先考虑使外接程序不仅在台式机上可用，而且在平板电脑或智能手机上也可用？</span><span class="sxs-lookup"><span data-stu-id="d5600-p103">Will customers use the add-in while viewing or composing an email message or appointment? Is being able to expose the add-in according to the current context important? Is making the add-in available on not just the desktop, but also on tablets and phones a priority?</span></span>

    <span data-ttu-id="d5600-p104">如果上述任一问题的答案是肯定的，请考虑创建 **Outlook 加载项**。然后，确定加载项的触发上下文（例如，撰写表单中的用户、特定消息类型、是否有附件、地址、任务建议、会议建议，或电子邮件或约会内容中的特定字符串模式）。</span><span class="sxs-lookup"><span data-stu-id="d5600-p104">If you answer yes to any of these questions, consider creating an **Outlook add-in**. Identify the context that will trigger your add-in (for example, the user being in a compose form, specific message types, the presence of an attachment, address, task suggestion, or meeting suggestion, or certain string patterns in the contents of an email or appointment).</span></span>

    <span data-ttu-id="d5600-125">若要了解如何根据上下文激活 Outlook 加载项，请参阅 [Outlook 加载项的激活规则](/outlook/add-ins/activation-rules)。</span><span class="sxs-lookup"><span data-stu-id="d5600-125">To find out how you can contextually activate the Outlook add-in, see [Activation rules for Outlook add-ins](/outlook/add-ins/activation-rules).</span></span>

- <span data-ttu-id="d5600-p105">客户是否要使用加载项来增强文档的查看或创作体验？如果是，建议考虑创建**任务窗格加载项**。</span><span class="sxs-lookup"><span data-stu-id="d5600-p105">Will customers use the add-in to enhance the viewing or authoring experience of a document? If so, you may want to consider creating a **task pane add-in**.</span></span>

<span data-ttu-id="d5600-128">（Windows、Mac、Web、Mobile）上运行的 Office 应用程序和平台之间的某些加载项 API 可能不同。</span><span class="sxs-lookup"><span data-stu-id="d5600-128">Support for certain add-in APIs may differ between Office applications and the platform they are running on (Windows, Mac, Web, Mobile).</span></span> <span data-ttu-id="d5600-129">若要查看客户端和平台的当前 API 覆盖范围，请参阅我们的 [Office 加载项主机和平台可用性](../overview/office-add-in-availability.md)页。</span><span class="sxs-lookup"><span data-stu-id="d5600-129">To see the current API coverage by client and platform, see our [Office Add-in host and platform availability](../overview/office-add-in-availability.md) page.</span></span>  


## <a name="4-design-and-implement-the-user-experience-and-user-interface-for-the-add-in"></a><span data-ttu-id="d5600-130">4. 为加载项设计和实施用户体验和用户界面</span><span class="sxs-lookup"><span data-stu-id="d5600-130">4. Design and implement the user experience and user interface for the add-in</span></span>

<span data-ttu-id="d5600-p107">设计快速、流畅的用户体验，不仅非常一致，还易于学习，主要方案只需几个步骤即可完成。根据加载项的用途，利用第三方 API 或 Web 服务。</span><span class="sxs-lookup"><span data-stu-id="d5600-p107">Design a fast and fluid user experience that is consistent, easy to learn, with primary scenarios that require only a few steps to complete. Depending on the purpose of the add-in, make use of third-party APIs or web services.</span></span>

<span data-ttu-id="d5600-133">可从各种 Web 开发工具中进行选择，并使用 HTML 和 JavaScript 实现用户界面。</span><span class="sxs-lookup"><span data-stu-id="d5600-133">You can choose from a variety of web development tools and use HTML and JavaScript to implement the user interface.</span></span>


## <a name="5-create-an-xml-manifest-file-based-on-the-office-add-ins-manifest-schema"></a><span data-ttu-id="d5600-134">5. 根据 Office 加载项清单架构创建 XML 清单文件</span><span class="sxs-lookup"><span data-stu-id="d5600-134">5. Create an XML manifest file based on the Office Add-ins manifest schema</span></span>

<span data-ttu-id="d5600-135">创建 XML 清单，以确定加载项及其要求，指定加载项使用的 HTML 以及任何 JavaScript 和 CSS 文件的位置，并根据加载项的类型指定默认大小和权限。</span><span class="sxs-lookup"><span data-stu-id="d5600-135">Create an XML manifest to identify the add-in and its requirements, specify the locations of the HTML and any JavaScript and CSS files that the add-in uses, and depending on the type of the add-in, the default size and permissions.</span></span>

<span data-ttu-id="d5600-p108">对于 Outlook 加载项，可以根据当前邮件或约会指定上下文，加载项在其中不仅相关，还可供 Outlook 在 UI 中使用。还可以确定希望加载项支持的设备。在清单中，将上下文指定为激活规则和受支持的设备。</span><span class="sxs-lookup"><span data-stu-id="d5600-p108">For Outlook add-ins, you can specify the context, based on the current message or appointment, under which your add-in is relevant and you would like Outlook to make available in the UI. You can also decide which devices you want the add-in to support. In the manifest, specify the context as activation rules and the supported devices.</span></span>


## <a name="6-install-and-test-the-add-in"></a><span data-ttu-id="d5600-139">6. 安装和测试加载项</span><span class="sxs-lookup"><span data-stu-id="d5600-139">6. Install and test the add-in</span></span>

<span data-ttu-id="d5600-p109">将 HTML 文件以及任何 JavaScript 和 CSS 文件放在外接程序清单文件中指定的 Web 服务器上。安装外接程序的过程取决于外接程序的类型。有关详细信息，请参阅[旁加载 Office 外接程序进行测试](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="d5600-p109">Place the HTML files and any JavaScript and CSS files on the web servers that are specified in the add-in manifest file. The process to install an add-in depends on the type of the add-in. For details, see [Sideload Office Add-ins for testing](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

<span data-ttu-id="d5600-p110">对于 Outlook 外接程序，将其安装在 Exchange 邮箱中，并指定 Exchange 管理中心 (EAC) 中外接程序清单文件的位置。有关详细信息，请参阅[部署和安装 Outlook 外接程序以供测试](/outlook/add-ins/testing-and-tips)。</span><span class="sxs-lookup"><span data-stu-id="d5600-p110">For Outlook add-ins, install it in an Exchange mailbox, and specify the location of the add-in manifest file in the Exchange Admin Center (EAC). For more information, see [Deploy and install Outlook add-ins for testing](/outlook/add-ins/testing-and-tips).</span></span>


## <a name="7-publish-the-add-in"></a><span data-ttu-id="d5600-145">7. 发布加载项</span><span class="sxs-lookup"><span data-stu-id="d5600-145">7. Publish the add-in</span></span>

<span data-ttu-id="d5600-p111">可以将加载项提交到 AppSource，客户从中能够安装加载项。此外，还可以向 SharePoint 上的应用目录或共享网络文件夹发布任务窗格和内容加载项，并在组织的 Exchange 服务器上直接部署 Outlook 加载项。有关详细信息，请参阅[发布 Office 加载项](../publish/publish.md)。</span><span class="sxs-lookup"><span data-stu-id="d5600-p111">You can submit the add-in to AppSource, from which customers can install the add-in. In addition, you can publish task pane and content add-ins to a private folder add-in catalog on SharePoint or to a shared network folder, and you can deploy an Outlook add-in directly on an Exchange server for your organization. For details, see [Publish your Office Add-in](../publish/publish.md).</span></span>


## <a name="8-maintain-the-add-in"></a><span data-ttu-id="d5600-149">8. 维护加载项</span><span class="sxs-lookup"><span data-stu-id="d5600-149">8. Maintain the add-in</span></span>

<span data-ttu-id="d5600-p112">如果外接程序调用 Web 服务，且在发布外接程序后对 Web 服务进行了更新，则无需重新发布外接程序。 但是，如果你对提交的外接程序的任何项目或数据进行了更改（如外接程序清单、屏幕截图、图标、HTML 或 JavaScript 文件），则需重新发布外接程序。</span><span class="sxs-lookup"><span data-stu-id="d5600-p112">If your add-in calls a web service, and if you make updates to the web service after publishing the add-in, you do not have to republish the add-in. However, if you change any items or data you submitted for your add-in, such as the add-in manifest, screenshots, icons, HTML or JavaScript files, you will need to republish the add-in.</span></span> 

<span data-ttu-id="d5600-p113">特别是，如果已将加载项发布到 AppSource，需要重新提交加载项，以便 AppSource 能够实现这些更改。必须重新提交加载项，并附带包含新版本号的更新后加载项清单。还必须确保将提交表单中的加载项版本号更新为，与新清单版本号一致。对于 Outlook 加载项，应确保 [Id](/office/dev/add-ins/reference/manifest/id) 元素包含加载项清单中的不同 UUID。</span><span class="sxs-lookup"><span data-stu-id="d5600-p113">In particular, if you have published the add-in to AppSource, you'll need to resubmit your add-in so that AppSource can implement those changes. You must resubmit your add-in with an updated add-in manifest that includes a new version number. You must also make sure to update the add-in version number in the submission form to match the new manifest's version number. For Outlook add-ins, you should make sure the [Id](/office/dev/add-ins/reference/manifest/id) element contains a different UUID in the add-in manifest.</span></span>
