---
title: Office 加载项的隐私和安全
description: ''
ms.date: 09/26/2019
localization_priority: Normal
ms.openlocfilehash: 5782cc7fcf23190cca73cc91a35a73e82d182261
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42323838"
---
# <a name="privacy-and-security-for-office-add-ins"></a><span data-ttu-id="658ce-102">Office 加载项的隐私和安全</span><span class="sxs-lookup"><span data-stu-id="658ce-102">Privacy and security for Office Add-ins</span></span>

## <a name="understanding-the-add-in-runtime"></a><span data-ttu-id="658ce-103">了解加载项运行时</span><span class="sxs-lookup"><span data-stu-id="658ce-103">Understanding the add-in runtime</span></span>

<span data-ttu-id="658ce-p101">Office 外接程序受到外接程序运行时环境、多层权限模型和性能调控器的保护。这一框架通过以下方式保护用户体验：</span><span class="sxs-lookup"><span data-stu-id="658ce-p101">Office Add-ins are secured by an add-in runtime environment, a multiple-tier permissions model, and performance governors. This framework protects the user's experience in the following ways:</span></span> 

- <span data-ttu-id="658ce-106">管理对主机应用程序的 UI 框架的访问。</span><span class="sxs-lookup"><span data-stu-id="658ce-106">Access to the host application's UI frame is managed.</span></span>

- <span data-ttu-id="658ce-107">只允许间接访问主机应用的 UI 线程。</span><span class="sxs-lookup"><span data-stu-id="658ce-107">Only indirect access to the host application's UI thread is allowed.</span></span>

- <span data-ttu-id="658ce-108">不允许模式交互（例如，不允许调用 JavaScript `alert`、 `confirm`和`prompt`函数，因为它们是模式的。</span><span class="sxs-lookup"><span data-stu-id="658ce-108">Modal interactions aren't allowed - for example, calls to JavaScript `alert`, `confirm`, and `prompt` functions aren't allowed because they're modal.</span></span>

<span data-ttu-id="658ce-109">此外，为了确保 Office 加载项不会损害用户环境，运行时框架还提供以下优势：</span><span class="sxs-lookup"><span data-stu-id="658ce-109">Further, the runtime framework provides the following benefits to ensure that an Office Add-in can't damage the user's environment:</span></span>

- <span data-ttu-id="658ce-110">隔离运行加载项的进程。</span><span class="sxs-lookup"><span data-stu-id="658ce-110">Isolates the process the add-in runs in.</span></span>

- <span data-ttu-id="658ce-111">不需要 .dll 或 .exe 替换项或 ActiveX 组件。</span><span class="sxs-lookup"><span data-stu-id="658ce-111">Doesn't require .dll or .exe replacement or ActiveX components.</span></span>

- <span data-ttu-id="658ce-112">可以轻松安装和卸载加载项。</span><span class="sxs-lookup"><span data-stu-id="658ce-112">Makes add-ins easy to install and uninstall.</span></span>

<span data-ttu-id="658ce-113">此外，还可以调控 Office 外接程序使用的内存、CPU 和网络资源，以确保维持良好的性能和可靠性。</span><span class="sxs-lookup"><span data-stu-id="658ce-113">Also, the use of memory, CPU, and network resources by Office Add-ins is governable to ensure that good performance and reliability are maintained.</span></span>

<span data-ttu-id="658ce-114">以下各节简要介绍运行时体系结构如何支持在基于 Windows 的设备上的 Office 客户端、OS X Mac 设备以及 Web 浏览器中运行加载项。</span><span class="sxs-lookup"><span data-stu-id="658ce-114">The following sections briefly describe how the runtime architecture supports running add-ins in Office clients on Windows-based devices, on OS X Mac devices, and in web browsers.</span></span>

### <a name="clients-on-windows-and-os-x-devices"></a><span data-ttu-id="658ce-115">Windows 和 OS X 设备上的客户端</span><span class="sxs-lookup"><span data-stu-id="658ce-115">Clients on Windows and OS X devices</span></span>

<span data-ttu-id="658ce-p102">在支持的台式机和平板电脑设备的客户端（如 Windows 版 Excel、Windows 版 Outlook 和 Mac 版 Outlook）中，通过集成进程内组件 Office 加载项运行时来支持 Office 加载项，该组件管理加载项的生命周期，并实现加载项和客户端应用程序之间的互操作性。加载项网页本身托管在进程外。如图 1 中所示，在 Windows 台式机或平板电脑设备上，[加载项网页托管在 Internet Explorer 或 Microsoft Edge 控件内部](browsers-used-by-office-web-add-ins.md)，而 Internet Explorer 控件托管在加载项运行时进程内部，提供安全和性能隔离。</span><span class="sxs-lookup"><span data-stu-id="658ce-p102">In supported clients for desktop and tablet devices, such as Excel on Windows, and Outlook on Windows and Mac, Office Add-ins are supported by integrating an in-process component, the Office Add-ins runtime, which manages the add-in lifecycle and enables interoperability between the add-in and the client application. The add-in webpage itself is hosted out-of-process. As shown in figure 1, on a Windows desktop or tablet device, [the add-in webpage is hosted inside an Internet Explorer or Microsoft Edge control](browsers-used-by-office-web-add-ins.md) which, in turn, is hosted inside an add-in runtime process that provides security and performance isolation.</span></span>

<span data-ttu-id="658ce-p103">在 Windows 桌面设备上，必须为受限网站区域启用 Internet Explorer 保护模式。通常情况下，此模式默认启用。如果禁用，则会在尝试启动加载项时[看到错误消息](https://support.microsoft.com/help/2761180/apps-for-office-don-t-start-if-you-disable-protected-mode-for-the-restricted-sites-zone-in-internet-explorer)。</span><span class="sxs-lookup"><span data-stu-id="658ce-p103">On Windows desktops, Protected Mode in Internet Explorer must be enabled for the Restricted Site Zone. This is typically enabled by default. If it is disabled, an [error will occur](https://support.microsoft.com/help/2761180/apps-for-office-don-t-start-if-you-disable-protected-mode-for-the-restricted-sites-zone-in-internet-explorer) when you try to launch an add-in.</span></span>

<span data-ttu-id="658ce-122">*图 1.基于 Windows 的台式机和平板电脑客户端中的 Office 外接程序运行时环境*</span><span class="sxs-lookup"><span data-stu-id="658ce-122">*Figure 1. Office Add-ins runtime environment in Windows-based desktop and tablet clients*</span></span>

![富客户端基础结构](../images/dk2-agave-overview-02.png)

<span data-ttu-id="658ce-124">如下图所示，在 OS X Mac 台式机上，加载项网页托管在沙盒 WebKit 运行时主机进程内部，这有助于提供类似级别的安全和性能保护。</span><span class="sxs-lookup"><span data-stu-id="658ce-124">As shown in the following figure, on an OS X Mac desktop, the add-in web page is hosted inside a sandboxed WebKit runtime host process which helps provide similar level of security and performance protection.</span></span> 

<span data-ttu-id="658ce-125">*图 2：OS X Mac 客户端中的 Office 加载项运行时环境*</span><span class="sxs-lookup"><span data-stu-id="658ce-125">*Figure 2. Office Add-ins runtime environment in OS X Mac clients*</span></span>

![OS X Mac 上的 Office 相关应用程序运行时环境](../images/dk2-agave-overview-mac-02.png)

<span data-ttu-id="658ce-127">Office 外接程序运行时管理进程间通信、JavaScript API 调用和事件到本机调用和事件的转换以及 UI 远程处理支持，从而使加载项能够呈现在文档内、任务窗格中或电子邮件、会议请求或约会旁边。</span><span class="sxs-lookup"><span data-stu-id="658ce-127">The Office Add-ins runtime manages interprocess communication, the translation of JavaScript API calls and events into native ones, as well as UI remoting support to enable the add-in to be rendered inside the document, in a task pane, or adjacent to an email message, meeting request, or appointment.</span></span>

### <a name="web-clients"></a><span data-ttu-id="658ce-128">Web 客户端</span><span class="sxs-lookup"><span data-stu-id="658ce-128">Web clients</span></span>

<span data-ttu-id="658ce-129">在受支持的 Web 客户端中，Office 外接程序承载在使用 HTML5**沙盒**属性运行的**iframe**中。</span><span class="sxs-lookup"><span data-stu-id="658ce-129">In supported Web clients, Office Add-ins are hosted in an **iframe** that runs using the HTML5 **sandbox** attribute.</span></span> <span data-ttu-id="658ce-130">不允许使用 ActiveX 组件或导航 Web 客户端主页。</span><span class="sxs-lookup"><span data-stu-id="658ce-130">ActiveX components or navigating the main page of the web client are not allowed.</span></span> <span data-ttu-id="658ce-131">通过集成适用于 Office 的 JavaScript API 在 Web 客户端中实现 Office 外接程序支持。</span><span class="sxs-lookup"><span data-stu-id="658ce-131">Office Add-ins support is enabled in the web clients by the integration of the JavaScript API for Office.</span></span> <span data-ttu-id="658ce-132">同理，对于桌面客户端应用程序，JavaScript API 管理加载项生命周期和加载项与 Web 客户端间的互操作性。</span><span class="sxs-lookup"><span data-stu-id="658ce-132">In a similar way to the desktop client applications, the JavaScript API manages the add-in lifecycle and interoperability between the add-in and the web client.</span></span> <span data-ttu-id="658ce-133">这种互操作性通过特殊的跨框架发布消息通信基础结构实现。</span><span class="sxs-lookup"><span data-stu-id="658ce-133">This interoperability is implemented by using a special cross-frame post message communication infrastructure.</span></span> <span data-ttu-id="658ce-134">桌面客户端上使用的同一 JavaScript 库 (Office.js) 可用来与 Web 客户端交互。</span><span class="sxs-lookup"><span data-stu-id="658ce-134">The same JavaScript library (Office.js) that is used on desktop clients is available to interact with the web client.</span></span> <span data-ttu-id="658ce-135">下图显示了支持在浏览器中运行的 Office 中的外接程序的基础结构，以及支持它们所需的相关组件（web 客户端、 **iframe**、Office 外接程序运行时和适用于 Office 的 JavaScript API）。</span><span class="sxs-lookup"><span data-stu-id="658ce-135">The following figure shows the infrastructure that supports add-ins in Office running in the browser, and the relevant components (the web client, **iframe**, Office Add-ins runtime, and JavaScript API for Office) that are required to support them.</span></span>


<span data-ttu-id="658ce-136">*图 3：支持 Office Web 客户端中 Office 加载项的基础结构*</span><span class="sxs-lookup"><span data-stu-id="658ce-136">*Figure 3. Infrastructure that supports Office Add-ins in Office web clients*</span></span>

![Web 客户端基础结构](../images/dk2-agave-overview-03.png)

## <a name="add-in-integrity-in-appsource"></a><span data-ttu-id="658ce-138">AppSource 中的加载项完整性</span><span class="sxs-lookup"><span data-stu-id="658ce-138">Add-in integrity in AppSource</span></span>

<span data-ttu-id="658ce-p105">若要向受众提供 Office 加载项，可以在 AppSource 中发布它们。AppSource 强制执行以下措施来维护加载项完整性：</span><span class="sxs-lookup"><span data-stu-id="658ce-p105">You can make your Office Add-ins available to the public by publishing them to AppSource. AppSource enforces the following measures to maintain the integrity of add-ins:</span></span>


- <span data-ttu-id="658ce-141">要求 Office 加载项的主机服务器始终使用安全套接字层 (SSL) 进行通信。</span><span class="sxs-lookup"><span data-stu-id="658ce-141">Requires the host server of an Office Add-in to always use Secure Sockets Layer (SSL) to communicate.</span></span>

- <span data-ttu-id="658ce-142">要求开发人员在提交加载项时提供身份证明、合约协议和适合的隐私策略。</span><span class="sxs-lookup"><span data-stu-id="658ce-142">Requires a developer to provide proof of identity, a contractual agreement, and a compliant privacy policy to submit add-ins.</span></span>

- <span data-ttu-id="658ce-143">确保加载项的源在只读模式下可访问。</span><span class="sxs-lookup"><span data-stu-id="658ce-143">Ensures that the source of add-ins is accessible in read-only mode.</span></span>

- <span data-ttu-id="658ce-144">支持针对可用加载项的用户审阅系统以推广自我管理的社区。</span><span class="sxs-lookup"><span data-stu-id="658ce-144">Supports a user-review system for available add-ins to promote a self-policing community.</span></span>

## <a name="addressing-end-users-privacy-concerns"></a><span data-ttu-id="658ce-145">解决最终用户的隐私问题</span><span class="sxs-lookup"><span data-stu-id="658ce-145">Addressing end users' privacy concerns</span></span>

<span data-ttu-id="658ce-146">此部分从客户（最终用户）的角度出发介绍了 Office 外接程序平台提供的保护，并介绍了有关如何达到用户的预期以及如何安全处理用户个人身份信息 (PII) 的指南。</span><span class="sxs-lookup"><span data-stu-id="658ce-146">This section describes the protection offered by the Office Add-ins platform from the customer's (end user's) perspective, and provides guidelines for how to support users' expectations and how to securely handle users' personally identifiable information (PII).</span></span>

### <a name="end-users-perspective"></a><span data-ttu-id="658ce-147">从最终用户的角度出发</span><span class="sxs-lookup"><span data-stu-id="658ce-147">End users' perspective</span></span>

<span data-ttu-id="658ce-p106">Office 加载项是使用浏览器控件或 **iframe** 中运行的 Web 技术而生成。因此，使用加载项与转到 Internet 或 Intranet 上的网站类似。加载项可以位于组织外部（如果从 AppSource 获取加载项的话），也可以位于组织内部（如果从 Exchange Server 加载项目录、SharePoint 应用目录或组织网络上的文件共享获取加载项的话）。加载项对网络的访问权限受限，大部分加载项都可以对活动文档或邮件项执行读取或写入操作。在用户或管理员安装或启动加载项前，加载项平台就施加了特定约束。不过，与任何扩展性模型一样，用户在启动未知加载项之前应非常谨慎。</span><span class="sxs-lookup"><span data-stu-id="658ce-p106">Office Add-ins are built using web technologies that run in a browser control or **iframe**. Because of this, using add-ins is similar to browsing to web sites on the Internet or intranet. Add-ins can be external to an organization (if you acquire the add-in from AppSource) or internal (if you acquire the add-in from an Exchange Server add-in catalog, SharePoint app catalog, or file share on an organization's network). Add-ins have limited access to the network and most add-ins can read or write to the active document or mail item. The add-in platform applies certain constraints before a user or administrator installs or starts an add-in. But as with any extensibility model, users should be cautious before starting an unknown add-in.</span></span>

<span data-ttu-id="658ce-154">加载项平台解决了最终用户的隐私问题，具体方式如下：</span><span class="sxs-lookup"><span data-stu-id="658ce-154">The add-in platform addresses end users' privacy concerns in the following ways:</span></span>

- <span data-ttu-id="658ce-155">与托管内容、Outlook 或任务窗格外接程序的 Web 服务器通信的数据以及外接程序与其使用的任何 Web 服务之间的通信必须使用安全套接字层 (SSL) 协议加密。</span><span class="sxs-lookup"><span data-stu-id="658ce-155">Data communicated with the web server that hosts a content, Outlook or task pane add-in as well as communication between the add-in and any web services it uses must be encrypted using the Secure Socket Layer (SSL) protocol.</span></span>

- <span data-ttu-id="658ce-p107">安装 AppSource 中的加载项前，用户可以查看相应加载项的隐私策略和要求。此外，与用户邮箱进行交互的 Outlook 加载项还指明了所需的特定权限；用户可以在安装 Outlook 加载项前，先查看使用条款、请求的权限和隐私策略。</span><span class="sxs-lookup"><span data-stu-id="658ce-p107">Before a user installs an add-in from AppSource, the user can view the privacy policy and requirements of that add-in. In addition, Outlook add-ins that interact with users' mailboxes surface the specific permissions that they require; the user can review the terms of use, requested permissions and privacy policy before installing an Outlook add-in.</span></span>

- <span data-ttu-id="658ce-p108">在共享一个文档时，用户也会共享已插入该文档或与该文档关联的加载项。如果用户打开一个包含其之前未使用的加载项的文档，则主机应用程序会提示用户向加载项授予在文档中运行的权限。在组织环境中，如果文档来自外部源，则 Office 主机应用程序也会提示用户。</span><span class="sxs-lookup"><span data-stu-id="658ce-p108">When sharing a document, users also share add-ins that have been inserted in or associated with that document. If a user opens a document that contains an add-in that the user hasn't used before, the host application prompts the user to grant permission for the add-in to run in the document. In an organizational environment, the Office host application also prompts the user if the document comes from an external source.</span></span>

- <span data-ttu-id="658ce-161">用户可启用或禁用对 AppSource 的访问。</span><span class="sxs-lookup"><span data-stu-id="658ce-161">Users can enable or disable the access to AppSource.</span></span> <span data-ttu-id="658ce-162">对于内容和任务窗格外接程序，用户可以在主机 Office 客户端（从**文件** > **选项** > "**信任中心** > **信任中心" 设置** > **受信任的外接程序目录**中打开）管理对受信任的加载项和**目录的访问**。</span><span class="sxs-lookup"><span data-stu-id="658ce-162">For content and task pane add-ins, users manage access to trusted add-ins and catalogs from the **Trust Center** on the host Office client (opened from **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**).</span></span> <span data-ttu-id="658ce-163">对于 Outlook 外接程序，使用可以通过选择 "**管理外接程序**" 按钮来管理外接程序：在 Windows 中的 outlook 中，选择 "**文件** > **管理外接程序**"。在 Mac 上的 Outlook 中，选择外接程序栏上的 "**管理外接程序**" 按钮。</span><span class="sxs-lookup"><span data-stu-id="658ce-163">For Outlook add-ins, uses can manage add-ins by choosing the **Manage Add-ins** button: in Outlook on Windows, choose **File** > **Manage Add-ins**. In Outlook on Mac, choose the **Manage Add-ins** button on the add-in bar.</span></span> <span data-ttu-id="658ce-164">在 Outlook 网页版中，依次选择“**设置**”菜单（齿轮图标）>“**管理加载项**”。管理员还可以[通过使用组策略](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office)来管理此访问。</span><span class="sxs-lookup"><span data-stu-id="658ce-164">In Outlook on the web, choose the **Settings** menu (gear icon) > **Manage add-ins**. Administrators can also manage this access [by using group policy](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span></span>

- <span data-ttu-id="658ce-165">加载项平台的设计在以下方面为最终用户提供了安全和性能保障：</span><span class="sxs-lookup"><span data-stu-id="658ce-165">The design of the add-in platform provides security and performance for end users in the following ways:</span></span>

  - <span data-ttu-id="658ce-p110">Office 外接程序在托管在独立于 Office 主机应用程序的加载项运行时环境中的 Web 浏览器控件中运行。此设计提供了与主机应用程序的安全和性能隔离。</span><span class="sxs-lookup"><span data-stu-id="658ce-p110">An Office Add-in runs in a web browser control that is hosted in an add-in runtime environment separate from the Office host application. This design provides both security and performance isolation from the host application.</span></span>

  - <span data-ttu-id="658ce-168">在 Web 浏览器控件中运行可允许加载项完成在浏览器中运行的常规网页可执行的所有操作，但同时将限制加载项遵守针对域隔离和安全区域的同源策略。</span><span class="sxs-lookup"><span data-stu-id="658ce-168">Running in a web browser control allows the add-in to do almost anything a regular web page running in a browser can do but, at the same time, restricts the add-in to observe the same-origin policy for domain isolation and security zones.</span></span>

<span data-ttu-id="658ce-p111">Outlook 外接程序通过特定的资源使用率监视提供额外安全和性能功能。有关详细信息，请参阅 [Outlook 外接程序的隐私、权限和安全性](../outlook/privacy-and-security.md)。</span><span class="sxs-lookup"><span data-stu-id="658ce-p111">Outlook add-ins provide additional security and performance features through Outlook add-in specific resource usage monitoring. For more information, see [Privacy, permissions, and security for Outlook add-ins](../outlook/privacy-and-security.md).</span></span>

### <a name="developer-guidelines-to-handle-pii"></a><span data-ttu-id="658ce-171">开发人员处理 PII 的准则</span><span class="sxs-lookup"><span data-stu-id="658ce-171">Developer guidelines to handle PII</span></span>

<span data-ttu-id="658ce-172">下面列出了一些特定于 Office 加载项开发人员的 PII 保护准则：</span><span class="sxs-lookup"><span data-stu-id="658ce-172">The following lists some specific PII protection guidelines for you as a developer of Office Add-ins:</span></span>

- <span data-ttu-id="658ce-p112">[Settings](/javascript/api/office/office.settings) 对象旨在保存内容加载项或任务窗格加载项的会话之间的加载项设置和状态数据，但不会在 **Settings** 对象中存储密码和其他敏感 PII。最终用户无法查看 **Settings** 对象中的数据，但该数据存储为文档的易于访问的文件格式的一部分。你应该限制加载项对 PII 的使用，并将加载项所需的任何 PII 存储在将加载项作为用户保护的资源托管的服务器上。</span><span class="sxs-lookup"><span data-stu-id="658ce-p112">The [Settings](/javascript/api/office/office.settings) object is intended for persisting add-in settings and state data across sessions for a content or task pane add-in, but don't store passwords and other sensitive PII in the **Settings** object. The data in the **Settings** object isn't visible to end users, but it is stored as part of the document's file format which is readily accessible. You should limit your add-in's use of PII and store any PII required by your add-in on the server hosting your add-in as a user-secured resource.</span></span>

- <span data-ttu-id="658ce-p113">使用某些应用程序可能会泄露 PII。请确保安全地存储用户的身份、位置、访问时间和任何其他凭据数据，以便该加载项的其他用户无法访问该数据。</span><span class="sxs-lookup"><span data-stu-id="658ce-p113">Using some applications can reveal PII. Make sure that you securely store data for your users' identity, location, access times, and any other credentials so that data won't become available to other users of the add-in.</span></span>

- <span data-ttu-id="658ce-p114">如果加载项已在 AppSource 中发布，HTTPS 的 AppSource 要求会保护 Web 服务器与客户端计算机或设备之间传输的 PII。不过，如果将此类数据重新传输到其他服务器，请务必遵守相同级别的保护要求。</span><span class="sxs-lookup"><span data-stu-id="658ce-p114">If your add-in is available in AppSource, the AppSource requirement for HTTPS protects PII transmitted between your web server and the client computer or device. However, if you re-transmit that data to other servers, make sure you observe the same level of protection.</span></span>

- <span data-ttu-id="658ce-p115">如果存储用户的 PII，请务必向用户提示这一点，并向用户提供检查和删除此类信息的方法。如果将加载项提交到 AppSource，可以在隐私声明中概述所收集的数据及其用途。</span><span class="sxs-lookup"><span data-stu-id="658ce-p115">If you store users' PII, make sure you reveal that fact, and provide a way for users to inspect and delete it. If you submit your add-in to AppSource, you can outline the data you collect and how it's used in the privacy statement.</span></span>

## <a name="developers-permission-choices-and-security-practices"></a><span data-ttu-id="658ce-182">开发人员的权限选择和安全做法</span><span class="sxs-lookup"><span data-stu-id="658ce-182">Developers' permission choices and security practices</span></span>

<span data-ttu-id="658ce-183">遵循这些常规指南以支持 Office 外接程序的安全模型，并进一步了解有关每种加载项类型的更多详细信息。</span><span class="sxs-lookup"><span data-stu-id="658ce-183">Follow these general guidelines to support the security model of Office Add-ins, and drill down on more details for each add-in type.</span></span>

### <a name="permissions-choices"></a><span data-ttu-id="658ce-184">权限选择</span><span class="sxs-lookup"><span data-stu-id="658ce-184">Permissions choices</span></span>

<span data-ttu-id="658ce-185">加载项平台中提供了一个权限模型，供加载项用于声明实现其功能所需的对用数据的访问级别。</span><span class="sxs-lookup"><span data-stu-id="658ce-185">The add-in platform provides a permissions model that your add-in uses to declare the level of access to a user's data that it requires for its features.</span></span> <span data-ttu-id="658ce-186">每个权限级别对应适用于 Office 的 JavaScript API 的子集，加载项通过这些权限级别实现其功能。</span><span class="sxs-lookup"><span data-stu-id="658ce-186">Each permission level corresponds to the subset of the JavaScript API for Office your add-in is allowed to use for its features.</span></span> <span data-ttu-id="658ce-187">例如，内容和任务窗格外接程序的**WriteDocument**权限允许访问[document.setselecteddataasync](/javascript/api/office/office.document)方法，该方法允许外接程序向用户的文档中写入数据，但不允许访问从文档读取数据的任何方法。</span><span class="sxs-lookup"><span data-stu-id="658ce-187">For example, the **WriteDocument** permission for content and task pane add-ins allows access to the [Document.setSelectedDataAsync](/javascript/api/office/office.document) method that lets an add-in write to the user's document, but doesn't allow access to any of the methods for reading data from the document.</span></span> <span data-ttu-id="658ce-188">此权限级别对于只需要对文档执行写入操作的加载项很有用，例如用户可以查询要插入到其文档的数据的加载项。</span><span class="sxs-lookup"><span data-stu-id="658ce-188">This permission level makes sense for add-ins that only need to write to a document, such as an add-in where the user can query for data to insert into their document.</span></span>

<span data-ttu-id="658ce-p117">最佳做法是应该基于“_最小特权_”原则请求权限。即应该请求外接程序正常运行所需的 API 的最小子集的访问权限。例如，如果外接程序只需要读取其功能的用户文档中的数据，则应仅请求“**ReadDocument**”权限。（但是，请注意如果请求权限不足，则会导致外接程序平台阻止外接程序使用部分 API 并将生成运行时错误。）</span><span class="sxs-lookup"><span data-stu-id="658ce-p117">As a best practice, you should request permissions based on the principle of  _least privilege_. That is, you should request permission to access only the minimum subset of the API that your add-in requires to function correctly. For example, if your add-in needs only to read data in a user's document for its features, you should request no more than the **ReadDocument** permission. (But, keep in mind that requesting insufficient permissions will result in the add-in platform blocking your add-in's use of some APIs and will generate errors at run time.)</span></span>

<span data-ttu-id="658ce-193">您可以在加载项清单中指定权限（如下面的示例中所示），最终用户可以在决定在首次安装或激活外接程序之前查看该外接程序的请求权限级别。</span><span class="sxs-lookup"><span data-stu-id="658ce-193">You specify permissions in the manifest of your add-in, as shown in the example in this section below, and end users can see the requested permission level of an add-in before they decide to install or activate the add-in for the first time.</span></span> <span data-ttu-id="658ce-194">此外，请求**ReadWriteMailbox**权限的 Outlook 外接程序需要显式管理员权限才能安装。</span><span class="sxs-lookup"><span data-stu-id="658ce-194">Additionally, Outlook add-ins that request the **ReadWriteMailbox** permission require explicit administrator privilege to install.</span></span>

<span data-ttu-id="658ce-195">下面的示例演示任务窗格加载项如何在其清单中指定**ReadDocument**权限。</span><span class="sxs-lookup"><span data-stu-id="658ce-195">The following example shows how a task pane add-in specifies the **ReadDocument** permission in its manifest.</span></span> <span data-ttu-id="658ce-196">为重点关注权限，清单中的其他元素将不显示。</span><span class="sxs-lookup"><span data-stu-id="658ce-196">To keep permissions as the focus, other elements in the manifest aren't displayed.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
           xmlns:ver="http://schemas.microsoft.com/office/appforoffice/1.0"
           xsi:type="TaskPaneApp">

... <!-- To keep permissions as the focus, not displaying other elements. -->
  <Permissions>ReadDocument</Permissions>
...
</OfficeApp>
```

<span data-ttu-id="658ce-197">有关任务窗格和内容加载项的权限的详细信息，请参阅[在加载项中请求获取 API 使用权限](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="658ce-197">For more information about permissions for task pane and content add-ins, see [Requesting permissions for API use in add-ins](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).</span></span>

<span data-ttu-id="658ce-198">若要详细了解 Outlook 加载项权限，请参阅以下主题：</span><span class="sxs-lookup"><span data-stu-id="658ce-198">For more information about permissions for Outlook add-ins, see the following topics:</span></span>

- [<span data-ttu-id="658ce-199">Outlook 加载项的隐私、权限和安全</span><span class="sxs-lookup"><span data-stu-id="658ce-199">Privacy, permissions, and security for Outlook add-ins</span></span>](../outlook/privacy-and-security.md)

- [<span data-ttu-id="658ce-200">了解 Outlook 外接程序权限</span><span class="sxs-lookup"><span data-stu-id="658ce-200">Understanding Outlook add-in permissions</span></span>](../outlook/understanding-outlook-add-in-permissions.md)

### <a name="same-origin-policy"></a><span data-ttu-id="658ce-201">同源策略</span><span class="sxs-lookup"><span data-stu-id="658ce-201">Same origin policy</span></span>

<span data-ttu-id="658ce-202">由于 Office 外接程序是在 Web 浏览器控件中运行的网页，因此，它们必须遵守浏览器强制实施的同源策略：默认情况下，一个域中的网页无法执行它的域之外的其他域进行 [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) Web 服务调用。</span><span class="sxs-lookup"><span data-stu-id="658ce-202">Because Office Add-ins are webpages that run in a web browser control, they must follow the same-origin policy enforced by the browser: by default, a webpage in one domain can't make [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) web service calls to another domain other than the one where it is hosted.</span></span>

<span data-ttu-id="658ce-203">解决此限制的一种方法是使用 JSON/P--通过包含包含**src**属性的**脚本**标记（指向其他域中承载的某个脚本）来提供 web 服务的代理。</span><span class="sxs-lookup"><span data-stu-id="658ce-203">One way to overcome this limitation is to use JSON/P -- provide a proxy for the web service by including a **script** tag with a **src** attribute that points to some script hosted on another domain.</span></span> <span data-ttu-id="658ce-204">你可以编程方式创建 **script** 标记，动态创建 **src** 属性所指向的 URL，并通过 URI 查询参数将参数传递到 URL。</span><span class="sxs-lookup"><span data-stu-id="658ce-204">You can programmatically create the **script** tags, dynamically creating the URL to which to point the **src** attribute, and passing parameters to the URL via URI query parameters.</span></span> <span data-ttu-id="658ce-205">Web 服务提供程序在特定的 URL 位置创建和托管 JavaScript 代码，并根据 URI 查询参数返回不同的脚本。</span><span class="sxs-lookup"><span data-stu-id="658ce-205">Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters.</span></span> <span data-ttu-id="658ce-206">这些脚本然后在插入位置执行并按照预期的方式工作。</span><span class="sxs-lookup"><span data-stu-id="658ce-206">These scripts then execute where they are inserted and work as expected.</span></span>

<span data-ttu-id="658ce-207">以下是 Outlook 外接程序示例中的 JSON/P 的示例。</span><span class="sxs-lookup"><span data-stu-id="658ce-207">The following is an example of JSON/P in the Outlook add-in example.</span></span> 

```js
// Dynamically create an HTML SCRIPT element that obtains the details for the specified video.
function loadVideoDetails(videoIndex) {
    // Dynamically create a new HTML SCRIPT element in the webpage.
    var script = document.createElement("script");
    // Specify the URL to retrieve the indicated video from a feed of a current list of videos,
    // as the value of the src attribute of the SCRIPT element. 
    script.setAttribute("src", "https://gdata.youtube.com/feeds/api/videos/" + 
        videos[videoIndex].Id + "?alt=json-in-script&amp;callback=videoDetailsLoaded");
    // Insert the SCRIPT element at the end of the HEAD section.
    document.getElementsByTagName('head')[0].appendChild(script);
}
```

<span data-ttu-id="658ce-p121">Exchange 和 SharePoint 提供了客户端代理以实现跨域访问。通常，Intranet 上的同源策略没有 Internet 上的同源策略那样严格。有关详细信息，请参阅[同源策略第 1 部分：不准偷看](/archive/blogs/ieinternals/same-origin-policy-part-1-no-peeking)和[解决 Office 加载项中的同源策略限制](../develop/addressing-same-origin-policy-limitations.md)。</span><span class="sxs-lookup"><span data-stu-id="658ce-p121">Exchange and SharePoint provide client-side proxies to enable cross-domain access. In general, same origin policy on an intranet isn't as strict as on the Internet. For more information, see [Same Origin Policy Part 1: No Peeking](/archive/blogs/ieinternals/same-origin-policy-part-1-no-peeking) and [Addressing same-origin policy limitations in Office Add-ins](../develop/addressing-same-origin-policy-limitations.md).</span></span>

### <a name="tips-to-prevent-malicious-cross-site-scripting"></a><span data-ttu-id="658ce-211">防止恶意跨站点脚本的提示</span><span class="sxs-lookup"><span data-stu-id="658ce-211">Tips to prevent malicious cross-site scripting</span></span>

<span data-ttu-id="658ce-212">恶意用户可能会通过文档或加载项中的字段输入恶意脚本，以此来攻击加载项源。</span><span class="sxs-lookup"><span data-stu-id="658ce-212">An ill-intentioned user could attack the origin of an add-in by entering malicious script through the document or fields in the add-in.</span></span> <span data-ttu-id="658ce-213">开发人员应处理用户输入以避免在其域中执行恶意用户的 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="658ce-213">A developer should process user input to avoid executing a malicious user's JavaScript within their domain.</span></span> <span data-ttu-id="658ce-214">以下是从文档或邮件中或者通过加载项中的字段处理用户输入可遵循的一些良好做法：</span><span class="sxs-lookup"><span data-stu-id="658ce-214">The following are some good practices to follow to handle user input from a document or mail message, or via fields in an add-in:</span></span>


- <span data-ttu-id="658ce-p123">根据需要使用 [innerText](https://developer.mozilla.org/docs/Web/API/Element/innerHTML) 和 [textContent](https://developer.mozilla.org/docs/DOM/Node.textContent) 属性，而非 DOM 属性 [innerHTML](https://developer.mozilla.org/docs/Web/API/Node/innerText)。执行以下操作获取 Internet Explorer 和 Firefox 跨浏览器支持：</span><span class="sxs-lookup"><span data-stu-id="658ce-p123">Instead of the DOM property [innerHTML](https://developer.mozilla.org/docs/Web/API/Element/innerHTML), use the [innerText](https://developer.mozilla.org/docs/Web/API/Node/innerText) and [textContent](https://developer.mozilla.org/docs/DOM/Node.textContent) properties where appropriate. Do the following for Internet Explorer and Firefox cross-browser support:</span></span>

    ```js
     var text = x.innerText || x.textContent
    ```

    <span data-ttu-id="658ce-217">有关**innerText**和**textContent**之间的差异的信息，请参阅[textContent](https://developer.mozilla.org/docs/DOM/Node.textContent)。</span><span class="sxs-lookup"><span data-stu-id="658ce-217">For information about the differences between **innerText** and **textContent**, see [Node.textContent](https://developer.mozilla.org/docs/DOM/Node.textContent).</span></span> <span data-ttu-id="658ce-218">有关常见浏览器间 DOM 兼容性的详细信息，请参阅 [W3C DOM 兼容性 - HTML](https://www.quirksmode.org/dom/w3c_html.html#t07)。</span><span class="sxs-lookup"><span data-stu-id="658ce-218">For more information about DOM compatibility across common browsers, see [W3C DOM Compatibility - HTML](https://www.quirksmode.org/dom/w3c_html.html#t07).</span></span>

- <span data-ttu-id="658ce-219">如果必须使用**innerHTML**，则在将用户的输入传递到**innerHTML**之前，请确保该用户的输入不包含恶意内容。</span><span class="sxs-lookup"><span data-stu-id="658ce-219">If you must use **innerHTML**, make sure the user's input doesn't contain malicious content before passing it to **innerHTML**.</span></span> <span data-ttu-id="658ce-220">有关如何安全使用**innerHTML**的详细信息和示例，请参阅[innerHTML](https://developer.mozilla.org/docs/Web/API/Element/innerHTML)属性。</span><span class="sxs-lookup"><span data-stu-id="658ce-220">For more information and an example of how to use **innerHTML** safely, see [innerHTML](https://developer.mozilla.org/docs/Web/API/Element/innerHTML) property.</span></span>

- <span data-ttu-id="658ce-221">如果要使用 jQuery，请使用 [.text()](https://api.jquery.com/text/) 方法，而非 [.html()](https://api.jquery.com/html/) 方法。</span><span class="sxs-lookup"><span data-stu-id="658ce-221">If you are using jQuery, use the [.text()](https://api.jquery.com/text/) method instead of the [.html()](https://api.jquery.com/html/) method.</span></span>

- <span data-ttu-id="658ce-222">使用 [toStaticHTML](https://developer.mozilla.org/en-US/docs/Web/HTML/Reference) 方法可在将用户输入传递到 **innerHTML** 之前删除用户输入中的所有动态 HTML 元素和属性。</span><span class="sxs-lookup"><span data-stu-id="658ce-222">Use the [toStaticHTML](https://developer.mozilla.org/en-US/docs/Web/HTML/Reference) method to remove any dynamic HTML elements and attributes in users' input before passing it to **innerHTML**.</span></span>

- <span data-ttu-id="658ce-223">使用 [encodeURIComponent](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/encodeuricomponent) 或 [encodeURI](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/encodeuri) 函数可对应为来自用户输入或包含用户输入的 URL 的文本进行编码。</span><span class="sxs-lookup"><span data-stu-id="658ce-223">Use the [encodeURIComponent](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/encodeuricomponent) or [encodeURI](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/encodeuri) function to encode text that is intended to be a URL that comes from or contains user input.</span></span>

- <span data-ttu-id="658ce-224">有关创建更安全的 Web 解决方案的更多最佳做法，请参阅[开发安全加载项](/previous-versions/windows/apps/hh849625(v=win.10))。</span><span class="sxs-lookup"><span data-stu-id="658ce-224">See [Developing secure add-ins](/previous-versions/windows/apps/hh849625(v=win.10)) for more best practices to create more secure web solutions.</span></span>

### <a name="tips-to-prevent-clickjacking"></a><span data-ttu-id="658ce-225">防止“点击劫持”的提示</span><span class="sxs-lookup"><span data-stu-id="658ce-225">Tips to prevent "Clickjacking"</span></span>

<span data-ttu-id="658ce-226">由于 Office 加载项通过 Office 主机应用程序运行在浏览器中时呈现在 iframe 中，请使用以下提示来尽量降低[点击劫持](https://en.wikipedia.org/wiki/Clickjacking)（一种黑客用来欺骗用户泄露机密信息的技术）的风险。</span><span class="sxs-lookup"><span data-stu-id="658ce-226">Because Office Add-ins are rendered in an iframe when running in a browser with Office host applications, use the following tips to minimize the risk of [clickjacking](https://en.wikipedia.org/wiki/Clickjacking) -- a technique used by hackers to fool users into revealing confidential information.</span></span>

<span data-ttu-id="658ce-p126">首先，确定您的加载项可以执行的敏感操作。其中包括未授权的用户可能恶意使用的任何操作，如启动金融交易或发布敏感数据。例如，您的加载项可能让用户将款项发送到用户定义的接收人。</span><span class="sxs-lookup"><span data-stu-id="658ce-p126">First, identify sensitive actions that your add-in can perform. These include any actions that an unauthorized user could use with malicious intent, such as initiating a financial transaction or publishing sensitive data. For example, your add-in might let the user send a payment to a user-defined recipient.</span></span>

<span data-ttu-id="658ce-p127">其次，对于敏感操作，你的加载项应在执行操作之前向用户确认。该确认应详细说明该操作将产生的影响。此外，如有必要，还应详细说明用户如何能够防止该操作，是通过选择标记为“不允许”的特定按钮，还是忽略确认。</span><span class="sxs-lookup"><span data-stu-id="658ce-p127">Second, for sensitive actions, your add-in should confirm with the user before it executes the action. This confirmation should detail what effect the action will have. It should also detail how the user can prevent the action, if necessary, whether by choosing a specific button marked "Don't Allow" or by ignoring the confirmation.</span></span>

<span data-ttu-id="658ce-233">第三，为了确保没有任何潜在的攻击者可以隐藏或掩盖确认，您应将其显示在加载项上下文以外（即，不在 HTML 对话框中）。</span><span class="sxs-lookup"><span data-stu-id="658ce-233">Third, to ensure that no potential attacker can hide or mask the confirmation, you should display it outside the context of the add-in (that is, not in an HTML dialog box).</span></span>

<span data-ttu-id="658ce-234">下面是如何获取确认的一些示例：</span><span class="sxs-lookup"><span data-stu-id="658ce-234">Here are some examples of how you could get confirmation:</span></span>

- <span data-ttu-id="658ce-235">向用户发送包含确认链接的电子邮件。</span><span class="sxs-lookup"><span data-stu-id="658ce-235">Send an email to the user that contains a confirmation link.</span></span>

- <span data-ttu-id="658ce-236">向用户发送短信，其中包含用户可在外接程序中输入的确认码。</span><span class="sxs-lookup"><span data-stu-id="658ce-236">Send a text message to the user that includes a confirmation code that the user can enter in the add-in.</span></span>

- <span data-ttu-id="658ce-p128">对于无法应用 iframe 的页面，在新浏览器窗口中打开确认对话框。这通常是登录页采用的模式。使用[对话框 API](../develop/dialog-api-in-office-add-ins.md) 新建对话框。</span><span class="sxs-lookup"><span data-stu-id="658ce-p128">Open a confirmation dialog in a new browser window to a page that cannot be iframed. This is typically the pattern that is used by login pages. Use the [dialog api](../develop/dialog-api-in-office-add-ins.md) to create a new dialog.</span></span>

<span data-ttu-id="658ce-p129">此外，请确保您用于与用户联系的地址不能由潜在的攻击者提供。例如，对于付款确认，使用经授权用户帐户的文件中的地址。</span><span class="sxs-lookup"><span data-stu-id="658ce-p129">Also, ensure that the address you use for contacting the user couldn't have been provided by a potential attacker. For example, for payment confirmations use the address on file for the authorized user's account.</span></span>

### <a name="other-security-practices"></a><span data-ttu-id="658ce-242">其他安全实践</span><span class="sxs-lookup"><span data-stu-id="658ce-242">Other security practices</span></span>

<span data-ttu-id="658ce-243">开发人员还应记下以下安全实践：</span><span class="sxs-lookup"><span data-stu-id="658ce-243">Developers should also take note of the following security practices:</span></span>


- <span data-ttu-id="658ce-244">开发人员不得在 Office 加载项中使用 ActiveX 控件，因为 ActiveX 控件不支持加载项平台的跨平台特性。</span><span class="sxs-lookup"><span data-stu-id="658ce-244">Developers shouldn't use ActiveX controls in Office Add-ins as ActiveX controls don't support the cross-platform nature of the add-in platform.</span></span>

- <span data-ttu-id="658ce-p130">内容和任务窗格加载项使用浏览器默认使用的相同 SSL 设置，并允许大部分内容仅通过 SSL 传送。Outlook 加载项要求所有内容都通过 SSL 传送。开发人员必须在加载项清单的 **SourceLocation** 元素中指定使用 HTTPS 的 URL，以标识加载项的 HTML 文件位置。</span><span class="sxs-lookup"><span data-stu-id="658ce-p130">Content and task pane add-ins assume the same SSL settings that the browser uses by default, and allows most content to be delivered only by SSL. Outlook add-ins require all content to be delivered by SSL. Developers must specify in the **SourceLocation** element of the add-in manifest a URL that uses HTTPS, to identify the location of the HTML file for the add-in.</span></span>

    <span data-ttu-id="658ce-248">若要确保加载项不使用 HTTP 交付内容，在测试加载项时，开发人员应确保在“**控制窗格**”中的“**Internet 选项**”中选择以下设置且其测试方案中不显示任何安全警告：</span><span class="sxs-lookup"><span data-stu-id="658ce-248">To make sure add-ins aren't delivering content by using HTTP, when testing add-ins, developers should make sure the following settings are selected in **Internet Options** in **Control Panel** and no security warnings appear in their test scenarios:</span></span>

    - <span data-ttu-id="658ce-249">确保针对“Internet”\*\*\*\* 区域的安全设置“显示混合内容”\*\*\*\* 设置为“提示”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="658ce-249">Make sure the security setting, **Display mixed content**, for the **Internet** zone is set to **Prompt**.</span></span> <span data-ttu-id="658ce-250">若要执行此操作，请在 " **Internet 选项**" 中选择以下选项：在 "**安全**" 选项卡上，选择 " **internet**区域"，选择 "**自定义级别**"，滚动查找 "**显示混合内容**"，然后选择 "**提示**" （如果尚未选中）。</span><span class="sxs-lookup"><span data-stu-id="658ce-250">You can do that by selecting the following in **Internet Options**: on the **Security** tab, select the **Internet** zone, select **Custom level**, scroll to look for **Display mixed content**, and select **Prompt** if it isn't already selected.</span></span>

    - <span data-ttu-id="658ce-251">确保在“Internet 选项”\*\*\*\* 对话框的“高级”\*\*\*\* 选项卡中，选中了“在安全和非安全模式之间转换时发出警告”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="658ce-251">Make sure **Warn if Changing between Secure and not secure mode** is selected in the **Advanced** tab of the **Internet Options** dialog box.</span></span>

- <span data-ttu-id="658ce-p132">为了确保加载项不使用过多的 CPU 内核或内存资源且不导致客户端计算机上出现任何拒绝服务的情况，加载项平台建立了资源使用率限制。作为测试的一部分，开发人员应验证加载项平台是否遵循了资源使用率限制。</span><span class="sxs-lookup"><span data-stu-id="658ce-p132">To make sure that add-ins don't use excessive CPU core or memory resources and cause any denial of service on a client computer, the add-in platform establishes resource usage limits. As part of testing, developers should verify whether an add-in performs within the resource usage limits.</span></span>

- <span data-ttu-id="658ce-254">在发布加载项之前，开发人员应确保在其加载项文件中公开的任何个人身份信息是安全的。</span><span class="sxs-lookup"><span data-stu-id="658ce-254">Before publishing an add-in, developers should make sure that any personal identifiable information that they expose in their add-in files is secure.</span></span>

- <span data-ttu-id="658ce-p133">开发人员不应嵌入用于直接在加载项的 HTML 页面中访问第三方 API 或服务（例如 Bing、Google 或 Facebook）的密钥。相反，他们应该创建自定义 Web 服务或安全 Web 存储的其他某些窗体中创建自定义 Web 服务，他们可以调用这些服务，将键值传递到加载项。</span><span class="sxs-lookup"><span data-stu-id="658ce-p133">Developers shouldn't embed keys that they use to access third-party APIs or services (such as Bing, Google, or Facebook) directly in the HTML pages of their add-in. Instead, they should create a custom web service or store the keys in some other form of secure web storage that they can then call to pass the key value to their add-in.</span></span>

- <span data-ttu-id="658ce-257">开发人员应在将加载项提交到 AppSource 时执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="658ce-257">Developers should do the following when submitting an add-in to AppSource:</span></span>

  - <span data-ttu-id="658ce-258">在支持 SSL 的 Web 服务器上托管要提交的加载项。</span><span class="sxs-lookup"><span data-stu-id="658ce-258">Host the add-in they are submitting on a web server that supports SSL.</span></span>
  - <span data-ttu-id="658ce-259">制定概述遵从性隐私策略的声明。</span><span class="sxs-lookup"><span data-stu-id="658ce-259">Produce a statement outlining a compliant privacy policy.</span></span>
  - <span data-ttu-id="658ce-260">准备好在提交加载项后签订合约协议。</span><span class="sxs-lookup"><span data-stu-id="658ce-260">Be ready to sign a contractual agreement upon submitting the add-in.</span></span>

<span data-ttu-id="658ce-p134">除资源使用率规则之外，Outlook 外接程序的开发人员还应确保其外接程序遵守有关指定激活规则和使用 JavaScript API 的限制。有关详细信息，请参阅[激活限制和适用于 Outlook 外接程序的 JavaScript API](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="658ce-p134">Other than resource usage rules, developers for Outlook add-ins should also make sure their add-ins observe limits for specifying activation rules and using the JavaScript API. For more information, see [Limits for activation and JavaScript API for Outlook add-ins](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md).</span></span>

## <a name="it-administrators-control"></a><span data-ttu-id="658ce-263">IT 管理员控制</span><span class="sxs-lookup"><span data-stu-id="658ce-263">IT administrators' control</span></span>

<span data-ttu-id="658ce-264">在企业设置中，对于启用或禁用对 AppSource 和任何专用目录的访问权限，IT 管理员拥有最高权限。</span><span class="sxs-lookup"><span data-stu-id="658ce-264">In a corporate setting, IT administrators have ultimate authority over enabling or disabling access to AppSource and any private catalogs.</span></span>

<span data-ttu-id="658ce-265">Office 设置的管理和执行由组策略设置完成。</span><span class="sxs-lookup"><span data-stu-id="658ce-265">The management and enforcement of Office settings is done with group policy settings.</span></span> <span data-ttu-id="658ce-266">这些操作可通过 [Office 部署工具](/deployoffice/overview-of-the-office-2016-deployment-tool)和 [Office 自定义工具](/deployoffice/overview-of-the-office-customization-tool-for-click-to-run)进行配置。</span><span class="sxs-lookup"><span data-stu-id="658ce-266">These are configurable through the [Office Deployment Tool](/deployoffice/overview-of-the-office-2016-deployment-tool), in conjunction with the [Office Customization Tool](/deployoffice/overview-of-the-office-customization-tool-for-click-to-run).</span></span>

| <span data-ttu-id="658ce-267">设置名称</span><span class="sxs-lookup"><span data-stu-id="658ce-267">Setting name</span></span> | <span data-ttu-id="658ce-268">说明</span><span class="sxs-lookup"><span data-stu-id="658ce-268">Description</span></span> |
|--------------|-------------|
| <span data-ttu-id="658ce-269">允许不安全的 Web 加载项和目录</span><span class="sxs-lookup"><span data-stu-id="658ce-269">Allow Unsecure web add-ins and Catalogs</span></span> | <span data-ttu-id="658ce-270">允许用户运行不安全的加载项，这些加载项的网页或目录位置不受 SSL 保护 (https://) 且不在用户的 Internet 区域中。</span><span class="sxs-lookup"><span data-stu-id="658ce-270">Allows users to run non-secure add-ins, which are add-ins that have webpage or catalog locations that are not SSL-secured (https://) and are not in users' Internet zones.</span></span> |
| <span data-ttu-id="658ce-271">阻止 Web 加载项</span><span class="sxs-lookup"><span data-stu-id="658ce-271">Block Web Add-ins</span></span> | <span data-ttu-id="658ce-272">允许阻止用户使用 Web 加载项。</span><span class="sxs-lookup"><span data-stu-id="658ce-272">Allows you to prevent users from using web add-ins.</span></span> |
| <span data-ttu-id="658ce-273">阻止 Office 应用商店</span><span class="sxs-lookup"><span data-stu-id="658ce-273">Block the Office Store</span></span> |  <span data-ttu-id="658ce-274">允许阻止用户使用或插入来自 Office 应用商店的 Web 加载项。</span><span class="sxs-lookup"><span data-stu-id="658ce-274">Allows you to prevent users from using or inserting web add-ins that come from the Office Store.</span></span> |

> [!IMPORTANT]
> <span data-ttu-id="658ce-275">如果你的工作组正在使用 Office 的多个版本，则必须为每个版本配置组策略设置。</span><span class="sxs-lookup"><span data-stu-id="658ce-275">If your working groups are using multiple releases of Office, group policy settings must be configured for each release.</span></span> <span data-ttu-id="658ce-276">要详细了解针对 Office 2013 的组策略设置，请参阅 [Office 2013 相关应用概述](/previous-versions/office/office-2013-resource-kit/jj219429(v%3doffice.15))一文中的[使用组策略来管理用户可如何安装和使用 Office 相关应用](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office)。</span><span class="sxs-lookup"><span data-stu-id="658ce-276">Please refer to the [Using Group Policy to manage how users can install and use apps for Office](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office) of the [Overview of apps for Office 2013](/previous-versions/office/office-2013-resource-kit/jj219429(v%3doffice.15)) article for details on group policy settings for Office 2013.</span></span>

## <a name="see-also"></a><span data-ttu-id="658ce-277">另请参阅</span><span class="sxs-lookup"><span data-stu-id="658ce-277">See also</span></span>

- [<span data-ttu-id="658ce-278">在加载项中请求获取 API 使用权限</span><span class="sxs-lookup"><span data-stu-id="658ce-278">Requesting permissions for API use in add-ins</span></span>](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
- [<span data-ttu-id="658ce-279">Outlook 外接程序的隐私、权限和安全性</span><span class="sxs-lookup"><span data-stu-id="658ce-279">Privacy, permissions, and security for Outlook add-ins</span></span>](../outlook/privacy-and-security.md)
- [<span data-ttu-id="658ce-280">了解 Outlook 外接程序权限</span><span class="sxs-lookup"><span data-stu-id="658ce-280">Understanding Outlook add-in permissions</span></span>](../outlook/understanding-outlook-add-in-permissions.md)
- [<span data-ttu-id="658ce-281">Outlook 外接程序的激活和 JavaScript API 限制</span><span class="sxs-lookup"><span data-stu-id="658ce-281">Limits for activation and JavaScript API for Outlook add-ins</span></span>](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="658ce-282">解决 Office 外接程序中的同源策略限制</span><span class="sxs-lookup"><span data-stu-id="658ce-282">Addressing same-origin policy limitations in Office Add-ins</span></span>](../develop/addressing-same-origin-policy-limitations.md)
- [<span data-ttu-id="658ce-283">同源策略</span><span class="sxs-lookup"><span data-stu-id="658ce-283">Same Origin Policy</span></span>](https://www.w3.org/Security/wiki/Same_Origin_Policy)
- [<span data-ttu-id="658ce-284">同源策略第 1 部分：不准偷看</span><span class="sxs-lookup"><span data-stu-id="658ce-284">Same Origin Policy Part 1: No Peeking</span></span>](/archive/blogs/ieinternals/same-origin-policy-part-1-no-peeking)
- [<span data-ttu-id="658ce-285">针对 JavaScript 的同源策略</span><span class="sxs-lookup"><span data-stu-id="658ce-285">Same origin policy for JavaScript</span></span>](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy)
- [<span data-ttu-id="658ce-286">IE 保护模式</span><span class="sxs-lookup"><span data-stu-id="658ce-286">IE Protect Mode</span></span>](https://support.microsoft.com/help/2761180/apps-for-office-don-t-start-if-you-disable-protected-mode-for-the-restricted-sites-zone-in-internet-explorer)
