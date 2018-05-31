---
title: Office 加载项的隐私和安全
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 326c8095b6ced105cc21492dc290a443212b3d3f
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437729"
---
# <a name="privacy-and-security-for-office-add-ins"></a><span data-ttu-id="e7e4b-102">Office 加载项的隐私和安全</span><span class="sxs-lookup"><span data-stu-id="e7e4b-102">Privacy and security for Office Add-ins</span></span>

## <a name="understanding-the-add-in-runtime"></a><span data-ttu-id="e7e4b-103">了解加载项运行时</span><span class="sxs-lookup"><span data-stu-id="e7e4b-103">Understanding the add-in runtime</span></span>

<span data-ttu-id="e7e4b-p101">Office 外接程序受到外接程序运行时环境、多层权限模型和性能调控器的保护。这一框架通过以下方式保护用户体验：</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p101">Office Add-ins are secured by an add-in runtime environment, a multiple-tier permissions model, and performance governors. This framework protects the user's experience in the following ways:</span></span> 

- <span data-ttu-id="e7e4b-106">管理对主机应用程序的 UI 框架的访问。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-106">Access to the host application's UI frame is managed.</span></span>

- <span data-ttu-id="e7e4b-107">只允许间接访问主机应用的 UI 线程。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-107">Only indirect access to the host application's UI thread is allowed.</span></span>

- <span data-ttu-id="e7e4b-108">不允许模式交互。例如，不允许调用 JavaScript **alert**、**confirm** 和 **prompt** 函数，因为它们是模式函数。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-108">Modal interactions aren't allowed - for example, calls to JavaScript  **alert**, **confirm**, and **prompt** functions aren't allowed because they're modal.</span></span>

<span data-ttu-id="e7e4b-109">此外，为了确保 Office 加载项不会损害用户环境，运行时框架还提供以下优势：</span><span class="sxs-lookup"><span data-stu-id="e7e4b-109">Further, the runtime framework provides the following benefits to ensure that an Office Add-in can't damage the user's environment:</span></span>

- <span data-ttu-id="e7e4b-110">隔离运行加载项的进程。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-110">Isolates the process the add-in runs in.</span></span>

- <span data-ttu-id="e7e4b-111">不需要 .dll 或 .exe 替换项或 ActiveX 组件。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-111">Doesn't require .dll or .exe replacement or ActiveX components.</span></span>

- <span data-ttu-id="e7e4b-112">可以轻松安装和卸载加载项。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-112">Makes add-ins easy to install and uninstall.</span></span>

<span data-ttu-id="e7e4b-113">此外，还可以调控 Office 外接程序使用的内存、CPU 和网络资源，以确保维持良好的性能和可靠性。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-113">Also, the use of memory, CPU, and network resources by Office Add-ins is governable to ensure that good performance and reliability are maintained.</span></span> 

<span data-ttu-id="e7e4b-114">以下各节简要介绍在基于 Windows 的设备上、OS X Mac 设备上以及 Web 上的 Office Online 客户端中，运行时体系结构如何支持在 Office 客户端中运行加载项。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-114">The following sections briefly describe how the runtime architecture supports running add-ins in Office clients on Windows-based devices, on OS X Mac devices, and in Office Online clients on the web.</span></span>

> <span data-ttu-id="e7e4b-115">**注意**：若要了解如何将 WIP 和 Intune 与 Office 加载项结合使用，请参阅[使用 WIP 和 Intune 保护运行 Office 加载项的文档中的企业数据](https://docs.microsoft.com/en-us/microsoft-365-enterprise/office-add-ins-wip)。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-115">**NOTE**  To learn about using WIP and Intune with Office Add-ins, see [Use WIP and Intune to protect enterprise data in documents running Office Add-ins](https://docs.microsoft.com/en-us/microsoft-365-enterprise/office-add-ins-wip).</span></span>

### <a name="clients-for-windows-and-os-x-devices"></a><span data-ttu-id="e7e4b-116">适用于 Windows 和 OS X 设备的客户端</span><span class="sxs-lookup"><span data-stu-id="e7e4b-116">Clients for Windows and OS X devices</span></span>

<span data-ttu-id="e7e4b-p102">在支持的台式机和平板电脑设备的客户端（如 Excel、Outlook 和适用于 Mac 的 Outlook）中，通过集成进程内组件 Office 外接程序运行时来支持 Office 外接程序，该组件管理外接程序的生命周期，并实现外接程序和客户端应用程序之间的互操作性。外接程序网页本身托管在进程外。如图 1 中所示，在 Windows 台式机或平板电脑设备上，外接程序网页托管在 Internet Explorer 控件内部，而 Internet Explorer 控件托管在外接程序运行时进程内部，提供安全和性能隔离。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p102">In supported clients for desktop and tablet devices, such as Excel, Outlook, and Outlook for Mac, Office Add-ins are supported by integrating an in-process component, the Office Add-ins runtime, which manages the add-in lifecycle and enables interoperability between the add-in and the client application. The add-in webpage itself is hosted out-of-process. As shown in figure 1, on a Windows desktop or tablet device, the add-in webpage is hosted inside an Internet Explorer control which, in turn, is hosted inside an add-in runtime process that provides security and performance isolation.</span></span>

<span data-ttu-id="e7e4b-p103">在 Windows 桌面设备上，必须为受限网站区域启用 Internet Explorer 保护模式。通常情况下，此模式默认启用。如果禁用，则会在尝试启动外接程序时[看到错误消息](https://support.microsoft.com/en-us/help/2761180/apps-for-office-don-t-start-if-you-disable-protected-mode-for-the-restricted-sites-zone-in-internet-explorer)。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p103">On Windows desktops, Protect Mode in Internet Explorer must be enabled for the Restricted Site Zone. This is typically enabled by default. If it is disabled, an [error will occur](https://support.microsoft.com/en-us/help/2761180/apps-for-office-don-t-start-if-you-disable-protected-mode-for-the-restricted-sites-zone-in-internet-explorer) when you try to launch an add-in.</span></span>

<span data-ttu-id="e7e4b-123">*图 1.基于 Windows 的台式机和平板电脑客户端中的 Office 外接程序运行时环境*</span><span class="sxs-lookup"><span data-stu-id="e7e4b-123">*Figure 1. Office Add-ins runtime environment in Windows-based desktop and tablet clients*</span></span>

![富客户端基础结构](../images/dk2-agave-overview-02.png)

<span data-ttu-id="e7e4b-125">如下图所示，在 OS X Mac 台式机上，加载项网页托管在沙盒 WebKit 运行时主机进程内部，这有助于提供类似级别的安全和性能保护。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-125">As shown in the following figure, on an OS X Mac desktop, the add-in web page is hosted inside a sandboxed WebKit runtime host process which helps provide similar level of security and performance protection.</span></span> 

<span data-ttu-id="e7e4b-126">*图 2：OS X Mac 客户端中的 Office 加载项运行时环境*</span><span class="sxs-lookup"><span data-stu-id="e7e4b-126">*Figure 2. Office Add-ins runtime environment in OS X Mac clients*</span></span>

![OS X Mac 上的 Office 相关应用程序运行时环境](../images/dk2-agave-overview-mac-02.png)

<span data-ttu-id="e7e4b-128">Office 外接程序运行时管理进程间通信、JavaScript API 调用和事件到本机调用和事件的转换以及 UI 远程处理支持，从而使加载项能够呈现在文档内、任务窗格中或电子邮件、会议请求或约会旁边。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-128">The Office Add-ins runtime manages interprocess communication, the translation of JavaScript API calls and events into native ones, as well as UI remoting support to enable the add-in to be rendered inside the document, in a task pane, or adjacent to an email message, meeting request, or appointment.</span></span>

### <a name="web-clients"></a><span data-ttu-id="e7e4b-129">Web 客户端</span><span class="sxs-lookup"><span data-stu-id="e7e4b-129">Web clients</span></span>

<span data-ttu-id="e7e4b-p104">在支持的 Web 客户端（如 Excel Online 和 Outlook Web App）中，Office 加载项托管在使用 HTML5 **sandbox** 属性运行的 **iframe** 中。不得使用 ActiveX 组件或导航 Web 客户端的主页。Web 客户端通过集成适用于 Office 的 JavaScript API，启用 Office 加载项支持。JavaScript API 管理加载项生命周期，以及加载项与 Web 客户端之间的互操作性，与桌面客户端应用采取的方式类似。此互操作性是使用特殊的交叉框架公告消息通信基础结构进行实现。在桌面客户端上使用的相同 JavaScript 库 (Office.js) 可用于与 Web 客户端进行交互。下图展示了支持 Office Online 中 Office 加载项（在浏览器中运行）的基础结构，以及支持这些加载项所需的相关组件（Web 客户端、**iframe**、Office 加载项运行时和适用于 Office 的 JavaScript API）。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p104">In supported Web clients, such as Excel Online and Outlook Web App, Office Add-ins are hosted in an  **iframe** that runs using the HTML5 **sandbox** attribute. ActiveX components or navigating the main page of the web client are not allowed. Office Add-ins support is enabled in the web clients by the integration of the JavaScript API for Office. In a similar way to the desktop client applications, the JavaScript API manages the add-in lifecycle and interoperability between the add-in and the web client. This interoperability is implemented by using a special cross-frame post message communication infrastructure. The same JavaScript library (Office.js) that is used on desktop clients is available to interact with the web client. The following figure shows the infrastructure that supports Office Add-ins in Office Online (running in the browser), and the relevant components (the web client, **iframe**, Office Add-ins runtime, and JavaScript API for Office) that are required to support them.</span></span>

<span data-ttu-id="e7e4b-137">*图 3：支持 Office Web 客户端中 Office 加载项的基础结构*</span><span class="sxs-lookup"><span data-stu-id="e7e4b-137">*Figure 3. Infrastructure that supports Office Add-ins in Office web clients*</span></span>

![Web 客户端基础结构](../images/dk2-agave-overview-03.png)

## <a name="add-in-integrity-in-appsource"></a><span data-ttu-id="e7e4b-139">AppSource 中的加载项完整性</span><span class="sxs-lookup"><span data-stu-id="e7e4b-139">Add-in integrity in AppSource</span></span>

<span data-ttu-id="e7e4b-p105">若要向受众提供 Office 加载项，可以在 AppSource 中发布它们。AppSource 强制执行以下措施来维护加载项完整性：</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p105">You can make your Office Add-ins available to the public by publishing them to AppSource. AppSource enforces the following measures to maintain the integrity of add-ins:</span></span>


- <span data-ttu-id="e7e4b-142">要求 Office 加载项的主机服务器始终使用安全套接字层 (SSL) 进行通信。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-142">Requires the host server of an Office Add-in to always use Secure Sockets Layer (SSL) to communicate.</span></span>

- <span data-ttu-id="e7e4b-143">要求开发人员在提交加载项时提供身份证明、合约协议和适合的隐私策略。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-143">Requires a developer to provide proof of identity, a contractual agreement, and a compliant privacy policy to submit add-ins.</span></span>

- <span data-ttu-id="e7e4b-144">确保加载项的源在只读模式下可访问。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-144">Ensures that the source of add-ins is accessible in read-only mode.</span></span>

- <span data-ttu-id="e7e4b-145">支持针对可用加载项的用户审阅系统以推广自我管理的社区。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-145">Supports a user-review system for available add-ins to promote a self-policing community.</span></span>

## <a name="addressing-end-users-privacy-concerns"></a><span data-ttu-id="e7e4b-146">解决最终用户的隐私问题</span><span class="sxs-lookup"><span data-stu-id="e7e4b-146">Addressing end users' privacy concerns</span></span>

<span data-ttu-id="e7e4b-147">此部分从客户（最终用户）的角度出发介绍了 Office 外接程序平台提供的保护，并介绍了有关如何达到用户的预期以及如何安全处理用户个人身份信息 (PII) 的指南。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-147">This section describes the protection offered by the Office Add-ins platform from the customer's (end user's) perspective, and provides guidelines for how to support users' expectations and how to securely handle users' personally identifiable information (PII).</span></span>

### <a name="end-users-perspective"></a><span data-ttu-id="e7e4b-148">从最终用户的角度出发</span><span class="sxs-lookup"><span data-stu-id="e7e4b-148">End users' perspective</span></span>

<span data-ttu-id="e7e4b-p106">Office 加载项是使用浏览器控件或 **iframe** 中运行的 Web 技术而生成。因此，使用加载项与转到 Internet 或 Intranet 上的网站类似。加载项可以位于组织外部（如果从 AppSource 获取加载项的话），也可以位于组织内部（如果从 Exchange Server 加载项目录、SharePoint 加载项目录或组织网络上的文件共享获取加载项的话）。加载项对网络的访问权限受限，大部分加载项都可以对活动文档或邮件项执行读取或写入操作。在用户或管理员安装或启动加载项前，加载项平台就施加了特定约束。不过，与任何扩展性模型一样，用户在启动未知加载项之前应非常谨慎。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p106">Office Add-ins are built using web technologies that run in a browser control or **iframe**. Because of this, using add-ins is similar to browsing to web sites on the Internet or intranet. Add-ins can be external to an organization (if you acquire the add-in from AppSource) or internal (if you acquire the add-in from an Exchange Server add-in catalog, SharePoint add-in catalog, or file share on an organization's network). Add-ins have limited access to the network and most add-ins can read or write to the active document or mail item. The add-in platform applies certain constraints before a user or administrator installs or starts an add-in. But as with any extensibility model, users should be cautious before starting an unknown add-in.</span></span>

<span data-ttu-id="e7e4b-155">加载项平台解决了最终用户的隐私问题，具体方式如下：</span><span class="sxs-lookup"><span data-stu-id="e7e4b-155">The add-in platform addresses end users' privacy concerns in the following ways:</span></span>

- <span data-ttu-id="e7e4b-156">与托管内容、Outlook 或任务窗格外接程序的 Web 服务器通信的数据以及外接程序与其使用的任何 Web 服务之间的通信必须使用安全套接字层 (SSL) 协议加密。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-156">Data communicated with the web server that hosts a content, Outlook or task pane add-in as well as communication between the add-in and any web services it uses must be encrypted using the Secure Socket Layer (SSL) protocol.</span></span>

- <span data-ttu-id="e7e4b-p107">安装 AppSource 中的加载项前，用户可以查看相应加载项的隐私策略和要求。此外，与用户邮箱进行交互的 Outlook 加载项还指明了所需的特定权限；用户可以在安装 Outlook 加载项前，先查看使用条款、请求的权限和隐私策略。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p107">Before a user installs an add-in from AppSource, the user can view the privacy policy and requirements of that add-in. In addition, Outlook add-ins that interact with users' mailboxes surface the specific permissions that they require; the user can review the terms of use, requested permissions and privacy policy before installing an Outlook add-in.</span></span>

- <span data-ttu-id="e7e4b-p108">在共享一个文档时，用户也会共享已插入该文档或与该文档关联的加载项。如果用户打开一个包含其之前未使用的加载项的文档，则主机应用程序会提示用户向加载项授予在文档中运行的权限。在组织环境中，如果文档来自外部源，则 Office 主机应用程序也会提示用户。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p108">When sharing a document, users also share add-ins that have been inserted in or associated with that document. If a user opens a document that contains an add-in that the user hasn't used before, the host application prompts the user to grant permission for the add-in to run in the document. In an organizational environment, the Office host application also prompts the user if the document comes from an external source.</span></span>

- <span data-ttu-id="e7e4b-p109">用户可以启用或禁用对 AppSource 的访问权限。对于内容和任务窗格加载项，用户通过主机 Office 客户端上的“信任中心”****（通过“文件”**** > “选项”**** > “信任中心”**** > “信任中心设置”**** > “受信任的加载项目录”**** 打开），管理对受信任的加载项和目录的访问权限。对于 Outlook 加载项，用户可以通过选择“管理加载项”**** 按钮管理加载项，具体操作为：在 Outlook for Windows 中，依次选择“文件”**** > “管理加载项”****；在 Outlook for Mac 中，选择加载项栏上的“管理加载项”**** 按钮；在 Outlook Web App 中，依次选择“设置”**** 菜单（齿轮图标）>“管理加载项”****。管理员还可以通过[使用组策略](http://technet.microsoft.com/en-us/library/jj219429.aspx#BKMK_Managing)管理此访问权限。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p109">Users can enable or disable the access to AppSource. For content and task pane add-ins, users manage access to trusted add-ins and catalogs from the  **Trust Center** on the host Office client (opened from **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**). For Outlook add-ins, uses can manage add-ins by choosing the  **Manage Add-ins** button: in Outlook for Windows, choose **File** > **Manage Add-ins**. In Outlook for Mac, choose the  **Manage Add-ins** button on the add-in bar. In Outlook Web App choose the **Settings** menu (gear icon) > **Manage add-ins**. Administrators can also manage this access [by using group policy](http://technet.microsoft.com/en-us/library/jj219429.aspx#BKMK_Managing).</span></span>

- <span data-ttu-id="e7e4b-166">加载项平台的设计在以下方面为最终用户提供了安全和性能保障：</span><span class="sxs-lookup"><span data-stu-id="e7e4b-166">The design of the add-in platform provides security and performance for end users in the following ways:</span></span>

  - <span data-ttu-id="e7e4b-p110">Office 外接程序在托管在独立于 Office 主机应用程序的加载项运行时环境中的 Web 浏览器控件中运行。此设计提供了与主机应用程序的安全和性能隔离。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p110">An Office Add-in runs in a web browser control that is hosted in an add-in runtime environment separate from the Office host application. This design provides both security and performance isolation from the host application.</span></span>

  - <span data-ttu-id="e7e4b-169">在 Web 浏览器控件中运行可允许加载项完成在浏览器中运行的常规网页可执行的所有操作，但同时将限制加载项遵守针对域隔离和安全区域的同源策略。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-169">Running in a web browser control allows the add-in to do almost anything a regular web page running in a browser can do but, at the same time, restricts the add-in to observe the same-origin policy for domain isolation and security zones.</span></span>

<span data-ttu-id="e7e4b-p111">Outlook 外接程序通过特定的资源使用率监视提供额外安全和性能功能。有关详细信息，请参阅 [Outlook 外接程序的隐私、权限和安全性](https://docs.microsoft.com/en-us/outlook/add-ins/privacy-and-security)。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p111">Outlook add-ins provide additional security and performance features through Outlook add-in specific resource usage monitoring. For more information, see [Privacy, permissions, and security for Outlook add-ins](https://docs.microsoft.com/en-us/outlook/add-ins/privacy-and-security).</span></span>

### <a name="developer-guidelines-to-handle-pii"></a><span data-ttu-id="e7e4b-172">开发人员处理 PII 的准则</span><span class="sxs-lookup"><span data-stu-id="e7e4b-172">Developer guidelines to handle PII</span></span>

<span data-ttu-id="e7e4b-p112">你可以在[保护人力资源应用程序开发和测试中的隐私](http://technet.microsoft.com/en-us/library/gg447064.aspx)中阅读 IT 管理员和开发人员用于保护 PII 的通用准则。下面列出了对 Office 外接程序开发人员的一些特定 PII 保护准则：</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p112">You can read general PII protection guidelines for IT administrators and developers in [Protecting Privacy in the Development and Testing of Human Resources Applications](http://technet.microsoft.com/en-us/library/gg447064.aspx). The following lists some specific PII protection guidelines for you as a developer of Office Add-ins:</span></span>

- <span data-ttu-id="e7e4b-p113">[Settings](https://dev.office.com/reference/add-ins/shared/settings) 对象旨在保存内容加载项或任务窗格加载项的会话之间的加载项设置和状态数据，但不会在 **Settings** 对象中存储密码和其他敏感 PII。最终用户无法查看 **Settings** 对象中的数据，但该数据存储为文档的易于访问的文件格式的一部分。你应该限制加载项对 PII 的使用，并将加载项所需的任何 PII 存储在将加载项作为用户保护的资源托管的服务器上。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p113">The [Settings](https://dev.office.com/reference/add-ins/shared/settings) object is intended for persisting add-in settings and state data across sessions for a content or task pane add-in, but don't store passwords and other sensitive PII in the **Settings** object. The data in the **Settings** object isn't visible to end users, but it is stored as part of the document's file format which is readily accessible. You should limit your add-in's use of PII and store any PII required by your add-in on the server hosting your add-in as a user-secured resource.</span></span>

- <span data-ttu-id="e7e4b-p114">使用某些应用程序可能会泄露 PII。请确保安全地存储用户的身份、位置、访问时间和任何其他凭据数据，以便该加载项的其他用户无法访问该数据。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p114">Using some applications can reveal PII. Make sure that you securely store data for your users' identity, location, access times, and any other credentials so that data won't become available to other users of the add-in.</span></span>

- <span data-ttu-id="e7e4b-p115">如果加载项已在 AppSource 中发布，HTTPS 的 AppSource 要求会保护 Web 服务器与客户端计算机或设备之间传输的 PII。不过，如果将此类数据重新传输到其他服务器，请务必遵守相同级别的保护要求。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p115">If your add-in is available in AppSource, the AppSource requirement for HTTPS protects PII transmitted between your web server and the client computer or device. However, if you re-transmit that data to other servers, make sure you observe the same level of protection.</span></span>

- <span data-ttu-id="e7e4b-p116">如果存储用户的 PII，请务必向用户提示这一点，并向用户提供检查和删除此类信息的方法。如果将加载项提交到 AppSource，可以在隐私声明中概述所收集的数据及其用途。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p116">If you store users' PII, make sure you reveal that fact, and provide a way for users to inspect and delete it. If you submit your add-in to AppSource, you can outline the data you collect and how it's used in the privacy statement.</span></span>

## <a name="developers-permission-choices-and-security-practices"></a><span data-ttu-id="e7e4b-184">开发人员的权限选择和安全做法</span><span class="sxs-lookup"><span data-stu-id="e7e4b-184">Developers' permission choices and security practices</span></span>

<span data-ttu-id="e7e4b-185">遵循这些常规指南以支持 Office 外接程序的安全模型，并进一步了解有关每种加载项类型的更多详细信息。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-185">Follow these general guidelines to support the security model of Office Add-ins, and drill down on more details for each add-in type.</span></span>

### <a name="permissions-choices"></a><span data-ttu-id="e7e4b-186">权限选择</span><span class="sxs-lookup"><span data-stu-id="e7e4b-186">Permissions choices</span></span>

<span data-ttu-id="e7e4b-187">外接程序平台提供了一种权限模型，外接程序可以使用它来声明对其功能所需的用户数据的访问级别。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-187">The add-in platform provides a permissions model that your add-in uses to declare the level of access to a user's data that it requires for its features.</span></span> <span data-ttu-id="e7e4b-188">每个权限级别都对应于外接程序可以用于其功能的适用于 Office 的 JavaScript API 的子集。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-188">Each permission level corresponds to the subset of the JavaScript API for Office your add-in is allowed to use for its features.</span></span> <span data-ttu-id="e7e4b-189">例如，内容和任务窗格外接程序的 ** WriteDocument**  权限允许访问 [ Document.setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) 方法，该方法允许将附加内容写入用户的文档，但不允许访问从文档中读取数据的任何方法。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-189">For example, the  **WriteDocument** permission for content and task pane add-ins allows access to the [Document.setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) method that lets an add-in write to the user's document, but doesn't allow access to any of the methods for reading data from the document.</span></span> <span data-ttu-id="e7e4b-190">此权限级别对于只需写入文档的外接程序很有意义，例如用户可以查询要插入其文档的数据的外接程序。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-190">This permission level makes sense for add-ins that only need to write to a document, such as an add-in where the user can query for data to insert into their document.</span></span>

<span data-ttu-id="e7e4b-p118">最佳做法是应该基于“_最小特权_”原则请求权限。即应该请求外接程序正常运行所需的 API 的最小子集的访问权限。例如，如果外接程序只需要读取其功能的用户文档中的数据，则应仅请求“**ReadDocument**”权限。（但是，请注意如果请求权限不足，则会导致外接程序平台阻止外接程序使用部分 API 并将生成运行时错误。）</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p118">As a best practice, you should request permissions based on the principle of  _least privilege_. That is, you should request permission to access only the minimum subset of the API that your add-in requires to function correctly. For example, if your add-in needs only to read data in a user's document for its features, you should request no more than the **ReadDocument** permission. (But, keep in mind that requesting insufficient permissions will result in the add-in platform blocking your add-in's use of some APIs and will generate errors at run time.)</span></span>

<span data-ttu-id="e7e4b-p119">你在外接程序清单中指定权限，如本节下面的示例所示。最终用户可以在首次决定安装或激活外接程序之前查看外接程序请求的权限级别。此外，请求 **ReadWriteMailbox** 权限的 Outlook 外接程序需要明确的管理员权限才能安装。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p119">You specify permissions in the manifest of your add-in, as shown in the example in this section below, and end users can see the requested permission level of an add-in before they decide to install or activate the add-in for the first time. Additionally, Outlook add-ins that request the  **ReadWriteMailbox** permission require explicit administrator privilege to install.</span></span>

<span data-ttu-id="e7e4b-p120">以下示例演示任务窗格加载项如何在其清单中指定 **ReadDocument** 权限。为重点关注权限，清单中的其他元素将不显示。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p120">The following example shows how a task pane add-in specifies the  **ReadDocument** permission in its manifest. To keep permissions as the focus, other elements in the manifest aren't displayed.</span></span>

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

<span data-ttu-id="e7e4b-199">若要详细了解任务窗格和内容加载项权限，请参阅[在内容和任务窗格加载项中请求获取 API 使用权限](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-199">For more information about permissions for task pane and content add-ins, see [Requesting permissions for API use in content and task pane add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins).</span></span>

<span data-ttu-id="e7e4b-200">若要详细了解 Outlook 加载项权限，请参阅以下主题：</span><span class="sxs-lookup"><span data-stu-id="e7e4b-200">For more information about permissions for Outlook add-ins, see the following topics:</span></span>

- [<span data-ttu-id="e7e4b-201">Outlook 加载项的隐私、权限和安全</span><span class="sxs-lookup"><span data-stu-id="e7e4b-201">Privacy, permissions, and security for Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/privacy-and-security)

- [<span data-ttu-id="e7e4b-202">了解 Outlook 外接程序权限</span><span class="sxs-lookup"><span data-stu-id="e7e4b-202">Understanding Outlook add-in permissions</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)

### <a name="same-origin-policy"></a><span data-ttu-id="e7e4b-203">同源策略</span><span class="sxs-lookup"><span data-stu-id="e7e4b-203">Same origin policy</span></span>

<span data-ttu-id="e7e4b-204">由于 Office 外接程序是在 Web 浏览器控件中运行的网页，因此，它们必须遵守浏览器强制实施的同源策略：默认情况下，一个域中的网页无法执行它的域之外的其他域进行 [XmlHttpRequest](http://www.w3.org/TR/XMLHttpRequest/) Web 服务调用。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-204">Because Office Add-ins are webpages that run in a web browser control, they must follow the same-origin policy enforced by the browser: by default, a webpage in one domain can't make [XmlHttpRequest](http://www.w3.org/TR/XMLHttpRequest/) web service calls to another domain other than the one where it is hosted.</span></span>

<span data-ttu-id="e7e4b-p121">消除此限制的一种方法是使用 JSON/P - 通过包括一个带指向承载于其他域上的某个脚本的 **src** 属性的 **script** 标记来为 Web 服务提供代理。你可以编程方式创建 **script** 标记，动态创建 **src** 属性所指向的 URL，并通过 URI 查询参数将参数传递到 URL。Web 服务提供程序创建和承载特定 URL 上的 JavaScript 代码，然后根据 URI 查询参数返回不同的脚本。随后，这些脚本将在其插入到的位置执行并按预期工作。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p121">One way to overcome this limitation is to use JSON/P -- provide a proxy for the web service by including a  **script** tag with a **src** attribute that points to some script hosted on another domain. You can programmatically create the **script** tags, dynamically creating the URL to which to point the **src** attribute, and passing parameters to the URL via URI query parameters. Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters. These scripts then execute where they are inserted and work as expected.</span></span>

<span data-ttu-id="e7e4b-209">以下是 Outlook 外接程序示例中的 JSON/P 的示例。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-209">The following is an example of JSON/P in the Outlook add-in example.</span></span> 

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

<span data-ttu-id="e7e4b-p122">Exchange 和 SharePoint 提供了客户端代理以实现跨域访问。通常，Intranet 上的同源策略没有 Internet 上的同源策略那样严格。有关详细信息，请参阅[同源策略第 1 部分：不准偷看](http://blogs.msdn.com/b/ieinternals/archive/2009/08/28/explaining-same-origin-policy-part-1-deny-read.aspx)和[解决 Office 加载项中的同源策略限制](../develop/addressing-same-origin-policy-limitations.md)。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p122">Exchange and SharePoint provide client-side proxies to enable cross-domain access. In general, same origin policy on an intranet isn't as strict as on the Internet. For more information, see [Same Origin Policy Part 1: No Peeking](http://blogs.msdn.com/b/ieinternals/archive/2009/08/28/explaining-same-origin-policy-part-1-deny-read.aspx) and [Addressing same-origin policy limitations in Office Add-ins](../develop/addressing-same-origin-policy-limitations.md).</span></span>

### <a name="tips-to-prevent-malicious-cross-site-scripting"></a><span data-ttu-id="e7e4b-213">防止恶意跨站点脚本的提示</span><span class="sxs-lookup"><span data-stu-id="e7e4b-213">Tips to prevent malicious cross-site scripting</span></span>

<span data-ttu-id="e7e4b-214">恶意用户可以通过在文档或外接程序字段中输入恶意脚本来攻击外接程序的来源。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-214">An ill-intentioned user could attack the origin of an add-in by entering malicious script through the document or fields in the add-in.</span></span> <span data-ttu-id="e7e4b-215">开发人员应该处理用户输入，以避免在其域中执行恶意用户的JavaScript。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-215">A developer should process user input to avoid executing a malicious user's JavaScript within their domain.</span></span> <span data-ttu-id="e7e4b-216">以下是处理用户通过文档、邮件消息或外接程序字段提供的输入的一些良好实践：</span><span class="sxs-lookup"><span data-stu-id="e7e4b-216">The following are some good practices to follow to handle user input from a document or mail message, or via fields in an add-in:</span></span>


- <span data-ttu-id="e7e4b-p124">根据需要使用 [innerText](http://msdn.microsoft.com/en-us/library/ie/ms533897.aspx) 和 [textContent](https://msdn.microsoft.com/library/ms533899.aspx) 属性，而非 DOM 属性 [innerHTML](https://developer.mozilla.org/en-US/docs/DOM/Node.textContent)。执行以下操作获取 Internet Explorer 和 Firefox 跨浏览器支持：</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p124">Instead of the DOM property [innerHTML](http://msdn.microsoft.com/en-us/library/ie/ms533897.aspx), use the [innerText](https://msdn.microsoft.com/library/ms533899.aspx) and [textContent](https://developer.mozilla.org/en-US/docs/DOM/Node.textContent) properties where appropriate. Do the following for Internet Explorer and Firefox cross-browser support:</span></span>

    ```js
     var text = x.innerText || x.textContent
    ```

    <span data-ttu-id="e7e4b-p125">有关 **innerText** 和 **textContent** 之间区别的信息，请参阅 [Node.textContent](https://developer.mozilla.org/en-US/docs/DOM/Node.textContent)。有关常见浏览器间 DOM 兼容性的详细信息，请参阅 [W3C DOM 兼容性 - HTML](http://www.quirksmode.org/dom/w3c_html.html#t07)。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p125">For information about the differences between  **innerText** and **textContent**, see [Node.textContent](https://developer.mozilla.org/en-US/docs/DOM/Node.textContent). For more information about DOM compatibility across common browsers, see [W3C DOM Compatibility - HTML](http://www.quirksmode.org/dom/w3c_html.html#t07).</span></span>

- <span data-ttu-id="e7e4b-p126">如果你必须使用 **innerHTML**，请在将用户输入传递到 **innerHTML** 之前确保用户输入不包含恶意内容。有关详细信息以及如何安全使用 **innerHTML** 的示例，请参阅 [innerHTML](http://msdn.microsoft.com/en-us/library/ie/ms533897.aspx) 属性。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p126">If you must use  **innerHTML**, make sure the user's input doesn't contain malicious content before passing it to  **innerHTML**. For more information and an example of how to use  **innerHTML** safely, see [innerHTML](http://msdn.microsoft.com/en-us/library/ie/ms533897.aspx) property.</span></span>

- <span data-ttu-id="e7e4b-223">如果要使用 jQuery，请使用 [.text()](http://api.jquery.com/text/) 方法，而非 [.html()](http://api.jquery.com/html/) 方法。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-223">If you are using jQuery, use the [.text()](http://api.jquery.com/text/) method instead of the [.html()](http://api.jquery.com/html/) method.</span></span>

- <span data-ttu-id="e7e4b-224">使用 [toStaticHTML](http://msdn.microsoft.com/en-us/library/ie/cc848922.aspx) 方法可在将用户输入传递到 **innerHTML** 之前删除用户输入中的所有动态 HTML 元素和属性。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-224">Use the [toStaticHTML](http://msdn.microsoft.com/en-us/library/ie/cc848922.aspx) method to remove any dynamic HTML elements and attributes in users' input before passing it to **innerHTML**.</span></span>

- <span data-ttu-id="e7e4b-225">使用 [encodeURIComponent](http://msdn.microsoft.com/en-us/library/8202bce6-1342-40dc-a5ef-ac6d210a7d15.aspx) 或 [encodeURI](http://msdn.microsoft.com/en-us/library/17bab5a2-bcd4-46c2-8b52-b2b5a0ed98a3.aspx) 函数可对应为来自用户输入或包含用户输入的 URL 的文本进行编码。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-225">Use the [encodeURIComponent](http://msdn.microsoft.com/en-us/library/8202bce6-1342-40dc-a5ef-ac6d210a7d15.aspx) or [encodeURI](http://msdn.microsoft.com/en-us/library/17bab5a2-bcd4-46c2-8b52-b2b5a0ed98a3.aspx) function to encode text that is intended to be a URL that comes from or contains user input.</span></span>

- <span data-ttu-id="e7e4b-226">有关创建更安全的 Web 解决方案的更多最佳做法，请参阅[开发安全加载项](http://msdn.microsoft.com/en-us/library/windows/apps/hh849625.aspx)。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-226">See [Developing secure add-ins](http://msdn.microsoft.com/en-us/library/windows/apps/hh849625.aspx) for more best practices to create more secure web solutions.</span></span>

### <a name="tips-to-prevent-clickjacking"></a><span data-ttu-id="e7e4b-227">防止“点击劫持”的提示</span><span class="sxs-lookup"><span data-stu-id="e7e4b-227">Tips to prevent "Clickjacking"</span></span>

<span data-ttu-id="e7e4b-228">由于 Office 加载项通过 Office Online 主机应用程序运行在浏览器中时呈现在 iframe 中，请使用以下提示来尽量降低[点击劫持](http://en.wikipedia.org/wiki/Clickjacking)（一种黑客用来欺骗用户泄露机密信息的技术）的风险。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-228">Because Office Add-ins are rendered in an iframe when running in a browser with Office Online host applications, use the following tips to minimize the risk of [clickjacking](http://en.wikipedia.org/wiki/Clickjacking) -- a technique used by hackers to fool users into revealing confidential information.</span></span>

<span data-ttu-id="e7e4b-p127">首先，确定您的加载项可以执行的敏感操作。其中包括未授权的用户可能恶意使用的任何操作，如启动金融交易或发布敏感数据。例如，您的加载项可能让用户将款项发送到用户定义的接收人。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p127">First, identify sensitive actions that your add-in can perform. These include any actions that an unauthorized user could use with malicious intent, such as initiating a financial transaction or publishing sensitive data. For example, your add-in might let the user send a payment to a user-defined recipient.</span></span>

<span data-ttu-id="e7e4b-p128">其次，对于敏感操作，你的加载项应在执行操作之前向用户确认。该确认应详细说明该操作将产生的影响。此外，如有必要，还应详细说明用户如何能够防止该操作，是通过选择标记为“不允许”的特定按钮，还是忽略确认。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p128">Second, for sensitive actions, your add-in should confirm with the user before it executes the action. This confirmation should detail what effect the action will have. It should also detail how the user can prevent the action, if necessary, whether by choosing a specific button marked "Don't Allow" or by ignoring the confirmation.</span></span>

<span data-ttu-id="e7e4b-235">第三，为了确保没有任何潜在的攻击者可以隐藏或掩盖确认，您应将其显示在加载项上下文以外（即，不在 HTML 对话框中）。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-235">Third, to ensure that no potential attacker can hide or mask the confirmation, you should display it outside the context of the add-in (that is, not in an HTML dialog box).</span></span>

<span data-ttu-id="e7e4b-236">下面是如何获取确认的一些示例：</span><span class="sxs-lookup"><span data-stu-id="e7e4b-236">Here are some examples of how you could get confirmation:</span></span>

- <span data-ttu-id="e7e4b-237">向用户发送包含确认链接的电子邮件。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-237">Send an email to the user that contains a confirmation link.</span></span>

- <span data-ttu-id="e7e4b-238">向用户发送短信，其中包含用户可在外接程序中输入的确认码。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-238">Send a text message to the user that includes a confirmation code that the user can enter in the add-in.</span></span>

- <span data-ttu-id="e7e4b-p129">对于无法应用 iframe 的页面，在新浏览器窗口中打开确认对话框。这通常是登录页采用的模式。使用[对话框 API](../develop/dialog-api-in-office-add-ins.md) 新建对话框。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p129">Open a confirmation dialog in a new browser window to a page that cannot be iframed. This is typically the pattern that is used by login pages. Use the [dialog api](../develop/dialog-api-in-office-add-ins.md) to create a new dialog.</span></span>

<span data-ttu-id="e7e4b-p130">此外，请确保您用于与用户联系的地址不能由潜在的攻击者提供。例如，对于付款确认，使用经授权用户帐户的文件中的地址。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p130">Also, ensure that the address you use for contacting the user couldn't have been provided by a potential attacker. For example, for payment confirmations use the address on file for the authorized user's account.</span></span>

### <a name="other-security-practices"></a><span data-ttu-id="e7e4b-244">其他安全实践</span><span class="sxs-lookup"><span data-stu-id="e7e4b-244">Other security practices</span></span>

<span data-ttu-id="e7e4b-245">开发人员还应记下以下安全实践：</span><span class="sxs-lookup"><span data-stu-id="e7e4b-245">Developers should also take note of the following security practices:</span></span>


- <span data-ttu-id="e7e4b-246">开发人员不得在 Office 加载项中使用 ActiveX 控件，因为 ActiveX 控件不支持加载项平台的跨平台特性。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-246">Developers shouldn't use ActiveX controls in Office Add-ins as ActiveX controls don't support the cross-platform nature of the add-in platform.</span></span>

- <span data-ttu-id="e7e4b-p131">内容和任务窗格加载项使用 Internet Explorer 默认使用的相同 SSL 设置，并允许大部分内容仅通过 SSL 传送。Outlook 加载项要求所有内容都通过 SSL 传送。开发人员必须在加载项清单的 **SourceLocation** 元素中指定使用 HTTPS 的 URL，以标识加载项的 HTML 文件位置。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p131">Content and task pane add-ins assume the same SSL settings that Internet Explorer uses by default, and allows most content to be delivered only by SSL. Outlook add-ins require all content to be delivered by SSL. Developers must specify in the **SourceLocation** element of the add-in manifest a URL that uses HTTPS, to identify the location of the HTML file for the add-in.</span></span>

    <span data-ttu-id="e7e4b-250">若要确保加载项不使用 HTTP 交付内容，在测试加载项时，开发人员应确保在 Internet Explorer 中选择以下设置且其测试方案中不显示任何安全警告：</span><span class="sxs-lookup"><span data-stu-id="e7e4b-250">To make sure add-ins aren't delivering content by using HTTP, when testing add-ins, developers should make sure the following settings are selected in Internet Explorer and no security warnings appear in their test scenarios:</span></span>

    - <span data-ttu-id="e7e4b-p132">确保针对“**Internet**”区域的安全设置“**显示混合内容**”设置为“**提示**”。可以通过在 Internet Explorer 中选择以下项目来完成此设置：在“**Internet 选项**”对话框的“**安全**”选项卡上，选择“**Internet**”区域，然后选择“**自定义级别**”，滚动查找“**显示混合内容**”并选择“**提示**”（如果未选择）。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p132">Make sure the security setting,  **Display mixed content**, for the  **Internet** zone is set to **Prompt**. You can do that by selecting the following in Internet Explorer: on the  **Security** tab of the **Internet Options** dialog box, select the **Internet** zone, select **Custom level**, scroll to look for  **Display mixed content**, and select  **Prompt** if it isn't already selected.</span></span>

    - <span data-ttu-id="e7e4b-253">确保在“Internet 选项”**** 对话框的“高级”**** 选项卡中，选中了“在安全和非安全模式之间转换时发出警告”****。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-253">Make sure **Warn if Changing between Secure and not secure mode** is selected in the **Advanced** tab of the **Internet Options** dialog box.</span></span>

- <span data-ttu-id="e7e4b-p133">为了确保加载项不使用过多的 CPU 内核或内存资源且不导致客户端计算机上出现任何拒绝服务的情况，加载项平台建立了资源使用率限制。作为测试的一部分，开发人员应验证加载项平台是否遵循了资源使用率限制。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p133">To make sure that add-ins don't use excessive CPU core or memory resources and cause any denial of service on a client computer, the add-in platform establishes resource usage limits. As part of testing, developers should verify whether an add-in performs within the resource usage limits.</span></span>

- <span data-ttu-id="e7e4b-256">在发布加载项之前，开发人员应确保在其加载项文件中公开的任何个人身份信息是安全的。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-256">Before publishing an add-in, developers should make sure that any personal identifiable information that they expose in their add-in files is secure.</span></span>

- <span data-ttu-id="e7e4b-p134">开发人员不应嵌入用于直接在加载项的 HTML 页面中访问第三方 API 或服务（例如 Bing、Google 或 Facebook）的密钥。相反，他们应该创建自定义 Web 服务或安全 Web 存储的其他某些窗体中创建自定义 Web 服务，他们可以调用这些服务，将键值传递到加载项。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p134">Developers shouldn't embed keys that they use to access third-party APIs or services (such as Bing, Google, or Facebook) directly in the HTML pages of their add-in. Instead, they should create a custom web service or store the keys in some other form of secure web storage that they can then call to pass the key value to their add-in.</span></span>

- <span data-ttu-id="e7e4b-259">开发人员应在将加载项提交到 AppSource 时执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="e7e4b-259">Developers should do the following when submitting an add-in to AppSource:</span></span>

  - <span data-ttu-id="e7e4b-260">在支持 SSL 的 Web 服务器上托管要提交的加载项。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-260">Host the add-in they are submitting on a web server that supports SSL.</span></span>
  - <span data-ttu-id="e7e4b-261">制定概述遵从性隐私策略的声明。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-261">Produce a statement outlining a compliant privacy policy.</span></span>
  - <span data-ttu-id="e7e4b-262">准备好在提交加载项后签订合约协议。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-262">Be ready to sign a contractual agreement upon submitting the add-in.</span></span>

<span data-ttu-id="e7e4b-p135">除资源使用率规则之外，Outlook 外接程序的开发人员还应确保其外接程序遵守有关指定激活规则和使用 JavaScript API 的限制。有关详细信息，请参阅[激活限制和适用于 Outlook 外接程序的 JavaScript API](http://msdn.microsoft.com/library/e0c9e3d0-517e-4333-b8bd-e169c51a07f6.aspx)。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-p135">Other than resource usage rules, developers for Outlook add-ins should also make sure their add-ins observe limits for specifying activation rules and using the JavaScript API. For more information, see [Limits for activation and JavaScript API for Outlook add-ins](http://msdn.microsoft.com/library/e0c9e3d0-517e-4333-b8bd-e169c51a07f6.aspx).</span></span>

## <a name="it-administrators-control"></a><span data-ttu-id="e7e4b-265">IT 管理员控制</span><span class="sxs-lookup"><span data-stu-id="e7e4b-265">IT administrators' control</span></span>

<span data-ttu-id="e7e4b-266">在企业设置中，对于启用或禁用对 AppSource 和任何专用目录的访问权限，IT 管理员拥有最高权限。</span><span class="sxs-lookup"><span data-stu-id="e7e4b-266">In a corporate setting, IT administrators have ultimate authority over enabling or disabling access to AppSource and any private catalogs.</span></span>

## <a name="see-also"></a><span data-ttu-id="e7e4b-267">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e7e4b-267">See also</span></span>

- [<span data-ttu-id="e7e4b-268">在内容和任务窗格加载项中请求获取 API 使用权限</span><span class="sxs-lookup"><span data-stu-id="e7e4b-268">Requesting permissions for API use in content and task pane add-ins</span></span>](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd.aspx)
- [<span data-ttu-id="e7e4b-269">Outlook 外接程序的隐私、权限和安全性</span><span class="sxs-lookup"><span data-stu-id="e7e4b-269">Privacy, permissions, and security for Outlook add-ins</span></span>](http://msdn.microsoft.com/library/44208fc4-05d4-42d8-ab20-faa89624de1c.aspx)
- [<span data-ttu-id="e7e4b-270">了解 Outlook 外接程序权限</span><span class="sxs-lookup"><span data-stu-id="e7e4b-270">Understanding Outlook add-in permissions</span></span>](https://docs.microsoft.com/en-us/outlook/add-ins/understanding-outlook-add-in-permissions)
- [<span data-ttu-id="e7e4b-271">Outlook 外接程序的激活和 JavaScript API 限制</span><span class="sxs-lookup"><span data-stu-id="e7e4b-271">Limits for activation and JavaScript API for Outlook add-ins</span></span>](http://msdn.microsoft.com/library/e0c9e3d0-517e-4333-b8bd-e169c51a07f6.aspx)
- [<span data-ttu-id="e7e4b-272">解决 Office 外接程序中的同源策略限制</span><span class="sxs-lookup"><span data-stu-id="e7e4b-272">Addressing same-origin policy limitations in Office Add-ins</span></span>](http://msdn.microsoft.com/library/36c800ae-1dda-4ea8-a558-37c89ffb161b.aspx)
- [<span data-ttu-id="e7e4b-273">同源策略</span><span class="sxs-lookup"><span data-stu-id="e7e4b-273">Same Origin Policy</span></span>](http://www.w3.org/Security/wiki/Same_Origin_Policy)
- [<span data-ttu-id="e7e4b-274">同源策略第 1 部分：不准偷看</span><span class="sxs-lookup"><span data-stu-id="e7e4b-274">Same Origin Policy Part 1: No Peeking</span></span>](http://blogs.msdn.com/b/ieinternals/archive/2009/08/28/explaining-same-origin-policy-part-1-deny-read.aspx)
- [<span data-ttu-id="e7e4b-275">针对 JavaScript 的同源策略</span><span class="sxs-lookup"><span data-stu-id="e7e4b-275">Same origin policy for JavaScript</span></span>](https://developer.mozilla.org/En/Same_origin_policy_for_JavaScript)
- [<span data-ttu-id="e7e4b-276">IE 保护模式</span><span class="sxs-lookup"><span data-stu-id="e7e4b-276">IE Protect Mode</span></span>](https://support.microsoft.com/en-us/help/2761180/apps-for-office-don-t-start-if-you-disable-protected-mode-for-the-restricted-sites-zone-in-internet-explorer)
