---
title: Office 加载项术语表
description: 常在 Office 加载项文档中使用的术语表。
ms.date: 09/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: ef8df6e344698f7d67ebe7afe1759e13630b385d
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/30/2022
ms.locfileid: "68234912"
---
# <a name="office-add-ins-glossary"></a>Office 加载项术语表

这是常在 Office 加载项文档中使用的术语术语表。

## <a name="add-in"></a>加载项

Office 加载项是扩展 Office 应用程序的 Web 应用程序。 这些 Web 应用程序向 Office 应用程序添加新功能，例如引入外部数据、自动化进程或在 Office 文档中嵌入交互式对象。

Office 加载项不同于 VBA、COM 和 VSTO 外接程序，因为它们提供跨平台支持 (通常是 Web、Windows、Mac 和 iPad) ，并且基于标准 Web 技术 (HTML、CSS 和 JavaScript) 。 Office 外接程序的主要编程语言是 JavaScript 或 TypeScript。

## <a name="add-in-commands"></a>加载项命令

**外接程序命令** 是 UI 元素，例如按钮和菜单，可扩展外接程序的 Office UI。 当用户选择外接程序命令元素时，他们启动操作，例如运行 JavaScript 代码或在任务窗格中显示外接程序。 外接程序命令使您的外接程序看起来和感觉都像是 Office 的一部分，这使用户对您的外接程序更有信心。 有关详细信息，请参阅 [适用于 Outlook 的 Excel、PowerPoint、Word](../design/add-in-commands.md) 和 [加载项命令的加载项命令](../outlook/add-in-commands-for-outlook.md) 。

另请参阅： [功能区、功能区按钮](#ribbon-ribbon-button)。

## <a name="application"></a>应用程序

**应用程序** 是指 Office 应用程序。 支持 Office 加载项的 Office 应用程序包括 Excel、OneNote、Outlook、PowerPoint、Project 和 Word。

另请参阅： [客户端](#client)、 [主机](#host)、 [Office 应用程序、Office 客户端](#office-application-office-client)。

## <a name="application-specific-api"></a>特定于应用程序的 API

特定于应用程序的 API 提供与特定 Office 应用程序原生对象交互的强类型对象。 例如，调用 Excel JavaScript API 以访问工作表、范围、表、图表等。 特定于应用程序的 API 目前适用于 Excel、OneNote、PowerPoint、Visio 和 Word。 有关详细信息，请参阅 [特定于应用程序的 API 模型](../develop/application-specific-api-model.md) 。

另请参阅： [通用 API](#common-api)。

## <a name="client"></a>客户

**客户端** 通常引用 Office 应用程序。 支持 Office 外接程序的 Office 应用程序或客户端为 Excel、OneNote、Outlook、PowerPoint、Project 和 Word。

另请参阅： [应用程序](#application)、 [主机](#host)、 [Office 应用程序、Office 客户端](#office-application-office-client)。

## <a name="common-api"></a>通用 API

常用 API 用于访问多个 Office 应用程序中常见的 UI、对话框和客户端设置等功能。 此 API 模型使用的是[回调](https://developer.mozilla.org/docs/Glossary/Callback_function)，这样,你在发送给 Office 应用程序的每个请求中只能指定一个操作。

Office 2013 引入了常见 API，用于与 Office 2013 或更高版本交互。 某些常见 API 是 2010 年代初期的旧 API。 Excel、PowerPoint 和 Word 都具有通用 API 功能，但大部分此功能已被特定于应用程序的 API 模型取代或取代。 如果可能，首选特定于应用程序的 API。

其他常见 API（例如与 Outlook、UI 和身份验证相关的常见 API）是新式的首选 API。 有关 Common API 对象模型的详细信息，请参阅 [Common JavaScript API 对象模型](../develop/office-javascript-api-object-model.md)。

另请参阅： [特定于应用程序的 API](#application-specific-api)。

## <a name="content-add-in"></a>内容加载项

**内容加载项** 是直接嵌入到 Excel、OneNote 或 PowerPoint 文档中的 Web 视图或 Web 浏览器视图。 用户可以通过内容加载项访问界面控件，运行代码以修改文档或显示数据源中的数据。 在你要将功能直接嵌入文档时，请使用内容加载项。 有关详细信息，请参阅 [内容 Office 加载项](../design/content-add-ins.md) 。

另请参阅： [Webview](#webview)。

## <a name="content-delivery-network-cdn"></a>内容分发网络 (CDN) 

**内容分发网络** 或 **CDN** 是服务器和数据中心的分布式网络。 与单个服务器或数据中心相比，它通常提供更高的资源可用性和性能。

## <a name="contoso"></a>Contoso

**Contoso** Ltd. (又称 Contoso 和 Contoso university) 是 Microsoft 用作示例公司和域的虚构公司。

## <a name="custom-function"></a>自定义函数

**自定义函数** 是使用 Excel 加载项打包的用户定义函数。 通过将 JavaScript 中的这些函数定义为加载项的一部分，自定义函数使开发人员能够添加除典型 Excel 功能之外的新函数。 Excel 中的用户可以访问自定义函数，就像在 Excel 中访问任何本机函数一样。 有关详细信息，请参阅 [Excel 中的创建自定义函](../excel/custom-functions-overview.md) 数。

## <a name="custom-functions-runtime"></a>自定义函数运行时

**自定义函数运行时** 是 [仅限 JavaScript 的运行时](../testing/runtimes.md#javascript-only-runtime)，可在 Office 主机和平台的某些组合上运行自定义函数。 它没有 UI，无法与Office.js API 交互。 如果外接程序仅具有自定义函数，则这是一个不错的轻型运行时。 如果自定义函数需要与任务窗格或Office.js API 交互，请配置 [共享运行时](../testing/runtimes.md#shared-runtime)。 请参阅 [“配置 Office 外接程序”以使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md) 了解详细信息。

另请参阅： [运行时](#runtime)、 [共享运行时](#shared-runtime)。

## <a name="custom-functions-only-add-in"></a>仅自定义函数加载项

包含自定义函数但不包含任务窗格等 UI 的加载项。 此类外接程序中的自定义函数在 [仅限 JavaScript 的运行时中运行](../testing/runtimes.md#javascript-only-runtime)。 包含 UI 的自定义函数可以使用共享运行时或仅限 JavaScript 的运行时和支持 HTML 的运行时的组合。 我们建议，如果有 UI，则使用共享运行时。

另请参阅： [自定义函数](#custom-function)、 [自定义函数运行时](#custom-functions-runtime)。

## <a name="host"></a>host

**\<Host\>** 通常是指 Office 应用程序。 支持 Office 外接程序的 Office 应用程序或主机包括 Excel、OneNote、Outlook、PowerPoint、Project 和 Word。

另请参阅： [应用程序](#application)、 [客户端](#client)、 [Office 应用程序、Office 客户端](#office-application-office-client)。

## <a name="office-application-office-client"></a>Office 应用程序、Office 客户端

**Office 客户端** 是指 Office 应用程序。 支持 Office 外接程序的 Office 应用程序或客户端为 Excel、OneNote、Outlook、PowerPoint、Project 和 Word。

另请参阅： [应用程序](#application)、 [客户端](#client)、 [主机](#host)。

## <a name="perpetual"></a>永久

**永久性** 是指通过批量许可协议或零售渠道提供的 Office 版本。

其他 Microsoft 内容可能使用 **非订阅** 术语来表示此概念。

另请参阅： [零售、零售永久](#retail-retail-perpetual)、 [批量许可、批量许可永久、批量许可](#volume-licensed-volume-licensed-perpetual-volume-licensing)

## <a name="platform"></a>平台

**平台** 通常是指运行 Office 应用程序的操作系统。 支持 Office 外接程序的平台包括 Windows、Mac、iPad 和 Web 浏览器。

## <a name="quick-start"></a>快速入门

**快速入** 门是对特定程序的基本操作所需的关键技能和知识的高级说明。 在 Office 加载项文档中，快速入门介绍了如何为特定应用程序（如 Outlook）开发加载项。 快速入门包含一系列步骤，外接程序开发人员可以在大约 5 分钟内完成这些步骤，从而生成一个正常运行的加载项和功能开发环境。

另请参阅： [教程](#tutorial)。

## <a name="requirement-set"></a>要求集

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

## <a name="retail-retail-perpetual"></a>零售、零售永久

**零售** 是指通过零售渠道提供的 Office 的永久版本。 这些版本不包括 Microsoft 365 订阅提供的版本或批量许可协议。

其他 Microsoft 内容可以使用 **“一次性购买** ”一词或 **“使用者** ”一词来表示此概念。

另请参阅： [永久](#perpetual)

## <a name="ribbon-ribbon-button"></a>功能区，功能区按钮

**功能区** 是一个命令栏，用于将应用程序的功能组织到窗口顶部的一系列选项卡或按钮中。 **功能区按钮** 是本系列中的其中一个按钮。 有关详细信息，请参阅 [Office 中的显示或隐藏功能区](https://support.microsoft.com/office/d946b26e-0c8c-402d-a0f7-c6efa296b527#ID0EBBD=Newer_Versions) 。

## <a name="runtime"></a>运行

**运行时** 是主机环境 (包括 JavaScript 引擎，通常也是外接程序运行的 HTML 呈现引擎) 。 在 Office on Windows 和 Office on Mac 中，运行时是嵌入式浏览器控件 (或 Web 视图) ，例如 Internet Explorer、Edge Legacy、Edge WebView2 或 Safari。 外接程序的不同部分在单独的运行时中运行。 例如，外接程序命令、自定义函数和任务窗格代码通常使用单独的运行时，除非配置 [共享运行时](../testing/runtimes.md#shared-runtime)。 有关详细信息，请参阅 [Office 加载项使用的 Office 加载项](../testing/runtimes.md) 和 [浏览器中的](../concepts/browsers-used-by-office-web-add-ins.md) 运行时。

另请参阅： [自定义函数运行时](#custom-functions-runtime)、 [共享运行时](#shared-runtime)、 [Webview](#webview)。

## <a name="shared-runtime"></a>共享运行时

**共享运行时** 允许加载项中的所有代码（包括任务窗格、外接程序命令和自定义函数）在同一运行时运行，即使在任务窗格关闭时也继续运行。 有关[在 Office 加载项中使用共享运行时的共享运行时](https://devblogs.microsoft.com/microsoft365dev/tips-for-using-the-shared-javascript-runtime-in-your-office-add-in%e2%80%af/)，请参阅[共享运行时](../testing/runtimes.md#shared-runtime)和提示，了解详细信息。

另请参阅： [自定义函数运行时](#custom-functions-runtime)、 [运行时](#runtime)。

## <a name="subscription"></a>订阅

**订阅** 是指 Microsoft 365 订阅提供的 Office 版本。

## <a name="task-pane"></a>任务窗格

任务窗格是通常显示在 Excel、Outlook、PowerPoint 和 Word 中窗口右侧的接口图面或 Web 视图。 任务窗格允许用户访问界面控件，此类控件运行代码以修改文档或电子邮件，或显示数据源中的数据。 无需或不能将功能直接嵌入到文档中时，请使用任务窗格。 有关详细信息，请参阅 [Office 加载项中的任务窗格](../design/task-pane-add-ins.md) 。

另请参阅： [Webview](#webview)。

## <a name="tutorial"></a>教程

**教程** 是一种教学帮助，旨在帮助人们学习使用产品或过程。 在 Office 外接程序上下文中，教程指导外接程序开发人员完成特定应用程序（如 Excel）的完整外接程序开发过程。 这涉及到以下 20 个或更多步骤，比 [快速入门更投入](#quick-start)时间。

另请参阅： [快速入门](#quick-start)。

## <a name="volume-licensed-volume-licensed-perpetual-volume-licensing"></a>批量许可、批量许可永久、批量许可

**批量许可** 是指通过 Microsoft 与公司之间的批量许可协议提供的 Office 的永久版本。

其他 Microsoft 内容可能使用 **商业** 术语来表示此概念。

另请参阅： [永久](#perpetual)

## <a name="web-add-in"></a>Web 加载项

**Web 加载项** 是 Office 加载项的旧术语。 当 Microsoft 365 文档需要将新式 Office 加载项与其他类型的外接程序（如 VBA、COM 或 VSTO）区分开来时，可以使用此术语。

另请参阅： [加载项](#add-in)。

## <a name="webview"></a>webview

**Web 视图** 是显示应用程序内 Web 内容的元素或视图。 内容加载项和任务窗格都包含嵌入式 Web 浏览器，并且是 Office 外接程序中的 Web 视图示例。

另请参阅： [内容加载项](#content-add-in)、 [任务窗格](#task-pane)。

## <a name="xll"></a>XLL

**XLL** 外接程序是一个 Excel 外接程序文件，它提供用户定义的函数并具有文件扩展 **名 .xll**。 XLL 文件是一种动态链接库 (只能由 Excel 打开的 DLL) 文件。 XLL 加载项文件必须以 C 或 C++编写。 自定义函数是 XLL 用户定义函数的新式等效函数。 自定义函数提供跨平台的支持，并与 XLL 文件向后兼容。 有关详细信息，请参阅 [使用 XLL 用户定义函数扩展自定义函](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf) 数。

另请参阅： [自定义函数](#custom-function)。

## <a name="yeoman-generator-yo-office"></a>Yeoman 生成器，yo office

[Office 外接程序的 Yeoman 生成器](../develop/yeoman-generator-overview.md)使用 开放源代码 [Yeoman](https://github.com/yeoman/yo) 工具通过命令行生成 Office 加载项。 `yo office` 是运行 Office 加载项的 Yeoman 生成器的命令。Office 加载项快速入门和教程使用 Yeoman 生成器。

## <a name="see-also"></a>另请参阅

- [Office 加载项其他资源](resources-links-help.md)