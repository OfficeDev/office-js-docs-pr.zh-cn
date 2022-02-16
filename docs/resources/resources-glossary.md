---
title: Office加载项术语表
description: 整个加载项文档中常用的术语Office术语表。
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0c83f056f4eea9c8750bbf4c2d47a2888af96ec2
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855726"
---
# <a name="office-add-ins-glossary"></a>Office加载项术语表

这是整个加载项文档中常用的术语Office术语表。

## <a name="add-in"></a>加载项

Office外接程序是扩展 Office 应用程序的 Web 应用程序。 这些 Web 应用程序向 Office 应用程序添加新功能，例如引入外部数据、自动处理或将交互式对象嵌入Office文档中。

Office 加载项不同于 VBA、COM 和 VSTO 加载项，因为它们提供跨平台支持 (通常为 Web、Windows、Mac 和 iPad) ，并且基于标准 Web 技术 (HTML、CSS 和 JavaScript) 。 加载项的主要编程语言Office JavaScript 或 TypeScript。

## <a name="add-in-commands"></a>外接程序命令

**外接程序命令是** UI 元素（如按钮和菜单）Office外接程序的 UI。 当用户选择外接程序命令元素时，他们将启动操作，如运行 JavaScript 代码或在任务窗格中显示外接程序。 外接程序命令使外接程序的外观和感觉就像外接程序的一Office，从而让用户更放心地信任您的外接程序。 请参阅[外接程序命令了解Excel、](../design/add-in-commands.md)PowerPoint、Word 和外接程序命令Outlook了解更多信息。[](../outlook/add-in-commands-for-outlook.md)

另请参阅功能 [区、功能区按钮](#ribbon-ribbon-button)。

## <a name="application"></a>应用程序

**应用程序** 引用一个Office应用程序。 支持Office外接程序Office包括 Excel、OneNote、Outlook、PowerPoint、Project 和 Word。

另请参阅：[客户端](#client)[、主机](#host)[、Office应用程序、Office客户端](#office-application-office-client)。

## <a name="application-specific-api"></a>特定于应用程序的 API

特定于应用程序的 API 提供强类型对象，这些对象与特定应用程序本地Office交互。 例如，调用 Excel JavaScript API 以访问工作表、区域、表、图表等。 应用程序特定的 API 当前可用于 Excel、OneNote、PowerPoint、Visio 和 Word。 有关详细信息 [，请参阅特定于应用程序的 API](../develop/application-specific-api-model.md) 模型。

另请参阅： [通用 API](#common-api)。

## <a name="client"></a>client

**客户端** 通常引用一个Office应用程序。 支持 Office 外接程序的 Office 应用程序或客户端包括 Excel、OneNote、Outlook、PowerPoint、Project 和 Word。

另请参阅：[application](#application)、[host](#host)[、Office application、Office client](#office-application-office-client)。

## <a name="common-api"></a>通用 API

通用 API 用于访问跨多个应用程序通用的 UI、对话框和客户端Office设置。 此 API 模型使用的是[回调](https://developer.mozilla.org/docs/Glossary/Callback_function)，这样,你在发送给 Office 应用程序的每个请求中只能指定一个操作。

通用 API 是在 2013 Office引入的，用于与 Office 2013 或更高版本进行交互。 一些常见的 API 是 2010 年初的旧 API。 Excel、PowerPoint 和 Word 都具有通用 API 功能，但大部分此功能已被特定于应用程序的 API 模型所替换或取代。 如果可能，特定于应用程序的 API 是首选。

其他常见 API（如与 Outlook、UI 和身份验证相关的通用 API）也是用于这些用途的新式和首选 API。 有关通用 API 对象模型的详细信息，请参阅 [常见 JavaScript API 对象模型](../develop/office-javascript-api-object-model.md)。

另请参阅： [特定于应用程序的 API](#application-specific-api)。

## <a name="content-add-in"></a>内容加载项

**内容外接程序是** 直接嵌入文档、Excel、OneNote或PowerPoint Web 浏览器视图。 用户可以通过内容加载项访问界面控件，运行代码以修改文档或显示数据源中的数据。 在你要将功能直接嵌入文档时，请使用内容加载项。 若要[了解Office，请参阅 Content Office Add-ins](../design/content-add-ins.md)。

另请参阅： [webview](#webview)。

## <a name="content-delivery-network-cdn"></a>内容交付网络 (CDN) 

内容 **传送网络****或CDN** 是服务器和数据中心的分布式网络。 与单个服务器或数据中心相比，它通常提供更高的资源可用性和性能。

## <a name="contoso"></a>Contoso

**Contoso** Ltd. (Contoso 和 Contoso University) 是一家虚构公司，由 Microsoft 用作公司示例和域。

## <a name="custom-function"></a>自定义函数

**自定义函数** 是用户定义的函数，与加载项Excel打包。 自定义函数使开发人员能够添加除典型 Excel 功能之外的新函数，方法为在 JavaScript 中将这些功能定义为外接程序的一部分。 用户Excel中的用户可以访问自定义函数，就像他们访问自定义函数Excel。 有关详细信息[，请参阅在 Excel](../excel/custom-functions-overview.md) 中创建自定义函数。

## <a name="custom-functions-runtime"></a>自定义函数运行时

自定义 **函数运行时** 是仅运行自定义函数的 JavaScript 运行时。 它没有任何 UI，并且无法与 Office.js API 交互。 如果加载项只有自定义函数，这是一个很好的轻型运行时使用。 如果你的自定义函数需要与任务窗格或Office.js API 交互，请配置共享的 JavaScript 运行时。 查看 [将 Office 加载项配置为使用共享 JavaScript 运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md) 以了解更多信息。

另请参阅：[JavaScript 运行时](#javascript-runtime)[、共享 JavaScript 运行时、共享运行时](#shared-javascript-runtime-shared-runtime)。

## <a name="host"></a>host

**主机** 通常是指一个Office应用程序。 支持 Office 外接程序的 Office 应用程序或主机包括 Excel、OneNote、Outlook、PowerPoint、Project 和 Word。

另请参阅[：application](#application)、[client](#client)[、Office application、Office client](#office-application-office-client)。

## <a name="javascript-runtime"></a>JavaScript 运行时

**JavaScript 运行时** 是运行加载项的浏览器主机环境。 在 Office mac Windows 和 Office 上，JavaScript 运行时是嵌入式浏览器控件 (或 webview) 如 Internet Explorer、Edge Legacy、Edge WebView2 或 Safari。 外接程序的不同部分在单独的 JavaScript 运行时中运行。 例如，加载项命令、自定义函数和任务窗格代码通常使用单独的 JavaScript 运行时，除非你配置了共享的 JavaScript 运行时。 有关详细信息[，请参阅Office外接程序](../concepts/browsers-used-by-office-web-add-ins.md)使用的浏览器。

另请参阅：[自定义函数运行时](#custom-functions-runtime)[、共享 JavaScript 运行时、共享运行时](#shared-javascript-runtime-shared-runtime)[、webview](#webview)。

## <a name="office-application-office-client"></a>Office应用程序，Office客户端

**Office客户端** 引用Office应用程序。 支持 Office 外接程序的 Office 应用程序或客户端包括 Excel、OneNote、Outlook、PowerPoint、Project 和 Word。

另请参阅：[应用程序](#application)[、客户端](#client)、[主机](#host)。

## <a name="platform"></a>平台

**平台** 通常指运行应用程序Office操作系统。 支持外接程序Office包括 Windows、Mac、iPad 和 Web 浏览器。

## <a name="quick-start"></a>快速入门

**快速入门** 是特定程序的基本操作所需的关键技能和知识高级说明。 在Office外接程序"文档中，快速入门介绍了如何为特定应用程序（如 Outlook）开发外接程序。 快速入门包含一系列步骤，加载项开发人员可以在大约 5 分钟内完成这些步骤，从而生成正常运行的加载项和功能开发环境。

另请参阅： [教程](#tutorial)。

## <a name="requirement-set"></a>要求集

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

## <a name="ribbon-ribbon-button"></a>功能区，功能区按钮

**功能** 区是一个命令栏，用于将应用程序的功能组织到窗口顶部的一系列选项卡或按钮中。 功能 **区按钮** 是本系列中的按钮之一。 有关详细信息[，请参阅Office](https://support.microsoft.com/office/d946b26e-0c8c-402d-a0f7-c6efa296b527#ID0EBBD=Newer_Versions)或隐藏功能区。

## <a name="runtime"></a>运行时

请参阅： [JavaScript 运行时](#javascript-runtime)。

## <a name="shared-javascript-runtime-shared-runtime"></a>共享 JavaScript 运行时，共享运行时

共享 **JavaScript** 运行时（或共享运行时）允许外接程序中的所有代码（包括任务窗格、外接程序命令和自定义函数）在同一 JavaScript 运行时中运行，即使任务窗格已关闭，也继续运行。 请参阅[将Office外接程序](../develop/configure-your-add-in-to-use-a-shared-runtime.md)配置为使用共享 JavaScript 运行时和 使用技巧 以在 Office 外接程序中使用共享 [JavaScript](https://devblogs.microsoft.com/microsoft365dev/tips-for-using-the-shared-javascript-runtime-in-your-office-add-in%e2%80%af/) 运行时了解更多信息。

另请参阅自定义[函数运行时、](#custom-functions-runtime)[JavaScript 运行时](#javascript-runtime)。

## <a name="task-pane"></a>任务窗格

任务窗格是界面图面或 Web 视图，通常显示在 Excel、Outlook、PowerPoint 和 Word 内窗口的右侧。 任务窗格允许用户访问界面控件，此类控件运行代码以修改文档或电子邮件，或显示数据源中的数据。 当不需要或不能将功能直接嵌入文档时，请使用任务窗格。 有关详细信息[，请参阅Office外接程序](../design/task-pane-add-ins.md)中的任务窗格。

另请参阅： [webview](#webview)。

## <a name="tutorial"></a>教程

**教程是** 一种教学辅助工具，旨在帮助用户学习使用产品或过程。 在Office外接程序上下文中，教程指导外接程序开发人员完成特定应用程序（如 Excel）的完整外接程序开发过程。 这涉及以下 20 个或多个步骤，并且所投入的时间比快速入门 [要大](#quick-start)。

另请参阅： [快速入门](#quick-start)。

## <a name="ui-less-custom-function"></a>无 UI 自定义函数

**无 UI 的自定义函数** 在自定义函数运行时中运行。 它们没有 UI，并且无法与Office.js API 交互。

另请参阅自定义[函数、](#custom-function)[自定义函数运行时](#custom-functions-runtime)。

## <a name="web-add-in"></a>Web 加载项

**Web 外接程序是** 加载项的旧Office术语。 当文档需要区分新式 Microsoft 365 Office 外接程序和其他类型的外接程序（如 VBA、COM 或 VSTO）时，可能会使用此VSTO。

另请参阅 [：外接程序](#add-in)。

## <a name="webview"></a>webview

**Webview** 是一个在应用程序内显示 Web 内容的元素或视图。 内容加载项和任务窗格均包含嵌入式 Web 浏览器，都是加载项中的 web Office示例。

另请参阅 [：内容外接程序](#content-add-in)、 [任务窗格](#task-pane)。

## <a name="xll"></a>XLL

**XLL** 加载项是一种Excel定义函数且文件扩展名为 **.xll 的加载项文件**。 XLL 文件是动态链接库的一 (DLL) ，它只能由 Excel。 XLL 加载项文件必须使用 C 或 C++ 编写。 自定义函数是 XLL 用户定义函数的新式等效函数。 自定义函数跨平台提供支持，并且向后兼容 XLL 文件。 有关详细信息 [，请参阅使用 XLL 用户定义函数](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf) 扩展自定义函数。

另请参阅： [自定义函数](#custom-function)。

## <a name="yeoman-generator-yo-office"></a>Yeoman generator， yo office

适用于[加载项Office Yeoman](https://github.com/OfficeDev/generator-office) 生成器使用开源 [Yeoman](https://github.com/yeoman/yo) 工具通过命令行Office加载项生成加载项。 `yo office`是运行适用于加载项的 Yeoman Office的命令。加载项Office快速入门和教程使用 Yeoman 生成器。

## <a name="see-also"></a>另请参阅

- [Office 加载项其他资源](resources-links-help.md)