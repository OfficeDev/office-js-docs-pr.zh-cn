---
title: 开发适用于 iPad 的 Office 外接程序
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 6fca7948c97f0a12f46742846ed9faca4179f362
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449804"
---
# <a name="develop-office-add-ins-for-the-ipad"></a>开发适用于 iPad 的 Office 加载项


下表列出了开发 Office 外接程序要执行的任务，以使其能够在 Office for iPad 中运行。


|**任务**|**描述**|**资源**|
|:-----|:-----|:-----|
|更新外接程序以支持 Office.js 版本 1.1。|将 Office 外接程序项目中使用的 JavaScript 文件（Office.js 和特定于应用的 .js 文件）和外接程序清单验证文件更新到版本 1.1。|[JavaScript API 中的更改内容](/office/dev/add-ins/reference/what's-changed-in-the-javascript-api-for-office)|
|应用 UI 设计最佳实践。|将外接程序 UI 与 iOS 体验无缝集成。|[针对 iOS 进行设计](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|应用外接程序设计最佳实践。|确保外接程序提供明确值、正常运行并持续执行。|[开发 Office 外接程序的最佳做法](../concepts/add-in-development-best-practices.md)|
|针对触摸优化外接程序。|使 UI 响应触摸输入以及鼠标和键盘。|[应用 UX 设计原则](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|使外接程序免费。|Office on iPad 是一个通道，通过它您可以接触到更多用户并提升您的服务。这些新用户可能成为您的客户。|[验证策略 10.8](/office/dev/store/validation-policies#10-apps-and-add-ins-utilize-supported-capabilities)|
|确保加载项无商业内容。|加载项不得包括应用内购买、试用优惠、追加销售付费加载项的 UI 或任何在线商店（方便用户购买或获取其他内容、应用或加载项）链接。隐私策略和使用条款页面也不得包含任何商业 UI 或 AppSource 链接。|[验证策略 3.4](/office/dev/store/validation-policies#3-apps-and-add-ins-can-sell-additional-features-or-content-through-purchases-within-the-app-or-add-in)|
|将加载项重新提交到 AppSource。|在卖家面板中，选中“将此加载项添加到 iPad 上的 Office 加载项目录中”**** 复选框，并在“Apple ID”框中输入 Apple 开发人员 ID。请查看[应用提供商协议](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.htm)，以确保了解协议。|[将解决方案提交到 AppSource 和 Office 应用商店](/office/dev/store/submit-to-the-office-store)|

对于正在其他平台上运行的 Office 应用程序，您的外接程序可以保持原样。您还可以基于您的外接程序所运行的浏览器/设备提供不同的 UI 服务。若要检测您的外接程序是否正在 iPad 上运行，您可以使用以下 API：
- var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchenabled)
- var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed)


## <a name="best-practices-for-developing-office-add-ins-for-ios-and-mac"></a>开发适用于 iOS 和 Mac 的 Office 外接程序的最佳实践

应用以下用于开发在 iOS 上运行的外接程序的最佳实践：


-  **使用 Visual Studio 开发外接程序。**

    如果使用 Visual Studio 开发外接程序，则在 iPad 或 Mac 上旁加载外接程序前，可以在 Windows 上运行的 Office 主机应用程序中 [设置断点并调试其代码](../develop/create-and-debug-office-add-ins-in-visual-studio.md)。因为在 Office for iOS 或 Office for Mac 中运行的外接程序支持在 Office for Windows 中作为外接程序运行的同一 API，所以外接程序的代码在这两种平台上的运行方式应当是相同的。

-  **在外接程序清单中或通过运行时检查指定 API 要求。**

    在外接程序清单中指定 API 要求时，Office 将确定主机应用程序是否支持这些 API 成员。如果 API 成员在主机中可用，则外接程序在该主机应用程序中可用。或者，在外接程序中使用某方法前，可以执行运行时检查以确定该方法是否在主机中可用。运行时检查确保外接程序始终在主机中可用，并在方法可用时提供其他功能。有关详细信息，请参阅 [指定 Office 主机和 API 要求](specify-office-hosts-and-api-requirements.md)。

有关常规的加载项开发最佳做法，请参阅 [Office 加载项开发最佳做法](../concepts/add-in-development-best-practices.md)。


## <a name="see-also"></a>另请参阅

- [将 Office 外接程序旁加载到 iPad 和 Mac 上](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)  
- [在 iPad 和 Mac 上调试 Office 加载项](../testing/debug-office-add-ins-on-ipad-and-mac.md)
