---
title: 在iPad 中加载项的特殊要求
description: 了解创建在 Office 上运行的加载项的一iPad。
ms.date: 09/03/2020
ms.localizationpriority: medium
ms.openlocfilehash: 8a114c5fc4a17ee3f7282321d82ad1faa60d9d71
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148863"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>在iPad 中加载项的特殊要求

如果加载项仅使用Office受支持的 api，iPad可以在 iPad 上安装它。  (有关详细信息，请参阅指定 Office 应用程序和 [API](specify-office-hosts-and-api-requirements.md)要求。) 如果外接程序将通过 *[AppSource](https://appsource.microsoft.com)* 进行销售，则除了适用于所有 [Office](../concepts/add-in-development-best-practices.md)外接程序的最佳实践之外，还必须遵循一些适用于可安装在 iPad 上的外接程序的做法。

下表列出了要执行的任务。

> [!NOTE]
> 有关设计外观Outlook良好且在 Outlook Mobile 上良好工作的外接程序的信息，请参阅[Add-ins for Outlook Mobile](../outlook/outlook-mobile-addins.md)。

|任务|说明|资源|
|:-----|:-----|:-----|
|更新外接程序以支持 Office.js 版本 1.1。|将 Office 外接程序项目中使用的 JavaScript 文件（Office.js 和特定于应用的 .js 文件）和外接程序清单验证文件更新到版本 1.1。|[更新 API 和清单版本](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|应用 iOS 设计最佳实践。|将外接程序 UI 与 iOS 体验无缝集成。| 请参阅下面的注释。 |
|针对触摸优化外接程序。|使 UI 响应触摸输入以及鼠标和键盘。|[应用 UX 设计原则](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|使外接程序免费。|Office on iPad 是一个通道，通过它您可以接触到更多用户并提升您的服务。这些新用户可能成为您的客户。|[认证策略 1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|使加载项在商业上免费iPad。|当外接程序在 iPad 上运行时，外接程序不得包含应用内购买、试用产品/服务、旨在向上销售非免费版本的 UI，或指向任何在线商店（用户可以购买或获取其他内容、应用或外接程序）的链接。隐私策略和使用条款页面还必须没有任何商务 UI 或 AppSource 链接。|[认证策略 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>您的外接程序仍然可以在其他平台上进行商务。 为此，请测试[Office.context.commerceAllowed](/javascript/api/office/office.context#commerceAllowed)属性，在返回 时禁止所有商务 `false` 。|
|将加载项提交到 AppSource。|在合作伙伴中心的"产品设置"页面上，选中"在 iOS 和 Android (（如果适用) ）上提供我的产品 **"复选框，** 在"帐户设置"中提供 Apple 开发人员 ID。 查看 [应用程序提供商协议](https://go.microsoft.com/fwlink/?linkid=715691) 以确保您了解这些条款。|[将解决方案提交到 AppSource 和 Office 应用商店](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> 加载项可以基于运行它的设备提供备用 UI。 若要检测加载项是否在加载项上运行iPad，可以使用以下 API。
>
> - var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchEnabled)
> - var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceAllowed)
>
> 在iPad上， `touchEnabled` 返回 `true` 和 `commerceAllowed` 返回 `false` 。
>
> 有关适用于 iOS 的最佳 UI 设计iPad，请参阅[Dinging for iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)。

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>开发可运行Office加载项的最佳实践iPad

应用以下最佳实践，以开发在 iPad。

-  **在 mac 或 Windows 上开发和调试外接程序，并旁加载iPad。**

    无法直接在 iPad 上开发外接程序，但可以在 Windows 或 Mac 计算机上开发和调试它，并旁加载它iPad进行测试。 由于在 iOS 或 Mac 上的 Office 中运行的外接程序支持与在 Windows 上的 Office 中运行的外接程序相同的 API，因此外接程序的代码应在这些平台上以相同方式运行。 有关详细信息，[请参阅在](../testing/test-debug-office-add-ins.md)Office 和 Mac 上测试和调试加载项Office加载项iPad[和旁加载以进行测试](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)。

-  **在外接程序清单中或通过运行时检查指定 API 要求。**

    在加载项清单中指定 API 要求时，Office确定加载项Office是否支持这些 API 成员。 如果 API 成员在应用程序中可用，则你的外接程序将可用。 或者，您可以执行运行时检查，以确定方法在外接程序中之前在应用程序中是否可用。 运行时检查可确保外接程序在应用程序中始终可用，并提供其他功能（如果方法可用）。 有关详细信息，请参阅指定Office[和 API 要求](specify-office-hosts-and-api-requirements.md)。
