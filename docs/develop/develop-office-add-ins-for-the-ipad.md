---
title: IPad 上的外接程序的特殊要求
description: 了解有关创建在 iPad 上运行的 Office 外接程序的一些要求。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 25ac5767db3301352e1921411af833957c4644d0
ms.sourcegitcommit: 10463841a977e9b8415362a3ae91b0ae5eebbf89
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/04/2020
ms.locfileid: "47399569"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>IPad 上的外接程序的特殊要求

如果外接程序仅使用 iPad 上支持的 Office Api，则客户可以在 Ipad 上安装它。  (有关详细信息，请参阅[指定 Office 应用程序和 API 要求](specify-office-hosts-and-api-requirements.md)。 ) *如果外接程序将通过[AppSource](https://appsource.microsoft.com)进行营销*，则除了[适用于所有 Office 外接程序的最佳做法](../concepts/add-in-development-best-practices.md)外，还必须遵循可在 ipad 上安装的外接程序。

下表列出了要执行的任务。

> [!NOTE]
> 若要了解如何在 Outlook Mobile 上设计外观良好且工作良好的 Outlook 外接程序，请参阅 [适用于 Outlook mobile 的外接](../outlook/outlook-mobile-addins.md)程序。

|任务|说明|资源|
|:-----|:-----|:-----|
|更新外接程序以支持 Office.js 版本 1.1。|将 Office 外接程序项目中使用的 JavaScript 文件（Office.js 和特定于应用的 .js 文件）和外接程序清单验证文件更新到版本 1.1。|[更新 API 和清单版本](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|应用 iOS 设计最佳实践。|将外接程序 UI 与 iOS 体验无缝集成。| 请参阅下面的注释。 |
|针对触摸优化外接程序。|使 UI 响应触摸输入以及鼠标和键盘。|[应用 UX 设计原则](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|使外接程序免费。|Office on iPad 是一个通道，通过它您可以接触到更多用户并提升您的服务。这些新用户可能成为您的客户。|[认证策略1120。2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|在 iPad 上让外接程序商业可用。|当它在 iPad 上运行时，您的外接程序必须不在应用程序内购买、试用产品和 UI，这是旨在追加到非免费版本的 UI，或者是用户可以在其中购买或获取其他内容、应用程序或外接程序的任何在线商店的链接。您的隐私策略和使用条款页面也必须没有任何 commerce UI 或 AppSource 链接。|[认证策略1100。3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>你的外接程序仍可以在其他平台上使用 commerce。 若要执行此操作，请测试 commerceAllowed 属性并在其返回时禁止显示所有商业的 [内容](/javascript/api/office/office.context#commerceallowed) `false` 。|
|将外接程序提交到 AppSource。|在 "合作伙伴中心" 中的 " **产品安装程序** " 页上，选中 "在 **iOS 和 Android 上提供我的产品 (如果适用) ** " 复选框，并在 "帐户设置" 中提供您的 Apple 开发人员 ID。 请查看 [应用程序提供商协议](https://go.microsoft.com/fwlink/?linkid=715691) ，以确保您了解这些条款。|[将解决方案提交到 AppSource 和 Office 应用商店](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> 你的外接程序可以基于运行它的设备提供备用 UI。 若要检测您的外接程序是否在 iPad 上运行，您可以使用以下 Api。
>
> - var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchenabled)
> - var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed)
>
> 在 iPad 上， `touchEnabled` 返回 `true` 并 `commerceAllowed` 返回 `false` 。
>
> 有关 iPad 最佳 UI 设计实践的信息，请参阅为 [IOS 设计](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)。

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>开发可在 iPad 上运行的 Office 外接程序的最佳做法

适用于开发在 iPad 上运行的外接程序的以下最佳实践。

-  **在 Windows 或 Mac 上开发并调试外接程序，并将其旁加载到 iPad。**

    您不能直接在 iPad 上开发外接程序，但您可以在 Windows 或 Mac 计算机上开发并调试它，并将其旁加载到 iPad 进行测试。 由于在 iOS 或 Mac 上的 Office 中运行的外接程序支持与在 Windows 上运行的加载项相同的 Api，因此外接程序的代码在这些平台上的运行方式相同。 有关详细信息，请参阅在 iPad 和 Mac 上 [测试并调试 Office 外接程序](../testing/test-debug-office-add-ins.md) 和 [旁加载 office 外接程序以进行测试](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)。

-  **在外接程序清单中或通过运行时检查指定 API 要求。**

    当您在加载项清单中指定 API 要求时，Office 将确定 Office 客户端应用程序是否支持这些 API 成员。 如果 API 成员在应用程序中可用，则外接程序将可用。 或者，您可以执行运行时检查，以确定方法在外接程序中使用之前是否在应用程序中可用。 运行时检查可确保您的外接程序在应用程序中始终可用，并在方法可用时提供其他功能。 有关详细信息，请参阅 [指定 Office 应用程序和 API 要求](specify-office-hosts-and-api-requirements.md)。
