---
title: 在iPad 中加载项的特殊要求
description: 了解创建在 iPad 上运行的 Office 加载项的一些要求。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: fdb402f4302e7e81589d586fa1ecd5b30d4e515d
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237852"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>在iPad 中加载项的特殊要求

如果加载项仅使用 iPad 上支持的 Office API，则客户可以在 iPad 上安装它。  (有关详细信息，请参阅"指定 Office 应用程序和 [API](specify-office-hosts-and-api-requirements.md)要求"。) 如果外接程序将通过 *[AppSource](https://appsource.microsoft.com)* 进行销售，则除了适用于所有 [Office](../concepts/add-in-development-best-practices.md)加载项的最佳实践外，还必须遵循一些适用于 iPad 的外接程序的做法。

下表列出了要执行的任务。

> [!NOTE]
> 有关设计在 Outlook Mobile 上外观良好且效果良好的 Outlook 外接程序的信息，请参阅 [Outlook Mobile 的外接程序](../outlook/outlook-mobile-addins.md)。

|任务|说明|资源|
|:-----|:-----|:-----|
|更新外接程序以支持 Office.js 版本 1.1。|将 Office 外接程序项目中使用的 JavaScript 文件（Office.js 和特定于应用的 .js 文件）和外接程序清单验证文件更新到版本 1.1。|[更新 API 和清单版本](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|应用 iOS 设计最佳实践。|将外接程序 UI 与 iOS 体验无缝集成。| 请参阅下面的注释。 |
|针对触摸优化外接程序。|使 UI 响应触摸输入以及鼠标和键盘。|[应用 UX 设计原则](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|使外接程序免费。|Office on iPad 是一个通道，通过它您可以接触到更多用户并提升您的服务。这些新用户可能成为您的客户。|[认证策略 1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|在 iPad 上使加载项商业免费。|当加载项在 iPad 上运行时，加载项不得包含应用内购买、试用产品/服务、旨在向上销售到非免费版本的 UI，或指向用户可以购买或获取其他内容、应用或外接程序的任何联机商店的链接。隐私策略和使用条款页面还必须没有任何商业 UI 或 AppSource 链接。|[认证策略 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>加载项仍可在其他平台上进行商务。 为此，请测试 [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed) 属性，并禁止返回时的所有商务 `false` 。|
|将加载项提交到 AppSource。|在合作伙伴中心的"产品设置"页上，选中"在 **iOS** 和 Android (上提供我的产品（如果适用) ）"复选框，在"帐户"设置中提供 Apple 开发人员 ID。 查看 [应用程序提供商协议](https://go.microsoft.com/fwlink/?linkid=715691) 以确保您了解这些条款。|[将解决方案提交到 AppSource 和 Office 应用商店](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> 加载项可以基于运行它的设备提供备用 UI。 若要检测加载项是否在 iPad 上运行，可以使用以下 API。
>
> - var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchenabled)
> - var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed)
>
> 在 iPad 上， `touchEnabled` 返回 `true` 和 `commerceAllowed` 返回 `false` 。
>
> 有关适用于 iPad 的最佳 UI 设计实践的信息，请参阅"针对[iOS 进行设计"。](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>开发可在 iPad 上运行的 Office 加载项的最佳实践

应用以下开发在 iPad 上运行的加载项的最佳实践。

-  **在 Windows 或 Mac 上开发和调试加载项，并旁加载到 iPad。**

    无法直接在 iPad 上开发外接程序，但可以在 Windows 或 Mac 计算机上开发和调试它，并旁加载它到 iPad 进行测试。 由于在 iOS 或 Mac 上的 Office 中运行的外接程序支持与在 Windows 上的 Office 中运行的外接程序相同的 API，因此加载项代码应在这些平台上以相同方式运行。 有关详细信息，请参阅 [测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md) ，以及将 Office 加载项旁加载在 iPad 和 Mac [上进行测试](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)。

-  **在外接程序清单中或通过运行时检查指定 API 要求。**

    在加载项清单中指定 API 要求时，Office 将确定 Office 客户端应用程序是否支持这些 API 成员。 如果 API 成员在应用程序中可用，则你的外接程序将可用。 或者，您可以执行运行时检查，以确定在加载项中使用方法之前，该方法在应用程序中是否可用。 运行时检查可确保外接程序在应用程序中始终可用，并且如果方法可用，则提供其他功能。 有关详细信息，请参阅"[指定 Office 应用程序和 API 要求"。](specify-office-hosts-and-api-requirements.md)
