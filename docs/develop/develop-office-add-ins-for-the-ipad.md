---
title: 在iPad 中加载项的特殊要求
description: 了解创建在 iPad 上运行的 Office 加载项的一些要求。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: cc75cc75daec756efcb066f3e3a77f865672e501
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889301"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>在iPad 中加载项的特殊要求

如果加载项仅使用 iPad 上支持的 Office API，则客户可以在 iPad 上安装它。  (请参阅有关详细信息 [的指定 Office 应用程序和 API 要求](specify-office-hosts-and-api-requirements.md)。) 如果外接程序 *将通过 [AppSource](https://appsource.microsoft.com) 进行营销*，则除了 [适用于所有 Office 外接程序的最佳做法](../concepts/add-in-development-best-practices.md)外，还必须遵循一些可安装在 iPad 上的外接程序的做法。

下表列出了要执行的任务。

> [!NOTE]
> 有关在 Outlook Mobile 上设计外观良好且运行良好的 Outlook 加载项的信息，请参阅 [Outlook Mobile 的加载项](../outlook/outlook-mobile-addins.md)。

|任务|说明|资源|
|:-----|:-----|:-----|
|更新外接程序以支持 Office.js 版本 1.1。|将 Office 外接程序项目中使用的 JavaScript 文件（Office.js 和特定于应用的 .js 文件）和外接程序清单验证文件更新到版本 1.1。|[更新 API 和清单版本](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|应用 iOS 设计最佳做法。|将外接程序 UI 与 iOS 体验无缝集成。| 请参阅下面的说明。 |
|针对触摸优化外接程序。|使 UI 响应触摸输入以及鼠标和键盘。|[应用 UX 设计原则](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|使外接程序免费。|Office on iPad 是一个通道，通过它您可以接触到更多用户并提升您的服务。这些新用户可能成为您的客户。|[认证策略 1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|在 iPad 上免费提供外接程序商务版。|在 iPad 上运行时，您的外接程序必须不提供应用内购买、试用产品/服务、旨在将产品升级到非免费版本的 UI，或指向用户可以购买或获取其他内容、应用或外接程序的任何在线商店的链接。隐私策略和使用条款页面还必须不包含任何商业 UI 或 AppSource 链接。|[认证策略 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>你的外接程序仍然可以在其他平台上进行商务访问。 为此，请测试 [Office.context.commerceAllowed 属性，](/javascript/api/office/office.context#office-office-context-commerceallowed-member) 并在它返回 `false`时禁止所有商业。|
|将加载项提交到 AppSource。|在合作伙伴中心 **“产品设置”** 页上，选中 **“使我的产品在 iOS 和 Android (上可用（如果适用）)** 复选框，并在”帐户“设置中提供 Apple 开发人员 ID。 查看应用程序 [提供程序协议](https://go.microsoft.com/fwlink/?linkid=715691) ，确保了解条款。|[将解决方案提交到 AppSource 和 Office 应用商店](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> 外接程序可以根据运行它的设备来提供备用 UI。 若要检测加载项是否在 iPad 上运行，可以使用以下 API。
>
> - const isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#office-office-context-touchenabled-member)
> - const allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#office-office-context-commerceallowed-member)
>
> 在 iPad 上， `touchEnabled` 返回 `true` 并 `commerceAllowed` 返回 `false`。
>
> 有关 iPad 的最佳 UI 设计做法的信息，请参阅 [iOS 设计](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)。

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>开发可在 iPad 上运行的 Office 加载项的最佳做法

应用以下最佳做法来开发在 iPad 上运行的加载项。

- **在 Windows 或 Mac 上开发和调试外接程序，并将其旁加载到 iPad。**

    不能直接在 iPad 上开发外接程序，但可以在 Windows 或 Mac 计算机上开发和调试它，并将其旁加载到 iPad 进行测试。 由于在 iOS 或 Mac 上的 Office 中运行的加载项支持与在 Windows 上的 Office 中运行的外接程序相同的 API，因此外接程序的代码应在这些平台上以相同的方式运行。 有关详细信息，请参阅 iPad 上的 [测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md) 和 [旁加载 Office 加载项以进行测试](../testing/sideload-an-office-add-in-on-ipad.md)。

- **在外接程序清单中或通过运行时检查指定 API 要求。**

    在加载项清单中指定 API 要求时，Office 将确定 Office 客户端应用程序是否支持这些 API 成员。 如果 API 成员在应用程序中可用，则外接程序将可用。 或者，在加载项中使用方法之前，可以执行运行时检查以确定应用程序中是否有可用的方法。 运行时检查可确保加载项始终在应用程序中可用，并在方法可用时提供其他功能。 有关详细信息，请参阅 [指定 Office 应用程序和 API 要求](specify-office-hosts-and-api-requirements.md)。
