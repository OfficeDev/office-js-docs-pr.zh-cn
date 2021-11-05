---
title: 共享运行时要求集
description: 指定支持 SharedRuntime Office的平台和应用程序。
ms.date: 11/03/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: a5f7d3c9394de047b358d7f190c5adae5b5199b1
ms.sourcegitcommit: 210251da940964b9eb28f1071977ea1fe80271b4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/05/2021
ms.locfileid: "60793601"
---
# <a name="shared-runtime-requirement-sets"></a>共享运行时要求集

要求集是指已命名的 API 成员组。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

运行 JavaScript 代码的 Office 外接程序的某些部分（如任务窗格、从外接程序命令启动的函数文件和 Excel 自定义函数）可以共享单个 JavaScript 运行时。 这允许所有部件共享一组全局变量、共享一组加载的库以及相互通信，而无需通过持久存储传递邮件。 有关详细信息，请参阅[将Office加载项配置为使用共享的 JavaScript 运行时](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

下表列出了 SharedRuntime 1.1 要求集、支持该要求集的 Office 客户端应用程序，以及 Office 应用程序的版本或版本号。

| 要求集 | Office 2021 年 1 月或Windows<br>（一次性购买） | Windows 版 Office<br>（关联至 Microsoft 365 订阅） | iPad 版 Office<br>（关联至 Microsoft 365 订阅） | Mac 版 Office<br>（关联至 Microsoft 365 订阅） | Office 网页版 | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1.1  | 内部版本 16.0.14326.20454 或更高版本 | 版本 2002 (内部版本 12527.20092) 或更高版本 | 不适用 | 16.35 或更高版本 | 2020 年 2 月 | 不适用 |

> [!IMPORTANT]
> 目前，iPad 或一次性购买版本的 Office 2019 或更早版本不支持共享 JavaScript 运行时。 有关其他支持详细信息，请参阅以下部分。

## <a name="support-for-version-11-on-excel"></a>支持版本 1.1 Excel

SharedRuntime 1.1 要求集针对 Excel web 版、Windows 和 Mac 发布。

## <a name="preview-support-for-version-11-on-word-and-powerpoint"></a>预览 Word 和 Word 版本 1.1 PowerPoint

下表列出了支持共享 JavaScript 运行时预览的附加应用程序版本。 共享运行时的预览版本可能会更改。 不支持在生产环境中使用。 要获取最新版本，你需要[加入 Office 预览体验计划](https://insider.office.com/join)。 试用预览版功能的好方法是使用 Microsoft 365 订阅。 如果还没有 Microsoft 365 订阅，可以通过加入[Microsoft 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获取一个订阅。

|Office 应用程序 |内部版本 |
|-------------------|------|
|Windows 版 PowerPoint |内部版本 16.0.13218.10000 或更高版本 |
|Windows 版 Word |内部版本 16.0.13218.10000 或更高版本 |
|Mac 版 Word |内部版本 16.46.207.0 或更高版本 |

## <a name="office-versions-and-build-numbers"></a>Office 版本和内部版本号

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="see-also"></a>另请参阅

- [将 Office 加载项配置为使用共享 JavaScript 运行时](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
