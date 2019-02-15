---
ms.date: 01/29/2019
description: 在 Excel 中使用自定义函数对用户进行身份验证。
title: 自定义函数的身份验证
ms.openlocfilehash: 260f15c39758b82a2145474f543c3c9ff5edd132
ms.sourcegitcommit: 70ef38a290c18a1d1a380fd02b263470207a5dc6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/15/2019
ms.locfileid: "30052733"
---
# <a name="authentication"></a>身份验证

在某些情况下, 自定义函数将需要对用户进行身份验证, 以便访问受保护的资源。 虽然自定义函数不需要特定的身份验证方法, 但您应注意, 自定义函数在单独的运行时中从任务窗格和外接程序的其他 UI 元素运行。 因此, 您需要使用`AsyncStorage`对象和对话框 API 在两个运行时之间来回传递数据。
  
## <a name="asyncstorage-object"></a>到 asyncstorage 对象

自定义函数运行时在全局`localStorage`窗口中没有可用的对象, 您通常可能会在其中存储数据。 相反, 应使用[OfficeRuntime](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage)来设置和获取数据, 从而在自定义函数和任务窗格之间共享数据。 

此外, 还提供了使用`AsyncStorage`的好处;它使用安全沙盒环境, 以便其他外接程序无法访问您的数据。  

### <a name="suggested-usage"></a>建议使用

当您需要从任务窗格或自定义函数进行身份验证时, 请检查到 asyncstorage 以查看是否已获取访问令牌。 如果不是, 请使用对话框 API 对用户进行身份验证、检索访问令牌, 然后在到 asyncstorage 中存储令牌以供将来使用。

## <a name="dialog-api"></a>对话框 API

如果令牌不存在, 则应使用对话框 API 要求用户登录。 用户输入凭据后, 生成的访问令牌可以存储在中`AsyncStorage`。

> [!NOTE]
> 自定义函数运行时使用与任务窗格使用的运行时中的对话框对象略有不同的 dialog 对象。 它们都称为 "对话框 API", 但用于`Officeruntime.Dialog`在自定义函数运行时中对用户进行身份验证。

有关如何使用的`OfficeRuntime.Dialog`信息, 请参阅[Custom 函数运行时](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box)。

在整体上构思整个身份验证过程时, 将加载项的任务窗格和 UI 元素以及外接程序的自定义函数部分视为可通过`AsyncStorage`相互通信的单独实体可能会有所帮助。

下图概述了此基本过程。 请注意, 点线表示在执行单独的操作时, 自定义函数和外接程序的任务窗格都是外接程序的整体部分。

1. 您从 Excel 工作簿中的单元格发出自定义函数调用。
2. 自定义函数`Officeruntime.Dialog`用于将您的用户凭据传递到网站。
3. 然后, 此网站将向自定义函数返回访问令牌。
4. 然后, 您的`AsyncStorage`自定义函数会将此访问令牌设置为。
5. 外接程序的任务窗格从`AsyncStorage`访问令牌。

![协同工作的自定义函数、OfficeRuntime 和任务窗格的关系图。](../images/Authdiagram.png "身份验证图。")

## <a name="general-guidance"></a>一般指南

Office 外接程序是基于 web 的, 您可以使用任何 web 身份验证技术。 使用自定义函数实现自己的身份验证时, 必须遵循任何特定的模式或方法。 您可能希望参考有关各种身份验证模式的文档, 从本文开始,[了解如何通过外部服务进行授权](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js)。  

在开发自定义函数时, 应避免使用以下位置来存储数据:  

- `localStorage`: 自定义函数不具有对全局`window`对象的访问权限, 因此无法访问存储在中`localStorage`的数据。
- `Office.context.document.settings`: 此位置不安全, 使用外接程序的任何人都可以提取信息。

## <a name="see-also"></a>另请参阅

* [自定义函数元数据](custom-functions-json.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [Excel 自定义函数教程](excel-tutorial-custom-functions.md)
