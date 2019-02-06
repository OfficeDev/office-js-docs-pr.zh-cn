---
ms.date: 1/29/2019
description: 在 Excel 中使用自定义函数的用户进行身份验证。
title: 身份验证的自定义的函数
ms.openlocfilehash: 0e42dbc93cb545660a8dbaae5bdb48724f3b7376
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/05/2019
ms.locfileid: "29745403"
---
# <a name="authentication"></a>身份验证

在某些情况下，您自定义的函数将需要对用户进行身份验证才能访问受保护资源。 自定义函数不需要特定的身份验证方法时, 应注意的自定义函数在单独的运行时从运行任务窗格和加载项的其他用户界面元素。 因此，您需要使用两个运行时之间来回传递数据`AsyncStorage`对象和对话框 API。
  
## <a name="asyncstorage-object"></a>AsyncStorage 对象

自定义函数运行时没有`localStorage`上全局窗口，其中可能通常存储数据的可用对象。 相反，您应之间共享数据自定义函数和任务窗格中，通过使用[OfficeRuntime.AsyncStorage](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage)设置和获取数据。 

此外，还会使用好处`AsyncStorage`;它使用安全沙盒环境，以便其他加载项无法访问您的数据。  

### <a name="suggested-usage"></a>建议使用情况

当您需要进行身份验证从任务窗格或自定义的函数时，检查 AsyncStorage 以查看是否已被收购访问令牌。 如果没有，请使用对话框 API 来验证用户身份和检索访问令牌，然后将该令牌存储在 AsyncStorage 以供将来使用。

## <a name="dialog-api"></a>对话框 API

如果令牌不存在，您应使用对话框 API 要求用户登录。 生成访问令牌用户输入凭据之后，可以存储在`AsyncStorage`。

> [!NOTE]
> 自定义函数运行时使用此对话框对象中运行时使用的任务窗格略有不同 Dialog 对象。 它们同时称为"对话框 API"，但使用`Officeruntime.Dialog`中的自定义函数的运行时的用户进行身份验证。

有关如何使用`OfficeRuntime.Dialog`，请参阅[自定义函数的运行时](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box)。

时构想作为一个整体整个身份验证过程，它可能需要考虑的任务窗格和加载项的 UI 元素和自定义为单独的实体，其中可以与通过每个其他通信功能的加载项部分`AsyncStorage`。

下图概述了此基本过程。 请注意，虚线表示他们执行单独操作，自定义的函数和外接程序的任务窗格的加载项作为一个整体两个部分。

1. 问题自定义的函数调用从 Excel 工作簿中的单元格。
2. 自定义的函数使用`Officeruntime.Dialog`可以将您的用户凭据传递到网站。
3. 然后，该网站将访问令牌返回到自定义的函数。
4. 然后，您自定义的函数将此访问令牌设置为`AsyncStorage`。
5. 外接程序的任务窗格访问来自令牌`AsyncStorage`。

![的自定义的函数、 OfficeRuntime 和协作的任务窗格的图表。](../images/Authdiagram.png "身份验证关系图。")

## <a name="general-guidance"></a>一般指导

Office 加载项是基于 web 的您可以使用任何 web 身份验证方法。 没有特定模式或方法必须遵循实现自己的身份验证与自定义函数。 您可能希望查阅有关各种身份验证模式，文档开头[有关授权通过外部服务这篇文章](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js)。  

避免使用的以下位置来开发自定义的函数时存储数据：  

- `localStorage`： 自定义函数不具有对全局访问`window`对象，并因此均没有访问权数据存储在`localStorage`。
- `Office.context.document.settings`： 此位置不安全，通过使用外接程序的任何人都可以提取信息。

## <a name="see-also"></a>另请参阅

* [自定义函数元数据](custom-functions-json.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [Excel 自定义函数教程](excel-tutorial-custom-functions.md)
