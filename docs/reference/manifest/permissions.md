---
title: 清单文件中的 Permissions 元素
description: Permissions 元素指定加载项的 API Office级别。
ms.date: 06/26/2020
localization_priority: Normal
ms.openlocfilehash: bc4cc2713d5a781c3407385470acd762910d17fd
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936967"
---
# <a name="permissions-element"></a>Permissions 元素

指定 Office 外接程序的 API 访问级别；您应基于最少特权的原则请求权限。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

对于内容和任务窗格外接程序：

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

对于邮件外接程序：

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注解

有关更多详细信息，请参阅在内容和任务窗格外接程序中请求[API](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)使用的权限和了解Outlook[外接程序权限](../../outlook/understanding-outlook-add-in-permissions.md)。
