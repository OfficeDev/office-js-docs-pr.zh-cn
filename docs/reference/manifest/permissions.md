---
title: 清单文件中的 Permissions 元素
description: Permissions 元素指定加载项的 API Office级别。
ms.date: 06/26/2020
localization_priority: Normal
ms.openlocfilehash: 2f2ccb4f6ec691b19cadea76a06520a9bad7a0b6c0e51699f2c8db67a3030de0
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089008"
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
