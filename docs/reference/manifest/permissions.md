---
title: 清单文件中的 Permissions 元素
description: 权限元素指定 Office 外接程序的 API 访问级别。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 603494b61ef126b35cb5cdff8c5f5b911bd25840
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611489"
---
# <a name="permissions-element"></a>Permissions 元素

指定 Office 外接程序的 API 访问级别；您应基于最少特权的原则请求权限。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

对于内容和任务窗格外接程序：

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

对于邮件外接程序

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注解

有关详细信息，请参阅在[外接程序中请求 API 的使用权限](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)和[了解 Outlook 外接程序权限](../../outlook/understanding-outlook-add-in-permissions.md)。
