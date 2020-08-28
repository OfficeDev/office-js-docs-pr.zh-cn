---
title: 清单文件中的 AllowSnapshot 元素
description: 指定是否将内容外接程序的快照图像与主机文档一起保存。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ea910e1ad747e304dbc6ab4fbdcf44a9610dab19
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294274"
---
# <a name="allowsnapshot-element"></a>AllowSnapshot 元素

指定是否将内容外接程序的快照图像与主机文档一起保存。

**外接程序类型：** 内容

## <a name="syntax"></a>语法

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注释

 > [!IMPORTANT]
 > **AllowSnapshot** 在默认情况下为 `true`。 这使得在不支持 Office 外接程序的 Office 应用程序版本中打开文档的用户可以看到加载项的图像，如果应用程序无法连接到承载外接程序的服务器，则会提供该外接程序的静态图像。 但是，这也意味着可以直接从托管该外接程序的文档访问显示在外接程序中的潜在敏感信息。
