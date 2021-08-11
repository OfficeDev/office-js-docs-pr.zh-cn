---
title: 清单文件中的 Version 元素
description: Version 元素指定Office外接程序版本。
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 9641153cbe6fa0284986b8dd286ba2114b32a82894bd5f8d33516e2a56c90be9
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57096325"
---
# <a name="version-element"></a>Version 元素

指定 Office 外接程序的版本。 版本号可以是 1、2、3 或 4 (，即 n、n.n、n.n 或 n.n.n) 。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注解

版本号的每个部分最多可包含 5 个数字。
