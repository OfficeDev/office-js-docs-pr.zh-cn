---
title: 清单文件中的 Version 元素
description: Version 元素指定 Office 外接程序版本。
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 48a2be94d95ece597e47468bb18db2a7962a51e9
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173932"
---
# <a name="version-element"></a>Version 元素

指定 Office 外接程序的版本。 版本号可以是 1、2、3 或 4 部分 (例如 n、n.n、n.n 或 n.n.n) 。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注解

版本号的每个部分最多为 5 位数字。
