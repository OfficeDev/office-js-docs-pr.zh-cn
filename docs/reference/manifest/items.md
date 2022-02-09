---
title: 清单文件的 Items 元素
description: 指定菜单中的项。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2249bc55db662a36cf3986ebb0b90353237d4985
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467899"
---
# <a name="items-element"></a>Items 元素

指定菜单中的项。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- 当父 **VersionOverrides** 的类型为 Taskpane [1.0 时，AddinCommands](../requirement-sets/add-in-commands-requirement-sets.md) 1.1。
- [当父](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) **VersionOverrides** 类型为 Mail 1.0 时，邮箱 1.3。
- [当父](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) **VersionOverrides** 类型为 Mail 1.1 时，邮箱 1.5。

## <a name="syntax"></a>语法

```XML
<Items>
...  
</Items>  
```

## <a name="contained-in"></a>包含于

[Menu 类型的 Control 元素](control-menu.md)

## <a name="must-contain"></a>必须包含

[项目](item.md)

## <a name="examples"></a>示例

有关示例，请参阅 [Menu 类型的控件](control-menu.md)。