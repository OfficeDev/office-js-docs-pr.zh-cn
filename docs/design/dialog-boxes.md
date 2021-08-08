---
title: Office 加载项中的对话框
description: 了解在加载项中可视化设计对话框Office最佳实践。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 84af1bf7d5574ef87d66f801f7e1f7e74934601fcebcc9c273b9e9e40e8f4e38
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57081959"
---
# <a name="dialog-boxes-in-office-add-ins"></a>Office 加载项中的对话框

对话框是浮动在活动的 Office 应用程序窗口之上的界面。你可以使用对话框为无法直接在任务窗格中打开的任务（例如登录页）提供额外的屏幕空间，或请求确认用户进行的操作，或显示如果局限在任务窗格中可能过小的视频。

*图 1：对话框典型布局*

![应用程序中显示的对话框的典型Office布局。](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a>最佳做法

|允许事项|禁止事项|
|:-----|:--------|
|<ul><li>包括包含外接程序名称以及当前任务的描述性标题。</li></ul>|<ul><li>请勿在标题中追加公司名称。</li></ul>|
||<ul><li>除非方案需要，否则请勿打开对话框。</li></ul>|

## <a name="implementation"></a>实现

有关实现对话框的示例，请参阅 GitHub 上的 [Office 加载项对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。

## <a name="see-also"></a>另请参阅 

- [Dialog 对象](/javascript/api/office/office.dialog)
- [适用于 Office 加载项的 UX 设计模式](../design/ux-design-pattern-templates.md)
