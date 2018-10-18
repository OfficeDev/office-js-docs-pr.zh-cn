---
title: Office 加载项中的对话框
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 3d2fe2767f2f0d2d044dd2a4c5b309ff35202384
ms.sourcegitcommit: 3da2038e827dc3f274d63a01dc1f34c98b04557e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/19/2018
ms.locfileid: "24016267"
---
# <a name="dialog-boxes-in-office-add-ins"></a>Office 加载项中的对话框
 
对话框是浮动在活动的 Office 应用程序窗口之上的界面。你可以使用对话框为无法直接在任务窗格中打开的任务（例如登录页）提供额外的屏幕空间，或请求确认用户进行的操作，或显示如果局限在任务窗格中可能过小的视频。

*图 1：对话框典型布局*

![显示对话框典型布局的示例图像](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a>最佳做法

|**允许事项**|**禁止事项**|
|:-----|:--------|
|<ul><li>包括包含外接程序名称以及当前任务的描述性标题。</li></ul>|<ul><li>请勿在标题中追加公司名称。</li></ul>|
||<ul><li>除非方案需要，否则请勿打开对话框。</li></ul>|

## <a name="implementation"></a>实现

有关实现对话框的示例，请参阅 GitHub 上的 [Office 加载项对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。

## <a name="see-also"></a>另请参阅

- [GitHub 开发资源](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Dialog 对象](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js)


