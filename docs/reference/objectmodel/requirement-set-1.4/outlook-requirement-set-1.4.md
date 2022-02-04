---
title: Outlook 加载项 API 要求集 1.4
description: 作为邮箱 API 1.4 Outlook外接程序和 Office JavaScript API 引入的功能和 API。
ms.date: 05/17/2021
ms.localizationpriority: medium
---

# <a name="outlook-add-in-api-requirement-set-14"></a>Outlook 外接程序 API 要求集 1.4

Outlook JavaScript API 的 Office API 子集包括可用于加载项的对象、方法、属性和Outlook事件。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。

## <a name="whats-new-in-14"></a>1.4 中的新增功能有哪些？

要求集 1.4 包括要求集 [1.3 的所有功能](../requirement-set-1.3/outlook-requirement-set-1.3.md)。 它添加了对 `Office.ui` 命名空间的访问权限。

### <a name="change-log"></a>更改日志

- 添加了 [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#office-office-ui-displaydialogasync-member(1))：在 Office 应用程序中显示对话框。
- 添加了 [Office.context.ui.messageParent](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#office-office-ui-messageparent-member(1))：将对话框中的消息传送到其父页/开始页。
- 添加了 [Dialog](/javascript/api/office/office.dialog?view=outlook-js-1.4&preserve-view=true) 对象：调用 [`displayDialogAsync`](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true#office-office-ui-displaydialogasync-member(1)) 方法时返回的对象。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
