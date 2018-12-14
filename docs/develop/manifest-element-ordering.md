---
title: 如何查找清单元素的正确顺序
description: 了解如何查找在父元素中放置子元素的正确顺序。
ms.date: 11/16/2018
ms.openlocfilehash: 3efc95926b7562b0e68bbb6f4b13c47cc4ae6824
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270612"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a>如何查找清单元素的正确顺序

Office 外接程序清单中的 XML 元素必须位于正确父元素下，*且*在父元素下以特定的相对顺序存在。

所需的排序在 [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) 文件夹的 XSD 文件中指定。 XSD 文件分类存放在对应任务窗格、内容和邮件三类外接程序的子文件夹中。

例如，在 `<OfficeApp>` 元素中，`<Id>`、`<Version>`、`<ProviderName>` 必须按此顺序出现。 如果添加了 `<AlternateId>` 元素，则其必须位于 `<Id>` 和 `<Version>` 元素之间。 如果任何元素的顺序出错，清单将无效并且你的外接程序将无法加载。

> [!NOTE]
> 当元素顺序被打乱时，[Office 外接程序验证程序](/office/dev/add-ins/testing/troubleshoot-manifest#validate-your-manifest-with-the-office-add-in-validator)将使用与元素位于错误父级下时相同的错误消息。 该错误消息会提示子元素不是父元素的有效子级。 如果出现此类错误，而子元素的参考文档却指示它对父级*是*有效的，则问题很可能是子级的放置顺序出现了错误。

若要查找给定父元素的子元素的正确顺序，请执行以下步骤。 （这是一个简化的过程，因为 XSD 文件非常复杂。 完全解析 XSD 文件不在本文的讨论范围之列。）

1. 打开 [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) 下的子文件夹，以获取你正在创建的外接程序的类型。 
2. 打开 XSD 文件，其中父元素被定义为复杂类型。 如果你不知道哪个文件具有该定义，则可能必须对多个文件执行步骤 3，直到找到它为止。
3. 搜索 `<xs:complexType name="PARENT_ELEMENT">`，其中 PARENT_ELEMENT 是该父元素的名称。
4. 在 PARENT_ELEMENT 的定义中，（通常）有一个名为 `<xs:sequence>` 的元素。 以下是 [TaskPaneAppVersionOverridesV1_0.xsd](https://raw.githubusercontent.com/OfficeDev/office-js-docs-pr/master/docs/overview/schemas/taskpane/TaskPaneAppVersionOverridesV1_0.xsd) 中对 `<SuperTip>` 的定义。

```xml
  <xs:complexType name="Supertip">
    <xs:annotation>
      <xs:documentation>
        Specifies the super tip for this control.
      </xs:documentation>
    </xs:annotation>
    <xs:sequence>
      <xs:element name="Title" type="bt:ShortResourceReference" minOccurs="1" maxOccurs="1" />
      <xs:element name="Description" type="bt:LongResourceReference" minOccurs="1" maxOccurs="1" />
    </xs:sequence>
  </xs:complexType>
```

`<xs:sequence>` *按照子元素必须出现的顺序*列出了可能的子元素。 但这并*不*意味着它们全都是必需的。 如果某个子元素的 `minOccurs` 值为 **0**，则该子元素是可选的。 *但如果该元素存在，则必须以由 `<xs:sequence>` 元素指定的顺序出现*。

如果没有 `<xs:sequence>` 元素，或者*有*该子元素但未列出（即使子元素的参考文档指示它对父级*是*有效的）；则在 XSD 文件中的其他位置通过其他子元素对父元素的复杂类型定义进行了扩展。 例如，`OfficeApp` 复杂类型的定义未将 `Requirements` 列为可能的子级。 但在文件的稍后部分（在 `TaskPaneApp` 复杂类型的定义中），对 `OfficeApp` 的定义进行了扩展，并添加了 `Requirements` 作为其他有效子级。

若要查找扩展的定义，请按照以下步骤操作：

1. 从文件的顶部开始，搜索 `<xs:extension base="PARENT_ELEMENT">`，其中 PARENT_ELEMENT 是父元素的名称。 可能存在多个扩展。
2. 查找与你正在使用的上下文相关的扩展。 例如，`OfficeApp` 复杂类型在 `ContentApp` 和 `MailApp` 复杂类型内进行了扩展，同时也在 `TaskPaneApp` 复杂类型内进行了扩展。

文件中的每个 `<xs:extension base="PARENT_ELEMENT">` 都有自己的 `<xs:sequence>`，后者会列出父级的其他有效子元素。 扩展列表上的子元素必须始终位于父级复杂类型定义中原始列表的子元素*之后*。

## <a name="see-also"></a>另请参阅

- [Office 外接程序清单的架构参考 (v1.1)](../develop/add-in-manifests.md)
