---
title: Excel JavaScript API 数据类型实体值卡
description: 了解如何将实体值卡与 Excel 外接程序中的数据类型配合使用。
ms.date: 07/14/2022
ms.topic: conceptual
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 7eb6251467b73af5e592d4cf013e899207944192
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889156"
---
# <a name="use-cards-with-entity-value-data-types-preview"></a>使用具有实体值数据类型的卡片 (预览) 

[!include[Data types preview availability note](../includes/excel-data-types-preview.md)]

本文介绍如何使用 [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 在 Excel UI 中使用实体值数据类型创建卡模式窗口。 这些卡片可以显示实体值中包含的其他信息，除了单元格中已经可见的信息，例如相关图像、产品类别信息和数据归属。

实体值（即 [EntityCellValue）](/javascript/api/excel/excel.entitycellvalue)是数据类型的容器，类似于面向对象的编程中的对象。 本文介绍如何使用实体值卡属性、布局选项和数据归属功能创建显示为卡片的实体值。

以下屏幕截图显示了打开实体值卡的示例，在此示例中，来自杂货店产品列表的 **Tofu** 产品。

:::image type="content" source="../images/excel-data-types-entity-card-tofu.png" alt-text="显示实体值数据类型的屏幕截图，其中显示了卡片窗口。":::

## <a name="card-properties"></a>卡片属性

使用实体值 [`properties`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-properties-member) 属性可以设置有关数据类型的自定义信息。 密 `properties` 钥接受嵌套数据类型。 每个嵌套属性或数据类型必须具有一个 `type` 和 `basicValue` 设置。

> [!IMPORTANT]
> 嵌套 `properties` 数据类型与后续文章部分中所述的 [卡片布局](#card-layout) 值结合使用。 在 `properties`定义嵌套数据类型后，必须在属性中 `layouts` 分配该数据类型才能在卡片上显示。

以下代码片段显示嵌套在其中的多个数据类型的实体值的 `properties`JSON。

> [!NOTE]
> 若要了解如何在完整的代码示例中使用此 JSON，请访问 [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/blob/main/samples/excel/85-preview-apis/data-types-entity-values.yaml) 存储库。

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        "Product ID": {
            type: Excel.CellValueType.string,
            basicValue: productID.toString() || ""
        },
        "Product Name": {
            type: Excel.CellValueType.string,
            basicValue: productName || ""
        },
        "Image": {
            type: Excel.CellValueType.webImage,
            address: product.productImage || ""
        },
        "Quantity Per Unit": {
            type: Excel.CellValueType.string,
            basicValue: product.quantityPerUnit || ""
        },
        "Unit Price": {
            type: Excel.CellValueType.formattedNumber,
            basicValue: product.unitPrice,
            numberFormat: "$* #,##0.00"
        },
        Discontinued: {
            type: Excel.CellValueType.boolean,
            basicValue: product.discontinued || false
        }
    },
    layouts: {
        // Enter layout settings here.
    }
};
```

以下屏幕截图显示了使用上述代码片段的实体值卡。 屏幕截图显示了前面代码片段中 **的产品 ID**、 **产品名称**、 **图像**、 **每单位数量** 和 **单价** 信息。

:::image type="content" source="../images/excel-data-types-entity-card-properties.png" alt-text="显示实体值数据类型的屏幕截图，其中显示了卡片布局窗口。该卡片显示产品名称、产品 ID、单位数量和单价信息。":::

## <a name="card-layout"></a>卡片布局

实体值 [`layouts`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-layouts-member) 属性为实体创建一个 [`card`](/javascript/api/excel/excel.entityviewlayouts) ，然后指定该卡片的外观，例如卡片的标题、卡片的图像和要显示的分区数。

> [!IMPORTANT]
> 嵌套 `layouts` 值与前面文章部分中所述的 [卡片属性](#card-properties) 数据类型结合使用。 必须先在其中定义 `properties` 嵌套数据类型，然后才能将其分配 `layouts` 到卡片上显示。

在属性中 `card` ，使用 [`CardLayoutStandardProperties`](/javascript/api/excel/excel.cardlayoutstandardproperties) 该对象定义卡片的组件，例如 `title`， `subTitle`以及 `sections`。

以下实体值 JSON 代码片段显示包含嵌`title`套和`mainImage`对象的布局，以及卡片中的三`sections`个`card`布局。 请注意，该 `title` 属性 `"Product Name"` 在前面的 [“卡片属性](#card-properties) ”文章部分中具有相应的数据类型。 该 `mainImage` 属性在上一部分中也有相应的 `"Image"` 数据类型。 该 `sections` 属性采用嵌套数组，并使用 [`CardLayoutSectionStandardProperties`](/javascript/api/excel/excel.cardlayoutsectionstandardproperties) 该对象来定义每个部分的外观。

在每个卡片部分中，可以指定元素，例如 `layout`， `title`以及 `properties`。 键 `layout` 使用 [`CardLayoutListSection`](/javascript/api/excel/excel.cardlayoutlistsection) 该对象并接受该值 `"List"`。 密 `properties` 钥接受字符串数组。 请注意， `properties` 这些值（例如 `"Product ID"`）在前面的 [“卡片属性](#card-properties) ”一文部分中具有相应的数据类型。 分区也可以折叠，并且可以在 Excel UI 中打开实体卡时使用布尔值定义为折叠或未折叠。

> [!NOTE]
> 若要了解如何在完整的代码示例中使用此 JSON，请访问 [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/blob/main/samples/excel/85-preview-apis/data-types-entity-values.yaml) 存储库。

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        // Enter property settings here.
    },
    layouts: {
        card: {
            title: { 
                property: "Product Name" 
            },
            mainImage: { 
                property: "Image" 
            },
            sections: [
                {
                    layout: "List",
                    properties: ["Product ID"]
                },
                {
                    layout: "List",
                    title: "Quantity and price",
                    collapsible: true,
                    collapsed: false, // This section will not be collapsed when the card is opened.
                    properties: ["Quantity Per Unit", "Unit Price"]
                },
                {
                    layout: "List",
                    title: "Additional information",
                    collapsible: true,
                    collapsed: true, // This section will be collapsed when the card is opened.
                    properties: ["Discontinued"]
                }
            ]
        }
    }
};
```

以下屏幕截图显示了使用上述代码片段的实体值卡。 屏幕截图显示顶部的 `mainImage` 对象，后跟 `title` 使用 **产品名称** 并设置为 **Tofu** 的对象。 屏幕截图还显示 `sections`。 **“数量”和“价格**”部分可折叠，包含 **每单位数量** 和 **单价**。 **“其他信息**”字段可折叠，并在卡片打开时折叠。

:::image type="content" source="../images/excel-data-types-entity-card-sections.png" alt-text="显示实体值数据类型的屏幕截图，其中显示了卡片布局窗口。该卡片显示卡片标题和分区。":::

## <a name="card-data-attribution"></a>卡片数据归因

实体值卡可以显示数据归因，以向实体卡中的信息提供者提供信用额度。 实体值[`provider`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-provider-member)属性使用[`CellValueProviderAttributes`](/javascript/api/excel/excel.cellvalueproviderattributes)对象，该对象定义`logoSourceAddress``description`值和`logoTargetAddress`值。

数据提供程序属性在实体卡的左下角显示图像。 它使用它 `logoSourceAddress` 来指定映像的源 URL。 如果选择了徽标映像，则该 `logoTargetAddress` 值定义 URL 目标。 将鼠标悬停在徽标上时，该 `description` 值将显示为工具提示。 如果`logoSourceAddress`未定义图像或图像的源地址已损坏，该`description`值还会显示为纯文本回退。

以下 JSON 代码片段显示一个实体值，该值使用该 `provider` 属性为实体指定数据提供程序属性。

> [!NOTE]
> 若要了解如何在完整的代码示例中使用此 JSON，请访问 [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/blob/main/samples/excel/85-preview-apis/data-types-entity-attribution.yaml) 存储库。

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        // Enter property settings here.
    },
    layouts: {
        // Enter layout settings here.
    },
    provider: {
        description: product.providerName, // Name of the data provider. Displays as a tooltip when hovering over the logo. Also displays as a fallback if the source address for the image is broken.
        logoSourceAddress: product.sourceAddress, // Source URL of the logo to display.
        logoTargetAddress: product.targetAddress // Destination URL that the logo navigates to when selected.
    }
};
```

以下屏幕截图显示了使用上述代码片段的实体值卡。 屏幕截图显示了左下角的数据提供程序归因。 在此实例中，数据提供程序为 Microsoft，并显示 Microsoft 徽标。

:::image type="content" source="../images/excel-data-types-entity-card-attribution.png" alt-text="显示实体值数据类型的屏幕截图，其中显示了卡片布局窗口。卡片显示左下角的数据提供程序归因。":::

## <a name="see-also"></a>另请参阅

- [ Excel 加载项中的数据类型的概述](excel-data-types-overview.md)
- [Excel 数据类型核心概念](excel-data-types-concepts.md)
- [Excel JavaScript API 参考](../reference/overview/excel-add-ins-reference-overview.md)