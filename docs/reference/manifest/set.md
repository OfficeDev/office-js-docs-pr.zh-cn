# <a name="set-element"></a>Set 元素

指定来自适用于 Office 的 JavaScript API 的要求集合，Office 外接程序需要该集才能激活。

**加载项类型：** Content、Task pane、mail

## <a name="syntax"></a>句法

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>包含在

[集](sets.md)

## <a name="attributes"></a>属性

|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|名称|String|必需|[要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)名称。|
|MinVersion|String|可选|指定加载项所需的 API 集的最低版本。如果 **DefaultMinVersion** 的值已在父 [Sets](sets.md) 元素中指定，则替代该值。|

## <a name="remarks"></a>备注

欲知要求集的详细信息，请参阅 [Office 版本和要求集合](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

欲知 **Set** 元素的 **MinVersion** 属性和 **Sets** 元素的 **DefaultMinVersion** 属性的详细信息，请参阅[在清单中设置要求元素](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)。

> [!IMPORTANT] 
> 对于邮件加载项，只有一个 `"Mailbox"` 要求集合可用。 此要求集合包含整个 outlook 邮件加载项中支持的 API 子集合，您必须指定邮件加载项清单中设置的 `"Mailbox"` 要求(对于内容和任务窗格加载项，它不是可选的)。 此外，无法声明支持邮件加载项中的特定模式。
