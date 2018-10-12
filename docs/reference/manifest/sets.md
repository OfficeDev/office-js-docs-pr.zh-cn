# <a name="sets-element"></a>Sets 元素

指定适用于 Office 的 JavaScript API 的最小子集，Office 外接程序需要该子集才能激活。

**外接程序类型：** Content、Task pane、Mail

## <a name="syntax"></a>语法

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>包含在

[要求](requirements.md)

## <a name="can-contain"></a>可以包含

[Set](set.md)

## <a name="attributes"></a>属性

|**属性**|**类型**|**是否必需**|**说明**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|字符串|可选|为所有子 **Set** 元素指定默认的 [MinVersion](set.md) 属性值。默认值为“1.1”。|

## <a name="remarks"></a>备注

有关要求集的详细信息，请参阅 [Office 版本和要求集合](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

有关 **集合** 元素的 **MinVersion** 属性和 **集合** 元素的 **DefaultMinVersion** 属性的详细信息，请参阅[在清单中设置要求元素](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)。

