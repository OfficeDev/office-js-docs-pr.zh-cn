# <a name="requirements-element"></a>要求元素

指定适用于 Office 的 JavaScript API 要求（[要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)和/或 方法）的最小集，Office 外接程序需要该集才能激活。

**外接程序类型：** 内容、任务窗格、邮件

## <a name="syntax"></a>语法

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>包含在

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>可以包含

|**元素**|**内容**|**邮件**|**任务窗格**|
|:-----|:-----|:-----|:-----|
|[集](sets.md)|x|x|x|
|[方法](methods.md)|x||x|

## <a name="remarks"></a>备注

有关要求集的详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

