# <a name="method-element"></a>Method 元素

指定来自适用于 Office 的 JavaScript API 的单个方法，Office 加载项需要该方法才能激活。

**加载项类型：** 内容、任务窗格

## <a name="syntax"></a>语法

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>包含在

[方法](methods.md)

## <a name="attributes"></a>属性

|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|名称|字符串|必需|指定由其父对象限定的所需方法的名称。例如，要指定 **getSelectedDataAsync** 方法，必须指定 `"Document.getSelectedDataAsync"`。|

## <a name="remarks"></a>备注

**Methods**  和 **Method** 元素不受邮件加载项的支持。有关要求集的详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets) 。

> [!IMPORTANT] 
> 因为无法指定单个方法的最低版本要求，所以为了确保在运行时提供可用的方法，当你在加载项的脚本中调用该方法时，还应该使用 **if** 语句。 有关如何执行此操作的详细信息，请参阅[了解适用于 Office 的 JavaScript API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)。

