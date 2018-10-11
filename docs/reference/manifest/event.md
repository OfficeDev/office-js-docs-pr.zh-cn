# <a name="event-element"></a>Event 元素

定义加载项中的事件处理程序。

> [!NOTE] 
> 目前，Office 365 中的  Outlook 网页版仅支持 `Event` 元素。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [类型](#type-attribute)  |  是  | 指定要处理的事件。 |
|  [FunctionExecution](#functionexecution-attribute)  |  是  | 指定事件处理程序的执行风格、异步或同步。目前仅支持同步事件处理程序。 |
|  [FunctionName](#functionname-attribute)  |  是  | 指定事件处理程序的函数名称。 |

### <a name="type-attribute"></a>类型属性

必需。指定哪个事件将调用该事件处理程序。此属性的可能值在下表中指定。

|  事件类型  |  说明  |
|:-----|:-----|
|  `ItemSend`  |  当用户发送邮件或会议邀请时将调用该事件处理程序。  |

### <a name="functionexecution-attribute"></a>FunctionExecution 属性

必需。 必须设置为 `synchronous`。

### <a name="functionname-attribute"></a>FunctionName 属性

必需。指定事件处理程序的函数名称。这个值必须与加载项的[函数文件](functionfile.md)中的函数名称相匹配。

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```