# <a name="coauthoring-in-excel-add-ins"></a>在 Excel 外接程序中共同创作  

借助[共同创作功能](https://support.office.com/en-US/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104)，多个人可以共同协作，并可以同时编辑同一个 Excel 工作簿。 其他合著者保存工作簿后，此工作簿的所有合著者均可立即看到此合著者的更改。 若要共同创作 Excel 工作簿，必须将工作簿存储在 OneDrive、OneDrive for Business 或 SharePoint Online 中。

> **重要说明：**在 Excel 2016 for Office 365 中，可在左上角看到“自动保存”功能。 启用“自动保存”后，将实时向合著者显示你的更改。 请考虑此行为对 Excel 外接程序设计的影响。 用户可以通过 Excel 窗口左上方的开关禁用“自动保存”。

共同创作功能在以下平台上可用：

- Excel Online
- Excel for Android
- Excel for iOS
- Excel Mobile for Windows 10
- 适用于 Office 365 客户的 Excel for Windows Desktop（Windows 桌面内部版本 16.0.8326.2076 或更高版本，当前的渠道客户自 2017 年 8 月起可获取这些版本）

## <a name="coauthoring-overview"></a>共同创作功能概述
 
当你更改工作簿的内容时，Excel 会自动向所有合著者同步这些更改。 合著者可以更改工作簿的内容，而 Excel 外接程序中运行的代码也可以更改此内容。 例如，在 Office 外接程序中运行以下 JavaScript 代码时，范围值会设置为 Contoso：


    range.values = [[‘Contoso’]];

向所有合著者同步“Contoso”后，同一个工作簿中运行的所有用户或外接程序均可看到新的范围值。 

共同创作功能仅同步共享工作簿中的内容。 不会同步 Excel 外接程序中从工作簿复制到 JavaScript 变量的值。 例如，如果外接程序将单元格的值（例如“Contoso”）存储在 JavaScript 变量中，然后一个合著者将此单元格的值更改为“Example”，则同步后所有合著者均可在单元格中看到“Example”。 但是 JavaScript 变量的值仍然设置为“Contoso”。 此外，当多个合著者使用同一个外接程序时，每个合著者会拥有自己的变量副本，并且此副本不会同步。 如果你使用的变量使用工作簿内容，那么，在使用此变量前，请务必查看工作簿中的更新值。 

## <a name="use-events-to-manage-the-in-memory-state-of-your-add-in"></a>使用事件管理外接程序的内存中状态
 
Excel 外接程序可以读取工作簿内容（通过隐藏工作表和设置对象），然后将内容存储在变量等数据结构中。 将原始值复制到其中的任意一个数据结构后，合著者可以更新原始的工作簿内容。 这表示现在数据结构中的复制值与工作簿内容不同步。 生成外接程序时，请务必考虑工作簿内容与数据结构中存储的值之间的这种独立性。

例如，你可能要生成一个显示自定义可视化效果的内容外接程序。 自定义可视化效果的状态可能保存在隐藏工作表中。 当合著者使用同一个工作簿时，可能会发生以下情况：

- 用户 A 打开文档，自定义可视化效果在工作簿中显示。 自定义可视化效果从隐藏工作表中读取数据（例如，将可视化效果的颜色设置为蓝色）。
- 用户 B 打开同一个文档，并开始修改自定义可视化效果。 用户 B 将自定义可视化效果的颜色设置为橙色。 橙色被保存到隐藏工作表中。
- 用户 A 的隐藏工作表更新为新值橙色。
- 用户 A 的自定义可视化效果仍然为蓝色。 

如果想要用户 A 的自定义可视化效果响应合著者对隐藏工作表的更改，请使用 [BindingDataChanged](http://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) 事件。 它可确保合著者对工作簿内容的更改反映到外接程序状态中。

## <a name="caveats-to-using-events-with-co-authoring"></a>使用事件进行共同创作的注意事项 

如上文所述，在某些情况下，针对所有合著者触发事件可改进用户体验。 但是，请注意在一些应用场景下，此行为可能会导致不良的用户体验。 

例如，在数据验证应用场景下，通常通过显示 UI 来响应事件。 本地用户或合著者（远程）通过绑定更改工作簿内容时，会运行前面部分中所述的 [BindingDataChanged](http://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) 事件。 如果 **BindingDataChanged** 事件的事件处理程序显示 UI，用户就会看到与他们在工作簿中进行的更改无关的 UI，从而导致不良的用户体验。 在外接程序中使用事件时，请避免显示 UI。

## <a name="see-also"></a>另请参阅 

- [Excel 中的共同创作功能的相关信息 (VBA)](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/about-coauthoring-in-excel) 
- [自动保存如何影响外接程序和宏 (VBA)](https://msdn.microsoft.com/en-us/vba/office-shared-vba/articles/how-autosave-impacts-addins-and-macros) 
