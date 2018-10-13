# <a name="officemenu-element"></a><span data-ttu-id="7cca6-101">OfficeMenu 元素</span><span class="sxs-lookup"><span data-stu-id="7cca6-101">OfficeMenu element</span></span>

<span data-ttu-id="7cca6-p101">定义要添加到 Office 上下文菜单的控件集合。适用于 Word、Excel、PowerPoint 和 OneNote 外接程序。</span><span class="sxs-lookup"><span data-stu-id="7cca6-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="7cca6-104">属性</span><span class="sxs-lookup"><span data-stu-id="7cca6-104">Attributes</span></span>

| <span data-ttu-id="7cca6-105">属性</span><span class="sxs-lookup"><span data-stu-id="7cca6-105">Attribute</span></span>            | <span data-ttu-id="7cca6-106">必需</span><span class="sxs-lookup"><span data-stu-id="7cca6-106">Required</span></span> | <span data-ttu-id="7cca6-107">说明</span><span class="sxs-lookup"><span data-stu-id="7cca6-107">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="7cca6-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="7cca6-108">xsi:type</span></span>](#xsitype) | <span data-ttu-id="7cca6-109">是</span><span class="sxs-lookup"><span data-stu-id="7cca6-109">Yes</span></span>      | <span data-ttu-id="7cca6-110">定义的 OfficeMenu 类型。</span><span class="sxs-lookup"><span data-stu-id="7cca6-110">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="7cca6-111">子元素</span><span class="sxs-lookup"><span data-stu-id="7cca6-111">Child elements</span></span>

|  <span data-ttu-id="7cca6-112">元素</span><span class="sxs-lookup"><span data-stu-id="7cca6-112">Element</span></span> |  <span data-ttu-id="7cca6-113">必需</span><span class="sxs-lookup"><span data-stu-id="7cca6-113">Required</span></span>  |  <span data-ttu-id="7cca6-114">说明</span><span class="sxs-lookup"><span data-stu-id="7cca6-114">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="7cca6-115">Control</span><span class="sxs-lookup"><span data-stu-id="7cca6-115">Control</span></span>](#control)    | <span data-ttu-id="7cca6-116">是</span><span class="sxs-lookup"><span data-stu-id="7cca6-116">Yes</span></span> |  <span data-ttu-id="7cca6-117">一个或多个 Control 对象的集合。</span><span class="sxs-lookup"><span data-stu-id="7cca6-117">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="7cca6-118">xsi:type</span><span class="sxs-lookup"><span data-stu-id="7cca6-118">xsi:type</span></span>

<span data-ttu-id="7cca6-119">指定要在其中添加此 Office 外接程序的 Office 客户端应用程序内置菜单。</span><span class="sxs-lookup"><span data-stu-id="7cca6-119">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="7cca6-p102">`ContextMenuText` - 当用户选定文本，然后打开（右键单击）选定文本上的上下文菜单时显示上下文菜单上的项。适用于 Word、Excel、PowerPoint 和 OneNote。</span><span class="sxs-lookup"><span data-stu-id="7cca6-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="7cca6-p103">`ContextMenuCell` - 当用户打开（右键单击）电子表格中的某个单元格上的上下文菜单时显示上下文菜单上的项。适用于 Excel。</span><span class="sxs-lookup"><span data-stu-id="7cca6-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="7cca6-124">Control</span><span class="sxs-lookup"><span data-stu-id="7cca6-124">Control</span></span>

<span data-ttu-id="7cca6-125">每个 OfficeMenu 元素都需要一个或多个 menu 控件。\*\* \*\* [ ](control.md#menu-dropdown-button-controls)</span><span class="sxs-lookup"><span data-stu-id="7cca6-125">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="7cca6-126">示例</span><span class="sxs-lookup"><span data-stu-id="7cca6-126">Example</span></span>

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```
