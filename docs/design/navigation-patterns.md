# <a name="navigation-patterns"></a><span data-ttu-id="c425e-101">导航模式</span><span class="sxs-lookup"><span data-stu-id="c425e-101">Navigation patterns</span></span>

<span data-ttu-id="c425e-102">加载项的主要功能通过特定的命令类型和有限的屏幕区域进行访问。</span><span class="sxs-lookup"><span data-stu-id="c425e-102">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="c425e-103">导航很直观，提供上下文并允许用户在整个加载项中轻松移动，这一点很重要。</span><span class="sxs-lookup"><span data-stu-id="c425e-103">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="c425e-104">最佳做法</span><span class="sxs-lookup"><span data-stu-id="c425e-104">Best practices</span></span>

| <span data-ttu-id="c425e-105">允许事项</span><span class="sxs-lookup"><span data-stu-id="c425e-105">Do</span></span>    | <span data-ttu-id="c425e-106">禁止事项</span><span class="sxs-lookup"><span data-stu-id="c425e-106">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="c425e-107">确保用户有一个清晰可见的导航选项。</span><span class="sxs-lookup"><span data-stu-id="c425e-107">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="c425e-108">不要使用非标准的 UI 来使导航流程复杂化。</span><span class="sxs-lookup"><span data-stu-id="c425e-108">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="c425e-109">根据适用情况，使用以下组件来允许用户浏览加载项。</span><span class="sxs-lookup"><span data-stu-id="c425e-109">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="c425e-110">不要让用户难以理解他们在加载项中的当前位置或上下文</span><span class="sxs-lookup"><span data-stu-id="c425e-110">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>



## <a name="command-bar"></a><span data-ttu-id="c425e-111">命令栏</span><span class="sxs-lookup"><span data-stu-id="c425e-111">command bar</span></span>

<span data-ttu-id="c425e-112">CommandBar 是一个表面，其中包含对上面所在的窗口、面板或父区域的内容进行操作的命令。</span><span class="sxs-lookup"><span data-stu-id="c425e-112">CommandBar is a surface that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="c425e-113">可选功能包括汉堡菜单访问点、搜索和侧面命令。</span><span class="sxs-lookup"><span data-stu-id="c425e-113">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![命令 - 桌面任务窗格的规范](../images/add-in-command-bar.png)



## <a name="tab-bar"></a><span data-ttu-id="c425e-115">选项卡栏</span><span class="sxs-lookup"><span data-stu-id="c425e-115">Tab bar</span></span>

<span data-ttu-id="c425e-116">data-id="undefined" class="unusedGlossaryTerm">选项卡栏</span><span class="sxs-lookup"><span data-stu-id="c425e-116">Tab bar - Shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="c425e-117">使用选项卡栏提供导航（使用简短的描述性标题的选项卡）。</span><span class="sxs-lookup"><span data-stu-id="c425e-117">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![选项卡栏 - 桌面任务窗格的规范](../images/add-in-tab-bar.png)


## <a name="back-button"></a><span data-ttu-id="c425e-119">后退按钮</span><span class="sxs-lookup"><span data-stu-id="c425e-119">Back button</span></span>

<span data-ttu-id="c425e-120">后退按钮允许用户从深化导航操作中恢复。</span><span class="sxs-lookup"><span data-stu-id="c425e-120">The back button allows users to recover from a drill down navigational action.</span></span> <span data-ttu-id="c425e-121">这个模式有助于确保用户遵循一系列有序的步骤。</span><span class="sxs-lookup"><span data-stu-id="c425e-121">Use this pattern to ensure users follow an ordered series of steps.</span></span>  

![后退按钮 - 桌面任务窗格的规范](../images/add-in-back-button.png)
