## Chap_07 宏

* 宏和模块：将Access数据库的对象组合起来，协同工作
* 宏
    > * 一个或多个操作组成的集合，其中每个操作都能实现特定的功能
    > * 独立宏、嵌入宏、数据宏

* 宏设计视图
    > * 宏生成器
    > * 操作目录
    >   * 程序流程：Comment、Group、If、Submacro
    >   * 操作 (共66条)：
    >       * 窗口管理
    >           CloseWindow：关闭指定窗口
    >           MaximizeWindow：最大化窗口
    >           MinimizeWindow：最小化窗口
    >           MoveAndSizeWindow：移动并调整窗口
    >           RestoreWindow：恢复窗口原来大小
    >       * 宏命令
    >           CancelEvent：取消导致该宏运行的Access事件
    >           ClearMacroError：清除MacroError对象的上一错误
    >           OnError：定义错误处理行为
    >           RunCode：执行VB Function过程
    >           RunDataMacro：运行数据宏
    >           RunMacro：运行一个宏
    >           RunMenuCommand：执行Access菜单命令
    >           StopAllMacros：终止所有正在运行的宏
    >           StopMacros：终止当前正在运行的宏
    >       * 筛选/查询/搜索
    >           FindRecord：查找符合指定条件的第一条或下一条记录
    >           OpenQuery：打开选择查询或交叉表查询，或者执行动作查询
    >       * 数据导入/导出
    >           ExportWithFoematting：将指定数据库对象中的数据输出为.xls、.rtf、.txt、.htm、.snp
    >       * 数据库对象
    >           GoToControl：把焦点移到激活数据表或窗体上指定的字段或控件上
    >           GoToRecord：把表、窗体或查询结果集中的指定记录成为当前记录
    >           OpenForm：在“窗体”、“设计”、“打印预览”、“数据表”视图打开窗体
    >           OpenReport：在“设计”视图、“打印预览”打开报表，或立即打印该报表
    >           OpenTable：在“数据表”、“设计”、“打印预览”视图打开表
    >           PrintObject：打印当前对象
    >           PrintPreview：当前对象的“打印预览”
    >           RepaintObject：完成所有未完成的屏幕更新或控件的重新计算
    >           SetProperty：设置控件属性
    >       * 数据输入操作
    >           DeleteRecord：删除当前记录
    >           EditListItems：编辑查阅列表中的项
    >           SaveRecord：保存当前已录
    >       * 系统命令
    >           Beep：使计算机发出嘟嘟声
    >           CloseDataBase：关闭当前数据库
    >           QuitAccess：退出Access
    >       * 用户界面命令
    >           AddMenu：为窗体或报表将菜单添加到自定义菜单栏
    >           MessageBox：显示含有警告或提示信息的消息框
    >           Redo：重复最近的用户操作
    >           UndoRecord：撤销最近的用户操作
    >   * 在此数据库中：当前数据库中已有的宏对象

* 创建宏
    > * 独立宏
    >   * 操作序列的独立宏一般只包含一条或多条操作和一个或多个注释，顺序执行
    > * 条件操作宏
    >   * 添加新操作 - If - 展开设计窗格 - 输入条件表达式 - End If
    >   * Else块
    > * 使用Group对宏中的操作进行分组
    >   * 不影响执行方式，仅标识一组操作
    >   * “Group” - “End Group”
    >   * 选择操作 - “生成分组程序块” - 输入名称
    > * 设置宏的操作参数
    >   * 不同操作所拥有的参数及参数个数是不同的
    > * 创建含子宏的独立宏
    >   * 引用子宏的格式是“宏名.子宏名”

    > * 嵌入宏
    >   * 嵌入在窗体或报表或其控件的事件属性中的宏
    >   * 使用控件向导创建控件
    >   * 对某对象的某事件属性使用宏生成器创建嵌入宏

    > * 数据宏
    >   * 通过使用数据宏将逻辑附加到数据中来增加代码的可维护性
    >   * 包括：插入后、更新后、删除后、删除前、更改前

    > * 创建自动执行的AutoExec独立宏
    >   * 对初始参量赋值、打开应用系统的“登录”窗体等
    >   * 自行设置名字

* 宏的运行
    > * 不含子宏的宏，直接指定该宏名运行
    > * 含有子宏的宏，用“宏名.子宏名”格式指定某个子宏
    > * 设计视图 - 运行
    > * 双击宏名/右键运行
    > * 在VBA过程中，使用DoCmd对象的RunMacro方法运行宏：DoCmd.RunMacro “宏名”
* 宏的调试
    > * 宏 - 单步 - 运行 单步执行/停止/继续

* 宏和VBA编程：取决于需要完成的任务
* 使用VBA的情况
    > * 使用内置函数或创建自己的函数
    > * 创建或操纵对象
    > * 执行系统级操作
    > * 一次一条的操纵记录
* 独立宏转换为VB代码