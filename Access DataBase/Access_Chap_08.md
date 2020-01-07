## Chap_08 模块与VBA程序设计

* 模块
    > * Access数据库中用于保存VBA程序代码的容器
    > * 由声明语句和 (Sub 和 Function)过程组成的集合
    > * 标准模块
    >   * 包含与任何其他对象都无关的常规过程 & 可以从数据库任何位置运行的经常使用的过程
    >   * 公共变量或公共过程具有全局特性，作用范围为整个应用程序
    > * 类模块
    >   * 可以包含新对象定义的模块
    >   * 窗体模块和报表模块都是类模块，具有局部特性
* 模块的组成
    > * 声明区域
    >   * Option声明
    >   * 变量或常量或自定义数据类型的声明
    > * Sub过程：子过程，无返回值
    >   * 关键字Call调用、引用过程名 (不带圆括号)
    > *  Function过程：函数过程，有返回值
    >   * 直接引用过程名 (必须带圆括号)

* VBA
    > * 面向对象
    > * 表、查询、窗体、报表、字段、控件都是对象
    > * 对象属性(静态操作)：描述对象的特征
    >   * 引用方式： 对象名.属性名
    > * 对象的方法(动态操作)：对象可以执行的行为
    >   * 改变对象的当前状态
    >   * 引用方式： 对象名.方法名
    > * DoCmd对象
    >   * 通过调用包含在内部的方法来实现VBA编程中对Access的操作
    >   * 如：DoCmd.OpenReport "课程信息"
    > * 对象的事件：窗体或控件等对象可“辨识”的动作
    >   * 用户操作的结果
    >   * 事件响应：
    >       * 使用宏对象来设置事件属性
    >       * 编写VBA代码：事件过程或事件响应代码
    > * 事件过程：事件处理程序
    >   * 响应由用户或程序引发的事件或由系统触发的事件而运行的过程

* VBE窗口 (Visual Basic Editor)
    > * VBA的编程环境
    > * 进入类模块的VBE
    >   * 属性表 - 事件 - ... - 代码生成器
    >   * 右击控件 - 事件生成器 - 代码生成器
    > * 进入标准模块的VBE
    >   * 创建 - 宏与代码 - 模块
    >   * 双击模块名
    > * 工程资源管理器窗口：工程窗口
    >   * 列表框包含应用程序的所有模块文件
    >   * 查看代码、查看对象、切换文件夹
    > * 属性窗口
    > * 代码窗口：编写、显示代码
    > * 立即窗口：测试代码
    > * 本地窗口：所有当前过程中的变量声明和变量值
    > * 监视窗口：监视表达式
    > * 提示信息、F1帮助信息

* 在模块中插入过程
    > * 插入 - 过程 - 添加过程 - 名称 - 类型(Sub/Func) - 范围(公共/私有)

* VBA数据类型 
    > * 标准数据类型：系统定义
    >   * 类型关键字、类型符(附加到变量名上的字符，指明数据类型)、前缀、取值范围

    > * Byt 字节型
    > * Int 整型 %
    > * Lng 长整型 &
    > * Sng 单精度 !
    > * Dbl 双精度 #
    > * Cur 货币型 @
    > * Str 字符串 $
    > * Bln 布尔型
    > * Dtm 日期型 
    > * Vnt 变体类型

    > * 变量名可以汉字命名，不区分大小写
    > * 显式变量
    >   * Dim i As integer, j As Single, s1 As String, s2 As String*8
    >   * Dim i%, j!, s1$, s2$*8
    > * 隐式变量：未直接定义，通过赋值建立的变量
    >   * m = 168.95
    >   * k% = 59
    > * 变量的作用域
    >   * 局部变量：过程声明的变量
    >   * 模块级变量：Dim...As声明变量
    >   * 全局变量：Public...As声明变量

    > * 一维数组
    >   * Dim E(1 to 4) As Single：下标从1开始到4的数组E
    >   * Dim F(5) As Integer：下标从0开始到5的数组F
    > * 二维数组
    >   * Dim G(2, 3) As Long：3(0~2)×4(0~3)的数组G
    > * 多维数组：与二维声明方法类似，最高60维
    > * 动态数组
    >   * Dim E() As Long：定义
    >   * ReDim E(9, 9, 9)：分配空间

    > * 自定义数据类型
    >   * 使用Type语句定义的数据类型
    > * 数据库对象变量
    >   * Forms!学生资料!学号.Value = "08031001"
    >   * Forms!学生资料!学号 = "08031001"
    >   * Forms!学生资料![学号] = "08031001"
    >   * Me![学号] = "08031001"