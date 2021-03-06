## Chap_04 查询

* 查询：
    > * 对表中数据进行检索、统计、分析、查看和更改
    > * 本质上是SQL语句
    > * 从不同的表中抽取数据并组合成一个动态数据表，以数据表视图的方式显示
    > * 保存的内容是查询的结构，并不保存动态数据表
    > * 表和查询是查询、窗体、报表的数据源
    > * 需要先建立表与表之间的关系
* 查询的运行和修改
    > * 运行 $\Rightarrow$ “数据表视图”
    > * 修改 $\Rightarrow$ “设计视图”
* 选择查询
    > * 可对记录进行分组或合计、技术、平均值等计算
* 查询条件的设置
    > * 表达式
    > * 比较运算符
    > * 逻辑运算符：Not、And、Or
    > * 特殊运算符：IN、BETWEEN...AND、IS NULL、IS NOT NULL、LIKE、NOT LIKE
    > * 常用字符串函数：Left(str, n)、Right(str, n)、Mid(str, p, n)、Len(str)、Ltrim(str)、Rtrim(str)、Trim(str)
    > * 常用日期时间函数：Day(date)、Month(date)、Year(date)、Weekday(date)、Date()、Time()、Hour(time)
    > * 组合条件：And、Or
* 设置查询的计算
    > * 预定义计算：通过聚合函数对查询中的分组记录或全部记录进行“总计”计算
    > * 自定义计算：使用一个或多个字段中的数据在每个记录上执行数值、日期或文本计算；将表达式输入到查询设计网格中的空“字段”单元格中
* 交叉表查询
    > * 计算并重新组织数据的结构
    > * 查询向导
    > * 设计视图
* 参数查询
    > * 设计视图创建单个参数的查询：“条件”行单元格输入：[xxxx]
    > * 设计视图创建多个参数的查询：使用表达式
* 操作查询
    > * 从数据源中查找符合条件的记录集
    > * 对多条记录进行更改和移动
    > * 生成表查询
    >   * 利用一个或多个表中的全部或部分数据创建新表
    >   * 继承字段名称、数据类型、字段大小，不继承表的主键
    > * 追加查询
    >   * 将一个或多个表的一组记录添加到另一个已存在的表的末尾
    > * 更新查询
    >   * 对表中的部分记录或全部记录做更改
    > * 删除查询
    >   * 从一个或多个表中删除一组记录
* SQL查询
    > * SQL (Structure Query Language) 结构化查询语言
    > * 查询对象本质上是一个SQL语言编写的命令，运行查询就是执行SQL命令
    > * SELECT语句：对数据库的表做选择运算的命令
    > * INSERT语句：添加记录
    > * UPDATE语句：修改更新数据表中记录
    > * DELETE语句：删除记录
    > * SQL特定查询：数据定义查询、传递查询、联合查询