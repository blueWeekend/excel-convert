# excel-convert

一个简单易用的 Go 语言 Excel 处理库，通过结构体标签（struct tags）自动映射 Excel 列与数据结构字段，让你无需手动处理繁琐的列索引操作。

特性优势：

声明式映射：只需在结构体字段上添加标签（如 excel:"姓名"），即可自动绑定 Excel 列与数据字段

自动类型转换：支持 string 和 integer 类型的自动读写转换，自定义字段写入需实现ExcelMarshaler方法，读取需实现ExcelUnmarshaler方法

灵活校验模式：提供严格、宽松和禁用三种模板校验模式

嵌套结构体支持：自动递归处理嵌套结构体字段

零手动索引操作：只需初始化时指定模板标题字段顺序，告别读取/写入时手动维护列索引的繁琐和易错问题

## 🚀 安装

```bash
go get github.com/blueWeekend/excel-convert/v1

1. 定义结构体并添加标签

type User struct {
    Name  string `excel:"姓名"`
    Age   int    `excel:"年龄"`
    Email string `excel:"邮箱"`
}

2. 写入/读取 Excel 文件
// 初始化要读取的excel文件表头
columns := []string{"姓名", "年龄", "邮箱"}
// 创建转换器
converter := excelConvert.NewExcelConverter(columns)
inputUsers := []User{
    {Name: "张三", Age: 25, Email: "zhangsan@example.com"},
    {Name: "李四", Age: 30, Email: "lisi@example.com"},
}
// 定义写入的表头
header := []string{"姓名", "年龄", "邮箱"}
err := converter.WriteExcel(header, "users.xlsx", "Sheet1", inputUsers)
if err != nil {
    log.Fatal(err)
}
// 修改表头只写入部分数据场景
header = []string{"姓名", "邮箱"}
err = converter.WriteExcel(header, "partUsers.xlsx", "Sheet1", inputUsers)
if err != nil {
    log.Fatal(err)
}
// 读取 Excel 文件
var outputUsers []User
err = converter.ReadAll("users.xlsx", &outputUsers)
if err != nil {
    log.Fatal(err)
}
fmt.Println("outputUsers:", outputUsers)
// 读取与初始化模板不匹配的文件
err = converter.ReadAll("partUsers.xlsx", &outputUsers)
if err != nil {
    //因文件只有姓名与邮箱与初始化表头不匹配因此报错invalid tmpl;可设置禁用模板校验正常读取：excelConvert.SetTmplCheckMode(excelConvert.TmplCheckDisable)
    log.Fatal(err)
}

⚙️ 配置选项

设置标签名
默认使用 excel 标签，可以通过 SetTagName 修改：
converter := excelConvert.NewExcelConverter(columns, excelConvert.SetTagName("json"))

设置模板校验模式
// 严格模式：Excel 标题必须与初始化数组完全一致
converter := excelConvert.NewExcelConverter(columns, excelConvert.SetTmplCheckMode(excelConvert.TmplCheckStrict))

// 宽松模式（默认）：Excel 标题需包含初始化数组中的标题
converter := excelConvert.NewExcelConverter(columns, excelConvert.SetTmplCheckMode(excelConvert.TmplCheckLenient))

// 禁用模板校验
converter := excelConvert.NewExcelConverter(columns, excelConvert.SetTmplCheckMode(excelConvert.TmplCheckDisable))