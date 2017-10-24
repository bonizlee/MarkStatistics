使用记事本打开config.json，修改配置内容
配置格式
"maxnumber":最大工位号，注意是工位号，不是选手人数
"project" :项目类文件名
"filetype":文件扩展名，xls或者xlsx
"subject" :数组表示分为多少个子评分项
	"filename" : 子评分项组别名称
        "judges" : 该子项评委数
        "markcell" : 第一个成绩所在单元格的标号
        "calculate" : 汇总方式。1为算术平均，2为去除最高最低再平均,3为去除偏差最厉害的一个成绩，剩余求平均

评分表格式命名  project+filename+judge序号

例：
{
    "maxnumber" : 15,
    "project" : "sec",
    "filetype" : "xlsx"
    "subject" : [
        {
            "filename" : "A",
            "judges" : 3,
            "markcell" : "D12",
            "calculate" : 1
        }
}
secA1.xls、secA2.xls、secA3.xls