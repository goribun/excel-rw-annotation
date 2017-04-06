# excel-rw-annotation
A simple tools for reading and writing excel by annotation

## Example

### annotation
```java
    @ExcelField(name = "姓名")
    private String name;

    @ExcelField(name = "年龄", tags = {1})
    private int age;

    @ExcelField(name = "体重", format = "0.00")
    private BigDecimal weight;

    @ExcelField(name = "手机", order = 10, tags = {2, 3})
    private String mobile;

    @ExcelField(name = "生日", format = "yyyy-MM-dd")
    private Date birthday;

    @ExcelField(name = "性别", defaultValue = "保密")
    private String sex;

    @ExcelField(name = "周度", string = "第{{value}}周")
    private int weekly;
```
### write

```java  
    byte[] bytes = ExcelHelper.write(personList, Person.class);
```  

```java
    //只导出注解tags包含2以及没有tags属性的字段
    byte[] bytes = ExcelHelper.write(persionList, Persion.class, 2);
```
### read

```java
    List<Persion> persionList = ExcelHelper.read(inputStream, Persion.class);
```

````java
    //读取excel文件为Map类型的列表
    List<Map<String, String>> list = ExcelHelper.read(inputStream);
````
    
