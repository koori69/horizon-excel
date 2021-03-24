# Horizon-Excel

Quickly import excel to POJO or export POJO to excel by annotation.

## Features

- import excel（xls、xlsx） to POJO
- export POJO to excel （xls、xlsx）
- LocalDateTime、LocalDate are supported
- custom strategy （define the specific behavior of  excel data to POJO / POJO to excel data） 
- cell merge data to POJO is supported
- support custom OutputStream
## Environmental

- JDK >= 8.0

## Annotation

### ExcelSheet

- name

  the name of excel sheet which need to import or export

- importIndex

  the index of sheet, default 0

- exportIndex

  the index of sheet, default 0

### ExcelColumn

- name

  the name of column

- width

  the width of column

- skip

  ignored when import or export, default false

- dateFormat

  format date to string, default yyyy-MM-dd HH:mm:ss

- precision

  Floating point accuracy, -1 means unlimited accuracy,  default -1

- round

  to the nearest whole number, default true

## Strategy

### DataStrategy

It's the interface that can be implemented to define the specific behavior of the data.

- getValue

  return processed data

- verify

  verify the correctness of the data

- getStrategy

  return the strategy

## Demo

### POJO

```java
@ExcelSheet(name = "用户", importIndex = 1)
public class User {
    @ExcelColumn(name = "序号")
    private Integer id;
    @ExcelColumn(name = "名字")
    private String name;
    @ExcelColumn(name = "时间")
    private LocalDateTime createTime;
    @ExcelColumn(name = "性别")
    private Boolean sex;

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public LocalDateTime getCreateTime() {
        return createTime;
    }

    public void setCreateTime(LocalDateTime createTime) {
        this.createTime = createTime;
    }

    public Boolean getSex() {
        return sex;
    }

    public void setSex(Boolean sex) {
        this.sex = sex;
    }
}
```

### Strategy

```java
public class SexStrategy implements DataStrategy {
    public static Map<String, Boolean> map;

    static {
        map = new HashMap<>();
        map.put("女", true);
        map.put("男", false);
    }

    private String key;
    private boolean value;

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public boolean isValue() {
        return value;
    }

    public void setValue(boolean value) {
        this.value = value;
    }

    @Override
    public Object getValue(String key) {
        return map.get(key);
    }

    @Override
    public boolean verify(Object key) {
        boolean flag = false;
        for (String k : map.keySet()) {
            if (k.equals(key)) {
                flag = true;
            }
        }
        return flag;
    }

    @Override
    public String getStrategy() {
        return map.keySet().stream().collect(Collectors.joining(", "));
    }
```



```java
public class SexExportStrategy implements DataStrategy {
    public static Map<Boolean, String> map;

    static {
        map = new HashMap<>();
        map.put(true, "女");
        map.put(false, "男");
    }

    private boolean key;
    private String value;

    public boolean getKey() {
        return key;
    }

    public void setKey(boolean key) {
        this.key = key;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    @Override
    public Object getValue(String key) {
        switch (key) {
            case "true":
                return map.get(true);
            case "false":
                return map.get(false);
            default:
                return null;
        }
    }
}
```

### import

```java
   File file = new File("D:/test.xlsx");
        try {
            ExcelDataFormatter edf = new ExcelDataFormatter();
            edf.set("sex", new SexStrategy());
            List<User> list = ExcelImportUtils.readFromFile(edf, file, User.class);
            if (ExcelImportUtils.getBuilder().length() > 0) {
                System.out.println(ExcelImportUtils.getBuilder().toString());
            }
            if (null != list) {
                for (User u : list) {
                    System.out.println("id:" + u.getId() + " name:" + u.getName() + " date:" + u.getCreateTime() + " " +
                        "sex: " + u.getSex());
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
```

### export

```java
List<User> list;
...
ExcelDataFormatter edf = new ExcelDataFormatter();
edf.set("sex", new SexExportStrategy());
ExcelExportUtils.writeSingleSheetToFile(edf, "D:/tt.xlsx", list);
```

