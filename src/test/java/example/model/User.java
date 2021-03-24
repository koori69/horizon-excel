package example.model;

import horizon.excel.annotation.ExcelColumn;
import horizon.excel.annotation.ExcelSheet;

import java.time.LocalDateTime;

/**
 * @author xxq
 * <p>
 * Create at 2018-06-14 11:27
 */
@ExcelSheet(name = "用户")
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
