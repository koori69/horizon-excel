package example.model;

import horizon.excel.annotation.ExcelColumn;
import horizon.excel.annotation.ExcelSheet;

/**
 * @author K·J
 * <p>
 * Create at 2018-06-19 15:29
 */
@ExcelSheet(name = "商品")
public class Product {
    @ExcelColumn(name = "名字")
    private String name;
    @ExcelColumn(name = "价格")
    private Double price;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Double getPrice() {
        return price;
    }

    public void setPrice(Double price) {
        this.price = price;
    }
}
