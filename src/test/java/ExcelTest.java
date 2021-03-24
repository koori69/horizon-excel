import example.model.Product;
import example.model.User;
import example.strategy.SexExportStrategy;
import example.strategy.SexStrategy;
import horizon.excel.formatter.ExcelDataFormatter;
import horizon.excel.utils.ExcelExportUtils;
import horizon.excel.utils.ExcelImportUtils;
import org.junit.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * @author K·J
 * <p>
 * Create at 2018-06-14 11:32
 */
public class ExcelTest {
    @Test
    public void testExcelImport() {
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
    }

    @Test
    public void testExcelExport() {
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
            edf = new ExcelDataFormatter();
            edf.set("sex", new SexExportStrategy());
            ExcelExportUtils.writeSingleSheetToFile(edf, "D:/tt.xlsx", list);

            List<Product> ps = new ArrayList<>();
            Product p = new Product();
            p.setName("test");
            p.setPrice(3.4D);
            ps.add(p);
            ExcelExportUtils.writeMultiSheetToFile(edf, "D:/ts.xlsx", ps, list);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
