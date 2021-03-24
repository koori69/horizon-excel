package horizon.excel.utils;

import horizon.excel.annotation.ExcelColumn;
import horizon.excel.annotation.ExcelSheet;
import horizon.excel.formatter.ExcelDataFormatter;
import horizon.excel.strategy.DataStrategy;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.*;

/**
 * @author K·J
 * <p>
 * Create at 2018-06-14 10:01
 */
public class ExcelImportUtils {
    private final static String XLSX = "xlsx";
    private final static String XLS = "xls";
    private static StringBuilder builder;
    private static int etimes = 0;
    private static int defaultTryTime = 7;

    public static <E> List<E> readFromStream(ExcelDataFormatter edf, InputStream is, String suffix, Class<?> clazz) throws Exception {
        builder = new StringBuilder();
        Field[] fields = ReflectUtils.getClassFieldsAndSuperClassFields(clazz);
        Map<String, String> keys = new HashMap<>();
        ExcelColumn column = null;
        ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);
        for (Field field : fields) {
            column = field.getAnnotation(ExcelColumn.class);
            if (null == column || column.skip()) {
                continue;
            }
            keys.put(column.name(), field.getName());
        }

        Workbook wb = null;
        if (suffix.endsWith(XLSX)) {
            wb = new XSSFWorkbook(is);
        } else if (suffix.endsWith(XLS)) {
            wb = new HSSFWorkbook(is);
        } else {
            return null;
        }
        int index = 0;
        if (null != excelSheet) {
            index = excelSheet.importIndex();
        }
        Sheet sheet = wb.getSheetAt(index);
        Row title = sheet.getRow(0);
        String[] titles = new String[title.getPhysicalNumberOfCells()];
        for (int i = 0; i < titles.length; i++) {
            titles[i] = title.getCell(i).getStringCellValue();
        }

        List<E> list = new ArrayList<>();
        E e = null;

        int rowIndex = 0;
        int columnCount = titles.length;
        Cell cell = null;
        Row row = null;

        for (Iterator<Row> it = sheet.rowIterator(); it.hasNext(); ) {
            row = it.next();
            if (rowIndex++ == 0) {
                continue;
            }
            if (row == null) {
                break;
            }
            e = (E)clazz.newInstance();
            for (int i = 0; i < columnCount; i++) {
                cell = row.getCell(i);
                if (ExcelUtils.isMergedRegion(sheet, row.getRowNum(), i)) {
                    Cell c = ExcelUtils.getMergedRegionCell(sheet, row.getRowNum(), i);
                    etimes = 0;
                    //解析column
                    readCellContent(keys.get(titles[i]), fields, c, e, edf);
                } else {
                    if (null == cell) {
                        continue;
                    }
                    etimes = 0;
                    //解析column
                    readCellContent(keys.get(titles[i]), fields, cell, e, edf);
                }
            }
            list.add(e);
        }
        return list;
    }

    public static <E> List<E> readFromFile(ExcelDataFormatter edf, File file, Class<?> clazz) throws Exception {
        builder = new StringBuilder();
        Field[] fields = ReflectUtils.getClassFieldsAndSuperClassFields(clazz);
        Map<String, String> keys = new HashMap<>();
        ExcelColumn column = null;
        ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);
        for (Field field : fields) {
            column = field.getAnnotation(ExcelColumn.class);
            if (null == column || column.skip()) {
                continue;
            }
            keys.put(column.name(), field.getName());
        }
        InputStream is = new FileInputStream(file);
        Workbook wb = null;
        if (file.getName().endsWith(XLSX)) {
            wb = new XSSFWorkbook(is);
        } else if (file.getName().endsWith(XLS)) {
            wb = new HSSFWorkbook(is);
        } else {
            return null;
        }
        int index = 0;
        if (null != excelSheet) {
            index = excelSheet.importIndex();
        }
        Sheet sheet = wb.getSheetAt(index);
        Row title = sheet.getRow(0);
        String[] titles = new String[title.getPhysicalNumberOfCells()];
        for (int i = 0; i < titles.length; i++) {
            titles[i] = title.getCell(i).getStringCellValue();
        }

        List<E> list = new ArrayList<>();
        E e = null;

        int rowIndex = 0;
        int columnCount = titles.length;
        Cell cell = null;
        Row row = null;

        for (Iterator<Row> it = sheet.rowIterator(); it.hasNext(); ) {
            row = it.next();
            if (rowIndex++ == 0) {
                continue;
            }
            if (row == null) {
                break;
            }
            e = (E)clazz.newInstance();
            for (int i = 0; i < columnCount; i++) {
                cell = row.getCell(i);
                if (ExcelUtils.isMergedRegion(sheet, row.getRowNum(), i)) {
                    Cell c = ExcelUtils.getMergedRegionCell(sheet, row.getRowNum(), i);
                    etimes = 0;
                    //解析column
                    readCellContent(keys.get(titles[i]), fields, c, e, edf);
                } else {
                    if (null == cell) {
                        continue;
                    }
                    etimes = 0;
                    //解析column
                    readCellContent(keys.get(titles[i]), fields, cell, e, edf);
                }
            }
            list.add(e);
        }
        return list;
    }

    /**
     * 从单元格读取数据，根据不同的数据类型，使用不同的方式读取
     * 根据Bean的数据类型进行相应转换
     * 如果找不到合适的处理类型，抛出异常
     *
     * @param key    column title
     * @param fields fileds
     * @param cell   cell
     * @param obj    object
     * @param edf    ExcelDataFormatter
     * @throws Exception error
     */
    public static void readCellContent(String key, Field[] fields, Cell cell, Object obj, ExcelDataFormatter edf) throws
            Exception {
        Object o = null;
        try {
            switch (cell.getCellType()) {
                case BOOLEAN: {
                    o = cell.getBooleanCellValue();
                    break;
                }
                case NUMERIC: {
                    o = cell.getNumericCellValue();
                    if (DateUtil.isCellDateFormatted(cell)) {
                        o = DateUtil.getJavaDate(cell.getNumericCellValue());
                    }
                    break;
                }
                case STRING: {
                    o = cell.getStringCellValue();
                    break;
                }
                case ERROR: {
                    o = cell.getErrorCellValue();
                    break;
                }
                case BLANK: {
                    o = null;
                    break;
                }
                case FORMULA: {
                    o = cell.getCellFormula();
                    break;
                }
                default:
                    o = "";
            }
            if (o == null) {
                return;
            }

            for (Field field : fields) {
                field.setAccessible(true);
                if (field.getName().equals(key)) {
                    Boolean bool = true;
                    DataStrategy strategy = null;
                    if (edf == null) {
                        bool = false;
                    } else {
                        strategy = edf.get(field.getName());
                        if (strategy == null) {
                            bool = false;
                        }
                    }

                    if (bool) {
                        if (!strategy.verify(o.toString())) {
                            builder.append(String.format("Error:[row:%d,column:%d,value:%s] does not compliant " +
                                            "the strategy [%s]\n", cell.getRowIndex() + 1,
                                    cell.getColumnIndex() + 1, o.toString(),
                                    strategy.getStrategy()));
                            continue;
                        }
                    }
                    if (field.getType().equals(Date.class) || field.getType().equals(LocalDateTime.class)) {
                        if (field.getType().equals(LocalDateTime.class)) {
                            field.set(obj, DateUtils.dateToLocalDateTime((Date)o));
                        } else if (o.getClass().equals(Date.class)) {
                            field.set(obj, o);
                        } else {
                            field.set(obj, LocalDateTime.parse(o.toString()));
                        }
                    } else if (field.getType().equals(String.class)) {
                        if (o.getClass().equals(String.class)) {
                            field.set(obj, o);
                        } else {
                            field.set(obj, o.toString());
                        }
                    } else if (field.getType().equals(Long.class)) {
                        if (o.getClass().equals(Long.class)) {
                            field.set(obj, o);
                        } else if (bool) {
                            field.set(obj, strategy.getValue(o.toString()));
                        } else {
                            // 检查是否需要转换
                            String ostr = o.toString();
                            ostr = ostr.split("\\.").length > 0 ? ostr.split("\\.")[0] : ostr;
                            try {
                                field.set(obj, Long.parseLong(ostr));
                            } catch (NumberFormatException ignored) {
                                builder.append(String.format("Error:[row:%d,column:%d,value:%s] does not compliant " +
                                                "the strategy [Long number]\n", cell.getRowIndex() + 1,
                                        cell.getColumnIndex() + 1, o.toString()));
                            }
                        }
                    } else if (field.getType().equals(Integer.class)) {
                        if (o.getClass().equals(Integer.class)) {
                            field.set(obj, o);
                        } else if (bool) {
                            field.set(obj, strategy.getValue(o.toString()));
                        } else {
                            // 检查是否需要转换
                            String ostr = o.toString();
                            ostr = ostr.split("\\.").length > 0 ? ostr.split("\\.")[0] : ostr;
                            try {
                                field.set(obj, Integer.parseInt(ostr));
                            } catch (NumberFormatException ignored) {
                                builder.append(String.format("Error:[row:%d,column:%d,value:%s] does not compliant " +
                                                "the strategy [Integer number]\n", cell.getRowIndex() + 1,
                                        cell.getColumnIndex() + 1, o.toString()));
                            }
                        }
                    } else if (field.getType().equals(BigDecimal.class)) {
                        if (o.getClass().equals(BigDecimal.class)) {
                            field.set(obj, o);
                        } else if (bool) {
                            field.set(obj, strategy.getValue(o.toString()));
                        } else {
                            field.set(obj, BigDecimal.valueOf(Double.parseDouble(o.toString())));
                        }
                    } else if (field.getType().equals(Boolean.class)) {
                        if (o.getClass().equals(Boolean.class)) {
                            field.set(obj, o);
                        } else if (bool) {
                            field.set(obj, strategy.getValue(o.toString()));
                        } else {
                            field.set(obj, Boolean.parseBoolean(o.toString()));
                        }
                    } else if (field.getType().equals(Float.class)) {
                        if (o.getClass().equals(Float.class)) {
                            field.set(obj, o);
                        } else if (bool) {
                            field.set(obj, strategy.getValue(o.toString()));
                        } else {
                            field.set(obj, Float.parseFloat(o.toString()));
                        }
                    } else if (field.getType().equals(Double.class)) {
                        if (o.getClass().equals(Double.class)) {
                            field.set(obj, o);
                        } else if (bool) {
                            field.set(obj, strategy.getValue(o.toString()));
                        } else {
                            field.set(obj, Double.parseDouble(o.toString()));
                        }
                    }
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
            // 如果还是读到的数据格式还是不对，只能放弃了
            if (etimes > defaultTryTime) {
                throw ex;
            }
            etimes++;
            if (o == null) {
                readCellContent(key, fields, cell, obj, edf);
            }
        }
    }

    public static StringBuilder getBuilder() {
        return builder;
    }

    public static void setBuilder(StringBuilder builder) {
        ExcelImportUtils.builder = builder;
    }
}
