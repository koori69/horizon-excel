package horizon.excel.utils;

import horizon.excel.annotation.ExcelColumn;
import horizon.excel.annotation.ExcelSheet;
import horizon.excel.formatter.ExcelDataFormatter;
import horizon.excel.strategy.DataStrategy;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.awt.Color;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.List;

/**
 * @author K·J
 * <p>
 * Create at 2018-06-14 13:39
 */
public class ExcelExportUtils {
    public static String defaultSheetName = "sheet";

    /**
     * 创建Sheet
     *
     * @param wb   Workbook
     * @param list data
     * @param edf  ExcelDataFormatter
     * @param <T>  Object
     * @throws Exception error
     */
    public static <T> void createSheet(Workbook wb, List<T> list,
                                       ExcelDataFormatter edf) throws Exception {
        if (null == wb || null == list || list.size() == 0) {
            return;
        }
        Sheet sheet = wb.createSheet();
        ExcelSheet excelSheet = list.get(0).getClass().getAnnotation(ExcelSheet.class);
        wb.setSheetName(excelSheet == null ? 0 : excelSheet.exportIndex(),
                excelSheet == null ? defaultSheetName : excelSheet.name());
        Row row = sheet.createRow(0);
        Cell cell = null;
        CreationHelper createHelper = wb.getCreationHelper();
        Field[] fields = ReflectUtils.getClassFieldsAndSuperClassFields(list.get(0).getClass());

        // 设置样式
        XSSFCellStyle titleStyle = (XSSFCellStyle)wb.createCellStyle();
        titleStyle.setFillForegroundColor(new XSSFColor(new Color(251, 226, 81)));
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 下边框
        titleStyle.setBorderBottom(BorderStyle.THIN);
        // 左边框
        titleStyle.setBorderLeft(BorderStyle.THIN);
        // 上边框
        titleStyle.setBorderTop(BorderStyle.THIN);
        // 右边框
        titleStyle.setBorderRight(BorderStyle.THIN);

        //写入标题
        int columnIndex = 0;
        ExcelColumn excelColumn = null;
        for (Field field : fields) {
            field.setAccessible(true);
            excelColumn = field.getAnnotation(ExcelColumn.class);
            if (excelColumn == null || excelColumn.skip()) {
                continue;
            }
            sheet.setColumnWidth(columnIndex, excelColumn.width() * 256);
            cell = row.createCell(columnIndex);
            cell.setCellStyle(titleStyle);
            cell.setCellValue(excelColumn.name());
            columnIndex++;
        }

        // 写入数据
        int rowIndex = 1;
        CellStyle cs = wb.createCellStyle();
        // 下边框
        cs.setBorderBottom(BorderStyle.THIN);
        // 左边框
        cs.setBorderLeft(BorderStyle.THIN);
        // 上边框
        cs.setBorderTop(BorderStyle.THIN);
        // 右边框
        cs.setBorderRight(BorderStyle.THIN);
        for (T t : list) {
            row = sheet.createRow(rowIndex);
            columnIndex = 0;
            Object o = null;
            for (Field field : fields) {
                field.setAccessible(true);
                excelColumn = field.getAnnotation(ExcelColumn.class);
                if (excelColumn == null || excelColumn.skip()) {
                    continue;
                }
                cell = row.createCell(columnIndex);
                cell.setCellStyle(cs);
                CellStyle cellStyle = wb.createCellStyle();
                cellStyle.cloneStyleFrom(cs);
                o = field.get(t);
                if (o == null) {
                    continue;
                }
                if (o instanceof LocalDateTime) {
                    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(excelColumn.dateFormat()));
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(DateUtils.localDateTimeToDate((LocalDateTime)field.get(t)));
                } else if (o instanceof Date) {
                    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(excelColumn.dateFormat()));
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue((Date)field.get(t));
                } else if (o instanceof Double || o instanceof Float) {
                    cell.setCellValue(field.get(t).toString());
                    if (excelColumn.precision() != -1) {
                        cell.setCellValue(new BigDecimal(field.get(t).toString()).setScale(excelColumn.precision(),
                                excelColumn.round() == true ? BigDecimal.ROUND_HALF_UP : BigDecimal.ROUND_FLOOR)
                                .toString());
                    }
                } else if (o instanceof BigDecimal) {
                    cell.setCellValue((field.get(t).toString()));
                    if (excelColumn.precision() != -1) {
                        cell.setCellValue(new BigDecimal(field.get(t).toString()).setScale(excelColumn.precision(),
                                excelColumn.round() == true ? BigDecimal.ROUND_HALF_UP : BigDecimal.ROUND_FLOOR)
                                .toString());
                    }
                } else if (o instanceof Boolean) {
                    Boolean bool = (Boolean)field.get(t);
                    if (edf == null) {
                        cell.setCellValue(bool);
                    } else {
                        DataStrategy strategy = edf.get(field.getName());
                        if (strategy == null) {
                            cell.setCellValue(bool);
                        } else {
                            cell.setCellValue((String)strategy.getValue(o.toString()));
                        }
                    }
                } else if (o instanceof Integer) {
                    Integer intValue = (Integer)field.get(t);
                    if (edf == null) {
                        cell.setCellValue(intValue);
                    } else {
                        DataStrategy strategy = edf.get(field.getName());
                        if (strategy == null) {
                            cell.setCellValue(intValue);
                        } else {
                            cell.setCellValue((String)strategy.getValue(o.toString()));
                        }
                    }
                } else {
                    cell.setCellValue(field.get(t).toString());
                }
                columnIndex++;
            }
            rowIndex++;
        }
    }

    /**
     * 获取Workbook对象
     *
     * @param list data
     * @param edf  ExcelDataFormatter
     * @param <T>  object
     * @return object
     * @throws Exception error
     */
    public static <T> Workbook getWorkBook(List<T> list,
                                           ExcelDataFormatter edf) throws Exception {
        // 创建工作簿
        Workbook wb = new SXSSFWorkbook();
        if (null == list || list.size() == 0) {
            return wb;
        }
        createSheet(wb, list, edf);
        return wb;
    }

    public static <T> Workbook getWorkBook(Workbook wb, List<T> list,
                                           ExcelDataFormatter edf) throws Exception {
        // 创建工作簿
        if (null == wb) {
            wb = new SXSSFWorkbook();
        }
        if (null == list || list.size() == 0) {
            return wb;
        }
        createSheet(wb, list, edf);
        return wb;
    }

    /**
     * 写入一个Sheet
     *
     * @param list     data
     * @param edf      ExcelDataFormatter
     * @param filePath filepath
     * @param <T>      object
     * @throws Exception error
     */
    public static <T> void writeSingleSheetToFile(ExcelDataFormatter edf, String filePath,
                                                  List<T> list) throws Exception {
        Workbook wb = getWorkBook(list, edf);
        // 写入到文件
        FileOutputStream out = new FileOutputStream(filePath);
        wb.write(out);
        out.close();
    }

    /**
     * 写入一个Sheet,不对os进行flush和close,调用者自己实现
     *
     * @param os   OutputStream
     * @param edf  ExcelDataFormatter
     * @param list data
     * @param <T>  object
     * @throws Exception error
     */
    public static <T> void writeSingleSheetToFile(OutputStream os, ExcelDataFormatter edf,
                                                  List<T> list) throws Exception {
        Workbook wb = getWorkBook(list, edf);
        wb.write(os);
    }

    /**
     * 写入多个sheet
     *
     * @param edf      ExcelDataFormatter
     * @param filePath filepath
     * @param lists    data
     * @throws Exception error
     */
    public static void writeMultiSheetToFile(ExcelDataFormatter edf, String filePath,
                                             List<?>... lists) throws Exception {
        Workbook wb = new SXSSFWorkbook();
        for (List list : lists) {
            getWorkBook(wb, list, edf);
        }
        // 写入到文件
        FileOutputStream out = new FileOutputStream(filePath);
        wb.write(out);
        out.close();
    }

    /**
     * 写入多个sheet,不对os进行flush和close,调用者自己实现
     *
     * @param os    OutputStream
     * @param edf   ExcelDataFormatter
     * @param lists data
     * @throws Exception error
     */
    public static void writeMultiSheetToFile(OutputStream os, ExcelDataFormatter edf,
                                             List<?>... lists) throws Exception {
        Workbook wb = new SXSSFWorkbook();
        for (List list : lists) {
            getWorkBook(wb, list, edf);
        }
        wb.write(os);
    }
}
