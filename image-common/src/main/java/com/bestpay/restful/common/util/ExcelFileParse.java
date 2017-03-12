package com.bestpay.restful.common.util;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel文件解析
 *
 * User: luohui Date: 2015/10/15 ProjectName: Base64-Image Version: 1.0
 */
public class ExcelFileParse {

    /**
     * 读取excel数据 适用excel2003
     * xls文件后缀
     *
     * @param file                      文件
     * @param startLine                 读取开始行数
     * @throws Exception
     */
    public List<Object[]> getDataFromExcel03(File file, Integer startLine)throws Exception {

        List<Object[]> objects = new ArrayList<>();
        InputStream inputstream = new FileInputStream(file);
        try {
            HSSFWorkbook hssfworkbook = new HSSFWorkbook(inputstream);
            HSSFSheet hssfsheet = hssfworkbook.getSheetAt(0);
            for (int j = 0; j < hssfsheet.getPhysicalNumberOfRows(); j++) {
                if (j < startLine) {
                    continue;
                }
                HSSFRow hssfrow = hssfsheet.getRow(j);
                if (hssfrow == null) {
                    break;
                }
                objects.add(getObject(hssfrow.getLastCellNum(), hssfrow));
            }
            return objects;
        } finally {
            IOUtils.closeQuietly(inputstream);
        }
    }

    /**
     * 读取excel数据 适用excel2007
     * xlsx文件后缀
     *
     * @param file                      文件
     * @param startLine                 读取开始行数
     * @throws Exception
     */
    public List<Object[]> getDataFromExcel07(File file, Integer startLine) throws Exception {

        InputStream inputstream = new FileInputStream(file);
        List<Object[]> objects = new ArrayList<>();
        try {
            XSSFWorkbook wb = new XSSFWorkbook(inputstream);
            XSSFSheet xssfsheet = wb.getSheetAt(0);
            for (int j = 0; j < xssfsheet.getPhysicalNumberOfRows(); j++) {
                //比较和传入的开始读取行数是否符合
                if (j < startLine) {
                    continue;
                }
                //开始循环行数
                XSSFRow xssfrow = xssfsheet.getRow(j);
                if (xssfrow == null) {
                    break;
                }
                objects.add(getObject(xssfrow.getLastCellNum(), xssfrow));
            }
            return objects;
        } finally {
            IOUtils.closeQuietly(inputstream);
        }
    }

    /**
     * 单元格值转换成String
     *
     * @param cell  单元格
     * @return      String值
     */
    private static String doubleFormat(Cell cell) {
        String cellValue;
        if (DateUtil.isCellDateFormatted(cell)) {
            DataFormatter formatter = new DataFormatter();
            cellValue = formatter.formatCellValue(cell);
        } else {
            double value = cell.getNumericCellValue();
            cellValue = new DecimalFormat("0").format(value);
        }

        return cellValue;
    }

    /**
     * 将excel的一行值转换成数组
     *
     * @param collCount     一行的列数
     * @param row           第几行（从第0行算起）
     * @return              转换后的数组结果
     */
    private Object[] getObject(short collCount, Row row) {
        Object[] objects = new Object[collCount];
        for (int k = 0; k < collCount; k++) {
            Cell cell = row.getCell((short) k);
            if (cell != null) {
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        objects[k] = doubleFormat(cell);
                        break;
                    case Cell.CELL_TYPE_STRING:
                        objects[k] = cell.getStringCellValue().trim();
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        objects[k] = cell.getCellFormula().trim();
                        break;
                    case Cell.CELL_TYPE_BLANK:
                        objects[k] = "";
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        objects[k] = cell.getBooleanCellValue();
                        break;
                    case Cell.CELL_TYPE_ERROR:
                        objects[k] = cell.getErrorCellValue();
                        break;
                    default:
                        objects[k] = "";
                        break;
                }
            } else {
                objects[k] = "";
            }
        }
        return objects;
    }

}
