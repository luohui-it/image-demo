package com.bestpay.image;


import com.bestpay.restful.common.util.Base64ToImage;
import com.bestpay.restful.common.util.ExcelFileParse;
import org.junit.Test;

import java.io.File;
import java.util.List;

/**
 * 文件解析测试
 *
 * User: luohui Date: 2015/10/15 ProjectName: image-system Version: 1.0
 */
public class ExcelFileParseTest {

    /**
     * 03Excel文件解析
     * @throws Exception
     */
    @Test
    public void getDataFromExcel03Test() throws Exception {
        String filePath = "D://国政通图片.xls";
        File file = new File(filePath);
        ExcelFileParse fileParse = new ExcelFileParse();
        List<Object[]> objects = fileParse.getDataFromExcel03(file, 1);
        for (Object[] object : objects){
            System.out.println(object[1] + ":" + object[2]);
        }
        System.out.println("解析总记录条数："+objects.size());
    }

    /**
     * 07Excel文件解析
     * @throws Exception
     */
    @Test
    public void getDataFromExcel07Test() throws Exception {
        String filePath = "D://国政通图片.xlsx";
        File file = new File(filePath);
        ExcelFileParse fileParse = new ExcelFileParse();
        List<Object[]> objects = fileParse.getDataFromExcel07(file, 1);
        for (Object[] object : objects){
            System.out.println(object[1] + ":" + object[2]);
        }
        System.out.println("解析总记录条数："+objects.size());
    }

    /**
     * Base64转换成图片
     * @throws Exception
     */
    @Test
    public void generateImage07Test() throws Exception {
        String filePath = "D://国政通图片.xlsx";
        File file = new File(filePath);
        ExcelFileParse fileParse = new ExcelFileParse();
        List<Object[]> objects = fileParse.getDataFromExcel07(file, 1);
        for (Object[] object : objects){
            System.out.println(object[0] + ":" + object[2]);
            Base64ToImage.generateImage(object[0].toString(),object[2].toString());
        }
        System.out.println("转换总记录条数："+objects.size());
    }

    /**
     * Base64转换成图片
     * @throws Exception
     */
    @Test
    public void generateImageTest() throws Exception {
        String filePath = "D://只有手持国政通.xls";
        File file = new File(filePath);
        ExcelFileParse fileParse = new ExcelFileParse();
        List<Object[]> objects = fileParse.getDataFromExcel03(file, 1);
        for (Object[] object : objects){
            System.out.println(object[0] + ":" + object[1]);
            Base64ToImage.generateImage(object[0].toString(),object[1].toString());
        }
        System.out.println("转换总记录条数："+objects.size());
    }
}
