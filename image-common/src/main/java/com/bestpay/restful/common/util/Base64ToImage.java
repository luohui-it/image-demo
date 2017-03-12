package com.bestpay.restful.common.util;

import sun.misc.BASE64Decoder;

import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * 将Base64的字符串转换成图片
 *
 * User: luohui Date: 2015/10/15 ProjectName: Base64-Image Version: 1.0
 */
public class Base64ToImage {

    /**
     * 对字节数组字符串进行Base64解码并生成图片
     * 生成的图片放在D盘的picture文件夹下
     *
     * @param imgStr    图片的Base64字符串
     * @return          转换结果
     */
    public static boolean generateImage(String fileName,String imgStr) {
        //图像数据为空
        if (imgStr == null) {
            return false;
        }
        BASE64Decoder decoder = new BASE64Decoder();
        try {
            //Base64解码
            byte[] b = decoder.decodeBuffer(imgStr);
            for (int i = 0; i < b.length; ++i) {
                //调整异常数据
                if (b[i] < 0) {
                    b[i] += 256;
                }
            }
            //生成jpg图片
            String imgFilePath = "D://pictures//"+fileName+".jpg";
            OutputStream out = new FileOutputStream(imgFilePath);
            out.write(b);
            out.flush();
            out.close();
            return true;
        } catch (Exception e) {
            return false;
        }
    }

}
