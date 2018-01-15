package com.convert;

import org.icepdf.core.exceptions.PDFException;
import org.icepdf.core.exceptions.PDFSecurityException;

import java.io.IOException;

public class Entry {

    public static void convert() {
        //Office源文件(word,excel,ppt)
        String inputFile = "C:\\Temp\\doc2.doc";
        //目标文件（准备存进哪里）
        String destFile = "C:\\Temp\\doc2.pdf";
        //Img目标文件（图片存储地方）
        String ImgdestFile = "C:\\Temp\\pic2\\";

        //获取返回码
        int code = OfficeToPDF.office2PDF(inputFile, destFile);
        if (code == -1) {
            System.out.println("找不着源文件");
        } else if (code == 1) {
            System.out.println("转换失败");
        } else if (code == 0) {
            System.out.println("转换成功");
            System.out.println("开始把pdf转图片...");
            try {
                PdfToJpg.tranfer(destFile, ImgdestFile);
            } catch (PDFException e) {
                e.printStackTrace();
            } catch (PDFSecurityException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        System.out.println("转换结束。");
    }


    public static void main(String[] arg0) {
        convert();
    }
}