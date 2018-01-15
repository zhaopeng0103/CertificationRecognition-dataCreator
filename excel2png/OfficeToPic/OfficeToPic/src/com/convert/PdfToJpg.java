package com.convert;

import org.icepdf.core.exceptions.PDFException;
import org.icepdf.core.exceptions.PDFSecurityException;
import org.icepdf.core.pobjects.Document;
import org.icepdf.core.pobjects.Page;
import org.icepdf.core.util.GraphicsRenderingHints;

import javax.imageio.IIOImage;
import javax.imageio.ImageIO;
import javax.imageio.ImageWriter;
import javax.imageio.stream.ImageOutputStream;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class PdfToJpg {
    public static final String FILETYPE_PNG = "png";

    /**
     * 将指定pdf文件的首页转换为指定路径的缩略图
     *
     * @param sourceFile 原文件路径，例如d:/test.pdf
     * @param destFile   图片生成路径，例如 d:/test-1.jpg
     */
    public static int tranfer(String sourceFile, String destFile) throws PDFException, PDFSecurityException, IOException {
        String FileName = sourceFile.substring(sourceFile.lastIndexOf("\\") + 1, sourceFile.lastIndexOf("."));

        Document document = null;
        BufferedImage img = null;
        float rotation = 0f;
        float zoom = 1.5f;

        //判断目录是否存在，如果不存在的话则创建
        File file = new File(destFile);
        if (!file.exists()) {
            file.mkdirs();
        }

        File inputFile = new File(sourceFile);
        if (!inputFile.exists()) {
            System.out.println("找不到源文件");
            return -1;// 找不到源文件, 则返回-1
        }
        document = new Document();

        document.setFile(sourceFile);

        int maxPages = document.getPageTree().getNumberOfPages();
        System.out.println("共计" + maxPages + "页");

        //进行pdf文件图片的转化
        for (int i = 0; i < document.getNumberOfPages(); i++) {
            img = (BufferedImage) document.getPageImage(i, GraphicsRenderingHints.SCREEN,
                    Page.BOUNDARY_CROPBOX, rotation, zoom);
            //设置图片的后缀名
            Iterator iter = ImageIO.getImageWritersBySuffix(FILETYPE_PNG);

            ImageWriter writer = (ImageWriter) iter.next();

            File outFile = new File(destFile + FileName + "_" + (i + 1) + ".png");

            FileOutputStream out = new FileOutputStream(outFile);

            ImageOutputStream outImage = ImageIO.createImageOutputStream(out);

            writer.setOutput(outImage);

            writer.write(new IIOImage(img, null, null));
        }
        img.flush();
        document.dispose();
        System.out.println("转化成功！！！ ");
        return 0;
    }

    public static void main(String[] args) {
        try {
            String sourceFile = "d:/rlesource/12.pdf";
            String destFile = "d:\\s\\";
            PdfToJpg.tranfer(sourceFile, destFile);
        } catch (PDFException e) {
            e.printStackTrace();
        } catch (PDFSecurityException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}