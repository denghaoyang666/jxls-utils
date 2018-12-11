package cn.hayye.JxlsUtils.utils;

import net.coobird.thumbnailator.Thumbnails;
import org.apache.commons.httpclient.HttpClient;
import org.apache.commons.httpclient.HttpMethod;
import org.apache.commons.httpclient.HttpStatus;
import org.apache.commons.httpclient.methods.GetMethod;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;

public class ImgUtils {
    static{
        ImageIO.scanForPlugins();
    }
    /**
     * 根据链接获取图片
     * @param url
     * @return ByteArrayOutputStream
     * @throws IOException
     */
    public static ByteArrayOutputStream getImg(String url) throws Exception {
        HttpClient client = new HttpClient();
        HttpMethod method=new GetMethod(url);
        client.executeMethod(method);
        if(method.getStatusCode()== HttpStatus.SC_OK){
            InputStream inputStream = method.getResponseBodyAsStream();
            BufferedInputStream br = new BufferedInputStream(inputStream);
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            int ch = 0;
            while((ch=br.read())!=-1) {
                bos.write(ch);
            }
            return bos;
        }else {
            throw new Exception(method.getStatusCode()+" Error: "+method.getStatusLine());
        }
    }
    /**
     * 压缩图片至，最大尺寸按定义的width等比缩放，质量缩减原图70%
     * @param fileByte
     * @param format 文件格式(源文件后缀)
     * @param width 缩放到指定宽度，高度等比例缩放
     * @return byte
     * @throws IOException
     */
    public static byte[] thumbImg (byte [] fileByte,String format,int width) {
        BufferedImage sourceImg;
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            sourceImg = ImageIO.read(new ByteArrayInputStream(fileByte));
            if(sourceImg==null) {
                return null;
            }
            BufferedImage targetImg = Thumbnails.of(sourceImg).size(width, width).outputQuality(0.7f).asBufferedImage();
            ImageIO.write(targetImg,format,out);

        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
        return out.toByteArray();
    }

}
