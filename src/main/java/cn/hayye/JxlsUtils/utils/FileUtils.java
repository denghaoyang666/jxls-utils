package cn.hayye.JxlsUtils.utils;

import java.io.*;

public class FileUtils {
    public static ByteArrayOutputStream fileToBos(File file) throws IOException {
        BufferedInputStream br = new BufferedInputStream(new FileInputStream(file));
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        int ch = 0;
        while((ch=br.read())!=-1) {
            bos.write(ch);
        }
        return bos;
    }
    public static File getResource(String path){
        if (path.contains("classpath")){
            return new File(path.replace("classpath:",getClassPath()));
        }
        return new File(path);
    }
    public static String getClassPath(){
        return Thread.currentThread().getContextClassLoader().getResource("").getPath();
    }
}
