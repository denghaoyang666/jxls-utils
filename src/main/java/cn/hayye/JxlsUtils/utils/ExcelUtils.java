package cn.hayye.JxlsUtils.utils;

import cn.hayye.JxlsUtils.entity.Excel;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.Properties;

public class ExcelUtils {

    private Excel<T> excel;
    public void setExcel(Excel<T> excel) {
        this.excel = excel;
    }

    /**
     * 生成转换数据为Workbook
     * @param excel
     * @return workbook
     * @throws Exception
     */
    public static Workbook buildExcel(Excel excel) throws Exception {
        //读取模版
        File templateFile = excel.getTemplateFile();
        XLSTransformer transformer = new XLSTransformer();
        //转换数据
        Workbook workbook = transformer.transformXLS(new BufferedInputStream(new FileInputStream(templateFile)), excel.getSheetMap());
        return workbook;
    }

    /**
     * 对指定的WorkBook生成Excel文件
     * @param workbook
     * @param excel
     * @return
     * @throws Exception
     */
    public static File getExcelFile(Workbook workbook,Excel excel,String savePath) throws Exception {
        // 生成文件
        File file = new File(savePath+excel.getFilename());
        if(!file.exists()){
            file.getParentFile().mkdirs();
        }
        //写入数据
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
        return file;
    }

    /**
     * 对通过解析Excel对象直接生成简单的Excel文件
     * @param excel
     * @return
     * @throws Exception
     */
    public static File getExcelFile(Excel excel,String savePath) throws Exception {
        // 创建文件
        File file = new File(savePath+excel.getFilename());
        if(!file.exists()){
            file.getParentFile().mkdirs();
        }
        //写入数据
        OutputStream outputStream = new FileOutputStream(file);
        buildExcel(excel).write(outputStream);
        outputStream.flush();
        outputStream.close();
        return file;
    }

}
