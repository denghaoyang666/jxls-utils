package cn.hayye.JxlsUtils.entity;

import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Excel<T> {
    private File templateFile;
    //String对应模板中遍历的对象名，List为遍历的属性列表，用Map封装代表可同时写数据进入多个sheet
    private Map<String,List<T>> sheetMap;
    private String filename;

    public File getTemplateFile() {
        return templateFile;
    }

    public void setTemplateFile(File templateFile) {
        this.templateFile = templateFile;
        String name = templateFile.getName();
        //this.fileSuffix = name.substring(name.lastIndexOf("."));
        this.filename += name.substring(name.lastIndexOf("."));
    }

    public Map<String, List<T>> getSheetMap() {
        return sheetMap;
    }

    public void setSheetMap(Map<String, List<T>> sheetMap) {
        this.sheetMap = sheetMap;
    }

    public String getFilename() {
        return filename;
    }

    public void setFilename(String filename) {
        this.filename = filename;
    }

    public static class Builder<T>{
        private File templateFile;
        private Map<String,List<T>> sheetMap;
        private String filename;

        public Builder templateFile(File templateFile){
            this.templateFile = templateFile;
            return this;
        }
        public Builder sheetMap(Map<String,List<T>> sheetMap){
            this.sheetMap = sheetMap;
            return this;
        }
        public Builder addSheet(String sheetName , List<T> dateList){
            if(this.sheetMap==null){
                this.sheetMap=new HashMap<>();
            }
            this.sheetMap.put(sheetName,dateList);
            return this;
        }
        public Builder filename(String filename){
            this.filename = filename;
            return this;
        }
        public Excel<T> build(){
            Excel<T> excel = new Excel<>();
            excel.setFilename(filename);
            excel.setSheetMap(sheetMap);
            excel.setTemplateFile(templateFile);
            return excel;
        }
    }
}
