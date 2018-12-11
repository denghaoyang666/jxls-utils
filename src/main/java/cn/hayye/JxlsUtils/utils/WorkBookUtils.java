package cn.hayye.JxlsUtils.utils;

import cn.hayye.JxlsUtils.entity.CellRange;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class WorkBookUtils {

    public static final String meragecolumn = "column";
    public static final String meragerow = "row";
    /**
     * 设置Url链接
     * @param cell
     * @param url
     */
    public static void setUrl(Cell cell, String url) {
        CreationHelper creationHelper = cell.getRow().getSheet().getWorkbook().getCreationHelper();
        setUrl(cell, url, creationHelper);
    }

    /**
     *
     * @param cell
     * @param url
     * @param creationHelper
     */
    public static void setUrl(Cell cell, String url,CreationHelper creationHelper){
        // 链接写入单元格
        Hyperlink link = creationHelper.createHyperlink(Hyperlink.LINK_URL);
        link.setAddress(url);
        cell.setHyperlink(link);
        cell.setCellValue(url);
    }

    public static void mergeCells(CellRange cellRange) {
        if (cellRange.getStartRow() == cellRange.getEndRow() && cellRange.getEndCol() == cellRange.getStartCol()) {
            return;
        } else {
            cellRange.getSheet().addMergedRegion(new CellRangeAddress(cellRange.getStartRow(), cellRange.getEndRow(), cellRange.getEndCol(), cellRange.getStartCol()));
        }
    }
    public static void mergeCells(Sheet sheet,int startRow,int endRow, int startCol, int endCol) {
        if (startRow == endRow && startCol == endCol) {
            return;
        } else {
            sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, startCol, endCol));
        }
    }

    /**
     * 合并指定列上相同的单元格
     * @param sheet
     * @param col
     */
    public static void merageColumnSameCells(Sheet sheet,int col){
        int totalRow = sheet.getLastRowNum();
        if (totalRow<0){
            return;
        }
        StringBuilder sb = new StringBuilder(getCellValue(sheet.getRow(0).getCell(col)).toString());
        int startRow=0;
        int endRow=0;
        for(int i=1;i<=totalRow;i++){
            Cell cell = sheet.getRow(i).getCell(col);
            String cellValue = getCellValue(cell).toString();
            if(cellValue!=""){
                if(cellValue.equals(sb.toString())){
                    cell.setCellValue("");
                    endRow=i;
                }else {
                    mergeCells(sheet, startRow, endRow, col, col);
                    startRow = i;
                    endRow = i;
                    sb.setLength(0);
                    sb.append(cellValue);
                }
            }
        }
        mergeCells(sheet,startRow, endRow, col, col);
    }

    /**
     * 合并指定行上相同的单元格
     * @param sheet
     * @param r
     */
    public static void merageRowSameCells(Sheet sheet,int r){
        Row row = sheet.getRow(r);
        int totalColomn = row.getPhysicalNumberOfCells();
        if (totalColomn<0){
            return;
        }
        StringBuilder sb = new StringBuilder(getCellValue(row.getCell(0)).toString());
        int startCol=0;
        int endCol=0;
        for(int i=1;i<totalColomn-1;i++){
            Cell cell = row.getCell(i);
            String cellValue = getCellValue(cell).toString();
            if(cellValue!=""){
                if(cellValue.equals(sb.toString())){
                    cell.setCellValue("");
                    endCol=i;
                }else {
                    mergeCells(sheet,r,r,startCol,endCol);
                    startCol = i;
                    endCol = i;
                    sb.setLength(0);
                    sb.append(cellValue);
                }
            }
        }
        mergeCells(sheet,r,r,startCol,endCol);
    }
/*
    public static void merageSameCells(Sheet sheet,int index,String rc){
        int total = -1;
        StringBuilder sb = new StringBuilder();
        switch (rc) {
            case meragecolumn:
                total = sheet.getLastRowNum();
                sb.append(getCellValue(sheet.getRow(0).getCell(index)).toString());
                break;
            case meragerow:
                total = sheet.getRow(index).getPhysicalNumberOfCells();
                sb.append(getCellValue(sheet.getRow(index).getCell(0)).toString());
                break;
            default:
                return;
        }
        if (total<0){
            return;
        }
        int start=0;
        int end=0;
        for(int i=1;i<total-1;i++){
            Cell cell = row.getCell(i);
            String cellValue = getCellValue(cell).toString();
        }
    }*/

    public static void insertPicOL(Cell cell,int width,int height,String url) throws Exception {
        ByteArrayOutputStream bos = ImgUtils.getImg(url);
        insertPic(cell,width,height,bos);
    }
    public static void insertPicFile(Cell cell, int width, int height, File file) throws Exception {
        ByteArrayOutputStream bos = FileUtils.fileToBos(file);
        insertPic(cell,width,height,bos);
    }

    public static void insertPic(Cell cell,int width,int height,ByteArrayOutputStream bos) throws Exception {
        Sheet sheet =  cell.getSheet();
        Drawing drawing = sheet.createDrawingPatriarch();
        int column = cell.getColumnIndex();
        int row = cell.getRowIndex();
        //插入位置
        ClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, column, row, column+1, row+1);
        // 设置单元格大小
        sheet.setColumnWidth(column, width * 32); // 解释：n*32则为n像素长度，x*256则为x实际宽度。
        sheet.getRow(row).setHeight((short) (height * 15)); // 解释：n*15则为x像素长度，x*20则为x实际高度。
        drawing.createPicture(anchor,sheet.getWorkbook().addPicture(bos.toByteArray(), Workbook.PICTURE_TYPE_JPEG));

    }

    public static Object getCellValue(Cell cell){
        switch (cell.getCellType()){
            case Cell.CELL_TYPE_NUMERIC:
                return cell.getNumericCellValue();
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            case Cell.CELL_TYPE_FORMULA:
                return cell.getCellFormula();
            case Cell.CELL_TYPE_BLANK:
                return "";
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_ERROR:
                return cell.getErrorCellValue();
            default:
                try{
                    return cell.getDateCellValue();
                }catch (Exception e){
                    try{
                        return cell.getRichStringCellValue().getString();
                    }catch (Exception e2){
                        return "";
                    }
                }
        }
    }

}
