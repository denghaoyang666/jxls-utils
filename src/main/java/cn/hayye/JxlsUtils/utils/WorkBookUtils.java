package cn.hayye.JxlsUtils.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;
import java.util.Map;

public class WorkBookUtils {

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

    public static void mergeCells(Sheet sheet,int firstRow, int lastRow, int firstCol, int lastCol){
        if(firstRow == lastRow && firstCol == lastCol){
            return;
        }else {
            sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
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
        mergeCells(sheet, startRow, endRow, col, col);
    }

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
