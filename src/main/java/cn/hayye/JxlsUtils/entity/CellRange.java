package cn.hayye.JxlsUtils.entity;

import org.apache.poi.ss.usermodel.Sheet;

public class CellRange {
    private Sheet sheet;
    private int startRow;
    private int endRow;
    private int startCol;
    private int endCol;

    public CellRange() {
    }

    public CellRange(Sheet sheet, int startRow, int endRow, int startCol, int endCol) {
        this.sheet = sheet;
        this.startRow = startRow;
        this.endRow = endRow;
        this.startCol = startCol;
        this.endCol = endCol;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int getEndRow() {
        return endRow;
    }

    public void setEndRow(int endRow) {
        this.endRow = endRow;
    }

    public int getStartCol() {
        return startCol;
    }

    public void setStartCol(int startCol) {
        this.startCol = startCol;
    }

    public int getEndCol() {
        return endCol;
    }

    public void setEndCol(int endCol) {
        this.endCol = endCol;
    }

    public static class Builder {
        private Sheet sheet;
        private int startRow;
        private int endRow;
        private int startCol;
        private int endCol;

        public Builder sheet(Sheet sheet) {
            this.sheet = sheet;
            return this;
        }

        public Builder startRow(int startRow) {
            this.startRow = startRow;
            return this;
        }

        public Builder endRow(int endRow) {
            this.endRow = endRow;
            return this;
        }

        public Builder startCol(int startCol) {
            this.startCol = startCol;
            return this;
        }

        public Builder endCol(int endCol) {
            this.endCol = endCol;
            return this;
        }

        public CellRange build() {
            return new CellRange(sheet, startRow, endRow, startCol, endCol);
        }

    }
}
