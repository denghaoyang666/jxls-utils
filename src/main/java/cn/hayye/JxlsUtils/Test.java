package cn.hayye.JxlsUtils;

import cn.hayye.JxlsUtils.entity.Excel;
import cn.hayye.JxlsUtils.utils.ExcelUtils;
import cn.hayye.JxlsUtils.utils.WorkBookUtils;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.util.LinkedList;
import java.util.List;

public class Test {
    public static void main(String[] args) throws Exception {
        List<Person> personList = new LinkedList<>();
        personList.add(new Person(1L,"张三","张三","男",10,"10"));
        personList.add(new Person(2L,"李四","张三","男",10,""));
        personList.add(new Person(3L,"张三","张三","男",10,""));
        personList.add(new Person(4L,"张三","张三","男",10,""));
        personList.add(new Person(5L,"张三","张三","男",10,""));
        personList.add(new Person(6L,"李四","张三","男",10,""));
        personList.add(new Person(7L,"张三","张三","男",10,""));
        personList.add(new Person(8L,"张三","张三","男",10,""));
        personList.add(new Person(9L,"李四","张三","男",10,""));
        personList.add(new Person(10L,"张三","张三","男",10,""));
        personList.add(new Person(11L,"李四","张三","男",10,""));
        personList.add(new Person(12L,"李四","张三","男",10,""));
        Excel<Person> excel = new Excel.Builder<Person>().addSheet("personList",personList)
                                                        .filename("p")
                                                        .templateFile(new File("E:/Test/test.xlsx"))
                                                        .build();
        Workbook workbook = ExcelUtils.buildExcel(excel);
        WorkBookUtils.merageColumnSameCells(workbook.getSheetAt(0),1);
        WorkBookUtils.merageRowSameCells(workbook.getSheetAt(0),3);
        ExcelUtils.getExcelFile(workbook,excel,"E:/Test/");
    }
}
