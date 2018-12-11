package cn.hayye;
import cn.hayye.JxlsUtils.entity.Excel;
import cn.hayye.JxlsUtils.utils.ExcelUtils;
import cn.hayye.JxlsUtils.utils.FileUtils;
import cn.hayye.JxlsUtils.utils.WorkBookUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.LinkedList;
import java.util.List;

public class Test {
    public static void main(String[] args) throws Exception {

        Excel<Person> excel = new Excel.Builder<Person>().addSheet("personList",getData())
                                                        .filename("测试")
                                                        .templateFile(FileUtils.getResource("classpath:template.xlsx"))
                                                        .build();
        Workbook workbook = ExcelUtils.buildExcel(excel);
        WorkBookUtils.merageColumnSameCells(workbook.getSheetAt(0),1);
       // WorkBookUtils.merageRowSameCells(workbook.getSheetAt(0),3);
        Cell cell = workbook.getSheetAt(0).getRow(1).getCell(6);
        WorkBookUtils.insertPicOL(cell,200,200,"http://blog.hayye.cn/images/avatar.jpg");
        ExcelUtils.getExcelFile(workbook,excel,"E:/Test/");
    }

    public static List<Person> getData(){
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

        return personList;
    }
}
