/*
 * This Java source file was generated by the Gradle 'init' task.
 */
package poi.excel;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import poi.excel.vo.ExcelUtilVO;
import poi.excel.vo.TargetVO;

public class App {
    public String getGreeting() {
        return "Hello world.";
    }

    public static void main(String[] args) {
        System.out.println(new App().getGreeting());

        File file = new File("/Users/tarpha/Documents/git/poi-excel/test.xlsx");

        List<ExcelUtilVO> list = new ArrayList<ExcelUtilVO>();

        ExcelUtilVO vo = new ExcelUtilVO();
        vo.setSourceFieldName("Field2");
        vo.setTargetFieldName("targetFieldName");
        vo.setTargetFieldType(String.class);

        list.add(vo);

        vo = new ExcelUtilVO();
        vo.setSourceFieldName("Field1");
        vo.setTargetFieldName("sourceFieldName");
        vo.setTargetFieldType(String.class);

        list.add(vo);

        TargetVO trg = new TargetVO();

        List<List<ExcelUtilVO>> arrList = new ArrayList<List<ExcelUtilVO>>();

        arrList.add(list);

        list = new ArrayList<ExcelUtilVO>();

        vo = new ExcelUtilVO();
        vo.setSourceFieldName("Field7");
        vo.setTargetFieldName("targetFieldName");
        vo.setTargetFieldType(String.class);

        list.add(vo);

        vo = new ExcelUtilVO();
        vo.setSourceFieldName("Field8");
        vo.setTargetFieldName("sourceFieldName");
        vo.setTargetFieldType(String.class);

        list.add(vo);

        arrList.add(list);

        list = new ArrayList<ExcelUtilVO>();

        vo = new ExcelUtilVO();
        vo.setSourceFieldName("Field9");
        vo.setTargetFieldName("targetFieldName");
        vo.setTargetFieldType(String.class);

        list.add(vo);

        vo = new ExcelUtilVO();
        vo.setSourceFieldName("Field10");
        vo.setTargetFieldName("sourceFieldName");
        vo.setTargetFieldType(String.class);

        list.add(vo);

        arrList.add(list);

        try {
            int cnt = ExcelUtil.executeServiceListArg(file, 0, arrList, ExcelUtilVO.class, trg, "setListCnt", 100);
            System.out.println("count:" + cnt);
        } catch (Exception e) {
            e.printStackTrace();
        }

        //System.out.println("trg: " + ((ExcelUtilVO)trg.getList().get(0)).getTargetFieldName());
    }
}