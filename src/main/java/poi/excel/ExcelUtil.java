package poi.excel;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import poi.excel.vo.ExcelUtilVO;

/**
 * Excel 공통 Library
 * 
 * @author LJW
 */
public class ExcelUtil {

    /**
     * Excel File을 읽어서 List를 인자로 받는 Service를 수행한다.
     * 
     * @author LJW
     * @param file            입력 Excel file
     * @param sheetInfo       대상 sheet name or index
     * @param mapperList      Excel dataset과 vo를 mapping하는 정보를 포함하는 vo list
     *                        하나의 list가 dataset에 해당하며, 여러 dataset을 지원하기 위하여 list의
     *                        list로 구성한다.
     * @param voClass         Method의 입력 List의 대상 Class형
     * @param service         Method 수행의 대상이 되는 instance
     * @param methodName      Method 명. 입력 인자는 List형식만 지원한다. (Service Method 명)
     * @param batchInsertSize Service의 List 인자에 넣을 최대 건 수. 초과 시 Service를 수행 후 비운다.
     * @return Service에서 리턴한 int 값의 합계
     * @throws IOException
     * @throws InvalidFormatException
     * @throws SecurityException
     * @throws NoSuchMethodException
     * @throws InvocationTargetException
     * @throws IllegalArgumentException
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    public static int executeServiceListArg(File file, Object sheetInfo, List<List<ExcelUtilVO>> mapperList, Class<?> voClass,
            Object service, String methodName, int batchInsertSize)
            throws InvalidFormatException, IOException, NoSuchMethodException, SecurityException,
            InstantiationException, IllegalAccessException, IllegalArgumentException, InvocationTargetException
             {
        
        //insert count (for return)
        int count = 0;

        //try-resource
        try (XSSFWorkbook workbook = new XSSFWorkbook(file);) {
            XSSFSheet sheet = null;

            //sheet명을 입력 받았을 경우
            if(sheetInfo instanceof String)
                sheet = workbook.getSheet((String)sheetInfo);
            //sheet index를 입력 받았을 경우
            else if (sheetInfo instanceof Integer)
                sheet = workbook.getSheetAt((int)sheetInfo);
            
            //sheet의 행 수를 가져온다.
            int rows = sheet.getLastRowNum();

            //batch insert 하기위한 list
            List<Object> list = new ArrayList<Object>();

            //batch insert 하기위한 method (입력인자로 동적 생성)
            Method serviceMethod = service.getClass().getMethod(methodName, List.class);

            /* sheet의 header column 명과 target vo의 변수 명의 리스트
              하나의 List가 dataset에 mapping 된다.
              sheet에 여러 dataset이 존재할 경우를 대비하여 List배열로 입력받는다. */
            for(List<ExcelUtilVO> mapper : mapperList) {
                //row looping
                for(int i = 0; i < rows; i++) {
                    //해당 row의 유효한 열 갯수
                    short cols = 0;
                    XSSFRow row = sheet.getRow(i);
                    if(row != null) cols = row.getLastCellNum();
                    
                    //인자로 입력받은 VO Class로 새 Instance를 생성한다. (Insert하기 위한 VO 생성)
                    Object voInstance = voClass.getDeclaredConstructor().newInstance();

                    //insert list에 추가할지 여부
                    Boolean valid = false;

                    //column looping
                    for(short j = 0; j < cols; j++) {
                        XSSFCell cell = row.getCell(j);
                        
                        //빈 cell일 경우 skip
                        if(cell == null || cell.getCellType() == CellType.BLANK || cell.getCellType() == CellType._NONE)
                            continue;
                        
                        //mapper에 정의된 field를 찾아서 vo에 넣어준다.
                        for(ExcelUtilVO vo : mapper) {
                            String stringValue = (String)getCellValue(cell, String.class);

                            //cell의 값이 없을 경우 skip
                            if(StringUtils.isBlank(stringValue)) 
                                continue;
                            
                            //System.out.println(">" + vo.getSourceFieldName() + ":" + stringValue);
                            
                            //header field를 찾거나, 이미 찾은 filed 인지 check 한다.
                            if(vo.getSourceFieldName().equals(stringValue) || vo.getIndex() == j) {
                                //header field일 경우 해당 col index를 vo에 저장하여 다음 row부터 value값을 vo에 넣어준다.
                                if(vo.getIndex() == -1) {
                                    vo.setIndex(j);
                                    continue;
                                }
                                
                                //인자로 받은 vo class로부터 setter를 동적 생성하여 넣어준다.
                                Method method = voClass.getMethod(getSetterName(vo.getTargetFieldName()), vo.getTargetFieldType());
                                method.invoke(voInstance, getCellValue(cell, vo.getTargetFieldType()));

                                //insert list에 추가할지 여부
                                valid = true;

                                //System.out.println(stringValue + ":" + vo.getIndex() + ":" + vo.getSourceFieldName());
                            }
                        }
                    }

                    //insert list에 추가
                    if(valid) list.add(voInstance);

                    //list의 크기가 인자로 받은 batch insert size보다 클 경우 insert method를 수행하고 list를 비운다.
                    if(list.size() > batchInsertSize) {
                        count += (int)serviceMethod.invoke(service, list);
                        list.clear();;
                    }
                }
            }

            //잔여 건이 있을 경우 Insert
            if(list.size() > 0)
                count += (int)serviceMethod.invoke(service, list);

        }

        return count;
    }

    /**
     * 대상 형식의 cell의 값을 가져온다.
     * 
     * @author                  LJW
     * @param cell              대상 cell
     * @param type              대상 형 (int/double/String/BigDecimal)
     * @return                  대상 형식의 cell의 값
     */
    public static Object getCellValue(XSSFCell cell, Class<?> type) {
        switch(cell.getCellType()) {
            case STRING:
                return getTypeValue(cell.getStringCellValue(), type);
            case NUMERIC:
                return getTypeValue(cell.getNumericCellValue(), type);
            case BOOLEAN:
                return getTypeValue(cell.getBooleanCellValue(), type);
            default:
                return getTypeValue(cell.getRawValue(), type);
        }
    }

    /**
     * 대상 형식의 object의 값을 가져온다.
     * 
     * @author                  LJW
     * @param obj               대상 object
     * @param type              대상 형 (int/double/String/BigDecimal)
     * @return                  대상 형식의 object 값
     */
    private static Object getTypeValue(Object obj, Class<?> type) {
        if(type.equals(int.class)) {
            if(obj instanceof String) return Integer.parseInt((String)obj);
            else return (int)obj;
        } else if(type.equals(double.class)) {
            if(obj instanceof String) return Double.parseDouble((String)obj);
            else return (double)obj;
        } else if(type.equals(BigDecimal.class)) {
            return new BigDecimal(String.valueOf(obj));
        } else {
            if(obj instanceof String) return (String)obj;
            else return String.valueOf(obj);
        }
    }

    /**
     * setter method 명을 가져온다.
     * 
     * @author                  LJW
     * @param name              setter 대상 member 변수명
     * @return                  setter 형식 method 명
     */
    private static String getSetterName(String name) {
        return "set" + StringUtils.capitalize(name);
    }

}