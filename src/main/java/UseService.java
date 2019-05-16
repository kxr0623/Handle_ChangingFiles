import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.List;

public class UseService{


    public void exportExcel(List<User> userList, OutputStream out) {
        // TODO Auto-generated method stub
        ExcelUtil.exportUserExcel(userList, out);
    }

    public void importExcel(

            File file, String excelFileName) {
        // TODO Auto-generated method stub
        //1.创建输入流
        try {
            FileInputStream inputStream = new FileInputStream(file);
            boolean is03Excel = excelFileName.matches("^.+\\.(?i)(xls)$");
            //1.读取工作簿
            Workbook workbook = is03Excel?new HSSFWorkbook(inputStream):new XSSFWorkbook(inputStream);
            //2.读取工作表
            Sheet sheet = workbook.getSheetAt(0);
            //3.读取行
            //判断行数大于二,是因为数据从第三行开始插入
            if(sheet.getPhysicalNumberOfRows() > 2)
            {
                User user = null;
                //跳过前两行
                for(int k=2;k<sheet.getPhysicalNumberOfRows();k++ )
                {
                    //读取单元格
                    Row row0 = sheet.getRow(k);
                    user = new User();
                    //用户名
                    Cell cell0 = row0.getCell(0);
                    user.setName(cell0.getStringCellValue());
                    //账号
                    Cell cell1 = row0.getCell(1);
                    user.setAccount(cell1.getStringCellValue());
                    //所属部门
                    Cell cell2 = row0.getCell(2);
                    user.setDept(cell2.getStringCellValue());
                    //设置性别
                    Cell cell3 = row0.getCell(3);
                    boolean gender = cell3.getStringCellValue() == "男"?true:false;
                    user.setGender(gender);
                    //设置手机
                    String mobile = "";
                    Cell cell4 = row0.getCell(4);
                    try {
                        mobile = cell4.getStringCellValue();
                    } catch (Exception e) {
                        // TODO Auto-generated catch block
                        double dmoblie = cell4.getNumericCellValue();
                        mobile = BigDecimal.valueOf(dmoblie).toString();
                    }
                    user.setMobile(mobile);
                    //设置电子邮箱
                    Cell cell5 = row0.getCell(5);
                    user.setEmail(cell5.getStringCellValue());
                    //默认用户密码是123456
                    user.setPassword("123456");
                    //用户默认状态是有效
                    user.setState(User.USER_STATE_VALIDE);
                    //保存用户
//                    save(user);
                }
            }
            workbook.close();
            inputStream.close();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }


}
