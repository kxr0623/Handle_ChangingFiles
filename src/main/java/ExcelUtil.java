import java.io.FileInputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.*;

//import javax.servlet.ServletOutputStream;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
    /**
     * 将用户的信息导入到excel文件中去
     *
     * @param userList 用户列表
     * @param out      输出表
     */
    public static void exportUserExcel(List<User> userList, OutputStream out) {
        try {
            //1.创建工作簿
            HSSFWorkbook workbook = new HSSFWorkbook();
            //1.1创建合并单元格对象
            CellRangeAddress callRangeAddress = new CellRangeAddress(0, 0, 0, 4);//起始行,结束行,起始列,结束列
            //1.2头标题样式
            HSSFCellStyle headStyle = createCellStyle(workbook, (short) 16);
            //1.3列标题样式
            HSSFCellStyle colStyle = createCellStyle(workbook, (short) 13);
            //2.创建工作表
            HSSFSheet sheet = workbook.createSheet("用户列表");
            //2.1加载合并单元格对象
            sheet.addMergedRegion(callRangeAddress);
            //设置默认列宽
            sheet.setDefaultColumnWidth(25);
            //3.创建行
            //3.1创建头标题行;并且设置头标题
            HSSFRow row = sheet.createRow(0);
            HSSFCell cell = row.createCell(0);

            //加载单元格样式
            cell.setCellStyle(headStyle);
            cell.setCellValue("用户列表");

            //3.2创建列标题;并且设置列标题
            HSSFRow row2 = sheet.createRow(1);
            String[] titles = {"用户名", "账号", "所属部门", "性别", "电子邮箱"};
            for (int i = 0; i < titles.length; i++) {
                HSSFCell cell2 = row2.createCell(i);
                //加载单元格样式
                cell2.setCellStyle(colStyle);
                cell2.setCellValue(titles[i]);
            }

        //4.操作单元格;将用户列表写入excel
        if(userList != null)
        {
            for(int j=0;j<userList.size();j++)
            {
                //创建数据行,前面有两行,头标题行和列标题行
                HSSFRow row3 = sheet.createRow(j+2);
                HSSFCell cell1 = row3.createCell(0);
                cell1.setCellValue(userList.get(j).getName());
                HSSFCell cell2 = row3.createCell(1);
                cell2.setCellValue(userList.get(j).getAccount());
                HSSFCell cell3 = row3.createCell(2);
                cell3.setCellValue(userList.get(j).getDept());
                HSSFCell cell4 = row3.createCell(3);
                cell4.setCellValue(userList.get(j).isGender()?"male":"female");
                HSSFCell cell5 = row3.createCell(4);
                cell5.setCellValue(userList.get(j).getEmail());
            }
        }
        //5.输出
        workbook.write(out);
        workbook.close();
        //out.close();
    }catch(Exception e)
    {
        e.printStackTrace();
    }
    }

    /**
     *
     * @param workbook
     * @param fontsize
     * @return 单元格样式
     */
    private static HSSFCellStyle createCellStyle(HSSFWorkbook workbook, short fontsize) {
        // TODO Auto-generated method stub
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
//        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//水平居中
//        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        //创建字体
        HSSFFont font = workbook.createFont();
//        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setBold(true);
        font.setFontHeightInPoints(fontsize);
        //加载字体
        style.setFont(font);
        return style;
    }

    public void read03and07(String filePath,int sheet_num, int row_num, int cell_num) throws Exception
    {
        //读取03或07的版本
//        String filePath = "D:\\银联国际\\其他主机业务\\SFTP_Configure定时任务变更\\3.12 - 自动填写变更单 - 执行过程优化\\biangengPlan.xls";

        if(filePath.matches("^.+\\.(?i)((xls)|(xlsx))$")) {
            FileInputStream fis = new FileInputStream(filePath);
            boolean is03Excell = filePath.matches("^.+\\.(?i)(xls)$")?true:false;
            Workbook workbook = is03Excell ? new HSSFWorkbook(fis):new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(sheet_num);
            Row row = sheet.getRow(row_num);
            Cell cell = row.getCell(cell_num);
            CellType cell_type=cell.getCellType();
            String cell_value = "";
            switch (cell_type){
                case _NONE:
                    throw new Exception("cannot get the type of cell value");
                case STRING:
                    cell_value=cell.getStringCellValue();
                    break;
                case NUMERIC:
                    // 判断当前的cell是否为Date
//                    TODO:
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        cell_value = sdf.format(date);
                    }
                    else{
                        cell_value = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                case BOOLEAN:
                    cell_value = String.valueOf(cell.getCellType());
                    break;
                case BLANK:
                    cell_value = String.valueOf(cell.getBooleanCellValue());
                    break;
                default:
                    cell_value = "";
                    break;

            }
            System.out.println("第"+row_num+"行第"+cell_num+"列的数据是:"+cell_value);
            /*
            Read all content in the excel sheet 0.
             */
            List<Map<String, Object>> Info = new ArrayList<Map<String, Object>>();
            // 得到Excel的行数
            int totalRows = sheet.getPhysicalNumberOfRows();
            int totalCells=0;
            // 得到Excel的列数(前提是有行数)
            if (totalRows > 1 && sheet.getRow(0) != null) {
                totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
            }
            System.out.println("totalrow:"+totalRows+"  totalCell"+totalCells);
            // 循环Excel行数
            for (int r = 0; r < totalRows; r++) {
                Row row_n = sheet.getRow(r);
                if (row_n == null) {
                    continue;
                }
                // 循环Excel的列
                Map<String, Object> map = new HashMap<String, Object>();
                for (int c = 0; c < totalCells; c++) {
                    Cell cell_n = row_n.getCell(c);
                    if (null != cell_n) {
                        // 如果是纯数字,比如你写的是25,cell.getNumericCellValue()获得是25.0,通过截取字符串去掉.0获得25
                        if (cell_n.getCellType() == CellType.NUMERIC) {
                            String name = "";
                            if (HSSFDateUtil.isCellDateFormatted(cell_n)) {
                                Date date = cell_n.getDateCellValue();
                                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                                name = sdf.format(date);
                            }
                            else{
                                name = String.valueOf(cell_n.getNumericCellValue());
                            }
                            map.put("R"+r+"C"+c , name.substring(0, name.length() - 2 > 0 ? name.length() - 2 : 1));// R1C1
                        } else {
                            map.put("R"+r+"C"+c, cell_n.getStringCellValue());// R1C1
                        }

                    }
                }
                // 添加到list
                Info.add(map);
            }
            String word1="";
            String word2="";
            for (Map<String,Object> item : Info){
                for (String key : item.keySet()) {
                    System.out.print("key:" + key);
                    if (String.valueOf(key).equals("R1C0")){
                        word1=String.valueOf(item.get(key)) ;
                    }
                    else {
                        word2=String.valueOf(item.get(key)) ;
                    }

                    System.out.println("  values:" + item.get(key));
                }
            }
            Random r=new Random();
            int num=(int)(Math.random()*2)+1;
            if (num==1){
                System.out.println("Result:"+word1);
            }else if(num==2) {
                System.out.println("Result:"+word2);
            }else {
                System.out.println("  num:" + num );
            }

        }
        else {
            throw new Exception("This is not a excel file: "+filePath);
        }
    }
}
