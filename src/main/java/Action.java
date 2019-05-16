public class Action {
//    public void exportExcel()
//    {
//        try
//        {
//            //1.查找用户列表
//            userList = userService.findObjects();
//            //2.导出
//            HttpServletResponse response = ServletActionContext.getResponse();
//            //这里设置的文件格式是application/x-excel
//            response.setContentType("application/x-excel");
//            response.setHeader("Content-Disposition", "attachment;filename=" + new String("用户列表.xls".getBytes(), "ISO-8859-1"));
//            ServletOutputStream outputStream = response.getOutputStream();
//            userService.exportExcel(userList, outputStream);
//            if(outputStream != null)
//                outputStream.close();
//        }catch(Exception e)
//        {
//            e.printStackTrace();
//        }
//    }

//    public String importExcel()
//    {
//        if(userExcel!= null)
//        {
//            //判断是否是Excel文件
//            if(userExcelFileName.matches("^.+\\.(?i)((xls)|(xlsx))$"))
//            {
//                userService.importExcel(userExcel, userExcelFileName);
//            }
//        }
//        return"list";
//    }
    public static void main(String[] args){
        ExcelUtil myexcel = new ExcelUtil();
        try {
            String filePath = "D:\\银联国际\\其他主机业务\\SFTP_Configure定时任务变更\\3.12 - 自动填写变更单 - 执行过程优化\\biangengPlan.xls";
            myexcel.read03and07(filePath,0,2,0);
        }catch (Exception e){
            System.out.println("Exception thrown  :" +e);
        }

    }


}
