import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Watchable;
import java.util.*;

// list tham khảo https://www.geeksforgeeks.org/reading-writing-data-excel-file-using-apache-poi/
// https://gpcoder.com/3144-huong-dan-doc-va-ghi-file-excel-trong-java-su-dung-thu-vien-apache-poi/
public class Main {
    public static void main(String[] args) throws FileNotFoundException, IOException {
       // writeExcelByMap();
        writeExcelByList();
    }
    public static void writeExcelByMap() throws FileNotFoundException, IOException{
        // tạo 1 workbook mới
        XSSFWorkbook workbook = new XSSFWorkbook();
        // tạo 1 sheet mới trong workbook, đặt tên cho sheet
        XSSFSheet sheet = workbook.createSheet("Test Sheet");
        // tạo 1 collection trống để lưu thông tin, ex sử dụng map
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        //ghi data vào collection sử dụng map thì sài phương thức put()
        data.put("1", new Object[]{"ID", "Tên","Họ"});
        data.put("2", new Object[]{1, "Cường","Nguyễn"});
        data.put("3", new Object[]{2, "Yên","Phạm"});
        data.put("4", new Object[]{3,"Trung","Lê"});
        //data.keySet() trả về kiểu Set<String>, lọc qua các key trong map để lấy value trong map là Object
        Set<String> keyset = data.keySet();
        int rowNum = 0; // số hàng, trong achea poi bắt đầu bằng index 0
        for (String key : keyset){
            //tạo hàng mới ở sheet
            XSSFRow row = sheet.createRow(rowNum++);
            //lấy value trong map thông qua key
            Object[] objects = data.get(key);
            int cellNum = 0;
            //quét qua value rồi ghi xuống theo từng cột trên hàng row
            for (Object obj : objects){
                XSSFCell cell = row.createCell(cellNum++);
                //nếu trong mảng là kiểu String thì ghi xuống cột cell thuộc hàng row kiểu String, ép sang kiểu String
                if (obj instanceof String){
                    cell.setCellValue((String)obj);
                } else if (obj instanceof Integer) {
                    cell.setCellValue((Integer)obj);
                }
            }
        }
        //tạo 1 file mới để ghi xuống
        FileOutputStream fileExcel = new FileOutputStream(new File("data/test_IO_ExcelFlie.xlsx "));
        //ghi file excel xuống ở workbook đã tạo
        workbook.write(fileExcel);
        fileExcel.close();
        System.out.println("Complete!");

    }


    //chuẩn bị data raw cho function writeExcelByList()
    private static String[] columns = {"Name", "Email", "Date of birth", "Days of work", "Salary Per Day", "Total Salary"};
    private static List<Employee> employees = new ArrayList<>();

    // Initializing employees data to insert into the excel file
    static {

        Calendar dateOfBirth = Calendar.getInstance();
        dateOfBirth.set(1995, 0, 8); // 0 means January

        //(String  name, String email, Date dateOfBirth, double daysOfWork, double salaryPerDay, Double   totalSalary)
        employees.add(new Employee("Tubean", "tubean@github.com", dateOfBirth.getTime(), 22, 100d, null));
        dateOfBirth.set(1998, 2, 15);
        employees.add(new Employee("Quynh", "vivichan@gmail.com", dateOfBirth.getTime(), 21, 120d, null));
    }
    public static void writeExcelByList() throws FileNotFoundException, IOException{
        XSSFWorkbook fileExcel = new XSSFWorkbook();
        XSSFSheet sheet = fileExcel.createSheet("Test Sheet 2");
        ///* CreationHelper giúp chúng ta tạo các thể hiện của nhiều thứ khác nhau như DataFormat
        //ở đây sử dụng CreationHelper để giúp định dạng ngày do contructor của Employee có định dạng Date
        CreationHelper createHelper = fileExcel.getCreationHelper();

        //tạo 1 hàng ở index row = 0, để điền thông tin các cột theo định dạng ở mảng
        // columns = {"Name", "Email", "Date of birth", "Days of work", "Salary Per Day", "Total Salary"};
        XSSFRow headerRow = sheet.createRow(0);
        //quét qua từng cột ở mảng
        for (int i =0;i<columns.length;i++){
            //tạo cột ở headerRow
            XSSFCell cell = headerRow.createCell(i);
            //ghi thông tin ở mảng columns = {"Name", "Email", "Date of birth", "Days of work",
            // "Salary Per Day", "Total Salary"} xuống từng ô
            cell.setCellValue(columns[i]);
            //cell.setCellStyle(); Cellstyle() là phương thức use để tạo font, màu...cho excel
        }
        CellStyle dateCellStyle = fileExcel.createCellStyle(); // khai báo style cho cell trong file
        //style set lại phần định dạng ngày cho file bằng cách sử dụng CreationHelper
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));
        //tạo hàng và cột mới, hàng bắt đầu từ index = 1 do index = 0 đã used
        int rowNum = 1;
        //quét hết mảng List<Employee> rồi ghi thông tin Employee xuống
        for (Employee employee : employees){
            XSSFRow row = sheet.createRow(rowNum++);
            //(String  name, String email, Date dateOfBirth, double daysOfWork, double salaryPerDay, Double   totalSalary)
            //ghi lần lượt các cột theo định dạng dữ kiệu employee
            // Employee's name (Column A)
            row.createCell(0).setCellValue(employee.getName());
            // Employee's email (Column B)
            row.createCell(1).setCellValue(employee.getEmail());
            // Employee's date of birth (Column C)
            // Column C sẽ ghi kiểu Date nên cần tìm cách định dạng kiểu Date cho cột C

            Cell dateOfBirth = row.createCell(2);
            dateOfBirth.setCellValue(employee.getDateOfBirth()); // ghi dateOfBirth kiểu Date của đối tượng xuống ô ở cột C, hàng thứ row
            dateOfBirth.setCellStyle(dateCellStyle); //định dạng style ở cột C cho kiểu Date theo form viết từ trước

            // Employee's days of Work (Column D)
            row.createCell(3,CellType.NUMERIC).setCellValue(employee.getDaysOfWork());
            // Employee's salary per day (Column E)
            row.createCell(4,CellType.NUMERIC).setCellValue(employee.getSalaryPerDay());
            // Employee's total salary (Column F = D * E)
            String formula = "D" + rowNum + " * E" + rowNum; //FORMULA: công thức
            //CellType thuộc kiểu FORMULA là kiểu công thức, ví dụ: =D1+E1 (="D" + rowNum + " * E" + rowNum;)
            row.createCell(5, CellType.FORMULA).setCellFormula(formula);
        }
        //điều chỉnh size các cột theo kích thước nội dung
//        for (int i = 0; i < columns.length; i++) {
//            sheet.autoSizeColumn(i);
//        }
        FileOutputStream fileOut = new FileOutputStream("data/test_IO_ExcelFlie.xlsx");
        fileExcel.write(fileOut);
        fileOut.close();
    }
}