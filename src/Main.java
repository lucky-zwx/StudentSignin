import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDateTime;
import java.util.Scanner;

public class Main {
    private static final String EXCEL_XLS = "xls";

    private static final String EXCEL_XLSX = "xlsx";
    //获得文件操作对象
    static File xlsxFile = new File("./data.xlsx");

    //初始化excel操作对象
    static Workbook workbook;

    static {
        try {
            workbook = getWorkbok(xlsxFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    static XSSFSheet sheet = (XSSFSheet) workbook.getSheet("Sheet1");

    static Scanner input = new Scanner(System.in);

    private static void Writetoexcel() {
        try {
            workbook.write(new FileOutputStream("./data.xlsx"));
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    //学生签到信息录入
    private static void Signin() {
        System.out.println("请输入学生学号：");
        String stuid = input.nextLine();

        System.out.println("请输入学生班级：");
        String stuclass = input.nextLine();

        System.out.println("请输入学生姓名：");
        String stuname = input.nextLine();

        System.out.println("请输入课堂名称：");
        String classname = input.nextLine();

        //签到日期
        String date = LocalDateTime.now().toLocalDate().toString();

        System.out.println("请输入授课教师：");
        String teachername = input.nextLine();

        //签到时间
        String qdtime = LocalDateTime.now().toLocalTime().toString();

        XSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());

        //学生学号写入到excel
        row.createCell(0).setCellValue(stuid);
        //学生班级写入到excel
        row.createCell(1).setCellValue(stuclass);
        //学生姓名写入到excel
        row.createCell(2).setCellValue(stuname);
        //课堂名称
        row.createCell(3).setCellValue(classname);
        //日期
        row.createCell(4).setCellValue(date);
        //授课教师
        row.createCell(5).setCellValue(teachername);
        row.createCell(6).setCellValue("");
        row.createCell(7).setCellValue("");
        row.createCell(8).setCellValue("");
        //签到时间
        row.createCell(9).setCellValue(qdtime);

        Writetoexcel();

    }

    //学生迟到信息录入
    private static void Late() {
        System.out.println("请输入学生学号：");
        String stuid = input.nextLine();
        System.out.println("请输入课堂名称：");
        String classname = input.nextLine();
        System.out.println("请输入日期：格式：(2019-12-28)");
        String date = input.nextLine();
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            String id = sheet.getRow(i).getCell(0).getStringCellValue();
            String class_ = sheet.getRow(i).getCell(3).getStringCellValue();
            String data_ = sheet.getRow(i).getCell(4).getStringCellValue();
            if (id.equals(stuid) && class_.equals(classname) && data_.equals(date)) {
                System.out.println("该学生是否迟到：");
                String cd = input.nextLine();
                sheet.getRow(i).createCell(6).setCellValue(cd);
                Writetoexcel();
                System.out.println("信息录入成功");
                return;
            }
        }
        System.out.println("没有该学生信息或相关的课堂信息");
    }

    //学生请假信息录入
    private static void Leave() {
        System.out.println("请输入学生学号：");
        String stuid = input.nextLine();
        System.out.println("请输入课堂名称：");
        String classname = input.nextLine();
        System.out.println("请输入日期：格式：(2019-12-28)");
        String date = input.nextLine();
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            String id = sheet.getRow(i).getCell(0).getStringCellValue();
            String class_ = sheet.getRow(i).getCell(3).getStringCellValue();
            String data_ = sheet.getRow(i).getCell(4).getStringCellValue();
            if (id.equals(stuid) && class_.equals(classname) && data_.equals(date)) {
                System.out.println("该学生是否请假：");
                String qj = input.nextLine();
                sheet.getRow(i).createCell(7).setCellValue(qj);
                Writetoexcel();
                System.out.println("信息录入成功");
                return;
            }
        }
        System.out.println("没有该学生信息或相关的课堂信息");
    }

    //学生早退信息录入
    private static void Leaveearly() {
        System.out.println("请输入学生学号：");
        String stuid = input.nextLine();
        System.out.println("请输入课堂名称：");
        String classname = input.nextLine();
        System.out.println("请输入日期：格式：(2019-12-28)");
        String date = input.nextLine();
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            String id = sheet.getRow(i).getCell(0).getStringCellValue();
            String class_ = sheet.getRow(i).getCell(3).getStringCellValue();
            String data_ = sheet.getRow(i).getCell(4).getStringCellValue();
            if (id.equals(stuid) && class_.equals(classname) && data_.equals(date)) {
                System.out.println("该学生是否早退：");
                String zt = input.nextLine();
                sheet.getRow(i).createCell(8).setCellValue(zt);
                Writetoexcel();
                System.out.println("信息录入成功");
                return;
            }
        }
        System.out.println("没有该学生信息或相关的课堂信息");
    }

    //学生相关信息查询
    private static void Search() {
        System.out.println("请输入学生学号：");
        String stuid = input.nextLine();
        boolean searched = false;
        for (int i = 1; i < sheet.getPhysicalNumberOfRows()+1; i++) {
            if (sheet.getRow(i) == null){
                continue;
            }
            String id = sheet.getRow(i).getCell(0).getStringCellValue();
            if (id.equals(stuid)) {
                searched = true;
                System.out.println("\t学生班级："+sheet.getRow(i).getCell(1).getStringCellValue()+"\t");
                System.out.println("\t学生姓名："+sheet.getRow(i).getCell(2).getStringCellValue()+"\t");
                System.out.println("\t课堂名称："+sheet.getRow(i).getCell(3).getStringCellValue()+"\t");
                System.out.println("\t日期："+sheet.getRow(i).getCell(4).getStringCellValue()+"\t");
                System.out.println("\t授课教师："+sheet.getRow(i).getCell(5).getStringCellValue()+"\t");
                System.out.println("\t是否迟到："+sheet.getRow(i).getCell(6).getStringCellValue()+"\t");
                System.out.println("\t是否请假："+sheet.getRow(i).getCell(7).getStringCellValue()+"\t");
                System.out.println("\t是否早退："+sheet.getRow(i).getCell(8).getStringCellValue()+"\t");
                System.out.println("\t签到时间："+sheet.getRow(i).getCell(9).getStringCellValue()+"\t");
            }
        }
        if (!searched){
            System.out.println("没有该学生信息或相关的课堂信息");
        }
    }

    //学生相关信息删除
    private static void Delete() {
        System.out.println("请输入学生学号：");
        String stuid = input.nextLine();
        System.out.println("请输入课堂名称：");
        String classname = input.nextLine();
        System.out.println("请输入日期：格式：(2019-12-28)");
        String date = input.nextLine();
        boolean searched = false;
        for (int i = 1; i < sheet.getPhysicalNumberOfRows()+1; i++) {
            if (sheet.getRow(i) == null){
                continue;
            }
            String id = sheet.getRow(i).getCell(0).getStringCellValue();
            String class_ = sheet.getRow(i).getCell(3).getStringCellValue();
            String data_ = sheet.getRow(i).getCell(4).getStringCellValue();
            if (id.equals(stuid) && class_.equals(classname) && data_.equals(date)) {
                searched = true;
                sheet.removeRow(sheet.getRow(i));
                Writetoexcel();
                return;
            }
        }
        if (!searched){
            System.out.println("没有该学生信息或相关的课堂信息");
        }
    }

    //学生相关信息修改
    private static void Update() {
        System.out.println("请输入学生学号：");
        String stuid = input.nextLine();
        System.out.println("请输入课堂名称：");
        String classname = input.nextLine();
        System.out.println("请输入日期：格式：(2019-12-28)");
        String date = input.nextLine();
        boolean searched = false;
        for (int i = 1; i < sheet.getPhysicalNumberOfRows()+1; i++) {
            if (sheet.getRow(i) == null){
                continue;
            }
            String id = sheet.getRow(i).getCell(0).getStringCellValue();
            String class_ = sheet.getRow(i).getCell(3).getStringCellValue();
            String data_ = sheet.getRow(i).getCell(4).getStringCellValue();
            if (id.equals(stuid) && class_.equals(classname) && data_.equals(date)) {
                searched = true;
                System.out.println("请选择你需要修改的选项：\n" +
                        "\t1.学生迟到信息\n" +
                        "\t2.学生请假信息\n" +
                        "\t3.学生早退信息");
                String selectid = input.nextLine();
                switch (selectid){
                    case "1":
                        System.out.println("请输入你的修改信息：");
                        String cd = input.nextLine();
                        sheet.getRow(i).getCell(6).setCellValue(cd);
                        break;
                    case "2":
                        System.out.println("请输入你的修改信息：");
                        String qj = input.nextLine();
                        sheet.getRow(i).getCell(7).setCellValue(qj);
                        break;
                    case "3":
                        System.out.println("请输入你的修改信息：");
                        String zt = input.nextLine();
                        sheet.getRow(i).getCell(7).setCellValue(zt);
                        break;
                    default:
                        System.out.println("请检查序号后重新输入!!!");
                        break;
                }
                Writetoexcel();
            }
        }
        if (!searched){
            System.out.println("没有该学生信息或相关的课堂信息");
        }
    }

    private static void Saveall() {
        System.out.println("感谢使用!");
        System.exit(1);
    }

    public static void main(String[] args) {
        System.out.println("\t\t欢迎使用学生签到系统\t\t\n" +
                "\t1.学生签到信息录入\n" +
                "\t2.学生迟到信息录入\n" +
                "\t3.学生请假信息录入\n" +
                "\t4.学生早退信息录入\n" +
                "\t5.学生相关信息查询\n" +
                "\t6.学生相关信息删除\n" +
                "\t7.学生相关信息修改\n" +
                "\t8.退出并保存");
        while (true) {
            System.out.println("请选择功能序号：");
            String id = input.nextLine();
            switch (id) {
                case "1":
                    Signin();
                    break;
                case "2":
                    Late();
                    break;
                case "3":
                    Leave();
                    break;
                case "4":
                    Leaveearly();
                    break;
                case "5":
                    Search();
                    break;
                case "6":
                    Delete();
                    break;
                case "7":
                    Update();
                    break;
                case "8":
                    Saveall();
                    break;
                default:
                    System.out.println("请检查序号后重新输入!!!");
                    break;
            }
        }
    }

    protected void finalize() {
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static Workbook getWorkbok(File file) throws IOException {
        Workbook wb = null;

        FileInputStream in = new FileInputStream(file);

        if (file.getName().endsWith(EXCEL_XLS)) {    //Excel 2003
            wb = new HSSFWorkbook(in);

        } else if (file.getName().endsWith(EXCEL_XLSX)) {   // Excel 2007/2010
            wb = new XSSFWorkbook(in);
        }

        return wb;
    }
}