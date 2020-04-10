import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import java.io.IOException;
import java.util.*;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class WriteDataToExcel {
    @SuppressWarnings("MismatchedReadAndWriteOfArray")
    private static Object[] header = new Object[9];
    private static int sum = 0;

    @SuppressWarnings("unused")
    static void fillHeader() {
        header[0] = "stamp_type_code";
        header[1] = "stamp_type";
        header[2] = "roll_nbr";
        header[3] = "series";
        header[4] = "first_stamp_nbr";
        header[5] = "last_stamp_nbr";
        header[6] = "qty";
        header[7] = "application_nbr";
        header[8] = "application_date";
    }

    private static String readFile(File file) {
        StringBuilder datab = new StringBuilder();
        try {
            Scanner scanner = new Scanner(file);
            while (scanner.hasNextLine()) {
                datab.append(scanner.nextLine());
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return datab.toString();
    }

    private static void fillInternalTab() {

    }

    private static Set<String> splitScanToRows(String data) {
        String deviding = "";
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < data.length(); i++) {
            if (data.charAt(i) != '&') {
                sb.append(data.charAt(i));
            } else {
                deviding = sb.toString();
                break;
            }
        }
        String[] rows = data.split(deviding);
        rows = ArrayUtils.remove(rows, 0);

        for (int i = 0; i < rows.length; i++) {
            rows[i] = deviding + rows[i];
        }


        return new HashSet<>(Arrays.asList(rows));
    }

    private static Map<String, Object[]> fillTable(Set<String> rows) {
        Map<String, Object[]> table = new HashMap<>();
//        fillHeader();
//        table.put("0", header);
        int valNum = 0;
        int rowNum = 1;
        String value;
        StringBuffer sb = new StringBuffer();
        for (String row : rows) {
            Object[] obj = new Object[9];
            for (int i = 0; i < row.length(); i++) {
                if (row.charAt(i) != '&') {
                    sb.append(row.charAt(i));
                } else if (row.charAt(i) == '&') {
                    if (valNum == 1 || valNum == 7) {
                        sb = new StringBuffer();
                        valNum++;
                    } else if (valNum == 2 || valNum == 5) {
                        valNum++;
                        value = sb.toString();
                        obj[valNum] = Integer.parseInt(value);
                        sb = new StringBuffer();
                        valNum++;
                    } else if (valNum == 0 || valNum == 3 || valNum == 4) {
                        value = sb.toString();
                        obj[valNum] = Integer.parseInt(value);
                        sb = new StringBuffer();
                        valNum++;
                    }
                }
            }

            int n1 = (int) obj[4];
            int n2 = (int) obj[6];
            sum = sum + n2;
            obj[5] = n1 + n2 - 1;
            table.put(Integer.toString(rowNum), obj);
            sb = new StringBuffer();
            valNum = 0;
            rowNum++;
        }
        return table;
    }

    private static void fillTableSum(Map<String, Object[]> table) {
        Object[] obj = new Object[9];
        int size = table.size() + 1;
        assert false;
        table.put(Integer.toString(size), obj);
        table.put(Integer.toString(size++), obj);
        obj = new Object[9];
        obj[6] = sum;
        table.put(Integer.toString(size + 1), obj);
    }

    private static CellStyle cellStyle(XSSFWorkbook workbook, Cell cell) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cell.setCellStyle(style);
        return style;
    }

    private static void addLine(Map<String, Object[]> table){
        Object[] obj = new Object[9];


    }
    private static void addInternalTable(XSSFWorkbook workbook, XSSFSheet spreadsheet) {
        XSSFRow headerRow = spreadsheet.getRow(2);
        XSSFRow valuesRow = spreadsheet.getRow(3);

        Object[] header = new Object[4];
        header[0] = "Литраж";
        header[1] = "ТН ВЭД ТС";
        header[2] = "Кол - во";
        header[3] = "Остаток";

        int cellid = 10;
        for (Object ob : header) {
            Cell cell = headerRow.createCell(cellid++);
            cell.setCellValue((String) ob);
            CellStyle style = cellStyle(workbook, cell);
            cell.setCellStyle(style);
        }

        Object[] values = new Object[4];
        cellid = 10;
        for (Object val : values) {
            Cell cell = valuesRow.createCell(cellid++);

            CellStyle style = cellStyle(workbook, cell);
            cell.setCellStyle(style);
        }

    }

    private static void fillExcel() throws IOException {
        Map<String, Object[]> table;
        Set<String> rows;

        File excelStamp = new File(Objects.requireNonNull(WriteDataToExcel.class
                .getClassLoader().getResource("testEX.xlsx")).getFile());

        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(excelStamp);
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

        assert workbook != null;
        XSSFSheet spreadsheet = workbook.getSheetAt(0);

        XSSFRow row;


        File file = new File(Objects.requireNonNull(WriteDataToExcel.class.getClassLoader().getResource("test.txt")).getFile());
        String data = readFile(file);
        rows = splitScanToRows(data);
        table = fillTable(rows);
        fillTableSum(table);

        //Iterate over data and write to sheet
        Set<String> keyid = table.keySet();
        int rowid = 1;

        for (String key : keyid) {
            row = spreadsheet.createRow(rowid++);
            Object[] objectArr = table.get(key);
            int cellid = 0;

            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellid++);
                if (obj != null) {
                    cell.setCellValue((Integer) obj);
                }
                if (rowid == table.size()-1){
                    CellStyle style = workbook.createCellStyle();
                    style.setBorderBottom(BorderStyle.THIN);
                    style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
                    cell.setCellStyle(style);
                }
            }
        }


        addInternalTable(workbook, spreadsheet);

        //Write the workbook in file system
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(
                    new File("D:\\testRes\\Writesheet.xlsx"));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        try {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        assert out != null;
        out.close();
        System.out.println("Writesheet.xlsx written successfully");
    }


    public static void main(String[] args) throws Exception {
        fillExcel();
    }
}