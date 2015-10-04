import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import com.sun.deploy.util.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Created by AndreyNadya on 10/4/2015.
 */
public class ParseXlsxFile2Csv {

    private File excelFile;
    private String outputPath="";
    private List<List<String>> table = new ArrayList<List<String>>();

    private void parseArgs(String[] args) {
        if (args.length<1) {
            System.out.println("Please provide the path to the excel file and optional output path");
            System.exit(-1);
        }
        excelFile = new File(args[0]);
        if (args.length>1) outputPath = args[1];
    }

    private void doMain(String[] args) throws IOException {
        parseArgs(args);
        FileInputStream fis = new FileInputStream(excelFile);
        XSSFWorkbook book = new XSSFWorkbook(fis);
        PrintWriter pw=null;
        for (int i=0; i<book.getNumberOfSheets(); i++) {

            XSSFSheet sheet = book.getSheetAt(i);
            Iterator<Row> itr = sheet.iterator(); // Iterating over Excel file in Java
            pw = new PrintWriter(outputPath + File.separator + excelFile.getName() +  i +".csv");
            while (itr.hasNext()) {
                Row row = itr.next();
                // Iterating over each column of Excel file
                Iterator<Cell> cellIterator = row.cellIterator();
                List<String> rowData = new ArrayList<String>();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    rowData.add(cell.toString());
//                    switch (cell.getCellType()) {
//                        case Cell.CELL_TYPE_STRING: rowData.add(cell.getStringCellValue()); break;
//                        case Cell.CELL_TYPE_NUMERIC: rowData.add(String.valueOf(cell.getNumericCellValue())); break;
//                        case Cell.CELL_TYPE_BOOLEAN: rowData.add(String.valueOf(cell.getBooleanCellValue())); break;
//                        default: }
                }
                pw.println(StringUtils.join(rowData, ","));
            }
            pw.close();
        }

    }
    public static void main(String[] args) throws IOException {
        new ParseXlsxFile2Csv().doMain(args);
    }
}
