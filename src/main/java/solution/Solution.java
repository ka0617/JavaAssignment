package solution;

import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

class Fruit
{
    public String name;
    public Double price;
    public Fruit(String name,Double price)
    {
        this.name=name;
        this.price=price;
    }
    public String toString()
    {
        return this.name + " "  + " "+ this.price;
    }
}
class FruitbyName implements Comparator<Fruit>
{
    public int compare(Fruit a, Fruit b)
    {

        return a.name.compareTo(b.name);
    }
}


public class Solution {

    public static ArrayList<Fruit> readXLSXfile(String fileName)
    {
        ArrayList<Fruit> ar = new ArrayList<Fruit>();
        try {

            FileInputStream file = new FileInputStream(new File(fileName));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            System.out.println("Reading Data from Excel file");
            String name="";
            Double price=0.0;
            int rows=0;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:
                            price = cell.getNumericCellValue();
                            break;
                        case Cell.CELL_TYPE_STRING:
                            name = cell.getStringCellValue();
                            break;
                    }
                }
                ar.add(new Fruit(name, price));
            }
            file.close();
            Collections.sort(ar, new FruitbyName());
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        return ar;
    }

    public static void writeXLSXFile(ArrayList<Fruit> ar,String outputFileName) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet spreadsheet = workbook.createSheet("Fruit");

        XSSFCellStyle  style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        int rownum=-1;
        for(int i=-1;i<ar.size();i++)
        {
            Row row = spreadsheet.createRow(++rownum);
            int columnCount = -1;
            if(i==-1)
            {
                Cell cell = row.createCell(++columnCount);
                cell.setCellValue((String) "Fruits");
                cell = row.createCell(++columnCount);
                cell.setCellValue((String) "Price per(Kg)");
                continue;
            }
            Cell cell = row.createCell(++columnCount);
            String name=ar.get(i).name;
            cell.setCellValue((String) name);
            if(ar.get(i).price>=50){
                cell.setCellStyle(style);
            }
            cell = row.createCell(++columnCount);
            Double price=ar.get(i).price;
            cell.setCellValue((Double) price);
            if(ar.get(i).price>=50) {
                cell.setCellStyle(style);
            }
        }
        //Storing into output file
        try (FileOutputStream outputStream = new FileOutputStream(outputFileName)) {
            workbook.write(outputStream);
        }
    }


    public static void main(String[] args) throws IOException {
        Scanner scan=new Scanner(System.in);
        // Asking user input file name
        System.out.println("Enter the path of Input file");
        String inputFileName=scan.next();
        //Reading data into List
        ArrayList<Fruit> ar=readXLSXfile(inputFileName);
        System.out.println("Read Completed");
        // Asking user Output file name
        System.out.println("Enter the path to Store output file");
        String outputFileName=scan.next();
        //Writing Data into Excel from list
        writeXLSXFile(ar,outputFileName);
        System.out.println("Successfully generated XLSX Output file at "+outputFileName);

    }

}