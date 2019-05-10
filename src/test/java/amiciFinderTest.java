import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;
import java.util.*;

public class amiciFinderTest {

    private static final String INPUT_WKBK_PATH = "C:\\Users\\mattm\\OneDrive\\Documents\\petitions.xlsx";
    private static final String OUTPUT_WKBK_PATH = "C:\\Users\\mattm\\OneDrive\\Desktop\\Output.xlsx";
    private static final String TILDA_SPLIT = "~~~Date~~~  ~~~~~~~Proceedings and Orders~~~~~~~~~~~~~~~~~~~~~;";

    private static final String REGEX_PARSER = "(?<=\\d{4})\\s{2}(?=[A-Z])|(?<=\\d|\\)|\\.);|;(?=[A-Z])|(?<=\\d)\\.(?=[A-Z])";

    @Test
    public void amiciFinder() {

        try {

            FileInputStream excelFile = new FileInputStream(new File(INPUT_WKBK_PATH));
            Workbook inputWkbk = new XSSFWorkbook(excelFile);
            Sheet inputSheet = inputWkbk.getSheetAt(0);
            Iterator<Row> iterator = inputSheet.iterator();
            Map<Integer, Object[]> finalData = new TreeMap<Integer, Object[]>();
            int keys = 0;

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {

                    String cellText = cellIterator.next().getStringCellValue();

                    if (cellText.contains("Brief amici curiae") || cellText.contains("Brief amicus curiae")) {

                        if (cellText.contains(TILDA_SPLIT)) {

                            cellText = cellText.split(TILDA_SPLIT)[1];
                        }

                        String cellTextArray[] = cellText.split(REGEX_PARSER);

                        for (int i = 0; i < cellTextArray.length; i++) {

                            if (cellTextArray[i].contains("Brief amici curiae") || cellTextArray[i].contains("Brief amicus curiae")) {

                                finalData.put(keys++, new Object[]{currentRow.getCell(0).toString(), cellTextArray[i - 1], cellTextArray[i]});
                                System.out.println(currentRow.getCell(0).toString() + " " + cellTextArray[i - 1] + " " + cellTextArray[i]);
                            }
                        }
                    }
                }
            }

            XSSFWorkbook outputWkbk = new XSSFWorkbook();
            XSSFSheet outputSheet = outputWkbk.createSheet("Parsed Data");
            Set<Integer> keySet = finalData.keySet();

            for (Integer key : keySet) {

                Row row = outputSheet.createRow(key);
                Object[] objArr = finalData.get(key);
                int cellnum = 0;

                for (Object obj : objArr) {

                    Cell cell = row.createCell(cellnum++);

                    if (obj instanceof String) {

                        cell.setCellValue((String)obj);

                    } else if (obj instanceof Integer) {

                        cell.setCellValue((Integer)obj);
                    }
                }
            }

            try {

                FileOutputStream out = new FileOutputStream(new File(OUTPUT_WKBK_PATH));
                outputWkbk.write(out);
                out.close();

            } catch (Exception e) {

                e.printStackTrace();
            }

        } catch (FileNotFoundException e) {

            e.printStackTrace();

        } catch (IOException e) {

            e.printStackTrace();
        }
    }
}
