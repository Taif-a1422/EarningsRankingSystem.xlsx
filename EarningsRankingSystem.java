package task45;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class EarningsRankingSystem {
    @Test
    public void testEarningsRanking() {
        String filePath = "EarningsRankingSystem.xlsx";
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            List<Map.Entry<Integer, Double>> earnings = new ArrayList<>();

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell earningCell = row.getCell(1);
                    if (earningCell != null && earningCell.getCellType() == CellType.NUMERIC) {
                        earnings.add(new AbstractMap.SimpleEntry<>(i, earningCell.getNumericCellValue()));
                    }
                }
            }

            earnings.sort(Comparator.comparingDouble(Map.Entry::getValue));

            int rank = 1;
            for (Map.Entry<Integer, Double> entry : earnings) {
                Row row = sheet.getRow(entry.getKey());
                Cell rankCell = row.createCell(2);
                rankCell.setCellValue(rank);
                rank++;
            }

            try (FileOutputStream fos = new FileOutputStream(new File(filePath))) {
                workbook.write(fos);
            }

            // Re-read to assert Wednesday's rank
            try (FileInputStream fis2 = new FileInputStream(new File(filePath));
                 Workbook workbook2 = new XSSFWorkbook(fis2)) {
                Sheet sheet2 = workbook2.getSheetAt(0);
                for (int i = 1; i <= sheet2.getLastRowNum(); i++) {
                    Row row = sheet2.getRow(i);
                    if (row != null) {
                        Cell dayCell = row.getCell(0);
                        if (dayCell != null && "Wednesday".equals(dayCell.getStringCellValue())) {
                            Cell rankCell = row.getCell(2);
                            assertEquals(1, (int) rankCell.getNumericCellValue(), "Wednesday should have rank 1");
                        }
                    }
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}