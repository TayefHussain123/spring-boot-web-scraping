package org.example.excel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import java.io.FileOutputStream;
import java.io.IOException;


public class ExcelSheet {

    public void createExcel() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Web scrapping");
        String[] headLine = {"Question sheet name."};
        String[] columHeading = {"Question :", "Option one : ", "Option two :", "Option three :", "Option four :", "Answer :", "Solving time"};
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.BLACK.index);
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(headerFont);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);

        Row headerRowOne = sheet.createRow(0);
        for (int i = 0; i < headLine.length; i++) {
            Cell cell = headerRowOne.createCell(i);
            cell.setCellValue(headLine[i]);
        }

        Row headerRow2 = sheet.createRow(2);
        for (int i = 0; i < columHeading.length; i++) {
            Cell cell = headerRow2.createCell(i);
            cell.setCellValue(columHeading[i]);
            cell.setCellStyle(headerStyle);
        }

        try {

            int rowNum = 4;

            String url = "https://example123.com";

            String answerUrl = "https://example123.com";

            Document document = Jsoup.connect(url).get();

            Document answerDocument = Jsoup.connect(answerUrl).get();

            Elements question = document.select(".sq_row1").tagName("h1");

            Elements optionOne = document.select(".inner_rows").select(".op1").select(".inner_inner").tagName("p");

            Elements optionTwo = document.select(".inner_rows").select(".op2").select(".inner_inner").tagName("p");

            Elements optionThree = document.select(".inner_rows").select(".op3").select(".inner_inner").tagName("p");

            Elements optionFour = document.select(".inner_rows").select(".op4").select(".inner_inner").tagName("p");

            Element answer = answerDocument.select(".tmplt_footer").tagName("span").tagName("p").first();

            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(question.text());
            row.createCell(1).setCellValue(optionOne.text());
            row.createCell(2).setCellValue(optionTwo.text());
            row.createCell(3).setCellValue(optionThree.text());
            row.createCell(4).setCellValue(optionFour.text());
            row.createCell(5).setCellValue(answer.text());

            for (int j = 0; j < columHeading.length; j++) {
                sheet.autoSizeColumn(j);
            }
            FileOutputStream fileOutputStream = new FileOutputStream("example.xlsx");
            workbook.write(fileOutputStream);

        } catch (Exception e) {
            System.out.println(e);
        }
    }

}