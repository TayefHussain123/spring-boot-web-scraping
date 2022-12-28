package org.example;


import org.example.excel.ExcelSheet;

import java.io.IOException;

public class WebScrapingApplication {
    public static void main(String[] args) throws IOException {
        ExcelSheet excelSheet = new ExcelSheet();
        excelSheet.createExcel();

    }
}


