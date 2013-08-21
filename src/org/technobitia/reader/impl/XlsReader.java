package org.technobitia.reader.impl;

import java.io.File;
import java.io.IOException;

import org.technobitia.reader.contract.Reader;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class XlsReader implements Reader {

    private String inputFile;

    public XlsReader(String inputFile) {
        this.inputFile = inputFile;
    }

    public void read() throws IOException {

        File inputWorkbook = new File(inputFile);
        Workbook w;
        try {
            w = Workbook.getWorkbook(inputWorkbook);

            // Get the first sheet
            Sheet sheet = w.getSheet(0);

            for (int i = 0; i < sheet.getRows(); i++) {
                for (int j = 0; j < sheet.getColumns(); j++) {

                    Cell cell = sheet.getCell(j, i);

                    CellType type = cell.getType();
                    if (type == CellType.LABEL) {
                        System.out.println("I got a label "
                                + cell.getContents());
                    }

                    if (type == CellType.NUMBER) {
                        System.out.println("I got a number "
                                + cell.getContents());
                    }

                }
            }
        } catch (BiffException e) {
            e.printStackTrace();
        }
    }

    // Getters and setters
    /**
     * @return the inputFile
     */
    public String getInputFile() {
        return inputFile;
    }

    /**
     * @param inputFile
     *            the inputFile to set
     */
    public void setInputFile(String inputFile) {
        this.inputFile = inputFile;
    }

    /**
     * @param args
     */
    public static void main(String[] args) {
        String resource = "resources" + File.separator + "profiles.xls";
        Reader reader = new XlsReader(resource);
        try {
            reader.read();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
