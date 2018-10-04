package com.bryanalexandercd.excel;

import java.io.File;
import java.io.FileInputStream;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {

	public static void main(String[] args) {

		try {

			FileInputStream fileInputStream = new FileInputStream(new File("P:\\GoogleDrive\\Meii.xlsx"));

			XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

			XSSFSheet sheet = workbook.getSheetAt(0);

			int numFilas = sheet.getLastRowNum();

			for (int x = 2; x < numFilas - 2; x++) {

				Row row = sheet.getRow(x);

				if (!Objects.equals(row, null)) {

					int numCols = row.getLastCellNum();

					for (int y = 0; y < numCols; y++) {

						Cell cell = row.getCell(y);

						if (!Objects.equals(cell, null)) {

							switch (cell.getCellTypeEnum().toString()) {

							case "NUMERIC":

								System.out.print(cell.getNumericCellValue() + " ");

								break;

							case "STRING":

								System.out.print(cell.getStringCellValue() + " ");

								break;

							case "FORMULA":

								System.out.print(cell.getCellFormula() + " ");

								break;

							default:
								break;
							}
						}

					}

					System.out.println("");

				}

			}

		} catch (Exception e) {

			e.printStackTrace();

		}
	}
}
