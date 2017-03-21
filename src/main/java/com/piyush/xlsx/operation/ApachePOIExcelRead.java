package com.piyush.xlsx.operation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.piyush.xlsx.utils.Utils;

public class ApachePOIExcelRead {

	private static final String FILE_NAME = "report.xlsx";
	private static final String TEMPLATE_FILE_NAME = "template.xlsx";

	@SuppressWarnings("unused")
	public static void main(String[] args) {

		try {

			FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet datatypeSheet = workbook.getSheetAt(0);
			
			
			FileInputStream templateExcelFile = new FileInputStream(new File(TEMPLATE_FILE_NAME));
			Workbook templateWorkbook = new XSSFWorkbook(templateExcelFile);
			Sheet templateDatatypeSheet = templateWorkbook.getSheetAt(0);
			
			
			//Utils.getCriteria(templateDatatypeSheet)
			
			List<Row> rowList= Utils.getRows(datatypeSheet, "ADMINSTRATION&LIASON", 3);
			
			
			List<Row> headerRow= Utils.generateHeaderRow(templateDatatypeSheet);
			
			List<Row> columnNameRow= Utils.generateColumnNameRow(templateDatatypeSheet);
			
			
			
			List<Row> listFinal = new ArrayList<Row>();
			
			listFinal.addAll(headerRow);
			listFinal.addAll(columnNameRow);
			listFinal.addAll(rowList);
			
			Utils.createExls(listFinal);
			
			//Utils.createExls(rowList);
			
			Iterator<Row> iterator = datatypeSheet.iterator();

			while (iterator.hasNext()) {

				Row currentRow = iterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();

				while (cellIterator.hasNext()) {

					Cell currentCell = cellIterator.next();

					// currentCell.getColumnIndex();

					// getCellTypeEnum shown as deprecated for version 3.15
					// getCellTypeEnum ill be renamed to getCellType starting
					// from version 4.0
					if (currentCell.getCellTypeEnum() == CellType.STRING) {
						System.out.print(currentCell.getStringCellValue() + "--");
					} else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {

						if (DateUtil.isCellDateFormatted(currentCell)) {
							System.out.print(currentCell.getDateCellValue() + "--");
						} else {
							System.out.print(currentCell.getNumericCellValue() + "--");
						}
					}

				}
				System.out.println();

			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
}