package com.piyush.xlsx.utils;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.piyush.xlsx.filtercriteria.bean.FilterCriteriaBean;

public class Utils {
	// http://www.concretepage.com/apache-api/how-to-create-date-cell-in-xlsx-using-poi-in-java

	/**
	 * return Row from the xls on the basis of given critea at particular
	 * column.
	 * 
	 * ex. return row of xls sheet where column 3 contains
	 * "ADMINSTRATION&LIASON"
	 * 
	 * @param sheet
	 * @param criteria
	 * @param column
	 * @return
	 */
	@SuppressWarnings("deprecation")
	public static List<Row> getRows(Sheet sheet, String criteria, int column) {
		List<Row> rowList = null;
		try {
			rowList = new ArrayList<Row>();
			Iterator<Row> iterator = sheet.iterator();

			while (iterator.hasNext()) {

				Row currentRow = iterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();

				while (cellIterator.hasNext()) {

					Cell currentCell = cellIterator.next();
					if (currentCell.getCellTypeEnum() == CellType.STRING
							&& currentCell.getStringCellValue().equals(criteria)
							&& currentCell.getColumnIndex() == column) {
						rowList.add(currentRow);
					}

				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		return rowList;
	}

	/**
	 * create xls for the given rows
	 * 
	 * @param rowList
	 */
	public static void createExls(List<Row> rowList) {
		final String FILE_NAME = "test.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		int rowNum = 0;
		for (Row row : rowList) {
			Row demoRow = sheet.createRow(rowNum++);
			int colNum = 0;
			Iterator<Cell> cellIterator = row.iterator();
			while (cellIterator.hasNext()) {
				Cell currentCell = cellIterator.next();
				Cell cell = demoRow.createCell(colNum++);
				// currentCell.getColumnIndex();
				if (currentCell.getCellTypeEnum() == CellType.STRING) {
					cell.setCellValue((String) currentCell.getStringCellValue());
					System.out.print(currentCell.getStringCellValue() + "--");
				} else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
					if (DateUtil.isCellDateFormatted(currentCell)) {
						XSSFCellStyle cellStyle = workbook.createCellStyle();
						CreationHelper createHelper = workbook.getCreationHelper();
						short dateFormat = createHelper.createDataFormat().getFormat("dd-mmm-yy");
						cellStyle.setDataFormat(dateFormat);
						Calendar c = toCalendar(currentCell.getDateCellValue());
						// cell.setCellValue(Calendar.getInstance());
						cell.setCellValue(c);
						cell.setCellStyle(cellStyle);
					} else {
						cell.setCellValue((int) currentCell.getNumericCellValue());
					}
				}

			}

		}
		try {
			FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
			workbook.write(outputStream);
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * when it will get three consecutive blank it will consider as end of
	 * header
	 * 
	 * @param sheet
	 * @return
	 */
	public static List<Row> generateHeaderRow(Sheet sheet) {

		List<Row> rowList = null;
		try {
			boolean flag = false;
			rowList = new ArrayList<Row>();
			Iterator<Row> iterator = sheet.iterator();

			while (iterator.hasNext()) {

				Row currentRow = iterator.next();

				// currentRow.

				Iterator<Cell> cellIterator = currentRow.iterator();

				while (cellIterator.hasNext()) {

					Cell currentCell = cellIterator.next();
					if (currentCell.getCellTypeEnum() == CellType.STRING
							&& currentCell.getStringCellValue().equals("HeaderEnds")
							&& currentCell.getColumnIndex() == 0) {
						flag = false;
						break;
					}
					if (flag) {
						rowList.add(currentRow);
					}
					if (currentCell.getCellTypeEnum() == CellType.STRING
							&& currentCell.getStringCellValue().equals("HeaderStarts")
							&& currentCell.getColumnIndex() == 0) {
						flag = true;
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		return rowList;
	}

	/**
	 * when it will get three consecutive blank row 4th row will be considered
	 * as Column Row
	 * 
	 * @param sheet
	 * @return
	 */
	public static List<Row> generateColumnNameRow(Sheet sheet) {

		List<Row> rowList = null;
		try {
			boolean flag = false;
			rowList = new ArrayList<Row>();
			Iterator<Row> iterator = sheet.iterator();
			while (iterator.hasNext()) {
				Row currentRow = iterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();
				while (cellIterator.hasNext()) {
					Cell currentCell = cellIterator.next();
					if (currentCell.getCellTypeEnum() == CellType.STRING
							&& currentCell.getStringCellValue().equals("ColumNameEnds")
							&& currentCell.getColumnIndex() == 0) {
						flag = false;
						break;
					}
					if (flag) {
						rowList.add(currentRow);
						break;
					}
					if (currentCell.getCellTypeEnum() == CellType.STRING
							&& currentCell.getStringCellValue().equals("ColumNameStarts")
							&& currentCell.getColumnIndex() == 0) {
						flag = true;
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		return rowList;
	}

	public static List<String> getEmailIds(Sheet sheet) {
		List<String> emailList = null;
		try {
			emailList = new ArrayList<String>();
			Iterator<Row> iterator = sheet.iterator();

			while (iterator.hasNext()) {

				Row currentRow = iterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();

				while (cellIterator.hasNext()) {

					Cell currentCell = cellIterator.next();
					if (currentCell.getCellTypeEnum() == CellType.STRING && currentCell.getColumnIndex() == 3) {
						emailList.add(currentCell.getStringCellValue());
					}

				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		return emailList;
	}

	public static int getColumn(Sheet sheet) {
		int coulumnNumber = 0;
		try {
			// coulumnNumber = new String();
			Iterator<Row> iterator = sheet.iterator();

			while (iterator.hasNext()) {

				Row currentRow = iterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();

				while (cellIterator.hasNext()) {

					Cell currentCell = cellIterator.next();
					if (currentCell.getCellTypeEnum() == CellType.NUMERIC && currentCell.getColumnIndex() == 2) {
						coulumnNumber = (int) currentCell.getNumericCellValue();
					}

				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		return coulumnNumber;
	}

	public static String getCriteria(Sheet sheet) {
		String criteria = null;
		try {
			criteria = new String();
			Iterator<Row> iterator = sheet.iterator();

			while (iterator.hasNext()) {

				Row currentRow = iterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();

				while (cellIterator.hasNext()) {
					Cell currentCell = cellIterator.next();
					if (currentCell.getCellTypeEnum() == CellType.STRING && currentCell.getColumnIndex() == 1) {
						criteria = currentCell.getStringCellValue();
					}

				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		return criteria;
	}

	public static List<FilterCriteriaBean> getFilterMap(Sheet sheet) {
		List<FilterCriteriaBean> filterListMap = new ArrayList<FilterCriteriaBean>();
		try {
			Iterator<Row> iterator = sheet.iterator();
			while (iterator.hasNext()) {

				Row currentRow = iterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();

				FilterCriteriaBean filterCriteriaBean = new FilterCriteriaBean();
				while (cellIterator.hasNext()) {
					Cell currentCell = cellIterator.next();
					if (currentCell.getCellTypeEnum() == CellType.STRING && currentCell.getColumnIndex() == 1) {
						filterCriteriaBean.setCriteria(currentCell.getStringCellValue());
						// filterMap.put("criteria", criteria);
					}
					if (currentCell.getCellTypeEnum() == CellType.NUMERIC && currentCell.getColumnIndex() == 2) {
						filterCriteriaBean.setColumn("" + (int) currentCell.getNumericCellValue());
						// filterMap.put("coulumnNumber", ""+coulumnNumber);
					}
					if (currentCell.getCellTypeEnum() == CellType.STRING && currentCell.getColumnIndex() == 3) {
						List<String> items = Arrays.asList(currentCell.getStringCellValue().split("\\s*,\\s*"));
						filterCriteriaBean.setEmailIds(items);
					}
				}
				filterListMap.add(filterCriteriaBean);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return filterListMap;
	}

	private static Calendar toCalendar(Date date) {
		Calendar cal = Calendar.getInstance();
		cal.setTime(date);
		return cal;
	}

}