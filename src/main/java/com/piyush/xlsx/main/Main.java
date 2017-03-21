package com.piyush.xlsx.main;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.piyush.xlsx.emailsender.SendMail;
import com.piyush.xlsx.filtercriteria.bean.FilterCriteriaBean;
import com.piyush.xlsx.utils.Utils;

public class Main {

	final static String FILE_NAME = "report.xlsx";
	final static String TEMPLATE_FILE_NAME = "template.xlsx";
	final static String FILTER_FILE_NAME = "filter.xlsx";

	@SuppressWarnings("static-access")
	public static void main(String[] args) {

		try {

			FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet dataSheet = workbook.getSheetAt(0);

			FileInputStream templateExcelFile = new FileInputStream(new File(TEMPLATE_FILE_NAME));
			Workbook templateWorkbook = new XSSFWorkbook(templateExcelFile);
			Sheet templateDatatypeSheet = templateWorkbook.getSheetAt(0);

			FileInputStream filterExcelFile = new FileInputStream(new File(FILTER_FILE_NAME));
			Workbook filterWorkbook = new XSSFWorkbook(filterExcelFile);
			Sheet filterSheet = filterWorkbook.getSheetAt(1);

			List<FilterCriteriaBean> filterCriteriaBeanMap = new ArrayList<FilterCriteriaBean>();
			filterCriteriaBeanMap = Utils.getFilterMap(filterSheet);

			filterCriteriaBeanMap.remove(0);

			// String criteria=Utils.getCriteria(filterSheet);
			// int column=Utils.getColumn(filterSheet);

			for (FilterCriteriaBean filterCriteriaBean : filterCriteriaBeanMap) {

				List<Row> rowList = Utils.getRows(dataSheet, filterCriteriaBean.getCriteria(),
						Integer.parseInt(filterCriteriaBean.getColumn()));

				List<Row> headerRow = Utils.generateHeaderRow(templateDatatypeSheet);

				List<Row> columnNameRow = Utils.generateColumnNameRow(templateDatatypeSheet);

				List<Row> listFinal = new ArrayList<Row>();

				listFinal.addAll(headerRow);
				listFinal.addAll(columnNameRow);
				listFinal.addAll(rowList);

				Utils.createExls(listFinal);

				SendMail sendMail = new SendMail();
				// sendMail.send(Utils.getEmailIds(filterSheet).get(1));

				for (String emailId : filterCriteriaBean.getEmailIds()) {
					sendMail.send(emailId);
					System.out.println("sent");
				}
			}
			// Utils.createExls(rowList);

		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}
