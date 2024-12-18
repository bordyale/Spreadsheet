package org.vulkan;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.apache.ofbiz.base.util.Debug;
import org.apache.ofbiz.base.util.UtilProperties;
import org.apache.ofbiz.base.util.UtilValidate;
import org.apache.ofbiz.product.spreadsheetimport.ImportProductHelper;
import org.apache.ofbiz.service.ServiceUtil;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Spreadsheet {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		System.out.println("ciao");

		String path = System.getProperty("user.dir") + "/spreadsheet";
		List<File> fileItems = new LinkedList<>();

		if (UtilValidate.isNotEmpty(path)) {
			File importDir = new File(path);
			if (importDir.isDirectory() && importDir.canRead()) {
				File[] files = importDir.listFiles();
				// loop for all the containing xls file in the spreadsheet
				// directory
				if (files == null) {
					System.out.println("error");
					return;
				}
				for (File file : files) {
					if (file.getName().toUpperCase(Locale.getDefault())
							.endsWith("XLS")) {
						fileItems.add(file);
					}
				}
			} else {
				System.out.println("error");
				return;
			}
		} else {
			System.out.println("error");
			return;
		}

		if (fileItems.size() < 1) {
			System.out.println("error");
			return;
		}

		for (File item : fileItems) {
			// read all xls file and create workbook one by one.
			List<Map<String, Object>> products = new LinkedList<>();
			List<Map<String, Object>> inventoryItems = new LinkedList<>();
			POIFSFileSystem fs = null;
			HSSFWorkbook wb = null;
			try {
				fs = new POIFSFileSystem(new FileInputStream(item));
				wb = new HSSFWorkbook(fs);
			} catch (IOException e) {
				System.out.println("error");
				return;
			}

			// get first sheet
			HSSFSheet sheet = wb.getSheetAt(0);
			wb.close();
			int sheetLastRowNumber = sheet.getLastRowNum();
			for (int j = 1; j <= sheetLastRowNumber; j++) {
				HSSFRow row = sheet.getRow(j);
				if (row != null) {
					// read productId from first column "sheet column index
					// starts from 0"
					HSSFCell cell2 = row.getCell(2);
					cell2.setCellType(CellType.STRING);
					String productId = cell2.getRichStringCellValue()
							.toString();
					// read QOH from ninth column
					HSSFCell cell5 = row.getCell(5);
					BigDecimal quantityOnHand = BigDecimal.ZERO;
					if (cell5 != null
							&& cell5.getcegetCellType() == CellType.NUMERIC) {
						quantityOnHand = new BigDecimal(
								cell5.getdgetNumericCellValue());
					}

				}
			}
		}
	}
}
