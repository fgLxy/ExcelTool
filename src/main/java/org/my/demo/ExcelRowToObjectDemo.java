package org.my.demo;

import java.io.File;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.my.entity.User;
import org.my.poi.wrap.WorkbookProxy;
import org.my.utils.ExcelRowToObjectUtil;

public class ExcelRowToObjectDemo {
	public static void main(String... args) throws Exception {
		try(
				Workbook workbook = new WorkbookProxy(new File("D://test/test1.xlsx"));
				) {
			Iterator<Sheet> sheetIter = workbook.iterator();
			
			while(sheetIter.hasNext()) {
				Sheet sheet = sheetIter.next();
				Iterator<Row> rowIter = sheet.iterator();
				while(rowIter.hasNext()) {
					Row row = rowIter.next();
					System.out.println(ExcelRowToObjectUtil.getObjectByRow(User.class, row));
				}
			}
		}
	}
}
