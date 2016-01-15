package org.my.demo;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Sheet;
import org.my.entity.User;
import org.my.poi.wrap.WorkbookProxy;
import org.my.poi.wrap.WorkbookProxy.ExcelType;
import org.my.utils.ObjectToExcelRowUtil;


public class ObjectToExcelRowDemo {
	public static void main(String... args) throws Exception {
		User user = constructUserBean();
		try(WorkbookProxy workbook = new WorkbookProxy(ExcelType.XSSF);) {
			Sheet sheet = workbook.createSheet("test");
			for(int i = 0; i < 10; i++)
				System.out.println(ObjectToExcelRowUtil.appendToSheet(sheet, user));
			workbook.save("D://test", "test1");
		}
		
	}
	
	private static User constructUserBean() {
		String pattern = "yyyy-MM-dd HH:mm:ss";
		DateFormat format = new SimpleDateFormat(pattern);
		User user = new User();
		user.setId(1);
		user.setStatus(1);
		user.setCreatetime(format.format(new Date()));
		user.setUsername("liuxiaoyang");
		user.setPassword("lxy2222");
		return user;
	}
}
