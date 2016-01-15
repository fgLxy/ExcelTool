package org.my.utils;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.my.annotation.ThisIgnore;

/**
 * 实体类转换成行数据
 * @author Administrator
 *
 */
public class ObjectToExcelRowUtil {
	
	/**
	 * 将数据插入到指定行号上
	 * @param sheet
	 * @param obj
	 * @param rowNum
	 * @return
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 */
	public static int insertIntoSheet(Sheet sheet, Object obj, int rowNum) throws IllegalArgumentException, IllegalAccessException {
		Row row = sheet.createRow(rowNum);
		Class<?> clazz = obj.getClass();
		Field[] fields = clazz.getDeclaredFields();
		int count = 0;
		for(Field field : fields) {
			if((Modifier.STATIC & field.getModifiers()) == 0) {
				ThisIgnore thisIgnore = field.getAnnotation(ThisIgnore.class);
				if(thisIgnore != null) continue;//忽略带ThisIgnore注解的属性
				Cell cell = row.createCell(count++);
				field.setAccessible(true);
				setCellValue(cell, field.get(obj));
			}
		}
		
		return sheet.getLastRowNum();
	}
	
	/**
	 * 向sheet末尾追加一行数据，并返回当前行数
	 * @param sheet
	 * @param obj
	 * @return
	 * @throws IllegalAccessException 
	 * @throws IllegalArgumentException 
	 */
	public static int appendToSheet(Sheet sheet, Object obj) throws IllegalArgumentException, IllegalAccessException {
		Row row = sheet.createRow(sheet.getLastRowNum() + 1);
		Class<?> clazz = obj.getClass();
		Field[] fields = clazz.getDeclaredFields();
		int count = 0;
		for(Field field : fields) {
			if((Modifier.STATIC & field.getModifiers()) == 0) {
				ThisIgnore thisIgnore = field.getAnnotation(ThisIgnore.class);
				if(thisIgnore != null) continue;//忽略带ThisIgnore注解的属性
				Cell cell = row.createCell(count++);
				field.setAccessible(true);
				setCellValue(cell, field.get(obj));
			}
		}
		
		return sheet.getLastRowNum();
	}

	private static void setCellValue(Cell cell, Object object) {
		String value = object.toString();
		cell.setCellValue(value);
	}
	
}
