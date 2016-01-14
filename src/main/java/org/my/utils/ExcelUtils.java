package org.my.utils;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Modifier;
import java.text.SimpleDateFormat;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.LocaleUtil;
import org.my.annotation.ThisIgnore;
import org.my.annotation.TimePattern;
import org.my.entity.User;
import org.my.exception.TransformException;


public class ExcelUtils {
	
	public static <T> T getObjectByRow(Class<T> clazz, Row row) throws InstantiationException, IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException, NoSuchFieldException {
		T obj = clazz.getConstructor().newInstance();
		Field[] fields = clazz.getDeclaredFields();
		int count = 0;
		for(Field field : fields) {
			//只转换实例属性
			if((Modifier.STATIC & field.getModifiers()) == 0) {
				ThisIgnore thisIgnore = field.getAnnotation(ThisIgnore.class);
				if(thisIgnore != null) continue;//忽略带ThisIgnore注解的属性
				Object value = getCellValue(field, row, count++);
				String fieldName = field.getName();
				String methodName = "set" + ((char)(fieldName.charAt(0) - 32)) + fieldName.substring(1);
				clazz.getMethod(methodName, field.getType()).invoke(obj, value);
			}
		}
		return obj;
	}
	
	private static  Object getCellValue(Field field, Row row, int index) {
		Class<?> type = field.getType();
		
		HSSFCell cell = (HSSFCell) row.getCell(index);
		int cellType = cell.getCellType();

		switch (cellType) {
        case Cell.CELL_TYPE_BLANK:
            return null;
        case Cell.CELL_TYPE_BOOLEAN:
            return cell.getBooleanCellValue();
        case Cell.CELL_TYPE_FORMULA:
            return cell.getCellFormula();
        case Cell.CELL_TYPE_NUMERIC:
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
            	TimePattern pattern = field.getAnnotation(TimePattern.class);
            	String patternStr = pattern == null ? "yyyy-MM-dd HH:mm:ss" : pattern.getPattern();
                SimpleDateFormat sdf = new SimpleDateFormat(patternStr, LocaleUtil.getUserLocale());
                sdf.setTimeZone(LocaleUtil.getUserTimeZone());
                return sdf.format(cell.getDateCellValue());
            }
			return getNumeric(type, cell.getNumericCellValue());
        case Cell.CELL_TYPE_STRING:
            return cell.getStringCellValue();
        default:
        	throw new TransformException("unsupport field type:" + type);
		}
		
	}

	private static Object getNumeric(Class<?> type, double value) {
		if(type.equals(short.class)) {
			return Double.valueOf(value).shortValue();
		} else if(type.equals(Short.class)) {
			return Short.valueOf(Double.valueOf(value).shortValue());
		} else if(type.equals(int.class)) {
			return Double.valueOf(value).intValue();
		} else if(type.equals(Integer.class)) {
			return Integer.valueOf(Double.valueOf(value).intValue());
		} else if(type.equals(float.class)) {
			return Double.valueOf(value).floatValue();
		} else if(type.equals(Float.class)) {
			return Float.valueOf(Double.valueOf(value).floatValue());
		} else if(type.equals(double.class)) {
			return value;
		} else if(type.equals(Double.class)) {
			return Double.valueOf(value);
		}
		return null;
	}
	
	public static void main(String... args) throws IOException, InstantiationException, IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException, NoSuchFieldException {
		try(
				NPOIFSFileSystem fileSystem = new NPOIFSFileSystem(new File("D://test.xls"));
				HSSFWorkbook document = new HSSFWorkbook(fileSystem);
				) {
			Iterator<Sheet> sheetIter = document.iterator();
			
			while(sheetIter.hasNext()) {
				Sheet sheet = sheetIter.next();
				Iterator<Row> rowIter = sheet.iterator();
				while(rowIter.hasNext()) {
					Row row = rowIter.next();
					System.out.println(getObjectByRow(User.class, row));
				}
			}
		}
	}
	
}
