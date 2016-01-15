package org.my.utils;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Modifier;
import java.text.SimpleDateFormat;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.LocaleUtil;
import org.my.annotation.ThisIgnore;
import org.my.annotation.TimePattern;
import org.my.exception.TransformException;


public class ExcelRowToObjectUtil {
	
	/**
	 * 根据excel的一行转化为对应的实体类（列号对应属性声明的顺序）
	 * @param clazz
	 * @param row
	 * @return
	 * @throws InstantiationException
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 * @throws NoSuchMethodException
	 * @throws SecurityException
	 * @throws NoSuchFieldException
	 */
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
	
	/**
	 * 从row中顺序取值，按照fieldNames中字段名称的顺序依次注入
	 * @param clazz
	 * @param row
	 * @param fieldNames
	 * @return
	 * @throws InstantiationException
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 * @throws NoSuchMethodException
	 * @throws SecurityException
	 * @throws NoSuchFieldException
	 */
	public static <T> T getObjectByRow(Class<T> clazz, Row row, String[] fieldNames) throws InstantiationException, IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException, NoSuchFieldException {
		T obj = clazz.getConstructor().newInstance();
		
		int count = 0;
		for(String fieldName : fieldNames) {
			Field field = clazz.getField(fieldName);
			//只转换实例属性
			if((Modifier.STATIC & field.getModifiers()) == 0) {
				ThisIgnore thisIgnore = field.getAnnotation(ThisIgnore.class);
				if(thisIgnore != null) continue;//忽略带ThisIgnore注解的属性
				Object value = getCellValue(field, row, count++);
				String methodName = "set" + ((char)(fieldName.charAt(0) - 32)) + fieldName.substring(1);
				clazz.getMethod(methodName, field.getType()).invoke(obj, value);
			}
		}
		
		return obj;
	}
	
	/**
	 * 根据属性-列号对应表注入实例
	 * @param clazz
	 * @param row
	 * @param fieldCellMap
	 * @return
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 * @throws NoSuchMethodException
	 * @throws SecurityException
	 * @throws InstantiationException
	 * @throws NoSuchFieldException
	 */
	public static <T> T getObjectByRow(Class<T> clazz, Row row, Map<String, Integer> fieldCellMap) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException, InstantiationException, NoSuchFieldException {
		T obj = clazz.getConstructor().newInstance();
		
		Set<Entry<String, Integer>> entrySet = fieldCellMap.entrySet();
		for(Entry<String, Integer> entry : entrySet) {
			int index = entry.getValue();
			String fieldName = entry.getKey();
			Field field = clazz.getField(fieldName);
			//只转换实例属性
			if((Modifier.STATIC & field.getModifiers()) == 0) {
				ThisIgnore thisIgnore = field.getAnnotation(ThisIgnore.class);
				if(thisIgnore != null) continue;//忽略带ThisIgnore注解的属性
				Object value = getCellValue(field, row, index);
				String methodName = "set" + ((char)(fieldName.charAt(0) - 32)) + fieldName.substring(1);
				clazz.getMethod(methodName, field.getType()).invoke(obj, value);
			}
		}
		
		return obj;
	}
	
	/**
	 * 读取单元格中数据，并转换为field需要的格式
	 * @param field
	 * @param row
	 * @param index
	 * @return
	 */
	private static  Object getCellValue(Field field, Row row, int index) {
		Class<?> type = field.getType();
		
		Cell cell = row.getCell(index);
		int cellType = cell.getCellType();
		Object value = null;
		switch (cellType) {
        case Cell.CELL_TYPE_BLANK: {
        	return value;
        }
        case Cell.CELL_TYPE_BOOLEAN: {
        	value = cell.getBooleanCellValue();
            break;
        }
        case Cell.CELL_TYPE_FORMULA: {
        	value = cell.getCellFormula();
        	break;
        }
        case Cell.CELL_TYPE_NUMERIC: {
        	if (HSSFDateUtil.isCellDateFormatted(cell)) {
            	TimePattern pattern = field.getAnnotation(TimePattern.class);
            	String patternStr = pattern == null ? "yyyy-MM-dd HH:mm:ss" : pattern.getPattern();
                SimpleDateFormat sdf = new SimpleDateFormat(patternStr, LocaleUtil.getUserLocale());
                sdf.setTimeZone(LocaleUtil.getUserTimeZone());
                value = sdf.format(cell.getDateCellValue());
            }
            value = getNumeric(type, cell.getNumericCellValue());
            break;
        }
        case Cell.CELL_TYPE_STRING: {
        	value = cell.getStringCellValue();
        	break;
        }
        default:
        	throw new TransformException("unsupport field type:" + type);
		}
		
		if(value == null || field.getType().equals(value.getClass())) {
			return value;
		}
		return transformToType(type, value);
		
	}
	
	/**
	 * value的类型只能是String类型了
	 * @param value
	 * @param type
	 * @return
	 */
	private static Object transformToType(Class<?> type, Object value) {
		Object rtn = null;
		try {
			double param = Double.valueOf(value.toString()).doubleValue();
			rtn = getNumeric(type, param);
		} catch(Exception e) {
			
		}
		if(rtn != null) return rtn;
		throw new TransformException("不能将" + value.getClass().getName() + "类的" + value + "转换为" + type.getName() + "类型");
	}

	/**
	 * 获取数字类型的单元格属性，并转换为对应的type类型
	 * @param type
	 * @param value
	 * @return
	 */
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
	
}
