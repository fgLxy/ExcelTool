package org.my.poi.wrap;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.my.exception.UnSupportFileTypeException;

/**
 * 对poi中HSSFWorkbook及XSSFWorkbook类型的代理类型。
 * 主要目的在于
 * 1、创建实例的时候无需考虑excel文件版本问题
 * 2、close的时候可以自动将打开时所用的InputStream也一同关闭
 * @author Administrator
 *
 */
public class WorkbookProxy implements Workbook {
	
	private InputStream input;
	
	private Workbook workbook;
	
	private static final String XSSF_TYPE = ".xlsx";
	
	private static final String HSSF_TYPE = ".xls";
	
	
	public WorkbookProxy(File file) throws UnSupportFileTypeException, IOException {
		String fileName = file.getName().toLowerCase();
		boolean isSuccess = false;
		try {
			this.input = new FileInputStream(file);
			if(fileName.endsWith(XSSF_TYPE)) {
				this.workbook = new XSSFWorkbook(input);
			} else if(fileName.endsWith(HSSF_TYPE)) {
				this.workbook = new HSSFWorkbook(input);
			} else {
				throw new UnSupportFileTypeException("不支持的文件:" + file.getName());
			}
			isSuccess = true;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			throw e;
		} catch (IOException e) {
			e.printStackTrace();
			throw e;
		} finally {
			if(!isSuccess && this.input != null) {
				this.input.close();
			}
		}
		
	}
	
	@Override
	public Iterator<Sheet> iterator() {
		return workbook.iterator();
	}

	@Override
	public int getActiveSheetIndex() {
		return workbook.getActiveSheetIndex();
	}

	@Override
	public void setActiveSheet(int sheetIndex) {
		workbook.setActiveSheet(sheetIndex);
	}

	@Override
	public int getFirstVisibleTab() {
		return workbook.getFirstVisibleTab();
	}

	@Override
	public void setFirstVisibleTab(int sheetIndex) {
		workbook.setFirstVisibleTab(sheetIndex);
	}

	@Override
	public void setSheetOrder(String sheetname, int pos) {
		workbook.setSheetOrder(sheetname, pos);
	}

	@Override
	public void setSelectedTab(int index) {
		workbook.setSelectedTab(index);
	}

	@Override
	public void setSheetName(int sheet, String name) {
		workbook.setSheetName(sheet, name);
	}

	@Override
	public String getSheetName(int sheet) {
		return workbook.getSheetName(sheet);
	}

	@Override
	public int getSheetIndex(String name) {
		return workbook.getSheetIndex(name);
	}

	@Override
	public int getSheetIndex(Sheet sheet) {
		return workbook.getSheetIndex(sheet);
	}

	@Override
	public Sheet createSheet() {
		return workbook.createSheet();
	}

	@Override
	public Sheet createSheet(String sheetname) {
		return workbook.createSheet(sheetname);
	}

	@Override
	public Sheet cloneSheet(int sheetNum) {
		return workbook.cloneSheet(sheetNum);
	}

	@Override
	public Iterator<Sheet> sheetIterator() {
		return workbook.sheetIterator();
	}

	@Override
	public int getNumberOfSheets() {
		return workbook.getNumberOfSheets();
	}

	@Override
	public Sheet getSheetAt(int index) {
		return workbook.getSheetAt(index);
	}

	@Override
	public Sheet getSheet(String name) {
		return workbook.getSheet(name);
	}

	@Override
	public void removeSheetAt(int index) {
		workbook.removeSheetAt(index);
	}

	@SuppressWarnings("deprecation")
	@Override
	public void setRepeatingRowsAndColumns(int sheetIndex, int startColumn,
			int endColumn, int startRow, int endRow) {
		workbook.setRepeatingRowsAndColumns(sheetIndex, startColumn, endColumn, startRow, endRow);
	}

	@Override
	public Font createFont() {
		return workbook.createFont();
	}

	@Override
	public Font findFont(short boldWeight, short color, short fontHeight,
			String name, boolean italic, boolean strikeout, short typeOffset,
			byte underline) {
		return workbook.findFont(boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
	}

	@Override
	public short getNumberOfFonts() {
		return workbook.getNumberOfFonts();
	}

	@Override
	public Font getFontAt(short idx) {
		return workbook.getFontAt(idx);
	}

	@Override
	public CellStyle createCellStyle() {
		return workbook.createCellStyle();
	}

	@Override
	public short getNumCellStyles() {
		return workbook.getNumCellStyles();
	}

	@Override
	public CellStyle getCellStyleAt(short idx) {
		return workbook.getCellStyleAt(idx);
	}

	@Override
	public void write(OutputStream stream) throws IOException {
		workbook.write(stream);
	}

	/**
	 * 关闭workbook的同时，关闭对应的inputstream
	 */
	@Override
	public void close() throws IOException {
		try {
			workbook.close();
		} catch(IOException e) {
			throw e;
		} finally {
			if(input != null) {
				input.close();
			}
		}
		
	}

	@Override
	public int getNumberOfNames() {
		return workbook.getNumberOfNames();
	}

	@Override
	public Name getName(String name) {
		return workbook.getName(name);
	}

	@Override
	public Name getNameAt(int nameIndex) {
		return workbook.getNameAt(nameIndex);
	}

	@Override
	public Name createName() {
		return workbook.createName();
	}

	@Override
	public int getNameIndex(String name) {
		return workbook.getNameIndex(name);
	}

	@Override
	public void removeName(int index) {
		workbook.removeName(index);
	}

	@Override
	public void removeName(String name) {
		workbook.removeName(name);
	}

	@Override
	public int linkExternalWorkbook(String name, Workbook workbook) {
		return workbook.linkExternalWorkbook(name, workbook);
	}

	@Override
	public void setPrintArea(int sheetIndex, String reference) {
		workbook.setPrintArea(sheetIndex, reference);
	}

	@Override
	public void setPrintArea(int sheetIndex, int startColumn, int endColumn,
			int startRow, int endRow) {
		workbook.setPrintArea(sheetIndex, startColumn, endColumn, startRow, endRow);
	}

	@Override
	public String getPrintArea(int sheetIndex) {
		return workbook.getPrintArea(sheetIndex);
	}

	@Override
	public void removePrintArea(int sheetIndex) {
		workbook.removePrintArea(sheetIndex);
	}

	@Override
	public MissingCellPolicy getMissingCellPolicy() {
		return workbook.getMissingCellPolicy();
	}

	@Override
	public void setMissingCellPolicy(MissingCellPolicy missingCellPolicy) {
		workbook.setMissingCellPolicy(missingCellPolicy);
	}

	@Override
	public DataFormat createDataFormat() {
		return workbook.createDataFormat();
	}

	@Override
	public int addPicture(byte[] pictureData, int format) {
		return workbook.addPicture(pictureData, format);
	}

	@Override
	public List<? extends PictureData> getAllPictures() {
		return workbook.getAllPictures();
	}

	@Override
	public CreationHelper getCreationHelper() {
		return workbook.getCreationHelper();
	}

	@Override
	public boolean isHidden() {
		return workbook.isHidden();
	}

	@Override
	public void setHidden(boolean hiddenFlag) {
		workbook.setHidden(hiddenFlag);
	}

	@Override
	public boolean isSheetHidden(int sheetIx) {
		return workbook.isSheetHidden(sheetIx);
	}

	@Override
	public boolean isSheetVeryHidden(int sheetIx) {
		return workbook.isSheetVeryHidden(sheetIx);
	}

	@Override
	public void setSheetHidden(int sheetIx, boolean hidden) {
		workbook.setSheetHidden(sheetIx, hidden);
	}

	@Override
	public void setSheetHidden(int sheetIx, int hidden) {
		workbook.setSheetHidden(sheetIx, hidden);
	}

	@Override
	public void addToolPack(UDFFinder toopack) {
		workbook.addToolPack(toopack);
	}

	@Override
	public void setForceFormulaRecalculation(boolean value) {
		workbook.setForceFormulaRecalculation(value);
	}

	@Override
	public boolean getForceFormulaRecalculation() {
		return workbook.getForceFormulaRecalculation();
	}
	
}
