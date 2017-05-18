package com.office.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.office.excel.annotation.ExcelField;
import com.office.excel.model.ComExcel;
import com.office.excel.model.Emp;

/**
 * excel导出工具
 * 
 * @author rgy
 *
 * @param <T>
 *            实体类
 */
public class ExcelTool<T> {

	/**
	 * 导出excel
	 * 
	 * @param filePath
	 *            文件路径
	 * @param fileName
	 *            文件名
	 * @param list
	 *            对象列表
	 * @param cls
	 *            实体类
	 */
	public void outExcel(String filePath, String fileName, List<T> list, Class<T> cls) {
		Workbook wb = createWorkbook(fileName);
		if(wb!=null){
			Sheet sheet = wb.createSheet();
			Row row = sheet.createRow(0);
			Field[] fields = cls.getDeclaredFields();
			Cell cell = null;
			for (int i = 0; i < fields.length; i++) {
				cell = row.createCell(i);
				cell.setCellValue(fields[i].getDeclaredAnnotation(ExcelField.class).title());
			}
			for (int i = 0; i < list.size(); i++) {
				row = sheet.createRow(i + 1);
				T t = list.get(i);
				for (int j = 0; j < fields.length; j++) {
					cell = row.createCell(j);
					Object value;
					try {
						fields[j].setAccessible(true);
						value = fields[j].get(t);
						cell = setCellValue(cell, value);
					} catch (IllegalArgumentException | IllegalAccessException e) {
						e.printStackTrace();
					}
				}
			}
			try {
				wb.write(new FileOutputStream(filePath + fileName));
			} catch (IOException e) {
				e.printStackTrace();
			} finally {
				try {
					if (wb != null)
						wb.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		
	}

	/**
	 * excel转成List<T>
	 * @param filePath 文件路径
	 * @param cls 实体类
	 * @return 封装号的实体类集合
	 */
	public List<T> excelToList(String filePath, Class<T> cls) {
		List result = new ArrayList<>();
		try {
			Workbook wb = new XSSFWorkbook(filePath);
			int sheetLength = wb.getNumberOfSheets();
			for (int i = 0; i < sheetLength; i++) {
				Sheet sheet = wb.getSheetAt(i);
				int rowLength = sheet.getPhysicalNumberOfRows();
				for (int j = 0; j < rowLength; j++) {
					T t = cls.newInstance();
					ComExcel comExcel = (ComExcel) t;
					Row row = sheet.getRow(j);
					int cellLength = row.getPhysicalNumberOfCells();
					for (int k = 0; k < cellLength; k++) {
						Cell cell = row.getCell(k);
						String cellValue = cell.getStringCellValue();
						Row title = sheet.getRow(0);
						Cell titleCell = title.getCell(k);
						String titleCellValue = titleCell.getStringCellValue();
						Field[] fields = Emp.class.getDeclaredFields();
						for (int l = 0; l < fields.length; l++) {
							Field field = fields[l];
							ExcelField ef = field.getDeclaredAnnotation(ExcelField.class);
							String anTitle = ef.title();
							if (titleCellValue.equals(anTitle)) {
								Field[] declaredFields = Emp.class.getDeclaredFields();
								for (Field f : declaredFields) {
									ExcelField declaredAnnotation = f.getDeclaredAnnotation(ExcelField.class);
									String extitle = declaredAnnotation.title();
									if (extitle.equals(titleCellValue)) {
										String fName = f.getName();
										fName = firstToUpper(fName);
										comExcel.getClass().getMethod("set" + fName, new Class[] { String.class })
												.invoke(comExcel, new Object[] { cellValue });
									}
								}
								continue;
							}
						}
					}
					result.add(comExcel);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}

	/**
	 * 
	 * @param str
	 * @return
	 */
	private String firstToUpper(String str) {
		String begin = str.substring(0, 1).toUpperCase();
		String end = str.substring(1);
		str = begin + end;
		return str;
	}

	/**
	 * 创建excel 根据文件格式判断创建xls还是xlsx
	 * 
	 * @param fileName
	 *            文件名
	 * @return excel
	 */
	private Workbook createWorkbook(String fileName) {
		Workbook workbook = null;
		String format = fileName.substring(fileName.indexOf("."));
		if (".xls".equals(format)) {
			workbook = new HSSFWorkbook();
			return workbook;
		} else if (".xlsx".equals(format)) {
			workbook = new XSSFWorkbook();
			return workbook;
		}
		return null;
	}

	/**
	 * 设置单元格的值
	 * 
	 * @param cell
	 *            单元格
	 * @param value
	 *            值
	 * @return 具有值的单元格
	 */
	private Cell setCellValue(Cell cell, Object value) {
		// 当数字时
		if (value instanceof Integer)
			cell.setCellValue((Integer) value);
		// 当为字符串时
		if (value instanceof String)
			cell.setCellValue((String) value);
		// 当为布尔时
		if (value instanceof Boolean)
			cell.setCellValue((Boolean) value);
		// 当为时间时
		if (value instanceof Date)
			cell.setCellValue(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format((Date) value));
		// 当为时间时
		if (value instanceof Calendar)
			cell.setCellValue((Calendar) value);
		// 当为小数时
		if (value instanceof Double)
			cell.setCellValue((Double) value);
		return cell;
	}

}
