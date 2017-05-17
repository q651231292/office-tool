package com.office.excel;

import java.util.List;

import com.office.excel.model.Emp;

public class ExcelToList {

	public static void main(String[] args) {
		List<Emp> list = new ExcelTool<Emp>().excelToList("resource/员工.xlsx",Emp.class);
		System.out.println(list);
	}
}
