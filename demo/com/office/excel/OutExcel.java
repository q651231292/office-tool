package com.office.excel;

import java.util.ArrayList;
import java.util.List;

import com.office.excel.model.Emp;

public class OutExcel {

	public static void main(String[] args) {
		List<Emp> emps = new ArrayList<>();
		emps.add(new Emp("001", "刘备"));
		emps.add(new Emp("002", "曹操"));
		new ExcelTool<Emp>().outExcel("D:/", "员工.xlsx", emps, Emp.class);
	}
}
