package com.office.excel.model;

import com.office.excel.annotation.ExcelField;

public class Emp {
	
	@ExcelField(title="员工ID")
	String empId;
	@ExcelField(title="员工姓名")
	String ename;
	
	public Emp(String empId,String ename){
		this.empId = empId;
		this.ename = ename;
	}
}
