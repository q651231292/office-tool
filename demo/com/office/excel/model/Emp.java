package com.office.excel.model;

import com.office.excel.annotation.ExcelField;

public class Emp implements ComExcel{
	
	@ExcelField(title="员工编号")
	String empId;
	@ExcelField(title="员工姓名")
	String ename;
	
	public Emp(){}
	
	public Emp(String empId,String ename){
		this.empId = empId;
		this.ename = ename;
	}

	public String getEmpId() {
		return empId;
	}

	public void setEmpId(String empId) {
		this.empId = empId;
	}

	public String getEname() {
		return ename;
	}

	public void setEname(String ename) {
		this.ename = ename;
	}

	@Override
	public String toString() {
		return "Emp [empId=" + empId + ", ename=" + ename + "]";
	}
	
}
