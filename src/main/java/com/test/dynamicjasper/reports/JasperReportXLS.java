package com.test.dynamicjasper.reports;

import static net.sf.dynamicreports.report.builder.DynamicReports.cmp;
import static net.sf.dynamicreports.report.builder.DynamicReports.export;
import static net.sf.dynamicreports.report.builder.DynamicReports.report;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;

import net.sf.dynamicreports.jasper.builder.JasperReportBuilder;
import net.sf.dynamicreports.jasper.builder.export.JasperXlsExporterBuilder;
import net.sf.dynamicreports.jasper.constant.JasperProperty;
import net.sf.dynamicreports.report.builder.column.Columns;
import net.sf.dynamicreports.report.builder.component.Components;
import net.sf.dynamicreports.report.builder.component.ImageBuilder;
import net.sf.dynamicreports.report.builder.component.TextFieldBuilder;
import net.sf.dynamicreports.report.builder.datatype.DataTypes;
import net.sf.dynamicreports.report.constant.HorizontalTextAlignment;
import net.sf.dynamicreports.report.exception.DRException;

public class JasperReportXLS {
	
	public static void main(String[] args) {
		build();
	}

	private static void build() {
		try {

			String query = "SELECT employee.ID as id, employee.NAME as name, employee.LOGIN_ID as loginId, employee.DEPARTMENT as department, "
					+ " employee.STATUS as status FROM test.employee;";

			JasperXlsExporterBuilder xlsExporter = export.xlsExporter("D:/EmployeeDetailsReport" + ".xls")
					.setDetectCellType(true).setIgnorePageMargins(true).setWhitePageBackground(false)
					.setRemoveEmptySpaceBetweenColumns(true)
			// .setPassword("pass1234")
			;

			JasperReportBuilder report = getReportColumns();
			report.setDataSource(query, getConnection());
			report.toXls(xlsExporter);

		} catch (DRException e) {
			e.printStackTrace();
		}
	}

	private static JasperReportBuilder getReportColumns() {

		File comapnyLogo = new File(
				new JasperReportXLS().getClass().getClassLoader().getResource("report/logo.png").getFile());
		ImageBuilder logo = cmp.image(comapnyLogo.getPath()).setFixedDimension(94, 40);

		// ImageBuilder logo = cmp.image("D:\\report\\logo.png").setFixedDimension(94, 40);

		TextFieldBuilder<String> tfb = cmp.text("Copyright 2018. All rights reserved.")
				.setHorizontalTextAlignment(HorizontalTextAlignment.LEFT);
		JasperReportBuilder report = report()
				// .setColumnTitleStyle(Templates.columnTitleStyle)
				.pageHeader(logo)
				.addPageHeader(Components.text("Employee Details\n-------------------")
						.setHorizontalTextAlignment(HorizontalTextAlignment.LEFT))
				.addPageFooter(tfb).addProperty(JasperProperty.EXPORT_XLS_FREEZE_ROW, "2").ignorePageWidth()
				.ignorePagination()
				.columns(Columns.column("Name", "name", DataTypes.stringType()).setFixedWidth(new Integer(150)),
						Columns.column("Login Id", "loginId", DataTypes.stringType()).setFixedWidth(new Integer(75)),
						Columns.column("Department", "department", DataTypes.stringType())
								.setFixedWidth(new Integer(150)),
						Columns.column("Status", "status", DataTypes.stringType()).setFixedWidth(new Integer(150)));
		return report;
	}
	
	private static Connection getConnection() {
		
		try {
			String driver = "com.mysql.jdbc.Driver";
			String url = "jdbc:mysql://localhost:3306/test";
			String username = "root";
			String password = "password"; // Change it to your Password
			System.setProperty(driver, "");
			Connection conn = DriverManager.getConnection(url, username,password);
			
			return conn;
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return null;
	}

}
