package generator_workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import report.Description;
import runner.RunReport;
import utils_report.UtilReport;

public class ALM {
	
	private List<String> listContentALM = new ArrayList<>();
	private File almFile;
	
	public ALM() {
		this.almFile = new File(RunReport.WAY_GENERATE_ALM + "\\ALM_ACCUMULATED_MACHINES.xlsx");
		UtilReport.copyFile(new File("workbooks//ALM.xlsx"), almFile);
		generetorALM();
	}
	
	public void generetorALM() {
		System.out.println("SAVING ALM'S MACHINE RESULTS...");
		System.out.println("READING FILES 'ALM.xlsx'...");
		final File[] foldersExistingAlm = new File(RunReport.WAY_GENERATE_ALM).listFiles();
		
		for (int posi = 0; posi < foldersExistingAlm.length; posi++) {
			File arqAlm = new File(foldersExistingAlm[posi].getAbsolutePath() + "\\ALM.xlsx");
			if(arqAlm.exists()) {
				try(XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(arqAlm))) {
					XSSFSheet sheet = wb.getSheetAt(0);
					for(Row row : sheet) {
						if(row.getCell(2).getStringCellValue().equals("Passed") && row.getCell(4).getStringCellValue() != null) {
							listContentALM.add(row.getCell(0).getStringCellValue() + ";" + row.getCell(1).getStringCellValue() + ";" + row.getCell(2).getStringCellValue() + ";NONE;" + row.getCell(4).getStringCellValue());
						}
					}
				} catch (IOException e) {
					System.err.println(e.getMessage() + "\n" + e);
					UtilReport.finish();
				}
			}
		}
	}
	
	public void writeAlmResult() {
		if(!almFile.exists()) {
			System.err.println("<<< FILE 'ALM.xlsx' NOT FOUND " + ALM.class + " >>>");
			UtilReport.finish();
		}
		Description.setAlmResult(listContentALM.size());
		FileOutputStream out = null;
		
		try(XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(almFile))) {
			XSSFSheet sheet = wb.getSheetAt(0);
			
			
			for(int posi = 0; posi < listContentALM.size(); posi++) {
				String[] values = listContentALM.get(posi).split(";");
				final int numLine = UtilReport.searchID(sheet, values[0]);
				if(numLine == -1) {
					System.err.println("<<< ID NOT FOUND " + ALM.class + " >>>");
					UtilReport.finish(); 
				}
				System.out.println("ALM - OK + WAY -- " + values[0] + " -- WITRING...");
				sheet.getRow(numLine).getCell(2).setCellValue(values[2]);
				
				sheet.getRow(numLine).createCell(4).setCellValue(values[4]);
				out = new FileOutputStream(almFile);
				wb.write(out);
			}
		} catch (IOException e) {
			System.err.println(e.getMessage() + "\n" + e);
			UtilReport.finish();
		} finally {
			try {
				if (out != null)
					out.close();
				System.out.println("Close File 'ALM.xlsx' on the Way --> " + almFile +  " -- " + new Date());
				System.out.println("=====================================================================================================>>");
			} catch (IOException e) {
				System.err.println(e.getMessage() + "\n" + e);
				UtilReport.finish();
			}
		}
	}
}