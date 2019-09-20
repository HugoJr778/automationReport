package generatorWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import enumService.Data;
import report.Report;
import runnerReport.RunReport;
import utilsReport.UtilReport;

public class GeneratorReport {
	
	private File fileCopyReport;
	
	public GeneratorReport() {
		this.fileCopyReport = new File(RunReport.WAY_FOLDER_REPORT + "\\RelatorioPorCenario.xlsx");
		UtilReport.copyFile(new File("E:\\URA RELATÓRIOS\\RelatorioPorCenario.xlsx"), this.fileCopyReport);
		editingReport();
	}
	
	public String dataReport(Data data) {
		return data.dataReport();
	}
	
	private void editingReport() {
		
		System.out.println("<<<<< FILE COPY! >>>>>");
		System.out.println("<<<<< EDITING OK'S IN 'RelatorioPorCenario.xlsx' ON DATE " + dataReport(Data.DATE) + " >>>>>");
		if(!fileCopyReport.exists()) {
			System.err.println("<<< ERROR FILE NOT FOUND " + GeneratorReport.class + ">>>");
			UtilReport.finish();
		}
		System.out.println("Open File 'RelatorioPorCenario.xlsx' -- " + new Date());
		
		System.out.println(">>>>>>>>>>>>>>>>>>>>>>>>>> SIZE: " + Report.getLineWBOK());
		System.out.println(">>>>>>>>>>>>>>>>>>>>>>>>>> FILE WAY: " + fileCopyReport);
		System.out.println(">>>>>>>>>>>>>>>>>>>>>>>>>> STRING[]: " + Report.oks);
		
		OutputStream os = null;
		
		try (InputStream fi = new FileInputStream(fileCopyReport);
				XSSFWorkbook wb = new XSSFWorkbook(fi)) {
			
			XSSFSheet planilha = wb.getSheetAt(0);
			for(int posi = 0; posi < Report.getListOK().size(); posi ++) {
				final int rowNum = UtilReport.searchID(planilha, Report.getListOK(posi));
				String[] x = Report.getOks(posi).split(";");
				
				System.out.println(">>>>>>>>>>" + Report.getOks(posi));
				for(int p = 0; p < x.length; p++) {
					System.out.println(">>>>>> X VALUES -- " + x[p]);
				}
				
				//Exception in thread "main" java.lang.IllegalArgumentException: Address of hyperlink must be a valid URI
				
//				CreationHelper helper = wb.getCreationHelper();
				
				planilha.getRow(rowNum).getCell(2).setCellValue(x[0]);
				planilha.getRow(rowNum).getCell(3).setCellValue(Double.parseDouble(x[1]));
				planilha.getRow(rowNum).getCell(4).setCellValue(x[2]);
				planilha.getRow(rowNum).getCell(5).setCellValue(x[3]);
				planilha.getRow(rowNum).getCell(6).setCellValue(x[4]);
				planilha.getRow(rowNum).getCell(7).setCellValue(x[5]);
				planilha.getRow(rowNum).getCell(8).setCellValue(x[6]);
//				Hyperlink link = helper.createHyperlink(HyperlinkType.FILE);
//				link.setAddress(x[7]);
				planilha.getRow(rowNum).getCell(10).setCellValue(x[7]);
				planilha.getRow(rowNum).getCell(11).setCellValue(x[8]);
//				Hyperlink link2 = helper.createHyperlink(HyperlinkType.FILE);
//				link2.setAddress(x[9]);
				planilha.getRow(rowNum).getCell(12).setCellValue(x[9]);
				
				os = new FileOutputStream(fileCopyReport);
				wb.write(os);
				System.out.println("OK - Scenario -- " +  Report.getListOK(posi) + " -- Successfully Modified...");
			}
		} catch (IOException e) {
			System.err.println(e.getMessage() + "\n" + e);
		} finally {
			try {
				if(os != null)
					os.close();
				System.out.println("Close File 'RelatorioPorCenario.xlsx' -- " + new Date());
				System.out.println("=====================================================================================================>>");
			} catch (IOException e) {
				System.err.println(e.getMessage() + "\n" + e);
			} 
		}
	}
}