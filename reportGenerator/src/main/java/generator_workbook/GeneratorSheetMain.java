package generator_workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import enum_service.Data;
import report.Report;
import runner.RunReport;
import utils_report.UtilReport;

public class GeneratorSheetMain {
	
	private File arqMain;
	private static final String SHEET_NAME_NOK = "URA NOK";
	private static final String SHEET_NAME_OK = "URA OK";
	private static final String LINE = "=====================================================================================================>>";
	private final int[] widthColuns = {
									   15 * 170, //ID
									   35 * 830, //Cenário
									   15 * 190, //Status
									   30 * 195, //Quat. Execuções
									   60 * 240, //Request Code Json
									   30 * 215, //Máquina Executada
									   30 * 200, //Tipo Erro
									   15 * 220, //IDMassa
									   35 * 750, //Erro 
									   30 * 200, //StatusDEV
									   53 * 270, //IDLog
									   53 * 240, //Data e Hora
									   53 * 270  //Evidência LOG
									   }; 

	public GeneratorSheetMain(String folderExisting) {
		this.arqMain = UtilReport.indentWay(dataReport(Data.FILE_REPORT_NAME));
		createWorkbook();
		getHeadersWorkbook(folderExisting);
		writerHeaders();
	}
	
	public static String getSheetNameNOK() {
		return SHEET_NAME_NOK;
	}
	
	public static String getSheetNameOK() {
		return SHEET_NAME_OK;
	}
	
	private final String dataReport(Data data) {
		return data.data();
	}

	private void createWorkbook() {
		System.out.println(LINE);
		System.out.println("Creating report file 'REPORT_URA' in: " + RunReport.WAY_FOLDER_REPORT + "...");
		FileOutputStream out = null;
		
		try(XSSFWorkbook workbook = new XSSFWorkbook();) {
			if(!arqMain.exists()) {
				out = new FileOutputStream(arqMain);
				workbook.createSheet(SHEET_NAME_NOK);
				workbook.createSheet(SHEET_NAME_OK);
				XSSFSheet sheet = workbook.getSheet(SHEET_NAME_NOK);
				
				for(int posi = 0; posi < 13; posi++) {
					sheet.setColumnWidth(posi, widthColuns[posi]);
				}
				
				sheet = workbook.getSheet(SHEET_NAME_OK);
				
				for(int posi = 0; posi < 13; posi++) {
					sheet.setColumnWidth(posi, widthColuns[posi]);
				}
				workbook.write(out);
				
			} else {
				System.err.println("<<<ARCHIVE 'REPORT-URA' ALREADY EXISTS " + GeneratorSheetMain.class + ">>>");
				UtilReport.finish();
			}
			System.out.println("Created File 'REPORT_URA'!");
			System.out.println(LINE);
		} catch (IOException e) {
			System.err.println(e.getMessage() + "\n" + e);
			UtilReport.finish();
		} finally {
			try {
				if (out != null) {
					out.flush();
					out.close();
				}
			} catch (IOException e) {
				System.err.println(e.getMessage() + "\n" + e);
				UtilReport.finish();
			}
		}
	} 
	
	private void getHeadersWorkbook(String folderExisting) {
		System.out.println("Open File 'RelatorioPorCenario.xlsx' in Folder --> " + folderExisting + " -- " + new Date());
		try (XSSFWorkbook workInfo = new XSSFWorkbook(new FileInputStream(UtilReport.indentWay(folderExisting + "\\RelatorioPorCenario.xlsx")))) {
			
			XSSFSheet sheet = workInfo.getSheetAt(0);
			System.out.println("Reading File Info...");
			
			for(Row row : sheet) {
				int posi = 0;
				while(posi < 13) {
					Report.addValuesReport(row.getCell(posi).getStringCellValue());
					Report.addStyleCell(row.getCell(posi).getCellStyle());
					posi++;
				}
				if(posi == 13)
					break;
			}
			System.out.println("Headers finds ------------------------------------------->");
			Report.printValuesReport();
			System.out.println("--------------------------------------------------------->");
			System.out.println("Close File 'RelatorioPorCenario.xlsx' in Folder --> " + folderExisting + " -- " + new Date());
			System.out.println(LINE);
		} catch (IOException e) {
			System.err.println(e.getMessage() + "\n" + e);
			UtilReport.finish();
		}
	}
	
	private void writerHeaders() {
		FileOutputStream outFile = null;
		File arqReport = UtilReport.indentWay(dataReport(Data.FILE_REPORT_NAME));
		
		try(XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(arqReport))) {
			System.out.println("Open file 'REPORT_URA' to Write Headers " + new Date());
			Row row = wb.getSheet(SHEET_NAME_NOK).createRow(0);
			
			System.out.println("Write Headers to File in Sheet '" + SHEET_NAME_NOK + "'...");
			UtilReport.populateCellsHeaders(row, wb);
			row = wb.getSheet(SHEET_NAME_OK).createRow(0);
			System.out.println("Write Headers to File in Sheet '" + SHEET_NAME_OK + "'...");
			UtilReport.populateCellsHeaders(row, wb);
			outFile = new FileOutputStream(RunReport.WAY_FOLDER_REPORT + "\\" + dataReport(Data.FILE_REPORT_NAME));
			wb.write(outFile);
			
		} catch (IOException e) {
			System.err.println(e.getMessage() + "\n" + e);	
		} finally {
			try {
				if (outFile != null)
					outFile.close();
				Report.setLineWBNOK(1);
				Report.setLineWBOK(1);
				System.out.println("Close file 'REPORT_URA' " + new Date());
				System.out.println(LINE);
			} catch (IOException e) {
				System.err.println(e.getMessage() + "\n" + e);
			}
		}
	}
}