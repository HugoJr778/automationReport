package sheetServiceNOK;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import enumService.Data;
import interfaceService.FaceReportInfo;
import report.Report;
import runnerReport.RunReport;
import utilsReport.UtilReport;
import workbookMain.GeneratorSheetMain;

public class SearchDataToReportNOK implements FaceReportInfo {
	
	private final String NAME_FOLDER;
	private final String[] LIST_HEADERS = {"id", "cenario", "status", "execucao", "json", "hostName", 
			"tipoErro", "massa", "erro", "statusDev", "idLog", "dataHora", "eviLog"};

	public SearchDataToReportNOK(String name) {
		this.NAME_FOLDER = name;
		searchDataReport();
	}
	
	@Override
	public String dataReport(Data data) {
		return data.dataReport();
	}
	
	@Override
	public void searchDataReport() {
		
		System.out.println("<<<<< COLLECTING NOK'S FROM REPORT '" + NAME_FOLDER + "' >>>>>");
		
		System.out.println("=====================================================================================================>>");
		System.out.println("Open File 'RelatorioPorCenario.xlsx' in Folder --> " + NAME_FOLDER + " -- " + new Date());
		File arq = new File(RunReport.WAY_FOLDER_REPORT + "\\" + NAME_FOLDER + "\\" + "RelatorioPorCenario.xlsx");
		if(!arq.exists()) {
			System.err.println("<<<FILE WAS NOT FOUND, THE NAME MAY BE DIVERGENT " + SearchDataToReportNOK.class + ">>>");
			UtilReport.finish();
		}
		FileOutputStream outFile = null;
		XSSFWorkbook wbInfo = null; 
		XSSFWorkbook wb = null;
		
		try {
			wbInfo = new XSSFWorkbook(new FileInputStream(arq));
			XSSFSheet sheet = wbInfo.getSheetAt(0);
			System.out.println("Reading File Info -- " + NAME_FOLDER + "...");
			HashMap<String, String> dct = new HashMap<String, String>(12);
			String value = "";
			
			for(Row row : sheet) {
				for(int posi = 0; posi < LIST_HEADERS.length; posi++) {
					Cell cell = row.getCell(posi);
					switch (cell.getCellTypeEnum()) {
						case STRING:
							value = cell.getStringCellValue();
							break;
						case NUMERIC:
							value = Integer.toString((int) cell.getNumericCellValue());
							break;
						case BLANK:
							value = "";
							break;
						case _NONE:
							value = "";
							break;
						default:
							System.out.println("Cell Not Found in Code Type --> " + cell.getCellTypeEnum());
							break;
					}
					dct.put(LIST_HEADERS[posi], value);
				}
				final boolean result = ServiceNOK.verifyContentWorkbookNOK(dct);
				
				if(result) {
					wb = UtilReport.getWbMain();
					XSSFSheet sheetMain = wb.getSheet(GeneratorSheetMain.getSheetNameNOK());
					Row rowWrite = sheetMain.createRow(Report.getLineWBNOK());
					System.out.println(NAME_FOLDER + " - NOK -- " + dct.get("id") + " -- WITRING...");
					ServiceNOK.populateCellsNOK(rowWrite, wb);
					Report.setLineWBNOK(1);
				} else {
					System.out.println(NAME_FOLDER + " - NOK -- JUMPING SCENARIO " + dct.get("id") + "...");
					Report.clearValuesReport();
				}
				Report.clearValuesReport();
				if(wb != null) {
					outFile = new FileOutputStream(RunReport.WAY_FOLDER_REPORT + "\\" + dataReport(Data.FILE_REPORT_NAME));
					wb.write(outFile);
				}
			}
		} catch (FileNotFoundException e) {
			System.err.println(e.getMessage() + "\n" + e);
			UtilReport.finish();
		} catch (IOException e) {
			System.err.println(e.getMessage() + "\n" + e);
			UtilReport.finish();
		} finally {
			try {
				if(wb != null)
				    wb.close();
				if(outFile != null)
				    outFile.close();
				if(wbInfo != null)
					wbInfo.close();
				System.out.println("Close File 'RelatorioPorCenario.xlsx' in Folder --> " + NAME_FOLDER +  " -- " + new Date());
				System.out.println("=====================================================================================================>>");
			} catch (IOException e) {
				System.err.println(e.getMessage() + "\n" + e);
				UtilReport.finish();
			}
		}
	}
}