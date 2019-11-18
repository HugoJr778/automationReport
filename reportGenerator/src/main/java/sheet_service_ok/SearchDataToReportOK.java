package sheet_service_ok;

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

import enum_service.Data;
import generator_workbook.GeneratorSheetMain;
import interface_service.FaceReportInfo;
import report.Report;
import runner.RunReport;
import utils_report.UtilReport;

public class SearchDataToReportOK implements FaceReportInfo {

	private final String NAME_FOLDER;
	private int posiFolder0;
	private final String[] LIST_HEADERS = {"id", "cenario", "status", "execucao", "json", "hostName", 
			"tipoErro", "massa", "erro", "statusDev", "idLog", "dataHora", "eviLog"};
	
	public SearchDataToReportOK(String name, int posiFolder0) {
		this.NAME_FOLDER = name;
		this.posiFolder0 = posiFolder0;
		searchDataReport();
	}
	
	private final boolean verifyID(String currentId) {
		boolean result = true;
		if(posiFolder0 == 0 && (!currentId.contains("ID"))) {
			Report.addListOK(currentId);
			return result;
		} else {
			if(Report.getListOK().contains(currentId)) {
				result = false;
			}
		}
		if(result) 
			Report.addListOK(currentId);
		
		return result;
	}
	
	@Override
	public String dataReport(Data data) {
		return data.dataReport();
	}

	@Override
	public void searchDataReport() {
		
		System.out.println("<<<<< COLLECTING OK'S FROM REPORT '" + NAME_FOLDER + "' >>>>>");
		
		System.out.println("=====================================================================================================>>");
		System.out.println("Open File 'RelatorioPorCenario.xlsx' in Folder --> " + NAME_FOLDER + " -- " + new Date());
		File arq = new File(RunReport.WAY_FOLDER_REPORT + "\\" + NAME_FOLDER + "\\" + "RelatorioPorCenario.xlsx");
		if(!arq.exists()) {
			System.err.println("<<<FILE WAS NOT FOUND, THE NAME MAY BE DIVERGENT " + SearchDataToReportOK.class + ">>>");
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
				final boolean result = ServiceOK.verifyContentWorkbookOK(dct);

				if(result) {
					if(verifyID(dct.get("id"))) {
						wb = UtilReport.getWbMain();
						XSSFSheet sheetMain = wb.getSheet(GeneratorSheetMain.getSheetNameOK());
						Row rowWrite = sheetMain.createRow(Report.getLineWBOK());
						System.out.println(NAME_FOLDER + " - OK -- " + dct.get("id") + " -- WITRING...");
						ServiceOK.populateCellsOK(rowWrite, wb);
						Report.setLineWBOK(1);
					}
				} else {
					System.out.println(NAME_FOLDER + " - OK -- JUMPING SCENARIO " + dct.get("id") + "...");
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