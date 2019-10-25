package runnerReport;

import java.io.File;

import generatorWorkbook.ALM;
import generatorWorkbook.GeneratorReport;
import generatorWorkbook.GeneratorSheetMain;
import sheetServiceNOK.SearchDataToReportNOK;
import sheetServiceOK.SearchDataToReportOK;
import utilsReport.UtilReport;

public class RunReport {

	public static final boolean GENERATE_ALM = true;
	public static final String WAY_MASSA = "C:\\testes\\PlanilhaUra.xls";
	public static final String WAY_FOLDER_REPORT = "D:\\URA RELATÓRIOS\\OUTUBRO\\4_Semana_Outubro\\25.10.19";
	public static long timeExecution = System.currentTimeMillis();
	
	public static void main(String[] args) {
		final File[] FOLDERS_EXISTING = UtilReport.indentWay(null).listFiles();
		new GeneratorSheetMain(FOLDERS_EXISTING[0].getName().contains("ALM") ? FOLDERS_EXISTING[1].getName() : FOLDERS_EXISTING[0].getName());
		
		for (int posi = 0; posi < FOLDERS_EXISTING.length; posi++) {
			final int value = UtilReport.returnsNumber(FOLDERS_EXISTING[posi].getName());

			if (value != 0) {
				new SearchDataToReportNOK(FOLDERS_EXISTING[posi].getName());
				new SearchDataToReportOK(FOLDERS_EXISTING[posi].getName(), value);
			} else 
				break;
		}
		new GeneratorReport();
		if(GENERATE_ALM)
			new ALM().writeAlmResult();
		UtilReport.printDescription();
	}
}