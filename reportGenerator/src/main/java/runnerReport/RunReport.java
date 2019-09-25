package runnerReport;

import generatorWorkbook.GeneratorReport;
import generatorWorkbook.GeneratorSheetMain;
import sheetServiceNOK.SearchDataToReportNOK;
import sheetServiceOK.SearchDataToReportOK;
import utilsReport.UtilReport;

public class RunReport {

	//CAMINHO PARA AS PASTAS COM OS RELATÓRIOS
	public static final String WAY_FOLDER_REPORT = "E:\\URA RELATÓRIOS\\SETEMBRO\\4_Semana_Setembro\\25.09.19\\TARDE";
	//CAMINHO PARA A PLANILHA DE MASSA
	public static final String WAY_MASSA = "C:\\testes\\PlanilhaUra.xls";
	//PASTAS EXISTENTES PARA O RELATÓRIO 
	public static final String[] FOLDERS_SEARCH_INFO_TO_REPORT =  "Aux1;Aux2;Romulo".split(";");
	public static long timeExecution = System.currentTimeMillis();
	
	public static void main(String[] args) {
		new GeneratorSheetMain(FOLDERS_SEARCH_INFO_TO_REPORT[0]);
		
		for (int posi = 0; posi < FOLDERS_SEARCH_INFO_TO_REPORT.length; posi++) {
			final int value = UtilReport.verifyExistsFolders(FOLDERS_SEARCH_INFO_TO_REPORT[posi]);

			if (value != 0) {
				new SearchDataToReportNOK(FOLDERS_SEARCH_INFO_TO_REPORT[posi]);
				new SearchDataToReportOK(FOLDERS_SEARCH_INFO_TO_REPORT[posi], value);
			} else 
				break;
		}
		new GeneratorReport();		
		UtilReport.printDescription();
	}
}