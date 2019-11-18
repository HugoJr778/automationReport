package runner;

import java.io.File;

import generator_workbook.ALM;
import generator_workbook.GeneratorReport;
import generator_workbook.GeneratorSheetMain;
import sheet_service_nok.SearchDataToReportNOK;
import sheet_service_ok.SearchDataToReportOK;
import utils_report.UtilReport;

public class RunReport {

	public static final String WAY_MASSA = "C:\\testes\\PlanilhaUra.xls";
	public static final String WAY_FOLDER_REPORT = "";
	public static final String WAY_GENERATE_ALM = "";
	public static final long TIME_EXECUTION = System.currentTimeMillis();
	
	public static void main(String[] args) {
		final File[] foldersExisting = UtilReport.indentWay(null).listFiles();
		new GeneratorSheetMain(foldersExisting[0].getName().contains("ALM") ? foldersExisting[1].getName() : foldersExisting[0].getName());
		
		for (int posi = 0; posi < foldersExisting.length; posi++) {
			final int value = UtilReport.returnsNumber(foldersExisting[posi].getName());

			if (value != 0) {
				new SearchDataToReportNOK(foldersExisting[posi].getName());
				new SearchDataToReportOK(foldersExisting[posi].getName(), value);
			} else {
				break;
			}
		}
		new GeneratorReport();
		if(!WAY_GENERATE_ALM.isEmpty())
			new ALM().writeAlmResult();
		UtilReport.printDescription();
	}
}