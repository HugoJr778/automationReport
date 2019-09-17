package utilsReport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import enumService.Data;
import enumService.DataMachines;
import junit.framework.Assert;
import report.Description;
import report.Report;
import runnerReport.RunReport;

public class UtilReport extends Report {
	
	private static final int dataReportMachine(DataMachines data) {
		return data.getMachine();
	}
	
	private static final String dataReport(Data data) {
		return data.dataReport();
	}

	public static String date(String formatHour, String formatDay) {
		Date dataHora = new Date();
		if(formatHour != null && (!formatHour.isEmpty()) &&
			 formatDay != null && (!formatDay.isEmpty())) {
			String hora = new SimpleDateFormat(formatHour).format(dataHora);
			String data = new SimpleDateFormat(formatDay).format(dataHora);
			return hora + ";" + data;
		} else if (formatHour != null && (!formatHour.isEmpty()) &&
					  (formatDay == null)) {
			return new SimpleDateFormat(formatHour).format(dataHora);
		} else if (formatDay != null && (!formatDay.isEmpty()) && 
					  (formatHour == null)) {
			return new SimpleDateFormat(formatDay).format(dataHora);
		} else {
			String hora = new SimpleDateFormat("HH:mm").format(dataHora);
			String data = new SimpleDateFormat("dd/MM/yyyy").format(dataHora);
			return hora + ";" + data;
		}
	}
	
	public static int verifyExistsFolders(String folder) {
		File arq = indentWay(folder);
		
		if(arq.exists() && folder.contains("Aux1")) {
			return dataReportMachine(DataMachines.HELP1);
		} else if (arq.exists() && folder.contains("Aux2")) {
			return dataReportMachine(DataMachines.HELP2);
		} else if (arq.exists() && folder.contains("Aux3")) {
			return dataReportMachine(DataMachines.HELP3);
		} else if (arq.exists() && folder.contains("Hugo")) {
			return dataReportMachine(DataMachines.MACHINEHUGO);
		} else if (arq.exists() && folder.contains("Romulo")) {
			return dataReportMachine(DataMachines.MACHINEROMULO);
		} else {
			return 0;
		}
	}
	
	public static File indentWay(String arqWay) {
		return new File(RunReport.WAY_FOLDER_REPORT + "\\" + arqWay);
	} 
	
	public static void finish() {
		Assert.assertTrue(false);
	}
	
	public static XSSFWorkbook getWbMain() {
		File arqReport = UtilReport.indentWay(dataReport(Data.FILE_REPORT_NAME));
		try {
			XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(arqReport));
			return wb;
		} catch (FileNotFoundException e) {
			System.err.println(e.getMessage() + "\n" + e);
		} catch (IOException e) {
			System.err.println(e.getMessage() + "\n" + e);
		}
		return null;
	}
	
	public static void populateCellsHeaders(Row row, XSSFWorkbook wb) {
		XSSFCellStyle style = wb.createCellStyle();
		Cell cell;
		
		for(int posi = 0; posi < 13; posi++) {
			style.cloneStyleFrom(getStyleCell(posi));
			cell = row.createCell(posi);
			cell.setCellStyle(style);
			cell.setCellValue(getValuesReport(posi));
		}
	}

	public static void printDescription() {
		System.out.println("<<<<< DESCRIPTION NOK'S AND OK'S >>>>>\n"
						 + "■ OK'S ► " + Report.getListOK().size() + "\n"
						 + "■ NOK'S ► " + (((Report.getLineWBNOK() - 1) > 786) ? "DUPLICATE NOK'S -- " + 
						 (Report.getLineWBNOK() - 1) : (Report.getLineWBNOK() - 1)) + "\n"
						 + "<<<<< DESCRIPTION ERROR >>>>>\n"
						 + "■ ENVIRONMENT ► " + Description.getEnvironment() + "\n"
						 + "■ MASSA ► " + Description.getPasta() + "\n"
						 + "■ RE_TEST ► " + Description.getReTest() + "\n"
						 + "■ APPLICATION ► " + Description.getApplication() + "\n"
						 + "■ SCENARIOS NOT AUTOMATED ► 287\n"
						 + "■ SCENARIOS AUTOMATED ► 786\n"
						 + "=====================================================================================================>>");
	}
}