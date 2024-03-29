package utils_report;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.channels.FileChannel;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import enum_service.Data;
import enum_service.DataMachines;
import junit.framework.Assert;
import report.Description;
import report.Report;
import runner.RunReport;

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
	
	public static int returnsNumber(String folder) {
		if(folder.contains("1")) {
			return dataReportMachine(DataMachines.HELP1);
		} else if (folder.contains("2")) {
			return dataReportMachine(DataMachines.HELP2);
		} else if (folder.contains("3")) {
			return dataReportMachine(DataMachines.HELP3);
		} else if (folder.contains("ug")) {
			return dataReportMachine(DataMachines.MACHINEHUGO);
		} else if (folder.contains("mul")) {
			return dataReportMachine(DataMachines.MACHINEROMULO);
		} else {
			return 0;
		}
	}
	
	public static File indentWay(String arqWay) {
		return ((arqWay == null || arqWay.isEmpty()) ? new File(RunReport.WAY_FOLDER_REPORT) : new File(RunReport.WAY_FOLDER_REPORT, arqWay));
	} 
	
	public static void finish() {
		Assert.assertTrue(false);
	}
	
	public static XSSFWorkbook getWbMain() {
		File arqReport = UtilReport.indentWay(dataReport(Data.FILE_REPORT_NAME));
		try {
			return new XSSFWorkbook(new FileInputStream(arqReport));
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
	
	public static void copyFile(File source, File destination) {
		try (FileChannel sourceOfc = new FileInputStream(source).getChannel();
				FileChannel destinationCopy = new FileOutputStream(destination).getChannel()) {
			sourceOfc.transferTo(0, sourceOfc.size(), destinationCopy);
			System.out.println("<<<<< FILE '" + source.getName() + "' COPIED IN: " + RunReport.WAY_FOLDER_REPORT + " >>>>>");
		} catch (IOException e) {
			System.err.println(e.getMessage() + "\n" + e);
		} 
	}
	
	public static int searchID(XSSFSheet planilha, String id) {
		int rowNumber = -1;
		for (Row row : planilha) {
			Cell cell = row.getCell(0);
			if (cell.toString().trim().equals(id)) {
				rowNumber = row.getRowNum();
				break;
			}
		}
		return rowNumber;
	}

	public static void printDescription() {
		System.out.println("<<<<< DESCRIPTION NOK'S AND OK'S >>>>>\n"
						 + (Description.getAlmResult() != 0 ? "■ ALM RESULT ► " + Description.getAlmResult() + "\n" : "")
						 + "■ OK'S ► " + Report.getListOK().size() + "\n"
						 + "■ NOK'S TOTAL ► " + (786 - Report.getListOK().size()) + "\n"
						 + "■ NOK'S REPORT ► " + (((Report.getLineWBNOK() - 1) > 786) ? "DUPLICATE NOK'S -- " + 
						 (Report.getLineWBNOK() - 1) : (Report.getLineWBNOK() - 1)) + "\n\n"
						 + "<<<<< DESCRIPTION ERROR >>>>>\n"
						 + "■ ENVIRONMENT ► " + Description.getEnvironment() + "\n"
						 + "■ MASSA ► " + Description.getPasta() + "\n"
						 + "■ RE_TEST ► " + Description.getReTest() + "\n"
						 + "■ APPLICATION ► " + Description.getApplication() + "\n"
						 + "■ SCENARIOS NOT AUTOMATED ► 287\n"
						 + "■ SCENARIOS AUTOMATED ► 786\n" 
						 + "■ TIME EXECUTION ► " + (new SimpleDateFormat("mm").format(new Date(System.currentTimeMillis() - RunReport.TIME_EXECUTION))) + " Minutes\n" 
						 + "=====================================================================================================>>");
	}
}