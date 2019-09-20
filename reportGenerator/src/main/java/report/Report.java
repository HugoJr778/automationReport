package report;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;

import utilsReport.UtilReport;

public class Report {
	
	private static List<CellStyle> styleCell = new ArrayList<CellStyle>(12);
	private static List<String> valuesReport = new ArrayList<String>(12);
	private static List<String> listOK = new ArrayList<String>();
	public static List<String> oks = new ArrayList<String>();
	private static int lineWBNOK = 0;
	private static int lineWBOK = 0;
	

	public static String getOks(int index) {
		return (index > oks.size()) ? "<<< POSITION INVALID >>>" : oks.get(index);
	}

	public static void addOks(String value) {
		oks.add(value);
		System.out.println(oks);
	}
	
	public static List<String> getListOK() {
		return listOK;
	}
	
	public static String getListOK(int index) {
		return listOK.get(index);
	}

	public static void addListOK(String element) {
		if(element.isEmpty()) {
			System.err.println("<<<ELEMENT IS EMPTY - LIST_OK " + Report.class + ">>>");
			UtilReport.finish();
		}
		listOK.add(element);
	}
	
	public static int getLineWBOK() {
		return lineWBOK;
	}

	public static void setLineWBOK(int lineWBOK) {
		Report.lineWBOK += lineWBOK;
	}
	
	public static int getLineWBNOK() {
		return lineWBNOK;
	}

	public static void setLineWBNOK(int add) {
		Report.lineWBNOK += add;
	}
	
	public static void addStyleCell(CellStyle value) {
		styleCell.add(value);
	}
	
	public static void addValuesReport(String value) {
		valuesReport.add(value);
	}
	
	public static CellStyle getStyleCell(int posi) {
		CellStyle value = styleCell.get(posi);
		if(value != null)
			return value;
		else 
			return null;		
	}
	
	public static String getValuesReport(int posi) {
		String value = valuesReport.get(posi);
		if(value == null) {
			return "";
		} else if(!value.isEmpty())
			return value;
		else 
			return "";		
	}
	
	public static void clearStyleCell() {
		styleCell.clear();
	}
 	
	public static void clearValuesReport() {
		valuesReport.clear();
	}
	
	public static void printValuesReport() {
		for(int posi = 0; posi < valuesReport.size(); posi++) {
			System.out.printf("ArrayList values position: %02d º --> %s%n", posi, getValuesReport(posi));
		}
	}
}