package sheet_service_nok;

import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import enum_service.StyleFormatCell;
import report.Description;
import report.Report;
import search_massa.PlanilhaDTO;
import search_massa.SearchInfoMassa;

public class ServiceNOK extends Report {
	
	private static final CellStyle dataStyleCell(StyleFormatCell cellStyle, XSSFWorkbook wb) {
		return cellStyle.dataStyleWBCELL(wb);
	}

	public static void populateCellsNOK(Row row, XSSFWorkbook wb) {

		// TYPE ERROR --> 0
		Cell typeError = row.createCell(6); // CELL --> 6
		typeError.setCellValue(getValuesReport(0));
		typeError.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));
		switch (getValuesReport(0)) {
			case "AMBIENTE":
				Description.setEnvironment();
				break;
			case "MASSA":
				Description.setPasta();
				break;
			case "APLICAÇÃO":
				Description.setApplication();
				break;
			case "RE_TEST":
				Description.setReTest();
				break;
		}

		// ERROR --> 1
		Cell error = row.createCell(8); // CELL --> 8
		error.setCellValue(getValuesReport(1));
		error.setCellStyle(dataStyleCell(StyleFormatCell.BORDER, wb));

		// IDCENÁRIO --> 2
		Cell idCenario = row.createCell(0); // CELL --> 0
		idCenario.setCellValue(getValuesReport(2));
		idCenario.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// CENÁRIO --> 3
		Cell nomeCenario = row.createCell(1); // CELL --> 1
		nomeCenario.setCellValue(getValuesReport(3));
		nomeCenario.setCellStyle(dataStyleCell(StyleFormatCell.BORDER, wb));

		// MASSA --> 4
		Cell idMassa = row.createCell(7); // CELL --> 7
		idMassa.setCellValue(getValuesReport(4));
		idMassa.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// STATUS --> 5
		Cell status = row.createCell(2); // CELL --> 2
		status.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));
		status.setCellValue(getValuesReport(5));

		// EXECUÇÃ0 --> 6
		Cell qtdExecucao = row.createCell(3); // CELL --> 3
		qtdExecucao.setCellValue(getValuesReport(6));
		qtdExecucao.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// JSON --> 7
		Cell codeRequest = row.createCell(4); // CELL --> 4
		codeRequest.setCellValue(getValuesReport(7));
		codeRequest.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// HOSTNAME --> 8
		Cell hostName = row.createCell(5); // CELL --> 5
		hostName.setCellValue(getValuesReport(8));
		hostName.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// STATUS DEV --> 9
		Cell statusDev = row.createCell(9); // CELL --> 9
		statusDev.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));
		statusDev.setCellValue(getValuesReport(9));

		// DATA HORA --> 11
		Cell dataHora = row.createCell(11); // CELL --> 11
		dataHora.setCellValue(getValuesReport(11));
		dataHora.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// ID LOG --> 10
		Cell idLog = row.createCell(10); // CELL --> 10
		idLog.setCellValue(getValuesReport(10));
		idLog.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// EVILOG --> 12
		Cell evidenciaLog = row.createCell(12); // CELL --> 12
		evidenciaLog.setCellValue(getValuesReport(12));
		evidenciaLog.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));
	}
	
	public static boolean verifyContentWorkbookNOK(HashMap<String, String> dct) {

		// VERIFICA SE OS VALORES SÃO HEADERS
		if (dct.get("id").equals("ID") && dct.get("status").equals("Status")
				&& dct.get("tipoErro").equals("Tipo Erro")) {
			return false;
		}

		// TIPOERRO -- NÃO OBRG 0
		if (dct.get("tipoErro").isEmpty() ? true : false) {
			addValuesReport("");
		} else {
			addValuesReport(dct.get("tipoErro"));
		}

		// ERRO - NÃO OBRG 1
		if (dct.get("erro").isEmpty() ? true : false) {
			addValuesReport("");
		} else {
			addValuesReport(dct.get("erro"));
		}

		// ID -- OBRG 2
		if (dct.get("id").isEmpty() ? true : false) {
			return false;
		} else {
			addValuesReport(dct.get("id"));
		}

		// CENARIO -- OBRG 3
		if (dct.get("cenario").isEmpty() ? true : false) {
			return false;
		} else {
			addValuesReport(dct.get("cenario"));
		}

		// MASSA -- NÃO OBRG 4
		if (dct.get("massa").isEmpty() ? true : false) {
			SearchInfoMassa.lerPlanilhaID(getValuesReport(2), getValuesReport(3));
			addValuesReport((PlanilhaDTO.getIdMassa().isEmpty() || PlanilhaDTO.getIdMassa() == null) ? "SEM MASSA"
					: PlanilhaDTO.getIdMassa());
		} else {
			addValuesReport(dct.get("massa"));
		}

		// STATUS -- OBRG 5
		if (dct.get("status").isEmpty() || dct.get("status").equals("ok") ? true : false) {
			return false;
		} else {
			addValuesReport(dct.get("status"));
		}

		// EXECUÇÃO -- OBRG 6
		if (dct.get("execucao").isEmpty() ? true : false) {
			return false;
		} else {
			addValuesReport(dct.get("execucao"));
		}

		// JSON -- OBRG 7
		if (dct.get("json").isEmpty() ? true : false) {
			return false;
		} else {
			addValuesReport(dct.get("json"));
		}

		// HOSTNAME -- OBRG 8
		if (dct.get("hostName").isEmpty() ? true : false) {
			return false;
		} else {
			addValuesReport(dct.get("hostName"));
		}

		// STATUS DEV -- OBRG 9
		if (dct.get("statusDev").isEmpty() ? true : false) {
			return false;
		} else {
			if (dct.get("statusDev").equals("Automatizado")) {
				addValuesReport("Automatizado");
			} else {
				return false;
			}
		}

		// IDLOG -- OBRG 10
		if (dct.get("idLog").isEmpty() ? true : false) {
			return false;
		} else {
			addValuesReport(dct.get("idLog"));
		}

		// DATA HORA -- OBRG 11
		if (dct.get("dataHora").isEmpty() ? true : false) {
			return false;
		} else {
			addValuesReport(dct.get("dataHora"));
		}

		// EVIDÊNCIA LOG -- OBRG 12
		if (dct.get("eviLog").isEmpty() ? true : false) {
			return false;
		} else {
			addValuesReport(dct.get("eviLog"));
		}
		return true;
	}
}