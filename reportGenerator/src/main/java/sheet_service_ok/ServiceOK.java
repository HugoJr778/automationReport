package sheet_service_ok;

import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import enum_service.StyleFormatCell;
import report.Report;
import search_massa.PlanilhaDTO;
import search_massa.SearchInfoMassa;

public class ServiceOK extends Report {
	
	private static final CellStyle dataStyleCell(StyleFormatCell cellStyle, XSSFWorkbook wb) {
		return cellStyle.dataStyleWBCELL(wb);
	}

	public static void populateCellsOK(Row row, XSSFWorkbook wb) {

		// STATUS --> 0
		Cell status = row.createCell(2); // CELL --> 2
		status.setCellValue(getValuesReport(0));
		status.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// IDCENÁRIO --> 1
		Cell idCenario = row.createCell(0); // CELL --> 0
		idCenario.setCellValue(getValuesReport(1));
		idCenario.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// CENÁRIO --> 2
		Cell nomeCenario = row.createCell(1); // CELL --> 1
		nomeCenario.setCellValue(getValuesReport(2));
		nomeCenario.setCellStyle(dataStyleCell(StyleFormatCell.BORDER, wb));

		// EXECUÇÃ0 --> 3
		Cell qtdExecucao = row.createCell(3); // CELL --> 3
		qtdExecucao.setCellValue(getValuesReport(3));
		qtdExecucao.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// JSON --> 4
		Cell codeRequest = row.createCell(4); // CELL --> 4
		codeRequest.setCellValue(getValuesReport(4));
		codeRequest.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// HOSTNAME --> 5
		Cell hostName = row.createCell(5); // CELL --> 5
		hostName.setCellValue(getValuesReport(5));
		hostName.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// MASSA --> 6
		Cell idMassa = row.createCell(7); // CELL --> 7
		idMassa.setCellValue(getValuesReport(6));
		idMassa.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// STATUS DEV --> 7
		Cell statusDev = row.createCell(9); // CELL --> 9
		statusDev.setCellValue(getValuesReport(7));
		statusDev.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// ID LOG --> 8
		Cell idLog = row.createCell(10); // CELL --> 10
		idLog.setCellValue(getValuesReport(8));
		idLog.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// DATA HORA --> 9
		Cell dataHora = row.createCell(11); // CELL --> 11
		dataHora.setCellValue(getValuesReport(9));
		dataHora.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// EVILOG --> 10
		Cell evidenciaLog = row.createCell(12); // CELL --> 12
		evidenciaLog.setCellValue(getValuesReport(10));
		evidenciaLog.setCellStyle(dataStyleCell(StyleFormatCell.CENTER, wb));

		// TYPE ERROR --> 11
		Cell typeError = row.createCell(6); // CELL --> 6
		typeError.setCellValue(getValuesReport(11));
		typeError.setCellStyle(dataStyleCell(StyleFormatCell.BORDER, wb));

		// ERROR --> 12
		Cell error = row.createCell(8); // CELL --> 8
		error.setCellValue(getValuesReport(12));
		error.setCellStyle(dataStyleCell(StyleFormatCell.BORDER, wb));
		
		Report.addOks(getValuesReport(0) + ";" + getValuesReport(3) + ";" + getValuesReport(4) + ";" + getValuesReport(5) + ";" + 
					  getValuesReport(11) + ";" + getValuesReport(6) + ";" + getValuesReport(12) + ";" + getValuesReport(8) + ";" + 
					  getValuesReport(9) + ";" + getValuesReport(10));
	}

	public static boolean verifyContentWorkbookOK(HashMap<String, String> dct) {

		// VERIFICA SE OS VALORES SÃO HEADERS
		if (dct.get("id").equals("ID") && dct.get("status").equals("Status")
				&& dct.get("tipoErro").equals("Tipo Erro")) {
			return false;
		}

		// STATUS -- OBRG 0
		if (dct.get("status").isEmpty() ? true : false) {
			return false;
		} else {
			if (dct.get("status").equals("ok") ? true : false) {
				addValuesReport(dct.get("status"));
			} else {
				return false;
			}
		}

		// ID -- OBRG 1
		if (dct.get("id").isEmpty() ? true : false) {
			return false;
		} else {
			addValuesReport(dct.get("id"));
		}

		// CENARIO -- OBRG 2
		if (dct.get("cenario").isEmpty() ? true : false) {
			return false;
		} else {
			addValuesReport(dct.get("cenario"));
		}

		// EXECUÇÃO -- NÃO OBRG 3
		if (dct.get("execucao").isEmpty() ? true : false) {
			addValuesReport("SEM Nº DE EXECUÇÕES");
		} else {
			addValuesReport(dct.get("execucao"));
		}

		// JSON -- NÃO OBRG 4
		if (dct.get("json").isEmpty() ? true : false) {
			addValuesReport("SEM RESQUEST CODE JSON");
		} else {
			addValuesReport(dct.get("json"));
		}

		// HOSTNAME -- NÃO OBRG 5
		if (dct.get("hostName").isEmpty() ? true : false) {
			addValuesReport("SEM HOSTNAME");
		} else {
			addValuesReport(dct.get("hostName"));
		}

		// MASSA -- NÃO OBRG 6
		if (dct.get("massa").isEmpty() ? true : false) {
			SearchInfoMassa.lerPlanilhaID(getValuesReport(1), getValuesReport(2));
			addValuesReport((PlanilhaDTO.getIdMassa().isEmpty() || PlanilhaDTO.getIdMassa() == null) ? "SEM MASSA"
					: PlanilhaDTO.getIdMassa());
		} else {
			addValuesReport(dct.get("massa"));
		}

		// STATUS DEV -- NÃO OBRG 7
		if (dct.get("statusDev").isEmpty() ? true : false) {
			return false;
		} else {
			if (dct.get("statusDev").equals("Automatizado")) {
				addValuesReport("Automatizado");
			} else {
				return false;
			}
		}

		// IDLOG -- NÃO OBRG 8
		if (dct.get("idLog").isEmpty() ? true : false) {
			addValuesReport("SEM ID E ÁUDIO DO LOG");
		} else {
			addValuesReport(dct.get("idLog"));
		}

		// DATA HORA -- NÃO OBRG 9
		if (dct.get("dataHora").isEmpty() ? true : false) {
			addValuesReport("SEM DATA E HORA");
		} else {
			addValuesReport(dct.get("dataHora"));
		}

		// EVIDÊNCIA LOG -- NÃO OBRG 10
		if (dct.get("eviLog").isEmpty() ? true : false) {
			addValuesReport("SEM EVIDÊNCIA DO LOG");
		} else {
			addValuesReport(dct.get("eviLog"));
		}

		// TIPOERRO -- NÃO OBRG 11
		addValuesReport("");

		// ERRO - NÃO OBRG 12
		addValuesReport("");

		return true;
	}
}