package search_massa;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import runner.RunReport;
import utils_report.UtilReport;

public final class SearchInfoMassa {

	public static synchronized void lerPlanilhaID(String id, String nomeCenario) {
		HSSFWorkbook workbook = null;
		HSSFSheet sheet = null;
		FileInputStream file = null;
		String idMassa = null;
		String idRenavam = null;
		String idCodigoBarras = null;
		PlanilhaDTO.setNomeCenario(nomeCenario);
		PlanilhaDTO.setIdCenario(id);
		
		try {
			file = new FileInputStream(new File(RunReport.WAY_MASSA));

			workbook = new HSSFWorkbook(file);
			sheet = workbook.getSheet("UR");

			for (Row row : sheet) {

				Cell cellID = row.getCell(0);

				if (cellID.toString().trim().equals(id)) {
					idMassa = row.getCell(2).toString();
					idCodigoBarras = row.getCell(3).toString();
					idRenavam = row.getCell(4).toString();

					PlanilhaDTO.setValor(row.getCell(5).toString());
					PlanilhaDTO.setTelefone(row.getCell(6).getStringCellValue());
					PlanilhaDTO.setDataPassada(row.getCell(7).toString().replaceAll("\\D", ""));
					PlanilhaDTO.setDataFutura(row.getCell(8).toString().replaceAll("\\D", ""));						
					PlanilhaDTO.setIdMassa(idMassa);
					break;
				}
			}
			sheet = workbook.getSheet("Massa");
			for (Row row : sheet) {

				Cell cellID = row.getCell(0);

				if (cellID.toString().trim().equals(idMassa)) {
					PlanilhaDTO.setCpf(row.getCell(1).getStringCellValue().replaceAll("\\D", ""));
					PlanilhaDTO.setDataNascimento(row.getCell(2).toString());
					PlanilhaDTO.setAgencia(row.getCell(4).toString().replaceAll("\\D", ""));
					PlanilhaDTO.setContaCorrente(row.getCell(5).toString().replaceAll("\\D", ""));
					PlanilhaDTO.setContaPoupanca(row.getCell(6).toString().replaceAll("\\D", ""));
					PlanilhaDTO.setNumeroCartao(row.getCell(7).toString().replaceAll("\\D", ""));
					PlanilhaDTO.setCvv(row.getCell(8).toString());
					if (row.getCell(9).toString().trim().toLowerCase().equals("bloqueado")) {
						PlanilhaDTO.setTemCartaoBloqueado(true);
					} else {
						PlanilhaDTO.setTemCartaoBloqueado(false);
					}
					String senha4dig = row.getCell(10).toString().replaceAll("\\D", "");
					PlanilhaDTO.setSenha4Dig((!senha4dig.isEmpty()) ? senha4dig.substring(0, 4) : "");
					String senha6dig = row.getCell(11).toString().replaceAll("\\D", "");
					PlanilhaDTO.setSenha6Dig((!senha6dig.isEmpty()) ? senha6dig.substring(0, 6) : "");
					String assinaturaEletronica = row.getCell(12).toString().replaceAll("\\D", "");
					PlanilhaDTO.setAssinaturaEletronica(
							(!assinaturaEletronica.isEmpty()) ? assinaturaEletronica.substring(0, 6) : "");
					PlanilhaDTO.setAgenciaFavorecido(row.getCell(13).toString());
					PlanilhaDTO.setContaFavorecido(row.getCell(14).toString());
					PlanilhaDTO.setDocTed(row.getCell(15).toString());
					PlanilhaDTO.setInvestimento(row.getCell(16).toString());
					PlanilhaDTO.setProdutos(row.getCell(17).toString());
					String talaoCheque = row.getCell(18).toString().replaceAll("\\D", "");
					PlanilhaDTO.setTalaoCheque((!talaoCheque.isEmpty()) ? talaoCheque.substring(0, 6) : "");
					break;
				}
			}
			sheet = workbook.getSheet("CodigoDeBarras");
			for (Row row : sheet) {

				Cell cellID = row.getCell(0);

				if (cellID.toString().trim().equals(idCodigoBarras)) {
					PlanilhaDTO.setCodigoBarra(row.getCell(2).toString().replaceAll("\\D", ""));
					break;
				}
			}
			sheet = workbook.getSheet("Renavam");
			for (Row row : sheet) {
				Cell cellID = row.getCell(0);
				if (cellID.toString().trim().equals(idRenavam)) {
					PlanilhaDTO.setRenavam1(row.getCell(1).toString());
					PlanilhaDTO.setRenavam2(row.getCell(2).toString());
					break;
				}
			} 
		} catch (IOException e) {
			System.err.println(e.getMessage() + "\n" + e);
			UtilReport.finish();
		} finally {
			try {
				if(workbook != null)
					workbook.close();
				if(file != null) 
					file.close();
			} catch (IOException e) {
				System.err.println(e.getMessage() + "\n" + e);
				UtilReport.finish();
			}
		}
	}
}