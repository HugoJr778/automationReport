package enumService;

import utilsReport.UtilReport;

public enum Data {
	
	FILE_REPORT_NAME {
		@Override
		public String dataReport() {
			return "REPORT_URA_" + hourDate[1] + EXTENSION_REPORT;
		}
	},
	
	DATE {
		@Override
		public String dataReport() {
			return hourDate[0];
		}
	};
	
	public final String data() {
		return dataReport();
	}
	
	private static String[] hourDate = UtilReport.date("HH.mm.ss", "dd.MM.yyyy").split(";");
	private static final String EXTENSION_REPORT = ".xlsx";
	public abstract String dataReport();
}