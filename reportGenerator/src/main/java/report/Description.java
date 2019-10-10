package report;

public class Description {
	
	private static int pasta = 0;
	private static int environment = 0;
	private static int reTest = 0;
	private static int application  = 0;
	private static int almResult = 0;
	
	public static int getPasta() {
		return pasta;
	}
	public static void setPasta() {
		pasta++;
	}
	public static int getEnvironment() {
		return environment;
	}
	public static void setEnvironment() {
		environment++;
	}
	public static int getReTest() {
		return reTest;
	}
	public static void setReTest() {
		reTest++;
	}
	public static int getApplication() {
		return application;
	}
	public static void setApplication() {
		application++;
	}
	public static int getAlmResult() {
		return almResult;
	}
	public static void setAlmResult(int x) {
		Description.almResult = x;
	}
}