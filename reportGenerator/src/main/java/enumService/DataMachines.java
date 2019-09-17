package enumService;

public enum DataMachines {

	HELP1(1),
	
	HELP2(2),
	
	HELP3(3),
	
	MACHINEHUGO(4),
	
	MACHINEROMULO(5);
	
	private int machine;
	
	public int getMachine() {
		return machine;
	}
	
	private DataMachines(final int machine) {
		this.machine = machine;
	} 
}