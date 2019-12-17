package Model;

public class Question {
	
	public static final String NAO_SEI = "Não sei / desconheço / não se aplica"; 
	public static final String RUIM = "Ruim"; 
	public static final String RAZOAVEL = "Razoável";
	public static final String BOM = "Bom";
	public static final String OTIMO = "Ótimo";
	public static final String SEM_RESPOSTA = "Sem resposta";

	
	private String questionDescription;
	private int naoSei;
	private int ruim;
	private int razoavel;
	private int bom;
	private int otimo;
	int semResposta;
	
	
	public String getQuestionDescription() {
		return questionDescription;
	}
	public void setQuestionDescription(String description) {
		this.questionDescription = description;
	}
	
	public int getNaoSei() {
		return naoSei;
	}
	public void increaseOneNaoSei() {
		this.naoSei++;
	}
	
	public int getRuim() {
		return ruim;
	}
	public void increaseOneRuim() {
		this.ruim++;
	}
	
	public int getRazoavel() {
		return razoavel;
	}
	public void increaseOneRazoavel() {
		this.razoavel++;
	}
	
	public int getBom() {
		return bom;
	}
	public void increaseOneBom() {
		this.bom++;
	}
	
	public int getOtimo() {
		return otimo;
	}
	public void increaseOneOtimo() {
		this.otimo++;
	}
	
	public int getSemResposta() {
		return semResposta;
	}
	public void increaseOneSemResposta() {
		this.semResposta++;
	}

}
