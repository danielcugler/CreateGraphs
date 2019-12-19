import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.NumberFormat;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileSystemView;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.labels.StandardCategoryItemLabelGenerator;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;

/*
 * Este programa gera graficos apenas para cenarios onde as respostas "nao sei/desconheco e ruim" ultrapassam os 40% das respostas
 * Pede-se 3 planilhas. Sao inseridos no grafico os 3 cursos 
 */

import Model.Question;


/*
 * This program was created to generate graphs for CPA (Comissao Propria de Avaliacao) do IFSP. 
 * This access a spreadsheet with questions in the first line, and answers in the vertical position. 
 */

public class MainGeraGraficosApenasCenariosMuitasRepostasRuinsTresCursos {

	public static void main(String[] args) {
		
		JOptionPane.showMessageDialog(null, "GERA APENAS GRAFICOS PARA CENÁRIOS COM PONTUAÇÃO RUIM\nEste programa foi desenvolvido para funcionar no Linux. \nSelecione a planilha desejada a seguir. Garanta que esta planilha tenha apenas uma aba e que todos os 3 arquivos TENHAM O MESMO NUMERO DE COLUNAS");
		
		//Opening ADS spreadsheet
		JOptionPane.showMessageDialog(null, "Selecione a planilha do curso ADS");		
		JFileChooser jfcADS = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
		int returnValueADS = jfcADS.showOpenDialog(null);
		File selectedFileADS = null;
		File selectedFileFolderADS = jfcADS.getCurrentDirectory();
		
		if (returnValueADS == JFileChooser.APPROVE_OPTION) {
			selectedFileADS = jfcADS.getSelectedFile();
		} else {
			JOptionPane.showMessageDialog(null, "Nenhum arquivo foi selecionado. O programa será fechado.");
			return;
		}
		
		FileInputStream fileADS;
		Workbook workbookADS = null;
		try {
			fileADS = new FileInputStream(new File(selectedFileADS.getAbsolutePath()));
			workbookADS = new XSSFWorkbook(fileADS);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			JOptionPane.showMessageDialog(null, "Error " + e.getMessage());
			e.printStackTrace();
		}
		Sheet sheetADS = workbookADS.getSheetAt(0);
		
		int numberOfColumns = sheetADS.getRow(0).getPhysicalNumberOfCells();
		System.out.println("Colunas: " + numberOfColumns);

		
		
		
		
		
		//Opening GPI spreadsheet
		JOptionPane.showMessageDialog(null, "Selecione a planilha do curso GPI");		
		JFileChooser jfcGPI = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
		int returnValueGPI = jfcGPI.showOpenDialog(null);
		File selectedFileGPI = null;
		File selectedFileFolderGPI = jfcGPI.getCurrentDirectory();
		
		if (returnValueGPI == JFileChooser.APPROVE_OPTION) {
			selectedFileGPI = jfcGPI.getSelectedFile();
		} else {
			JOptionPane.showMessageDialog(null, "Nenhum arquivo foi selecionado. O programa será fechado.");
			return;
		}
		
		FileInputStream fileGPI;
		Workbook workbookGPI = null;
		try {
			fileGPI = new FileInputStream(new File(selectedFileGPI.getAbsolutePath()));
			workbookGPI = new XSSFWorkbook(fileGPI);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			JOptionPane.showMessageDialog(null, "Error " + e.getMessage());
			e.printStackTrace();
		}
		Sheet sheetGPI = workbookGPI.getSheetAt(0);
		

		
		
		
		
		//Opening Pedagogia spreadsheet
		JOptionPane.showMessageDialog(null, "Selecione a planilha do curso Pedadogia");		
		JFileChooser jfcPedagogia = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
		int returnValuePedagogia = jfcPedagogia.showOpenDialog(null);
		File selectedFilePedagogia = null;
		File selectedFileFolderPedagogia = jfcPedagogia.getCurrentDirectory();
		
		if (returnValuePedagogia  == JFileChooser.APPROVE_OPTION) {
			selectedFilePedagogia  = jfcPedagogia .getSelectedFile();
		} else {
			JOptionPane.showMessageDialog(null, "Nenhum arquivo foi selecionado. O programa será fechado.");
			return;
		}		
		
		FileInputStream filePedagogia;
		Workbook workbookPedagogia = null;
		try {
			filePedagogia = new FileInputStream(new File(selectedFilePedagogia.getAbsolutePath()));
			workbookPedagogia = new XSSFWorkbook(filePedagogia);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			JOptionPane.showMessageDialog(null, "Error " + e.getMessage());
			e.printStackTrace();
		}
		Sheet sheetPedagogia = workbookPedagogia.getSheetAt(0);

		
		

		
		

		
		//for (int columnIndex=0; columnIndex < 1; columnIndex++) {
		for (int columnIndex=0; columnIndex < numberOfColumns; columnIndex++) { //number of columns is the number of questions
			System.out.println("*********************** QUESTION INDEX: " + columnIndex);

			int total = 0;

			//ADS
			Question currentQuestionADS = new Question();
			Cell questionCellADS = sheetADS.getRow(0).getCell(columnIndex);
			String questionADS = questionCellADS.getRichStringCellValue().getString();
			System.out.println("Question: " + questionADS);
			
			//GPI
			Question currentQuestionGPI = new Question();
			Cell questionCellGPI = sheetGPI.getRow(0).getCell(columnIndex);
			String questionGPI = questionCellGPI.getRichStringCellValue().getString();
			
			//Pedagogia
			Question currentQuestionPedagogia = new Question();
			Cell questionCellPedagogia = sheetPedagogia.getRow(0).getCell(columnIndex);
			String questionPedagogia = questionCellPedagogia.getRichStringCellValue().getString();

			
			//ADS
			for(Row row: sheetADS) {
				
				if(row.getRowNum() > 0) {
					Cell currentCell = row.getCell(columnIndex);
					String currentAnswer = null;
					if(currentCell != null) { 
						currentAnswer = currentCell.getRichStringCellValue().getString();
					} else {
						currentAnswer = "";
					}
					
					total++;
					
					currentQuestionADS.setQuestionDescription(questionADS);
					
					switch (currentAnswer) {
					case Question.NAO_SEI:
						currentQuestionADS.increaseOneNaoSei();
						break;
						
					case Question.RUIM:
						currentQuestionADS.increaseOneRuim();
						break;
						
					case Question.RAZOAVEL:
						currentQuestionADS.increaseOneRazoavel();
						break;
						
					case Question.BOM:
						currentQuestionADS.increaseOneBom();
						break;
						
					case Question.OTIMO:
						currentQuestionADS.increaseOneOtimo();
						break;
						
					case "":
						currentQuestionADS.increaseOneSemResposta();
						break;						
					}
				}
			}
			
			
			//GPI
			for(Row row: sheetGPI) {
				
				if(row.getRowNum() > 0) {
					Cell currentCell = row.getCell(columnIndex);
					String currentAnswer = null;
					if(currentCell != null) { 
						currentAnswer = currentCell.getRichStringCellValue().getString();
					} else {
						currentAnswer = "";
					}
					
					total++; //only increase it once, in the ADS
					
					currentQuestionGPI.setQuestionDescription(questionGPI);
					
					switch (currentAnswer) {
					case Question.NAO_SEI:
						currentQuestionGPI.increaseOneNaoSei();
						break;
						
					case Question.RUIM:
						currentQuestionGPI.increaseOneRuim();
						break;
						
					case Question.RAZOAVEL:
						currentQuestionGPI.increaseOneRazoavel();
						break;
						
					case Question.BOM:
						currentQuestionGPI.increaseOneBom();
						break;
						
					case Question.OTIMO:
						currentQuestionGPI.increaseOneOtimo();
						break;
						
					case "":
						currentQuestionGPI.increaseOneSemResposta();
						break;						
					}
				}
			}	
			
			
			//Pedagogia
			for(Row row: sheetPedagogia) {
				
				if(row.getRowNum() > 0) {
					Cell currentCell = row.getCell(columnIndex);
					String currentAnswer = null;
					if(currentCell != null) { 
						currentAnswer = currentCell.getRichStringCellValue().getString();
					} else {
						currentAnswer = "";
					}
					
					//total++; only increase it once, in the ADS
					
					currentQuestionPedagogia.setQuestionDescription(questionPedagogia);
					
					switch (currentAnswer) {
					case Question.NAO_SEI:
						currentQuestionPedagogia.increaseOneNaoSei();
						break;
						
					case Question.RUIM:
						currentQuestionPedagogia.increaseOneRuim();
						break;
						
					case Question.RAZOAVEL:
						currentQuestionPedagogia.increaseOneRazoavel();
						break;
						
					case Question.BOM:
						currentQuestionPedagogia.increaseOneBom();
						break;
						
					case Question.OTIMO:
						currentQuestionPedagogia.increaseOneOtimo();
						break;
						
					case "":
						currentQuestionPedagogia.increaseOneSemResposta();
						break;						
					}
				}
			}						
			
			
			
			System.out.println(Question.NAO_SEI + ": " + currentQuestionADS.getNaoSei() + " (" + (Float.valueOf(currentQuestionADS.getNaoSei())  / Float.valueOf(total)) + "%)");
			System.out.println(Question.RUIM + ": " + currentQuestionADS.getRuim() + " (" + (Float.valueOf(currentQuestionADS.getRuim())  / Float.valueOf(total)) + "%)");
			System.out.println(Question.RAZOAVEL + ": " + currentQuestionADS.getRazoavel() + " (" + (Float.valueOf(currentQuestionADS.getRazoavel())  / Float.valueOf(total)) + "%)");
			System.out.println(Question.BOM + ": " + currentQuestionADS.getBom() + " (" + (Float.valueOf(currentQuestionADS.getBom())  / Float.valueOf(total)) + "%)");
			System.out.println(Question.OTIMO + ": " + currentQuestionADS.getOtimo() + " (" + (Float.valueOf(currentQuestionADS.getOtimo())  / Float.valueOf(total)) + "%)");
			System.out.println(Question.SEM_RESPOSTA + ": " + currentQuestionADS.getSemResposta() + " (" + (Float.valueOf(currentQuestionADS.getSemResposta())  / Float.valueOf(total)) + "%)");
			System.out.println("TOTAL: " + total);
			
			//ADS
			float percentageNaoSeiPlusRuimADS;
			if(currentQuestionADS.getRazoavel() + currentQuestionADS.getBom() + currentQuestionADS.getOtimo() == 0) {
				percentageNaoSeiPlusRuimADS =  (currentQuestionADS.getNaoSei() + currentQuestionADS.getRuim()) / (currentQuestionADS.getRazoavel() + currentQuestionADS.getBom() + currentQuestionADS.getOtimo() + currentQuestionADS.getSemResposta());
			} else {
				percentageNaoSeiPlusRuimADS = (currentQuestionADS.getNaoSei() + currentQuestionADS.getRuim()) / (currentQuestionADS.getRazoavel() + currentQuestionADS.getBom() + currentQuestionADS.getOtimo());
			}
			
			//GPI
			float percentageNaoSeiPlusRuimGPI;
			if(currentQuestionGPI.getRazoavel() + currentQuestionGPI.getBom() + currentQuestionGPI.getOtimo() == 0) {
				percentageNaoSeiPlusRuimGPI =  (currentQuestionGPI.getNaoSei() + currentQuestionGPI.getRuim()) / (currentQuestionGPI.getRazoavel() + currentQuestionGPI.getBom() + currentQuestionGPI.getOtimo() + currentQuestionGPI.getSemResposta());
			} else {
				percentageNaoSeiPlusRuimGPI = (currentQuestionGPI.getNaoSei() + currentQuestionGPI.getRuim()) / (currentQuestionGPI.getRazoavel() + currentQuestionGPI.getBom() + currentQuestionGPI.getOtimo());
			}
			
			//Pedagogia
			float percentageNaoSeiPlusRuimPedagogia;
			if(currentQuestionPedagogia.getRazoavel() + currentQuestionPedagogia.getBom() + currentQuestionPedagogia.getOtimo() == 0) {
				percentageNaoSeiPlusRuimPedagogia =  (currentQuestionPedagogia.getNaoSei() + currentQuestionPedagogia.getRuim()) / (currentQuestionPedagogia.getRazoavel() + currentQuestionPedagogia.getBom() + currentQuestionPedagogia.getOtimo() + currentQuestionPedagogia.getSemResposta());
			} else {
				percentageNaoSeiPlusRuimPedagogia = (currentQuestionPedagogia.getNaoSei() + currentQuestionPedagogia.getRuim()) / (currentQuestionPedagogia.getRazoavel() + currentQuestionPedagogia.getBom() + currentQuestionPedagogia.getOtimo());
			}			
			
			

			
			
			
			if(percentageNaoSeiPlusRuimADS >= 0.40 || percentageNaoSeiPlusRuimGPI >= 0.40 || percentageNaoSeiPlusRuimPedagogia >= 0.40 ) {
				DefaultCategoryDataset dataset = new DefaultCategoryDataset();

				//ADS
				dataset.addValue(currentQuestionADS.getNaoSei(), Question.NAO_SEI, "ADS");
				dataset.addValue(currentQuestionADS.getRuim(), Question.RUIM, "ADS");
				dataset.addValue(currentQuestionADS.getRazoavel(), Question.RAZOAVEL, "ADS");
				dataset.addValue(currentQuestionADS.getBom(), Question.BOM, "ADS");
				dataset.addValue(currentQuestionADS.getOtimo(), Question.OTIMO, "ADS");
				dataset.addValue(currentQuestionADS.getSemResposta(), Question.SEM_RESPOSTA, "ADS");

				
				//GPI
				dataset.addValue(currentQuestionGPI.getNaoSei(), Question.NAO_SEI, "GPI");
				dataset.addValue(currentQuestionGPI.getRuim(), Question.RUIM, "GPI");
				dataset.addValue(currentQuestionGPI.getRazoavel(), Question.RAZOAVEL, "GPI");
				dataset.addValue(currentQuestionGPI.getBom(), Question.BOM, "GPI");
				dataset.addValue(currentQuestionGPI.getOtimo(), Question.OTIMO, "GPI");
				dataset.addValue(currentQuestionGPI.getSemResposta(), Question.SEM_RESPOSTA, "GPI");
				
				//Pedagogia
				dataset.addValue(currentQuestionPedagogia.getNaoSei(), Question.NAO_SEI, "Pegagogia");
				dataset.addValue(currentQuestionPedagogia.getRuim(), Question.RUIM, "Pegagogia");
				dataset.addValue(currentQuestionPedagogia.getRazoavel(), Question.RAZOAVEL, "Pegagogia");
				dataset.addValue(currentQuestionPedagogia.getBom(), Question.BOM, "Pegagogia");
				dataset.addValue(currentQuestionPedagogia.getOtimo(), Question.OTIMO, "Pegagogia");
				dataset.addValue(currentQuestionPedagogia.getSemResposta(), Question.SEM_RESPOSTA, "Pegagogia");				
				
				
				JFreeChart barChart = ChartFactory.createBarChart(
				         questionADS, 
				         "", 
				         "", 
				         dataset,
				         PlotOrientation.VERTICAL, 
				         true, 
				         true, 
				         false);
				

				
		        CategoryPlot plot = barChart.getCategoryPlot();

		        plot.getRenderer().setSeriesItemLabelsVisible(1, true);
		        plot.getRenderer().setDefaultItemLabelsVisible(true);
		        plot.getRenderer().setDefaultSeriesVisible(true);
		        barChart.getCategoryPlot().setRenderer(plot.getRenderer());

				
		        plot.getRenderer().setDefaultItemLabelGenerator(new StandardCategoryItemLabelGenerator("{3}", NumberFormat.getPercentInstance()));
		        plot.getRenderer().setDefaultItemLabelsVisible(true);
				
				
				
				
				         
			    int width = 640;    /* Width of the image */
			    int height = 480;   /* Height of the image */ 
			    questionADS = questionADS.replace("/", "-");
			    File myChartFile = new File(selectedFileFolderADS.getAbsolutePath() + "/" + questionADS + ".jpeg" ); 
			    
			    try {
					ChartUtils.saveChartAsJPEG(myChartFile , barChart , width , height);
			    } catch (IOException e) {
			    	JOptionPane.showMessageDialog(null, "Error " + e.getMessage());
			    	e.printStackTrace();
					
			    }				
			}
			

		}
	}

}
