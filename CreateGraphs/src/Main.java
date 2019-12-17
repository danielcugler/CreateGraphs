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

import Model.Question;


/*
 * This program was created to generate graphs for CPA (Comissao Propria de Avaliacao) do IFSP. 
 * This access a spreadsheet with questions in the first line, and answers in the vertical position. 
 */

public class Main {

	public static void main(String[] args) {
		
		JOptionPane.showMessageDialog(null, "Este programa foi desenvolvido para funcionar no Linux. \nSelecione a planilha desejada a seguir. Garanta que esta planilha tenha apenas uma aba.");
		
		
		
		JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());

		int returnValue = jfc.showOpenDialog(null);
		
		

		File selectedFile = null;
		File selectedFileFolder = jfc.getCurrentDirectory();
		
		if (returnValue == JFileChooser.APPROVE_OPTION) {
			selectedFile = jfc.getSelectedFile();
		} else {
			JOptionPane.showMessageDialog(null, "Nenhum arquivo foi selecionado. O programa será fechado.");
			return;
		}
		
		
		
		FileInputStream file;
		Workbook workbook = null;
		try {
			file = new FileInputStream(new File(selectedFile.getAbsolutePath()));
			workbook = new XSSFWorkbook(file);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			JOptionPane.showMessageDialog(null, "Error " + e.getMessage());
			e.printStackTrace();
		}
		

		Sheet sheet = workbook.getSheetAt(0);
		
		int numberOfColumns = sheet.getRow(0).getPhysicalNumberOfCells();
		System.out.println("Colunas: " + numberOfColumns);

		
		//for (int columnIndex=0; columnIndex < 1; columnIndex++) {
		for (int columnIndex=0; columnIndex < numberOfColumns; columnIndex++) {
			System.out.println("*********************** QUESTION INDEX: " + columnIndex);
			
			Question currentQuestion = new Question();
			int total = 0;

			Cell questionCell = sheet.getRow(0).getCell(columnIndex);
			String question = questionCell.getRichStringCellValue().getString();
			System.out.println("Question: " + question);
			
			for(Row row: sheet) {
				
				if(row.getRowNum() > 0) {
					Cell currentCell = row.getCell(columnIndex);
					String currentAnswer = null;
					if(currentCell != null) { 
						currentAnswer = currentCell.getRichStringCellValue().getString();
					} else {
						currentAnswer = "";
					}
					
					total++;
					
					currentQuestion.setQuestionDescription(question);
					
					switch (currentAnswer) {
					case Question.NAO_SEI:
						currentQuestion.increaseOneNaoSei();
						break;
						
					case Question.RUIM:
						currentQuestion.increaseOneRuim();
						break;
						
					case Question.RAZOAVEL:
						currentQuestion.increaseOneRazoavel();
						break;
						
					case Question.BOM:
						currentQuestion.increaseOneBom();
						break;
						
					case Question.OTIMO:
						currentQuestion.increaseOneOtimo();
						break;
						
					case "":
						currentQuestion.increaseOneSemResposta();
						break;						

					}

				}

			}
			System.out.println(Question.NAO_SEI + ": " + currentQuestion.getNaoSei() + " (" + (Float.valueOf(currentQuestion.getNaoSei())  / Float.valueOf(total)) + "%)");
			System.out.println(Question.RUIM + ": " + currentQuestion.getRuim() + " (" + (Float.valueOf(currentQuestion.getRuim())  / Float.valueOf(total)) + "%)");
			System.out.println(Question.RAZOAVEL + ": " + currentQuestion.getRazoavel() + " (" + (Float.valueOf(currentQuestion.getRazoavel())  / Float.valueOf(total)) + "%)");
			System.out.println(Question.BOM + ": " + currentQuestion.getBom() + " (" + (Float.valueOf(currentQuestion.getBom())  / Float.valueOf(total)) + "%)");
			System.out.println(Question.OTIMO + ": " + currentQuestion.getOtimo() + " (" + (Float.valueOf(currentQuestion.getOtimo())  / Float.valueOf(total)) + "%)");
			System.out.println(Question.SEM_RESPOSTA + ": " + currentQuestion.getSemResposta() + " (" + (Float.valueOf(currentQuestion.getSemResposta())  / Float.valueOf(total)) + "%)");
			System.out.println("TOTAL: " + total);
			
			
			
			
			
			DefaultCategoryDataset dataset = new DefaultCategoryDataset();
			
			dataset.addValue(currentQuestion.getNaoSei(), Question.NAO_SEI, "");
			dataset.addValue(currentQuestion.getRuim(), Question.RUIM, "");
			dataset.addValue(currentQuestion.getRazoavel(), Question.RAZOAVEL, "");
			dataset.addValue(currentQuestion.getBom(), Question.BOM, "");
			dataset.addValue(currentQuestion.getOtimo(), Question.OTIMO, "");
			dataset.addValue(currentQuestion.getSemResposta(), Question.SEM_RESPOSTA, "");
			
			JFreeChart barChart = ChartFactory.createBarChart(
			         question, 
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
		    question = question.replace("/", "-");
		    File myChartFile = new File(selectedFileFolder.getAbsolutePath() + "/" + question + ".jpeg" ); 
		    
		    try {
				ChartUtils.saveChartAsJPEG(myChartFile , barChart , width , height);
		    } catch (IOException e) {
		    	JOptionPane.showMessageDialog(null, "Error " + e.getMessage());
		    	e.printStackTrace();
				
		    }
		}
	}

}
