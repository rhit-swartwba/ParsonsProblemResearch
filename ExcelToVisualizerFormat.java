import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.*;
import java.io.FileWriter;


public class ExcelToVisualizerFormat {
	
	public static void main(String[] args) {
		
		/*
		 * Change these fields to the description as appropriately
		 * Then just run the class
		 * Other code should not have to be modified to work assuming base parameters
		 */
		
		//IMPORTANT: Make sure to change the excel file to 'workbook' type and sort from (COL H) A-Z using the 'Question' Column
		
		//the file path to the excel file
		String excelFileWithDataFullPath = "C:\\Users\\swartwba\\OneDrive - Rose-Hulman Institute of Technology\\Desktop\\ResearchProject\\Data\\Spring2324Data\\CSSE_132_2324SpringCSSE132_HW3_all_submissions.xlsx";
		//change the last \\ part to the desired file name and choose where to place the file and the name of the file
		String newFileWithInfoFullPath = "C:\\Users\\swartwba\\OneDrive - Rose-Hulman Institute of Technology\\Desktop\\ResearchProject\\Data\\Spring2324Data\\FinalExam\\HW3AllSubmissionsTest.txt";
		//for answers-names "" - put that here (needs to be the same across all questions of the homework assignment)
		String chosenLabelForQuestions = "Answers";
		getExcelDataToVisualizerFormat(excelFileWithDataFullPath, newFileWithInfoFullPath, chosenLabelForQuestions);
			
	}
	
	/*
	 * Takes the Excel final of all submissions and creates a file with the data inside it in the desired format
	 * 
	 * @paramexcelFileWithDataFullPath - the path to the excel file with the data
	 * @paramnewFileWithInfoFullPath - the path where the new file (make sure to name it)
	 * @paramnewFileWithInfoFullPath - the label for all the blocks in Prairie Learn, needs to be the same in each question

	 */
	static void getExcelDataToVisualizerFormat(String excelFileWithDataFullPath, String newFileWithInfoFullPath,
			String chosenLabelForQuestions) {
		try {
			//Add in excel file to pull from
			File file = new File(excelFileWithDataFullPath);
			FileInputStream fis = new FileInputStream(file); // obtaining bytes from the file

			//Add in file want to write to
			FileWriter myWriter = new FileWriter(newFileWithInfoFullPath);

			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheetAt(0); // creating a Sheet object to retrieve object
			Iterator<Row> itr = sheet.iterator(); // iterating over excel file

			itr.next();
			String currentQuestion = "notAQuestion";
			boolean isNewQuestion = false;
			boolean isEmptyAnswer = false;
			while (itr.hasNext()) {
				Row row = itr.next();
				Iterator<Cell> cellIterator = row.cellIterator(); // iterating over each column
				String toWrite = "";

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					// Name: 3, Problem: 7; DateTime: 15, Submission: 16, Correct: 26
					//Needs the columns to be set up in this standard way, otherwise change number to match column names
					int currColumn = cell.getColumnIndex();
					if (currColumn == 0 || currColumn == 7 || currColumn == 16 || currColumn == 26) {
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_STRING: // field that represents string cell type
							if (currColumn == 16) {
								// System.out.print(cell.getStringCellValue() + "\t\t\t");
								String answer = parseSubmission(cell.getStringCellValue(), chosenLabelForQuestions);
								if(answer.equals(""))
								{
									isEmptyAnswer = true;
								}
								toWrite += answer + ",";

							} 
							else if(currColumn == 7){
								//instead of adding in the question name, simply add it to the file beforehand
								if(!cell.getStringCellValue().equals(currentQuestion))
								{
									isNewQuestion = true;
									myWriter.write("\n Question: " + cell.getStringCellValue() + "\n\n");
									System.out.println("\n Question: " + cell.getStringCellValue() + "\n\n");
									currentQuestion = cell.getStringCellValue();
								}
							}
							else {
								toWrite += cell.getStringCellValue() + ",";
							}
							break;
						case Cell.CELL_TYPE_NUMERIC: // field that represents number cell type
							toWrite += cell.getNumericCellValue() + ",";
							break;
						case Cell.CELL_TYPE_BOOLEAN: // field that represents boolean cell type
							if (cell.getBooleanCellValue() == true) {
								toWrite += "1";
							} else {
								toWrite += "0";
							}
							break;
						default:
						}
					}

				}
				if(!isEmptyAnswer)
				{
					if(isNewQuestion)
					{
						isNewQuestion = false;
					}
					myWriter.write(toWrite + "\n");
					System.out.print(toWrite);
					System.out.println();
				}
				
			}
			myWriter.close();
			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	
	private static String parseSubmission(String stringCellValue, String chosenLabelForQuestions) {
		String subAnswer = "";
		if(!stringCellValue.contains("[]"))
		{
			JSONObject obj = new JSONObject(stringCellValue);
			JSONArray arr = obj.getJSONArray(chosenLabelForQuestions);
			for (int i = 0; i < arr.length(); i++) {
				//switching from using tag to ranking
				int order = arr.getJSONObject(i).getInt("ranking");
				subAnswer += " " + order;
			}
		}
		return subAnswer;
	}
	

	
}
	
	

