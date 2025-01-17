import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.*;
import java.io.FileWriter; // Import the FileWriter class

/*
 * Takes the excel all submissions from prairie learn for a homework assignment and then parses the answers to the file
 * in the desired format for the visualizer.
 * 
 * Desired format: Name, QuestionName, Ordering (1 3 4 2 5), Correctness (0 or 1)
 * Can be updated to add in Date/Time as well
 * Does not currently account for distracting values - error will null values (see parsing method) and not able to be distguished from each other
 * 
 * Assumptions:
 * This is for a single homework assignment.
 * All questions have the same answers-names string - need to put in manually
 * Grading method is ranking (not sure if necessary)
 * Single variant is true for each question
 * 
 * 
 * Author: Blaise Swartwood
 */
public class ExcelToDataFile {

	public static void main(String[] args) {
		
		/*
		 * Change these fields to the description as appropriately
		 * Then just run the class
		 * Other code should not have to be modified to work assuming base parameters
		 */
		
		//Make sure to change the excel file to 'workbook' type and sort from A-Z using the 'Question' Column
		
		//the file path to the excel file
		String excelFileWithDataFullPath = "C:\\Users\\swartwba\\OneDrive - Rose-Hulman Institute of Technology\\Desktop\\ResearchProject\\CSSE_132_Winter2024_HW2_all_submissions.xlsx";
		//change the last \\ part to the desired file name and choose where to place the file
		String newFileWithInfoFullPath = "C:\\Users\\swartwba\\OneDrive - Rose-Hulman Institute of Technology\\Desktop\\ResearchProject\\CSSE132HW2222.txt";
		//for answers-names "" - put that here (needs to be the same across all questions of the homework assignment)
		String chosenLabelForQuestions = "Answers";
		getExcelDataToFile(excelFileWithDataFullPath, newFileWithInfoFullPath, chosenLabelForQuestions);
			
	}

	/*
	 * Takes the Excel final of all submissions and creates a file with the data inside it in the desired format
	 * 
	 * @paramexcelFileWithDataFullPath - the path to the excel file with the data
	 * @paramnewFileWithInfoFullPath - the path where the new file (make sure to name it)
	 * @paramnewFileWithInfoFullPath - the label for all the blocks in prarire learn, needs to be the same in each question

	 */
	private static void getExcelDataToFile(String excelFileWithDataFullPath, String newFileWithInfoFullPath,
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
			while (itr.hasNext()) {
				Row row = itr.next();
				Iterator<Cell> cellIterator = row.cellIterator(); // iterating over each column
				String toWrite = "";
				boolean isEmptyAnswer = false;
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					// Name: 3, Problem: 7; DateTime: 15, Submission: 16, Correct: 26
					//Needs the columns to be set up in this standard way, otherwise change number to match column names
					int currColumn = cell.getColumnIndex();
					if (currColumn == 0 || currColumn == 7 || currColumn == 15 || currColumn == 16 || currColumn == 26) {
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

							} else {
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

	/*
	 * This method parses the JSON submitted from Prairie learn submission into the desired format 
	 * Currently uses the TAG indicator, maybe switch to UID to allow distracting values?
	 * 
	 * @paramstringCellValue - the JSON string inside the specific cell of the excel file
	 * @paramchosenLabelForQuestions - the label for all the blocks in prairie learn, needs to be the same in each question
	 * 
	 * Issue to address: tags for any distracting blocks are null, which a) are not
	 * distinct and cannot be identified through ordering, and b) cannot be grabbed
	 * using JSONgetString below
	 * 
	 */
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