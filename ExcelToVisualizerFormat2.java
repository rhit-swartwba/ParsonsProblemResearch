import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

public class ExcelToVisualizerFormat2 {

    public static void main(String[] args) {

        // the file path to the excel file
        String excelFileWithDataFullPath = "C:\\Users\\swartwba\\OneDrive - Rose-Hulman Institute of Technology\\Desktop\\ResearchProject\\Data\\Spring2324Data\\CSSE_132_2324SpringCSSE132_HW3_all_submissions.xlsx";
        // the directory where the output files will be saved
        String outputDirectory = "C:\\Users\\swartwba\\OneDrive - Rose-Hulman Institute of Technology\\Desktop\\ResearchProject\\Data\\Spring2324Data\\FinalExam\\2324SpringHW3";
        // for answers-names "" - put that here (needs to be the same across all questions of the homework assignment)
        String chosenLabelForQuestions = "Answers";

        getExcelDataToVisualizerFormat(excelFileWithDataFullPath, outputDirectory, chosenLabelForQuestions);
    }

    static void getExcelDataToVisualizerFormat(String excelFileWithDataFullPath, String outputDirectory,
                                               String chosenLabelForQuestions) {
        try {
            File file = new File(excelFileWithDataFullPath);
            FileInputStream fis = new FileInputStream(file);

            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);
            Iterator<Row> itr = sheet.iterator();

            itr.next(); // skip header row

            Map<String, FileWriter> fileWriters = new HashMap<>();
            String currentQuestion = "notAQuestion";
            boolean isEmptyAnswer;

            while (itr.hasNext()) {
                Row row = itr.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                String toWrite = "";
                isEmptyAnswer = false;

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    int currColumn = cell.getColumnIndex();
                    if (currColumn == 0 || currColumn == 7 || currColumn == 16 || currColumn == 26) {
                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_STRING:
                                if (currColumn == 16) {
                                    String answer = parseSubmission(cell.getStringCellValue(), chosenLabelForQuestions);
                                    if (answer.equals("")) {
                                        isEmptyAnswer = true;
                                    }
                                    toWrite += answer + ",";
                                } else if (currColumn == 7) {
                                    String questionName = cell.getStringCellValue();
                                    if (!questionName.equals(currentQuestion)) {
                                        currentQuestion = questionName;
                                        if (!fileWriters.containsKey(currentQuestion)) {
                                            FileWriter writer = new FileWriter(outputDirectory + currentQuestion + ".txt");
                                            fileWriters.put(currentQuestion, writer);
                                        }
                                    }
                                } else {
                                    toWrite += cell.getStringCellValue() + ",";
                                }
                                break;
                            case Cell.CELL_TYPE_NUMERIC:
                                toWrite += cell.getNumericCellValue() + ",";
                                break;
                            case Cell.CELL_TYPE_BOOLEAN:
                                toWrite += cell.getBooleanCellValue() ? "1," : "0,";
                                break;
                            default:
                                break;
                        }
                    }
                }

                if (!isEmptyAnswer) {
                    FileWriter writer = fileWriters.get(currentQuestion);
                    if (writer != null) {
                        writer.write(toWrite + "\n");
                    }
                }
            }

            for (FileWriter writer : fileWriters.values()) {
                writer.close();
            }

            wb.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String parseSubmission(String stringCellValue, String chosenLabelForQuestions) {
        StringBuilder subAnswer = new StringBuilder();
        if (!stringCellValue.contains("[]")) {
            JSONObject obj = new JSONObject(stringCellValue);
            JSONArray arr = obj.getJSONArray(chosenLabelForQuestions);
            for (int i = 0; i < arr.length(); i++) {
                int order = arr.getJSONObject(i).getInt("ranking");
                subAnswer.append(" ").append(order);
            }
        }
        return subAnswer.toString();
    }
}
