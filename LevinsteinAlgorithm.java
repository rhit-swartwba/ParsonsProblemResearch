
import java.util.Arrays;
 
public class LevinsteinAlgorithm {
    // Method to calculate Levenshtein distance using two matrix rows
    public static int levenshteinTwoMatrixRows(String str1, String str2) {
        int m = str1.length();
        int n = str2.length();
 
        // Initializing two arrays to store the current and previous row values
        int[] prevRow = new int[n + 1];
        int[] currRow = new int[n + 1];
 
        // Initializing the first row with increasing integers
        for (int j = 0; j <= n; j++) {
            prevRow[j] = j;
        }
 
        // Looping through each character of str1
        for (int i = 1; i <= m; i++) {
            // Initializing the first element of the current row with the row number
            currRow[0] = i;
 
            // Looping through each character of str2
            for (int j = 1; j <= n; j++) {
                // If characters are equal, no operation needed, take the diagonal value
                if (str1.charAt(i - 1) == str2.charAt(j - 1)) {
                    currRow[j] = prevRow[j - 1];
                } else {
                    // If characters are not equal, find the minimum value of insert, delete, or replace
                    currRow[j] = 1 + Math.min(currRow[j - 1], Math.min(prevRow[j], prevRow[j - 1]));
                }
            }
 
            // Update prevRow with currRow values
            prevRow = Arrays.copyOf(currRow, currRow.length);
        }
        // Return the final Levenshtein distance stored at the bottom-right corner of the matrix
        return currRow[n];
    }
    
    public static String parseAnswer(String ans, int correctLength) {
    	char[] parseAnswer = new char[correctLength];
    	int k = 0;
    	for(int i = 0; i < ans.length(); i++) {
    		if(ans.charAt(i) == ' ' || ans.charAt(i) == ',') {
    			continue;
    		}
    		parseAnswer[k] = ans.charAt(i);
    		k++;
    	}
		return new String(parseAnswer);
    	
    }
 
    // Main method for testing
    public static void main(String[] args) {
    	//manually specify the correct answer
        String correctAnswer = "0123345678";
        String ans = "0 3 1 2 3 4 5 6 7 8";
        String parsedAnswer = parseAnswer(ans, correctAnswer.length());
 
        // Function Call
        int distance = levenshteinTwoMatrixRows(correctAnswer, parsedAnswer);
        System.out.println("Levenshtein Distance: " + distance);
    }
}