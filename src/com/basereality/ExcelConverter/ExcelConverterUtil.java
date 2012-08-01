package com.basereality.ExcelConverter;

import java.util.ArrayList;

public class ExcelConverterUtil {

	/**
	 * get the Number of Times a Character Occurs In a String
	 * @param needle
	 * @param hayStack
	 * @return
	 */
	public static int getNumberTimesCharacterOccursInString(char needle, String hayStack){

		int numberTimes = 0;

		char[] charArray = hayStack.toCharArray();

		for(char charFromString : charArray){
			if(charFromString == needle){
				numberTimes++;
			}
		}

		return	numberTimes;
	}

	/**
	 * Parse a line of CSV data and return it as a list of strings
	 * @param data The CSV data to parse.
	 * @param delimiter The delimiter to use
	 * @return  The individual fields.
	 */
	public static ArrayList parseCSV(String data, String delimiter){

		String enclosure = "\"";
		String newline = "\n";

		int pos = -1;
		int last_pos = -1;
		int end = data.length();

		int row = 0;
		boolean quote_open = false;
		boolean trim_quote = false;

		ArrayList returnCells = new ArrayList();

		// Create a continuous loop
		for (int i = -1;; ++i){
			++pos;
			// Get the positions
			int comma_pos = data.indexOf(delimiter, pos);
			int quote_pos = data.indexOf(enclosure, pos);
			int newline_pos = data.indexOf(newline, pos);

			// Which one comes first?
			pos = min((comma_pos == -1) ? end : comma_pos, (quote_pos == -1) ? end : quote_pos, (newline_pos == -1) ? end : newline_pos);

			// Cache it
			//String character = (isset($data[$pos])) ? $data[$pos] : null;
			String character = "";

			boolean done = false;

			if(pos >= data.length() - 1){
				done = true;
			}
			else{
				character = data.substring(pos, pos + 1);
			}

			if(pos == end){
				done = true;
			}

			// It it a special character?
			if (done || (character.compareTo(delimiter) == 0) || character.compareTo(newline) == 0){

				// Ignore it as we're still in a quote
				if (quote_open && !done){
					continue;
				}

				int length = pos - ++last_pos;

				//# Is the last thing a newline?
				if( character.compareTo(newline) == 0 ){
					//# Well then get rid of it
					--length;
				}

				// Is the last thing a quote?
				if (trim_quote){
				// Well then get rid of it
					--length;
				}

				String nextCell = "";

				// Get all the contents of this column
				 if(length > 0){
					 nextCell = data.substring(last_pos, last_pos + length).replace(enclosure + enclosure, enclosure);
				 }

				 returnCells.add(nextCell);

				 //$return[$row][] = thing(i);

				// And we're done
				if (done){
					break;
				}

				// Save the last position
				last_pos = pos;

				// Next row?
				if (character.compareTo(newline) == 0){
					++row;
				}

				trim_quote = false;
			}
			// Our quote?
			else if (character.compareTo(enclosure) == 0){

				// Toggle it
				if (quote_open == false){
					// It's an opening quote
					quote_open = true;
					trim_quote = false;

					// Trim this opening quote?
					if (last_pos + 1 == pos){
						++last_pos;
					}
				}
				else {
					// It's a closing quote
					quote_open = false;

					// Trim the last quote?
					trim_quote = true;
				}
			}
		}

		return returnCells;
	}

	/**
	 * Return the minimum of three values.
	 * @param a
	 * @param b
	 * @param c
	 * @return
	 */
	public static int min(int a, int b, int c){
		int low1 = Math.min(a, b);
		int low2 = Math.min(b, c);
		return Math.min(low1, low2);
	}
}