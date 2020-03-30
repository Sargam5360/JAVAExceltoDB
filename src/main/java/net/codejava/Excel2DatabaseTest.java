package net.codejava;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Sample Java program that imports data from an Excel file to MySQL database.
 * 
 * @author Nam Ha Minh - https://www.codejava.net
 * 
 */
public class Excel2DatabaseTest {

	public static void main(String[] args) {
		String jdbcURL = "jdbc:mysql://localhost/mydb?useSSL=false";
		String username = "root";
		String password = "root";

		String excelFilePath = "Students.xlsx";

		int batchSize = 10;

		Connection connection = null;

		try {
			long start = System.currentTimeMillis();
			
			FileInputStream inputStream = new FileInputStream(excelFilePath);

			Workbook workbook = new XSSFWorkbook(inputStream);

			Sheet firstSheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = firstSheet.iterator();
			int n = firstSheet.getLastRowNum();
			System.out.println(n + "hahah");
            connection = DriverManager.getConnection(jdbcURL, username, password);
            connection.setAutoCommit(false);
 
            String sql = "INSERT INTO students (name, enrolled, progress) VALUES (?, ?, ?)";
            PreparedStatement statement = connection.prepareStatement(sql);		
			
            int count = 1;
            boolean flag = false;
            rowIterator.next(); // skip the header row
            
			while (rowIterator.hasNext() && rowIterator != null) {
				count++;
				Row nextRow = rowIterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();
				String name = null;
				Date enrollDate = null;
				int progress = 0;
				while (cellIterator.hasNext()) {
					Cell nextCell = cellIterator.next();

					int columnIndex = nextCell.getColumnIndex();

					switch (columnIndex) {
					case 0:
						name = nextCell.getStringCellValue();
						statement.setString(1, name);
						System.out.println("name" + name);
						break;
					case 1:
						enrollDate = nextCell.getDateCellValue();
						if(enrollDate == null) break;
						statement.setTimestamp(2, new Timestamp(enrollDate.getTime()));
					case 2:
						progress = (int) nextCell.getNumericCellValue();
						statement.setInt(3, progress);
					}

				}
				
                statement.addBatch();
                System.out.println(count);
                if (count % batchSize == 0) {
                    statement.executeBatch();
                }
                if(count==22) break;

			}

			workbook.close();
			
            // execute the remaining queries
            statement.executeBatch();
 
            connection.commit();
            connection.close();	
            
            long end = System.currentTimeMillis();
            System.out.printf("Import done in %d ms\n", (end - start));
            
		} catch (IOException ex1) {
			System.out.println("Error reading file");
			ex1.printStackTrace();
		} catch (SQLException ex2) {
			System.out.println("Database error");
			ex2.printStackTrace();
		}

	}

}
