package com.example.customermanagement.helper;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import com.example.customermanagement.model.Customer;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

public class ExcelHelper {

    public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    static String[] HEADERs = { "id", "name", "dateOfBirth", "nicNumber","phoneNumber",
            "addressLine1","addressLine2","city","country"};
    static String SHEET = "Customer";

    public static boolean hasExcelFormat(MultipartFile file) {

        if (!TYPE.equals(file.getContentType())) {
            return false;
        }

        return true;
    }

    public static List<Customer> excelToCustomers(InputStream is) {
        try {
            Workbook workbook = new XSSFWorkbook(is);

            Sheet sheet = workbook.getSheet(SHEET);
            Iterator<Row> rows = sheet.iterator();

            List<Customer> customers = new ArrayList<Customer>();

            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();

                // skip header
                if (rowNumber == 0) {
                    rowNumber++;
                    continue;
                }

                Iterator<Cell> cellsInRow = currentRow.iterator();

                Customer customer = new Customer();

                int cellIdx = 0;
                while (cellsInRow.hasNext()) {
                    Cell currentCell = cellsInRow.next();

                    switch (cellIdx) {
                        case 0:
                            customer.setId((long) currentCell.getNumericCellValue());
                            break;

                        case 1:
                            customer.setName(currentCell.getStringCellValue());
                            break;

                        case 2:
                            customer.setDateOfBirth(currentCell.getStringCellValue());
                            break;

                        case 3:
                            customer.setNicNumber(currentCell.getStringCellValue());
                            break;

                        case 4:
                            customer.setPhoneNumber(currentCell.getStringCellValue());
                            break;

                        case 5:
                            customer.setAddressLine1(currentCell.getStringCellValue());
                            break;

                        case 6:
                            customer.setAddressLine2(currentCell.getStringCellValue());
                            break;

                        case 7:
                            customer.setCity(currentCell.getStringCellValue());
                            break;

                        case 8:
                            customer.setCountry(currentCell.getStringCellValue());
                            break;

                        default:
                            break;
                    }

                    cellIdx++;
                }

                customers.add(customer);
            }

            workbook.close();

            return customers;
        } catch (IOException e) {
            throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
        }
    }
}
