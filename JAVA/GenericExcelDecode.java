package Java;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.StringWriter;
import java.io.PrintWriter;

import java.text.DateFormat;
import java.text.SimpleDateFormat;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * GenericExcelDecode is the implementation class to read Excel file, do basic
 * format level validations and convert the data to xml string, which can be
 * used in SOA.
 *
 * @author Deepika Teegapuram
 * @version 1.0
 * @since 1.0
 */
public class GenericExcelDecode implements GenericExcelDecodeIntf {

    /* Global Variable Declaration */
    public String decodeExcel(byte[] excelData, String readableSheetName, int linesStartingRowNum,
                              int headersStartingRowNum, String headerColumnNames, String lineColumnNames) {
        /* Local Variable Declaration */
        Workbook wb;
        String xml = "";
        String endXML = "" ;
        HashMap<Integer, String> columnElements = new HashMap<Integer, String>();
        List<String> lineColumnElementsList =
            (lineColumnNames.length() > 0) ?
            new ArrayList<String>(Arrays.asList(lineColumnNames.split("\\s*\\|\\s*"))) : null;
        List<String> headerColumnElementsList =
            (headerColumnNames.length() > 0) ?
            new ArrayList<String>(Arrays.asList(headerColumnNames.split("\\s*\\|\\s*"))) : null;
        int i = 0 ;
        int row = 0 ;
        int col = 0 ;
        try {
            // Read Excel Data in the form of Byte Array
            wb = WorkbookFactory.create(new ByteArrayInputStream(excelData));
            // Create xml start tag with namespace
            xml += "<Excel xmlns=\"http://xmlns.oracle.com/pcbpel/adapter/file/ExcelContent\">";
            endXML = "</Excel>" ;
            int columnNum;
            // Iterate through sheets
            for (i = 0; i < wb.getNumberOfSheets(); i++) {
                Sheet sheet = wb.getSheetAt(i);
                // if the sheet name is not equal to readableSheetName input parameter then create xml End tag
                // and continue to next sheet.
                if (sheet.getSheetName().trim().equalsIgnoreCase(readableSheetName.trim())) {
                    // create current sheet xml tag
                    xml += "<Sheet name=\"" + sheet.getSheetName() + "\" num=\"" + i + "\">";
                    endXML = "</Sheet></Excel>" ;

                    // if the sheet name is equal to readableSheetName input parameter, Iterate through rows of
                    // the sheet
                    
                    for (row = 0; row <= sheet.getLastRowNum(); row++) {
                        // declare a local string variable which will be used to
                        // create xml rows in output xml.
                        String xmlrow = "";

                        // iterate through all the cells in the sheet.
                        if (sheet.getRow(row) != null) {
                            int totalCellCount = sheet.getRow(row).getLastCellNum();
                            // read all cells in each row per iteration.
                            for (col = 0; col < sheet.getRow(row).getLastCellNum(); col++) {
                                // create temp variable cell and add the value of 0th cell of the row
                                Cell cell = sheet.getRow(row).getCell(col);
                                // increment the cell number
                                columnNum = col + 1;
                                String columnElement = "";
                                String columnValue = "";
                                // check the cell type and assign the value of the cell to columnValue var
                                if (cell != null) {
                                    if ((totalCellCount == 2) && (row >= headersStartingRowNum) && (col == 0) &&
                                        (!headerColumnElementsList.isEmpty()) &&
                                        (cell.getCellType() == Cell.CELL_TYPE_STRING)) {
                                        columnValue =
                                            cell.getStringCellValue().trim().replaceAll("[\\s+ #+ &+ .+ ?+ (+ ,+ )+ _+ //+]", "");
                                        
                                        for (String hColumnName : headerColumnElementsList) {
                                            hColumnName =
                                                hColumnName.trim().replaceAll("[\\s+ #+ &+ .+ ?+ (+ ,+ )+ _+ //+]", "");
                                            if (hColumnName.trim().equalsIgnoreCase(columnValue)) {
                                                columnElement = columnValue;
                                                break;
                                            }
                                        }
                                        col++;
                                        cell = sheet.getRow(row).getCell(col);
                                    }
                                    if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                                        columnValue = cell.getStringCellValue();
                                    } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC ||
                                               cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                                        if (DateUtil.isCellDateFormatted(cell)) {
                                            double dateValue = cell.getNumericCellValue();
                                            if (DateUtil.isValidExcelDate(dateValue)) {
                                                Date date = DateUtil.getJavaDate(dateValue);
                                                DateFormat df = new SimpleDateFormat("dd-MMM-yyyy");
                                                String dateString = df.format(date);
                                                columnValue = dateString;
                                            }
                                        } else {
                                            Double numValue = cell.getNumericCellValue();
                                            int numberValue = numValue.intValue();
                                            if ((numValue - numberValue) != 0) {
                                                columnValue = "" + numValue;
                                            } else {
                                                columnValue = "" + numberValue;
                                            }
                                        }
                                    } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
                                        columnValue = "" + cell.getBooleanCellValue();
                                    } else if (cell.getCellType() == Cell.CELL_TYPE_ERROR) {
                                        columnValue = "";
                                    } else if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
                                        columnValue = "";
                                    }
                                }
                                if (!lineColumnElementsList.isEmpty() && row == linesStartingRowNum) {
                                    columnValue = columnValue.trim().replaceAll("[\\s+ #+ &+ .+ ?+ (+ ,+ )+ _+ //+]", "");
                                    for (String lColumnName : lineColumnElementsList) {
                                        //String lColumnNameBeforeReplace =lColumnName;
                                        lColumnName =
                                            lColumnName.trim().replaceAll("[\\s+ #+ &+ .+ ?+ (+ ,+ )+ _+ //+]", "");
                                        if (lColumnName.trim().equalsIgnoreCase(columnValue)) {
                                            columnElement = columnValue;
                                            columnElements.put(columnNum, columnElement);
                                            break;
                                        }
                                    
                                    }
                                } else {
                                    if (totalCellCount > 2) {
                                        columnElement = columnElements.get(columnNum);
                                        // incase DVM values List is empty
                                        if (columnElement == null || columnElement == "") {
                                            columnElement = "COLUMN_" + columnNum;
                                            columnElements.put(columnNum, columnElement);
                                        }
                                    }

                                    // >,<,& signs in the values should not conflict with xml structure
                                    columnValue =
                                        columnValue.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">",
                                                                                                                "&gt;");
                                    // create xml element as header name and column value as xml value
                                    if (columnValue != null && cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK &&
                                        columnValue.length() > 0 && columnElement != "") {
                                        xmlrow += "<" + columnElement + ">";
                                        xmlrow += columnValue;
                                        xmlrow += "</" + columnElement + ">";
                                    }
                                }
                            }
                            
                            //Header starting tag
                            if (row == headersStartingRowNum && totalCellCount == 2) {
                                xml += "<Header num=\"" + row + "\">";
                            }
                            if (xmlrow != null && xmlrow.length() > 0) {
                               
                                if (totalCellCount == 2) {
                                    //  xml+="<Header num=\"" + row + "\">";
                                    xml += xmlrow;
                                  
                                } else {
                                    xml += "<Row num=\"" + row + "\">";
                                    xml += xmlrow;
                                    xml += "</Row>";
                                }
                            }
                            // Header ending tag
                            if (row == headerColumnElementsList.size() && totalCellCount == 2) {
                                xml += "</Header>";
                            }
                        }

                    }
                    xml += "</Sheet>";
                }

            }
            xml += "</Excel>";
        }
        catch (Exception e) {
            StringWriter sw = new StringWriter() ;
            PrintWriter pw = new PrintWriter(sw) ;
            e.printStackTrace(pw) ;
            xml += "<Failure><Exception>" + sw.toString() + "</Exception>";
            xml += "<curRow>" + row + "</curRow>" ;
            xml += "<curCol>" + col + "</curCol>" ;
            xml += "</Failure>" ;
            xml += endXML;
        }
        return xml;
    }
}

