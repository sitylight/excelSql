// Copyright (c) 1998-2019 Core Solutions Limited. All rights reserved.
// ============================================================================
// CURRENT VERSION CNT.5.0.1
// ============================================================================
// CHANGE LOG
// CNT.5.0.1 : 2019-02-25, derrick.liang, creation
// ============================================================================
package com.excel;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author derrick.liang
 */
public class ConvertToXml {
    public static void main(String[] args) throws IOException {
        convert("C:/cbxsoftware/personal-projects/excelSql/src/main/resources/excel/user_entity.xlsx");
    }

    private static void convert(String excelPath) throws IOException {
        final OutputFormat format = OutputFormat.createPrettyPrint();
        final XMLWriter output = new XMLWriter(
                new FileWriter("C:/cbxsoftware/personal-projects/excelSql/src/main/resources/xml/user.xml"), format);
        File file = new File(excelPath);
        String fileName = file.getName();
        final String prefix = fileName.substring(fileName.lastIndexOf("."));
        int num = prefix.length();
        final String fileOtherName = fileName.substring(0, fileName.length() - num);
        Workbook workbook = WorkbookFactory.create(file);
        Document document = DocumentHelper.createDocument();
        DataFormatter dataFormatter = new DataFormatter();
        Element root = document.getRootElement();
        if (root == null) {
            root = document.addElement(fileOtherName);
            root.addAttribute("position", fileName);
        }
        for (Sheet sheet : workbook) {
            Element firstElement = root.addElement("Sheet");
            firstElement.addAttribute("id", sheet.getSheetName());
            Element secondElement = null;
            boolean entityStart = false;
            boolean isEntityProperties = false;
            boolean isEntityPValue = false;
            boolean isFieldLabel = false;
            boolean isFieldValue = false;
            List<String> entityInfo = new ArrayList<>();
            Element thirdElement = null;
            Element forthElement = null;
            for (Row row : sheet) {
                boolean isFirstElement = true;
                int columnNum = row.getPhysicalNumberOfCells();
                for (Cell cell : row) {
                    int cellIndex = cell.getColumnIndex();
                    String cellStr = dataFormatter.formatCellValue(cell);
                    if (cellStr.startsWith("##Entity")) {
                        secondElement = firstElement.addElement("Entity");
                        entityStart = true;
                    } else if (cellStr.startsWith("#begin")) {
                        if (entityStart) {
                            isEntityProperties = true;
                            entityStart = false;
                        } else {
                            String id = StringUtils.substringAfter(cellStr, ":");
                            thirdElement = secondElement.addElement("elements");
                            thirdElement.addAttribute("id", id);
                            isFieldLabel = true;
                        }
                    } else if (cellStr.startsWith("#end")) {
                        isEntityProperties = false;
                        isEntityPValue = false;
                        isFieldValue = false;
                        entityInfo.clear();
                    } else if (isEntityProperties || isFieldLabel) {
                        entityInfo.add(cellStr);
                        if (entityInfo.size() == columnNum) {
                            if (isEntityProperties) {
                                isEntityPValue = true;
                                isEntityProperties = false;
                            } else {
                                isFieldValue = true;
                                isFieldLabel = false;
                            }
                        }
                    } else if (isEntityPValue) {
                        secondElement.addAttribute(entityInfo.get(cellIndex), cellStr);
                    } else if (isFieldValue) {
                        if (isFirstElement) {
                            forthElement = thirdElement.addElement("element");
                            Element fifthElement = forthElement.addElement(entityInfo.get(cellIndex));
                            fifthElement.setText(cellStr);
                            isFirstElement = false;
                        } else {
                            Element fifthElement = forthElement.addElement(entityInfo.get(cellIndex));
                            fifthElement.setText(cellStr);
                        }
                    }
                }
            }
        }
        output.write(document);
        output.flush();
        output.close();
    }
}
