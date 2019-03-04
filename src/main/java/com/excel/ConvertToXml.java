// Copyright (c) 1998-2019 Core Solutions Limited. All rights reserved.
// ============================================================================
// CURRENT VERSION CNT.5.0.1
// ============================================================================
// CHANGE LOG
// CNT.5.0.1 : 2019-02-25, derrick.liang, creation
// ============================================================================
package com.excel;

import com.sun.org.apache.xerces.internal.dom.DeepNodeListImpl;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author derrick.liang
 */
public class ConvertToXml {
    public static void main(final String[] args)
            throws IOException, ParserConfigurationException, TransformerException {
//        convert("C:/cbxsoftware/personal-projects/excelSql/src/main/resources/excel/user_entity.xlsx");
        w3cConvert("/Users/derrick/develop/projects/excelSql/src/main/resources/excel/user_entity.xlsx");
    }

    private static void w3cConvert(final String path)
            throws ParserConfigurationException, IOException, TransformerException {
        final DocumentBuilderFactory dbFactory =
                DocumentBuilderFactory.newInstance();
        final DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
        final Document doc = dBuilder.newDocument();
        final File file = new File(path);
        final String fileName = file.getName();
        final Workbook workbook = WorkbookFactory.create(file);
        final DataFormatter dataFormatter = new DataFormatter();
        final Element root = doc.createElement("entity");
        root.setAttribute("position", fileName);
        doc.appendChild(root);
        for (final Sheet sheet : workbook) {
            boolean isEntityStart = false;
            boolean isEntityArrLabel = false;
            boolean isEntityArrValue = false;
            boolean isFieldLabel = false;
            boolean isFieldValue = false;
            final List<String> labels = new ArrayList<>();
            Element entityElement = null;
            Element elements = null;
            Element element = null;
            final Element sheetElement = doc.createElement("sheet");
            sheetElement.setAttribute("id", sheet.getSheetName());
            root.appendChild(sheetElement);
            for (final Row row : sheet) {
                boolean isFirstElement = true;
                final int cellNum = row.getPhysicalNumberOfCells();
                for (final Cell cell : row) {
                    final String cellStr = dataFormatter.formatCellValue(cell);
                    if (cellStr.startsWith("##Entity")) {
                        entityElement = doc.createElement("Entity");
                        sheetElement.appendChild(entityElement);
                        isEntityStart = true;
                    } else if (cellStr.startsWith("#begin")) {
                        if (isEntityStart) {
                            isEntityArrLabel = true;
                            isEntityStart = false;
                        } else {
                            isFieldLabel = true;
                            elements = doc.createElement("elements");
                            final String id = StringUtils.substringAfter(cellStr, ":");
                            final Attr attr = doc.createAttribute("id");
                            attr.setValue(id);
                            elements.setAttributeNode(attr);
                            entityElement.appendChild(elements);
                        }
                    }else if (cellStr.startsWith("#end")) {
                        labels.clear();
                        isEntityArrValue = false;
                        isFieldValue = false;
                    } else if (isEntityArrLabel || isFieldLabel) {
                        labels.add(cellStr);
                        if (labels.size() == cellNum) {
                            if (isEntityArrLabel) {
                                isEntityArrLabel = false;
                                isEntityArrValue = true;
                            } else {
                                isFieldLabel = false;
                                isFieldValue = true;
                            }
                        }
                    } else if (isEntityArrValue) {
                        entityElement.setAttribute(labels.get(cell.getColumnIndex()), cellStr);
                    } else if (isFieldValue) {
                        if (isFirstElement) {
                            element = doc.createElement("element");
                            elements.appendChild(element);
                            final Element fieldElement = doc.createElement(labels.get(cell.getColumnIndex()));
                            fieldElement.setTextContent(cellStr);
                            element.appendChild(fieldElement);
                            isFirstElement = false;
                        } else {
                            final Element fieldElement = doc.createElement(labels.get(cell.getColumnIndex()));
                            fieldElement.setTextContent(cellStr);
                            element.appendChild(fieldElement);
                        }
                    }
                }
            }
        }
        final TransformerFactory transformerFactory = TransformerFactory.newInstance();
        final Transformer transformer = transformerFactory.newTransformer();
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty(OutputKeys.DOCTYPE_PUBLIC, "yes");
        transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
        final DOMSource source = new DOMSource(doc);
        final StreamResult result = new StreamResult(
                new File("/Users/derrick/develop/projects/excelSql/src/main/resources/xml/user.xml"));
        transformer.transform(source, result);
//        final StreamResult consoleResult = new StreamResult(System.out);
//        transformer.transform(source, consoleResult);
        final DeepNodeListImpl sheets = (DeepNodeListImpl) doc.getElementsByTagName("sheet");
        List<EntityDefinition> entityDefinitions = new ArrayList<>();
        for (int i = 0; i < sheets.getLength(); i++) {
            final Node sheetNode = sheets.item(i);
            final Node attr = sheetNode.getAttributes().getNamedItem("id");
            if ("entityDef".equals(attr.getTextContent())) {
                final NodeList nodeList = sheetNode.getChildNodes();
                for (int m = 0; m < nodeList.getLength(); m++) {
                    EntityDefinition entityDefinition = new EntityDefinition();
                    List<FieldDefinition> fieldDefinitions = new ArrayList<>();
                    NamedNodeMap attrs = nodeList.item(m).getAttributes();
                    for (int a = 0; a < attrs.getLength(); a++) {
                        Node node = attrs.item(a);
                        if ("name".equals(node.getNodeName())) {
                            entityDefinition.setEntityName(node.getTextContent());
                        } else if ("table_name".equals(node.getNodeName())) {
                            entityDefinition.setTableName(node.getTextContent());
                        }
                    }
                    final NodeList entityList = nodeList.item(m).getChildNodes();
                    for (int j = 0; j < entityList.getLength(); j++) {
//                        final String entityName = entityList.item(j).getAttributes().getNamedItem("name")
//                                .getTextContent();
                        final NodeList elements = entityList.item(j).getChildNodes();
                        for (int e = 0; e < elements.getLength(); e++) {
                            final NodeList elementAttrs = elements.item(e).getChildNodes();
                            final FieldDefinition fieldDefinition = new FieldDefinition();
                            for (int a = 0; a < elementAttrs.getLength(); a++) {
                                String nodeName = elementAttrs.item(a).getNodeName();
                                if ("field_id".equals(nodeName)) {
                                    fieldDefinition.setFieldId(elementAttrs.item(a).getTextContent());
                                } else if ("field_type".equals(nodeName)) {
                                    fieldDefinition.setFieldType(elementAttrs.item(a).getTextContent());
                                }
//                                System.out.print(elementAttrs.item(a).getNodeName());
//                                System.out.print("-----");
//                                System.out.print(elementAttrs.item(a).getTextContent());
//                                System.out.println();
                            }
                            fieldDefinitions.add(fieldDefinition);
                        }
                    }
                    entityDefinition.setFieldDefinitions(fieldDefinitions);
                    entityDefinitions.add(entityDefinition);
                }
            }
        }
        entityDefinitions.forEach(e -> System.out.println(e.toString()));
    }

//    private static void convert(final String excelPath) throws IOException {
//        final OutputFormat format = OutputFormat.createPrettyPrint();
//        final XMLWriter output = new XMLWriter(
//                new FileWriter("C:/cbxsoftware/personal-projects/excelSql/src/main/resources/xml/user.xml"), format);
//        final File file = new File(excelPath);
//        final String fileName = file.getName();
//        final String prefix = fileName.substring(fileName.lastIndexOf("."));
//        final int num = prefix.length();
//        final String fileOtherName = fileName.substring(0, fileName.length() - num);
//        final Workbook workbook = WorkbookFactory.create(file);
//        final Document document = DocumentHelper.createDocument();
//        final DataFormatter dataFormatter = new DataFormatter();
//        Element root = document.getRootElement();
//        if (root == null) {
//            root = document.addElement("entity");
//            root.addAttribute("position", fileName);
//        }
//        for (final Sheet sheet : workbook) {
//            final Element firstElement = root.addElement("Sheet");
//            firstElement.addAttribute("id", sheet.getSheetName());
//            Element secondElement = null;
//            boolean entityStart = false;
//            boolean isEntityProperties = false;
//            boolean isEntityPValue = false;
//            boolean isFieldLabel = false;
//            boolean isFieldValue = false;
//            final List<String> entityInfo = new ArrayList<>();
//            Element thirdElement = null;
//            Element forthElement = null;
//            for (final Row row : sheet) {
//                boolean isFirstElement = true;
//                final int columnNum = row.getPhysicalNumberOfCells();
//                for (final Cell cell : row) {
//                    final int cellIndex = cell.getColumnIndex();
//                    final String cellStr = dataFormatter.formatCellValue(cell);
//                    if (cellStr.startsWith("##Entity")) {
//                        secondElement = firstElement.addElement("Entity");
//                        entityStart = true;
//                    } else if (cellStr.startsWith("#begin")) {
//                        if (entityStart) {
//                            isEntityProperties = true;
//                            entityStart = false;
//                        } else {
//                            final String id = StringUtils.substringAfter(cellStr, ":");
//                            thirdElement = secondElement.addElement("elements");
//                            thirdElement.addAttribute("id", id);
//                            isFieldLabel = true;
//                        }
//                    } else if (cellStr.startsWith("#end")) {
//                        isEntityProperties = false;
//                        isEntityPValue = false;
//                        isFieldValue = false;
//                        entityInfo.clear();
//                    } else if (isEntityProperties || isFieldLabel) {
//                        entityInfo.add(cellStr);
//                        if (entityInfo.size() == columnNum) {
//                            if (isEntityProperties) {
//                                isEntityPValue = true;
//                                isEntityProperties = false;
//                            } else {
//                                isFieldValue = true;
//                                isFieldLabel = false;
//                            }
//                        }
//                    } else if (isEntityPValue) {
//                        secondElement.addAttribute(entityInfo.get(cellIndex), cellStr);
//                    } else if (isFieldValue) {
//                        if (isFirstElement) {
//                            forthElement = thirdElement.addElement("element");
//                            final Element fifthElement = forthElement.addElement(entityInfo.get(cellIndex));
//                            fifthElement.setText(cellStr);
//                            isFirstElement = false;
//                        } else {
//                            final Element fifthElement = forthElement.addElement(entityInfo.get(cellIndex));
//                            fifthElement.setText(cellStr);
//                        }
//                    }
//                }
//            }
//        }
//        output.write(document);
//        output.flush();
//        output.close();
//        readXml(document);
//    }

}
