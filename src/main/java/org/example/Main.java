package org.example;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;
public class Main {
    public static void main(String[] args) {
        String filePath = "C:\\Users\\lukin\\OneDrive\\Área de Trabalho\\Nova pasta\\teste.xlsx";
        String username = "root";
        String password = "@soma+";
        String table = "product";
        String url = "jdbc:mysql://localhost:3306/db000";
        String defaultValue = "";
        try (Connection connection = DriverManager.getConnection(url, username, password)) {
            FileInputStream fileInputStream = new FileInputStream(filePath); //Cria um arquivo com o caminho passado
            Workbook workbook = WorkbookFactory.create(fileInputStream); //
            Sheet sheet = workbook.getSheetAt(0); //Instancia a planilha do arquivo pegando a de indice 0 ou seja a primeira
            DataFormatter dataFormatter =  new DataFormatter();
            StringBuilder insertQuery = new StringBuilder("INSERT INTO " + table + " (internal_code, name, description, barcode, price, minimum_stock");
            StringBuilder valuePlaceholders = new StringBuilder(" VALUES (?,?,?,?,?,?");
            List<String> defaultValues = new ArrayList<>();
            DatabaseMetaData metaData = connection.getMetaData();
            ResultSet resultSet = metaData.getColumns(null, null, table, null);
            int totalColumnsInDatabase = 7;

            while (resultSet.next()) {
                String columnName = resultSet.getString("COLUMN_NAME");
                if (!columnName.equals("internal_code") && !columnName.equals("name") && !columnName.equals("description") && !columnName.equals("barcode") && !columnName.equals("price") && !columnName.equals("minimum_stock")) {
                    if (!columnName.equals("validity") && !columnName.equals("deleted_at") && !columnName.equals("delivery") && !columnName.equals("card") && !columnName.equals("balcony") && !columnName.equals("parameters")) {
                        if (totalColumnsInDatabase > 0) {
                            insertQuery.append(",");
                            valuePlaceholders.append(",");
                        }
                        insertQuery.append(columnName);
                        valuePlaceholders.append("?");
                        defaultValues.add(defaultValue);
                        totalColumnsInDatabase++;
                    }
                }
            }
            resultSet.close();
            insertQuery.append(")");
            valuePlaceholders.append(")");
            insertQuery.append(valuePlaceholders);

            int rowIndex;
            int totalLinhasInseridas = 0;
            for (rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell internalCodeCell = row.getCell(0);
                Cell nameCell = row.getCell(1);
                Cell descriptionCell = row.getCell(2);
                Cell barcodeCell = row.getCell(3);
                Cell priceCell = row.getCell(4);
                Cell minimumStockCell = row.getCell(5);
                Cell inativeCell = row.getCell(6);

                String inative = dataFormatter.formatCellValue(inativeCell);
                if  (inative.equals("S")) {
                    continue;
                }
                String internalCodeValue = dataFormatter.formatCellValue(internalCodeCell);
                String nameValue = dataFormatter.formatCellValue(nameCell);
                String description = dataFormatter.formatCellValue(descriptionCell);
                String barcodeValue = dataFormatter.formatCellValue(barcodeCell);
                String priceValue = dataFormatter.formatCellValue(priceCell).replace(",","");
                String minimumStockValue = dataFormatter.formatCellValue(minimumStockCell);
                PreparedStatement preparedStatement = connection.prepareStatement(insertQuery.toString());
                preparedStatement.setString(1, internalCodeValue);
                preparedStatement.setString(2, nameValue);
                preparedStatement.setString(3, description);
                preparedStatement.setString(4, barcodeValue);
                preparedStatement.setString(5, priceValue);
                preparedStatement.setString(6, minimumStockValue);
                for (int j = 0; j < defaultValues.size(); j++) {
                    String value = defaultValues.get(j);
                    if (value.isEmpty()) {
                        preparedStatement.setInt(j + 7, 0);
                    } else {
                        preparedStatement.setString(j + 7, value);
                    }
                }
                preparedStatement.execute();
                preparedStatement.close();
            }
        } catch (SQLException | IOException e) {
            throw new RuntimeException(e);
        }
    }
    }