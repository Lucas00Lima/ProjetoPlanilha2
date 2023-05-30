package org.example;

import org.apache.poi.ss.usermodel.*;
import javax.swing.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        JFileChooser fileChooser = new JFileChooser();
        int result = fileChooser.showOpenDialog(null);
        String filePath = null;
        if (result == JFileChooser.APPROVE_OPTION) {
            filePath = fileChooser.getSelectedFile().getAbsolutePath();
            JOptionPane.showMessageDialog(null, "Arquivo selecionado: " + filePath);
        } else if (result == JFileChooser.CANCEL_OPTION) {
            JOptionPane.showMessageDialog(null, "Seleção de arquivo cancelada.");
        } else if (result == JFileChooser.ERROR_OPTION) {
            JOptionPane.showMessageDialog(null, "Erro ao selecionar arquivo.");
        }
        String username = JOptionPane.showInputDialog("Insira o usuario do banco");
        String password = JOptionPane.showInputDialog("Insira a senha do banco");
        String table = JOptionPane.showInputDialog("Qual tabela deseja fazer a transferencia ?");
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

                String priceValue = dataFormatter.formatCellValue(priceCell);
                if (priceValue.contains(",")) {
                    int decimalIndex = priceValue.indexOf(",");
                    int decimalPlaces = priceValue.length() - decimalIndex - 1;
                    if (decimalPlaces == 1) {
                        priceValue += "0";
                    }
                }
                String priceValues = priceValue.replaceAll("," , "");

                String minimumStockValue = dataFormatter.formatCellValue(minimumStockCell);
                PreparedStatement preparedStatement = connection.prepareStatement(insertQuery.toString());
                preparedStatement.setString(1, internalCodeValue);
                preparedStatement.setString(2, nameValue);
                preparedStatement.setString(3, description);
                preparedStatement.setString(4, barcodeValue);
                preparedStatement.setString(5, priceValues);
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
                preparedStatement.addBatch("UPDATE " + table + " SET type = 1");
                preparedStatement.addBatch("UPDATE " + table + " SET department_id = 1");
                preparedStatement.addBatch("UPDATE " + table + " SET measure_unit = 'u'");
                preparedStatement.addBatch("UPDATE " + table + " SET production_group = 1");
                preparedStatement.addBatch("UPDATE " + table + " SET panel = 1");
                preparedStatement.addBatch("UPDATE " + table + " SET active = 1");
                preparedStatement.addBatch("UPDATE " + table + " SET hall_table = 1");
                preparedStatement.addBatch("UPDATE " + table + " SET category_id = 1");
                preparedStatement.executeBatch();
                totalLinhasInseridas++;
                preparedStatement.close();
            }
            connection.close();
            System.out.println("Row affected = " + totalLinhasInseridas);
        } catch (SQLException | IOException e) {
            throw new RuntimeException(e);
        }
    }
    }