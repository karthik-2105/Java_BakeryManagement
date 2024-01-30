import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BakeryManagementSystem {
    private static final String FILE_NAME = "BakeryOrders.xlsx";

    public static void main(String[] args) throws IOException {
        Workbook workbook = loadExistingWorkbookOrCreateNew();
        Scanner scanner = new Scanner(System.in);

        while (true) {
            System.out.println("Bakery Management System");
            System.out.println("1. Add Order");
            System.out.println("2. Get Details");
            System.out.println("3. Complete Order");
            System.out.println("4. Cancel Order");
            System.out.println("5. Exit");
            System.out.print("Enter your choice: ");

            int choice = scanner.nextInt();
            scanner.nextLine(); // Consume the newline character

            switch (choice) {
                case 1:
                    addOrder(workbook, scanner);
                    break;
                case 2:
                    getOrderDetails(workbook, scanner);
                    break;
                case 3:
                    completeOrder(workbook, scanner);
                    break;
                case 4:
                    cancelOrder(workbook, scanner);
                    break;
                case 5:
                    saveWorkbookToFile(workbook);
                    System.out.println("Exiting the Bakery Management System. Goodbye!");
                    System.exit(0);
                default:
                    System.out.println("Invalid choice. Please enter a number between 1 and 5.");
            }
        }
    }

    private static Workbook loadExistingWorkbookOrCreateNew() {
        File file = new File(FILE_NAME);
        Workbook workbook;

        if (file.exists()) {
            try (FileInputStream fis = new FileInputStream(file)) {
                workbook = new XSSFWorkbook(fis);
            } catch (IOException e) {
                e.printStackTrace();
                workbook = new XSSFWorkbook();
            }
        } else {
            workbook = new XSSFWorkbook();
        }

        return workbook;
    }

    private static void saveWorkbookToFile(Workbook workbook) {
        try (FileOutputStream fileOut = new FileOutputStream(FILE_NAME)) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private static void addOrder(Workbook workbook, Scanner scanner) throws IOException {
        Sheet sheet = createExcelFile(workbook);
        System.out.print("Enter Customer Name: ");
        String customerName = scanner.nextLine();

        System.out.print("Enter Order ID: ");
        int orderId = scanner.nextInt();

        // Check if the order ID already exists
        if (isOrderIdExist(sheet, orderId)) {
            System.out.println("Order ID already exists. Please choose a different one.");
            return;
        }

        scanner.nextLine(); // Consume the newline character

        System.out.print("Enter Items: ");
        String items = scanner.nextLine();

        System.out.print("Enter Quantity: ");
        int quantity = scanner.nextInt();

        System.out.print("Enter Price: ");
        double price = scanner.nextDouble();

        Row row = sheet.createRow(sheet.getPhysicalNumberOfRows());
        row.createCell(0).setCellValue(orderId);
        row.createCell(1).setCellValue(customerName);
        row.createCell(2).setCellValue(items);
        row.createCell(3).setCellValue(quantity);
        row.createCell(4).setCellValue(price);
        row.createCell(5).setCellValue("Pending");

        try (FileOutputStream fileOut = new FileOutputStream(FILE_NAME)) {
            workbook.write(fileOut);
        }

        System.out.println("Order added successfully!");
    }

    private static void getOrderDetails(Workbook workbook, Scanner scanner) {
        Sheet sheet = workbook.getSheet("Orders");

        System.out.print("Enter Order ID to get details: ");
        int orderId = getIntInput(scanner, "Enter prompt: ");

        Row headerRow = sheet.getRow(0);
        int orderIdColumnIndex = findColumnIndex(headerRow, "Order ID");

        for (Row row : sheet) {
            Cell cell = row.getCell(orderIdColumnIndex);

            if (cell != null && cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() == orderId) {
                printOrderDetails(headerRow, row);
                return;
            }
        }

        System.out.println("Order with ID " + orderId + " not found.");
    }

    private static void completeOrder(Workbook workbook, Scanner scanner) throws IOException {
        Sheet sheet = workbook.getSheet("Orders");

        System.out.print("Enter Order ID to mark as completed: ");
        int orderId = scanner.nextInt();

        Row headerRow = sheet.getRow(0);
        int orderIdColumnIndex = findColumnIndex(headerRow, "Order ID");
        int statusColumnIndex = findColumnIndex(headerRow, "Status");

        for (Row row : sheet) {
            Cell orderIdCell = row.getCell(orderIdColumnIndex);

            if (orderIdCell != null && orderIdCell.getCellType() == CellType.NUMERIC && orderIdCell.getNumericCellValue() == orderId) {
                Cell statusCell = row.getCell(statusColumnIndex);
                if (statusCell != null) {
                    statusCell.setCellValue("Completed");
                    try (FileOutputStream fileOut = new FileOutputStream(FILE_NAME)) {
                        workbook.write(fileOut);
                    }
                    System.out.println("Order with ID " + orderId + " marked as completed.");
                    return;
                }
            }
        }

        System.out.println("Order with ID " + orderId + " not found.");
    }

    private static void cancelOrder(Workbook workbook, Scanner scanner) throws IOException {
        Sheet sheet = workbook.getSheet("Orders");

        System.out.print("Enter Order ID to cancel: ");
        int orderId = scanner.nextInt();

        Row headerRow = sheet.getRow(0);
        int orderIdColumnIndex = findColumnIndex(headerRow, "Order ID");
        int statusColumnIndex = findColumnIndex(headerRow, "Status");

        for (Row row : sheet) {
            Cell orderIdCell = row.getCell(orderIdColumnIndex);

            if (orderIdCell != null && orderIdCell.getCellType() == CellType.NUMERIC && orderIdCell.getNumericCellValue() == orderId) {
                Cell statusCell = row.getCell(statusColumnIndex);
                if (statusCell != null) {
                    statusCell.setCellValue("Cancelled");
                    try (FileOutputStream fileOut = new FileOutputStream(FILE_NAME)) {
                        workbook.write(fileOut);
                    }
                    System.out.println("Order with ID " + orderId + " cancelled.");
                    return;
                }
            }
        }

        System.out.println("Order with ID " + orderId + " not found.");
    }

    private static void printOrderDetails(Row headerRow, Row orderRow) {
        System.out.println("Order Details:");
        for (Cell cell : orderRow) {
            int columnIndex = cell.getColumnIndex();
            System.out.println(headerRow.getCell(columnIndex).getStringCellValue() + ": " + getCellValueAsString(cell));
        }
    }

    private static boolean isOrderIdExist(Sheet sheet, int orderId) {
        Row headerRow = sheet.getRow(0);
        int orderIdColumnIndex = findColumnIndex(headerRow, "Order ID");

        for (Row row : sheet) {
            Cell cell = row.getCell(orderIdColumnIndex);

            if (cell != null && cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() == orderId) {
                return true;
            }
        }

        return false;
    }

    private static String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case BLANK:
                return "";
            default:
                return "";
        }
    }
    private static int getIntInput(Scanner scanner, String prompt) {
        while (true) {
            try {
                System.out.print(prompt);
                return scanner.nextInt();
            } catch (java.util.InputMismatchException e) {
                // Clear the invalid input from the scanner buffer
                scanner.nextLine();

                // Print an error message
                System.out.println("Invalid input. Please enter a valid integer.");
            }
        }
    }


    private static int findColumnIndex(Row headerRow, String columnName) {
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equals(columnName)) {
                return cell.getColumnIndex();
            }
        }
        return -1; // Return -1 if the column is not found
    }
    private static Sheet createExcelFile(Workbook workbook) {
        Sheet sheet = workbook.getSheet("Orders");

        if (sheet == null) {
            // If the sheet doesn't exist, create a new one
            sheet = workbook.createSheet("Orders");

            Row headerRow = sheet.createRow(0);
            String[] columns = {"Order ID", "Customer Name", "Items", "Quantity", "Price", "Status"};
            for (int i = 0; i < columns.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns[i]);
            }
        }

        return sheet;
    }

}
