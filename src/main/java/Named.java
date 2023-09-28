import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Named extends Application {

    private Map<String, String> namesMap = new HashMap<>();
    private Label nameLabel;

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Random Name Picker");

        nameLabel = new Label();
        nameLabel.getStyleClass().add("name-label");

        Button pickButton = new Button("随机点名");
        pickButton.setOnAction(e -> pickRandomName());

        VBox layout = new VBox(10);
        layout.setPadding(new Insets(10));
        layout.setAlignment(Pos.CENTER);
        layout.getChildren().addAll(nameLabel, pickButton);
        layout.getStyleClass().add("main-container");

        Scene scene = new Scene(layout, 300, 200);
        scene.getStylesheets().add(getClass().getResource("styles.css").toExternalForm());
        primaryStage.setScene(scene);
        primaryStage.show();

        readExcel("F:\\code\\Named_test1\\src\\main\\resources\\21大数据点名册.xlsx");
    }

    private void readExcel(String filePath) {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0); // 假设数据位于第一个工作表

            // 从第二行开始读取数据，第一行是标题行
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                Cell indexCell = row.getCell(0); // 序号所在的单元格
                Cell nameCell = row.getCell(1); // 姓名所在的单元格

                String index;
                String name;

                if (indexCell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                    index = String.valueOf((int) indexCell.getNumericCellValue());
                } else {
                    index = indexCell.getStringCellValue();
                }

                if (nameCell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                    name = String.valueOf((int) nameCell.getNumericCellValue());
                } else {
                    name = nameCell.getStringCellValue();
                }

                namesMap.put(index, name);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void pickRandomName() {
        if (namesMap.isEmpty()) {
            nameLabel.setText("没有可用的姓名");
            return;
        }

        int randomIndex = (int) (Math.random() * namesMap.size());
        String randomName = (String) namesMap.values().toArray()[randomIndex];
        nameLabel.setText(randomName);
    }
}