/**
 *
 * @author Daniel EAgy
 * @version 1.0
 * 
 * This program takes a spreadsheet provided by the client SAS 
 * and updates and formats it to match TriStar Software's Winserve
 * application for batch imports. 
 */

import java.io.*;
import org.apache.poi.ss.usermodel.*;
import java.nio.file.Files;

import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import org.controlsfx.control.spreadsheet.*;
import javafx.scene.layout.*;
import javafx.stage.*;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class SASFormatter extends Application {
    Formatter format = null;
    File inFile = null;
    File outFile = null;
    

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        launch(args);
    }

    @Override
    public void start(Stage stage) throws Exception {
        
        GridBase grid = new GridBase (15, 10);
        Button open = new Button();
        open.setText("Open");
        
        Button saveAs = new Button();
        saveAs.setText("Save As");
        
        Button save = new Button();
        save.setText("Save");
        
        TextField inPane = new TextField();
        inPane.setText("     ");
        inPane.setDisable(true);
        
        TextField outPane = new TextField();
        outPane.setText("     ");
        outPane.setDisable(true);
        
        open.setOnAction(new EventHandler<ActionEvent>() {
            
            @Override
            public void handle(ActionEvent event) {
                FileChooser fileChooser = new FileChooser();
                
                FileChooser.ExtensionFilter xlsFilter = new FileChooser.ExtensionFilter("Excel Workbook 1997-2003 (*.xls)", "*.xls");
                FileChooser.ExtensionFilter xlsxFilter = new FileChooser.ExtensionFilter("Excel Workbook (*.xlsx)", "*.xlsx");
                
                fileChooser.getExtensionFilters().add(xlsxFilter);
                fileChooser.getExtensionFilters().add(xlsFilter);
                
               
                
                inFile = fileChooser.showOpenDialog(stage);
                inPane.setText(inFile.toString());
            }
        });
        
        saveAs.setOnAction(new EventHandler<ActionEvent>() {
            
            @Override
            public void handle(ActionEvent event) {
                FileChooser fileChooser = new FileChooser();
                if(inFile != null) {
                    fileChooser.setInitialFileName(inFile.toString());
                }
                
                FileChooser.ExtensionFilter xlsFilter = new FileChooser.ExtensionFilter("Excel Workbook 1997-2003 (*.xls)", "*.xls");
                FileChooser.ExtensionFilter xlsxFilter = new FileChooser.ExtensionFilter("Excel Workbook (*.xlsx)", "*.xlsx");
                
                fileChooser.getExtensionFilters().add(xlsxFilter);
                fileChooser.getExtensionFilters().add(xlsFilter);
                
                
                
                outFile = fileChooser.showSaveDialog(stage);
                outPane.setText(outFile.toString());

            }
        });
        
        save.setOnAction(new EventHandler<ActionEvent>() {
            
            @Override
            public void handle(ActionEvent event) {
                format = new Formatter(inFile);
                format.format();
                
                format.save(outFile);
                
                inPane.setText(" ");
                outPane.setText(" ");
                
                inFile = null;
                outFile = null;
            }
        });
        
        SpreadsheetView ss = new SpreadsheetView();
        GridPane root = new GridPane();
        root.setPadding(new Insets(10, 10, 10, 10));
        root.add(inPane, 0, 1);
        root.add(open, 1, 1);
        root.add(outPane, 0, 2);
        root.add(saveAs, 1, 2);
        root.add(save, 2, 3);
        Scene scene = new Scene(root);
        
        stage.setTitle("SAS Spreadsheet Formatter");
        stage.setScene(scene);
        stage.show();
    }
    
}
