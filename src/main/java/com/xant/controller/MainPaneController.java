package com.xant.controller;

import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.input.MouseButton;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.StackPane;

public class MainPaneController {

    @FXML
    private StackPane contextArea;

    @FXML
    void routeBillAge(MouseEvent event) {
        if (MouseButton.PRIMARY != event.getButton()) {
            return;
        }
        try {
            Parent fxmlLoader = FXMLLoader.load(getClass().getResource("/template/BillAgePane.fxml"));
            contextArea.getChildren().removeAll();
            contextArea.getChildren().setAll(fxmlLoader);
        } catch (Exception e) {
            throw new RuntimeException("面板加载错误", e);
        }
    }

    @FXML
    void routeExcel2Word(MouseEvent event) {
        try {
            Parent fxmlLoader = FXMLLoader.load(getClass().getResource("/template/Excel2WordPane.fxml"));
            contextArea.getChildren().removeAll();
            contextArea.getChildren().setAll(fxmlLoader);
        } catch (Exception e) {
            throw new RuntimeException("面板加载错误", e);
        }
    }


}
