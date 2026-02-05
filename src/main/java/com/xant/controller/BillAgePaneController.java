package com.xant.controller;

import io.github.palexdev.materialfx.controls.MFXButton;
import io.github.palexdev.materialfx.controls.MFXProgressSpinner;
import io.github.palexdev.materialfx.controls.MFXRadioButton;
import io.github.palexdev.materialfx.controls.MFXTextField;
import javafx.fxml.FXML;
import javafx.scene.control.TextArea;
import javafx.scene.input.MouseEvent;

public class BillAgePaneController {

    @FXML
    private MFXButton defaultButton;

    @FXML
    private MFXButton generateButton;

    @FXML
    private MFXProgressSpinner generateSpinner;

    @FXML
    private MFXTextField inputFile;

    @FXML
    private MFXTextField inputFileAmountCol;

    @FXML
    private MFXTextField inputFileBalanceCol;

    @FXML
    private MFXTextField inputFileCompanyCol;

    @FXML
    private MFXRadioButton inputFileIsOrderByName;

    @FXML
    private MFXTextField inputFileStartRow;

    @FXML
    private MFXTextField inputFileYearCol;

    @FXML
    private TextArea logTextArea;

    @FXML
    private MFXTextField outputFile;

    @FXML
    private MFXTextField outputFileBillAgeFillStartCol;

    @FXML
    private MFXTextField outputFileCompanyCol;

    @FXML
    private MFXTextField outputFileSheetName;

    @FXML
    private MFXTextField outputFileStartRow;

    @FXML
    void defaultButtonSave(MouseEvent event) {

    }

    @FXML
    void generate(MouseEvent event) {

    }

    @FXML
    void inputFileChoose(MouseEvent event) {

    }

    @FXML
    void outputFileChoose(MouseEvent event) {

    }

}
