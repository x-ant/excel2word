package com.xant.controller;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.StrUtil;
import com.xant.component.log.TextAreaAppender;
import com.xant.entity.ConfigPO;
import com.xant.manager.ConfigManager;
import com.xant.util.Excel2WordUtil;
import io.github.palexdev.materialfx.controls.MFXButton;
import io.github.palexdev.materialfx.controls.MFXProgressSpinner;
import io.github.palexdev.materialfx.controls.MFXTextField;
import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.TextArea;
import javafx.scene.input.MouseButton;
import javafx.scene.input.MouseEvent;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Duration;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.util.Objects;
import java.util.concurrent.Executor;
import java.util.concurrent.Executors;

@Slf4j
public class MainPaneController {

    private Executor executor = Executors.newSingleThreadExecutor();
    @FXML
    private MFXButton generateButton;
    @FXML
    private MFXProgressSpinner generateSpinner;
    @FXML
    private MFXButton defaultButton;
    @FXML
    private MFXTextField templateFile;
    private boolean templateFileUserTyping = false;
    @FXML
    private MFXTextField inputDir;
    private boolean inputDirUserTyping = false;
    @FXML
    private MFXTextField outputDir;
    private boolean outputDirUserTyping = false;
    @FXML
    private MFXButton inputDirChooser;
    @FXML
    private MFXButton outputDirChooser;
    @FXML
    private MFXButton templateFileChooser;
    @FXML
    private TextArea logTextArea;

    @FXML
    public void initialize() {
        TextAreaAppender.setTextArea(logTextArea);

        ConfigPO configPO = ConfigManager.getSingletonConfigPO();
        templateFile.setText(configPO.getTemplateFile());
        inputDir.setText(configPO.getInputDir());
        outputDir.setText(configPO.getOutputDir());

        PauseTransition templateFilePause = new PauseTransition(Duration.millis(500));
        templateFile.textProperty().addListener((observable, oldValue, newValue) -> {
            if (!templateFileUserTyping) {
                templateFileUserTyping = true;
            }
            templateFilePause.setOnFinished(event -> {
                templateFileUserTyping = false;

                ConfigPO templateFileConfigPO = ConfigManager.getPrototypeConfigPO();
                templateFileConfigPO.setTemplateFile(newValue);
                setDefaultButton(templateFileConfigPO);
            });
            templateFilePause.playFromStart();
        });
        PauseTransition inputDirPause = new PauseTransition(Duration.millis(500));
        inputDir.textProperty().addListener((observable, oldValue, newValue) -> {
            if (!inputDirUserTyping) {
                inputDirUserTyping = true;
            }
            inputDirPause.setOnFinished(event -> {
                inputDirUserTyping = false;

                ConfigPO inputDirConfigPO = ConfigManager.getPrototypeConfigPO();
                inputDirConfigPO.setInputDir(newValue);
                setDefaultButton(inputDirConfigPO);
            });
            inputDirPause.playFromStart();
        });
        PauseTransition outputDirPause = new PauseTransition(Duration.millis(500));
        outputDir.textProperty().addListener((observable, oldValue, newValue) -> {
            if (!outputDirUserTyping) {
                outputDirUserTyping = true;
            }
            outputDirPause.setOnFinished(event -> {
                outputDirUserTyping = false;

                ConfigPO outputDirConfigPO = ConfigManager.getPrototypeConfigPO();
                outputDirConfigPO.setOutputDir(newValue);
                setDefaultButton(outputDirConfigPO);
            });
            outputDirPause.playFromStart();
        });
    }

    @FXML
    void generate(MouseEvent event) {
        if (MouseButton.PRIMARY != event.getButton()) {
            return;
        }
        logTextArea.clear();
        generateButton.setDisable(true);
        generateSpinner.setVisible(true);

        ConfigPO configPO = paddingCurrentValue(ConfigManager.getPrototypeConfigPO());
        executor.execute(() -> {
            Excel2WordUtil.generateWordFromExcel(configPO);
            Platform.runLater(() -> {
                generateButton.setDisable(false);
                generateSpinner.setVisible(false);
            });
        });
    }

    @FXML
    void defaultButtonSave(MouseEvent event) {
        if (MouseButton.PRIMARY != event.getButton()) {
            return;
        }
        ConfigPO configPO = paddingCurrentValue(ConfigManager.getSingletonConfigPO());
        ConfigManager.setConfigPO(configPO);
        defaultButton.setDisable(true);
    }

    @FXML
    void templateFileChoose(MouseEvent event) {
        if (MouseButton.PRIMARY != event.getButton()) {
            return;
        }
        Stage stage = new Stage();
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("word模板文件选择");
        String currentTemplateFile = templateFile.getText();
        if (StrUtil.isNotEmpty(currentTemplateFile) && FileUtil.exist(currentTemplateFile) && FileUtil.isFile(currentTemplateFile)) {
            fileChooser.setInitialDirectory(new File(currentTemplateFile).getParentFile());
        }
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Word文档", "*.doc", "*.docx", "*.docm", "*.dot", "*.dotx", ".*dotm"));
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("所有类型", "*.*"));
        File targetTemplateFile = fileChooser.showOpenDialog(stage);
        if (Objects.nonNull(targetTemplateFile)) {
            templateFile.setText(targetTemplateFile.getAbsolutePath());

            ConfigPO configPO = ConfigManager.getPrototypeConfigPO();
            configPO.setTemplateFile(targetTemplateFile.getAbsolutePath());
            setDefaultButton(configPO);
        }
    }

    @FXML
    void inputDirChoose(MouseEvent event) {
        if (MouseButton.PRIMARY != event.getButton()) {
            return;
        }
        Stage stage = new Stage();
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle("excel输入文件夹选择");
        String currentInputDir = inputDir.getText();
        if (StrUtil.isNotEmpty(currentInputDir) && FileUtil.exist(currentInputDir) && FileUtil.isDirectory(currentInputDir)) {
            directoryChooser.setInitialDirectory(new File(currentInputDir));
        }
        File targetInputDir = directoryChooser.showDialog(stage);
        if (Objects.nonNull(targetInputDir)) {
            inputDir.setText(targetInputDir.getAbsolutePath());

            ConfigPO configPO = ConfigManager.getPrototypeConfigPO();
            configPO.setInputDir(targetInputDir.getAbsolutePath());
            setDefaultButton(configPO);
        }
    }

    @FXML
    void outputDirChoose(MouseEvent event) {
        if (MouseButton.PRIMARY != event.getButton()) {
            return;
        }
        Stage stage = new Stage();
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle("word输出文件夹选择");
        String currentOutputDir = outputDir.getText();
        if (StrUtil.isNotEmpty(currentOutputDir) && FileUtil.exist(currentOutputDir) && FileUtil.isDirectory(currentOutputDir)) {
            directoryChooser.setInitialDirectory(new File(currentOutputDir));
        }
        File targetOutputDir = directoryChooser.showDialog(stage);
        if (Objects.nonNull(targetOutputDir)) {
            outputDir.setText(targetOutputDir.getAbsolutePath());

            ConfigPO configPO = ConfigManager.getPrototypeConfigPO();
            configPO.setOutputDir(targetOutputDir.getAbsolutePath());
            setDefaultButton(configPO);
        }
    }

    private void setDefaultButton(ConfigPO newConfigPO) {
        ConfigPO configPO = ConfigManager.getSingletonConfigPO();
        defaultButton.setDisable(configPO.equals(newConfigPO));
    }

    private ConfigPO paddingCurrentValue(ConfigPO configPO) {
        configPO.setTemplateFile(templateFile.getText());
        configPO.setInputDir(inputDir.getText());
        configPO.setOutputDir(outputDir.getText());
        return configPO;
    }

}
