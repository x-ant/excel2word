package com.xant.component;

import ch.qos.logback.classic.encoder.PatternLayoutEncoder;
import ch.qos.logback.classic.spi.ILoggingEvent;
import ch.qos.logback.core.AppenderBase;
import javafx.application.Platform;
import javafx.scene.control.TextArea;

/**
 * 自定义的 Logback Appender，用于将日志输出到 JavaFX TextArea。
 * 注意：此类必须在 Logback 配置文件 (logback.xml) 中引用。
 */
public class TextAreaAppender extends AppenderBase<ILoggingEvent> {

    /**
     * 核心：持有 Encoder 的引用，将由 Logback 框架自动注入
     */
    private PatternLayoutEncoder encoder;
    /**
     * 静态引用，以便在UI控制器中设置
     */
    private static TextArea targetTextArea;
    private static final int MAX_LOG_LENGTH = 10000;

    /**
     * 必须由 Logback 通过配置文件注入。
     */
    public void setEncoder(PatternLayoutEncoder encoder) {
        this.encoder = encoder;
    }

    /**
     * 必须由JavaFX应用线程调用，在界面初始化时设置目标TextArea。
     */
    public static void setTextArea(TextArea textArea) {
        targetTextArea = textArea;
    }

    @Override
    protected void append(ILoggingEvent event) {
        // 如果 TextArea 或 Encoder 未初始化，则忽略
        if (targetTextArea == null || encoder == null) {
            return;
        }
        // 使用 Encoder 将日志事件转换为字节数组，再解码为字符串
        byte[] encodedBytes = encoder.encode(event);
        // 默认使用系统编码，建议在 Encoder 中明确指定
        String formattedLog = new String(encodedBytes, encoder.getCharset());

        // 确保UI更新在JavaFX应用线程上执行
        Platform.runLater(() -> {
            if (targetTextArea != null) {
                // 追加日志
                targetTextArea.appendText(formattedLog);
                // 可选：自动滚动到最后一行
                targetTextArea.setScrollTop(Double.MAX_VALUE);
                // 可选：限制TextArea中的文本长度，防止无限增长
                enforceMaxLength();
            }
        });
    }

    /**
     * 可选方法：限制日志区域的最大字符数，以保持性能。
     */
    private void enforceMaxLength() {
        if (targetTextArea.getLength() > MAX_LOG_LENGTH) {
            int excess = targetTextArea.getLength() - MAX_LOG_LENGTH / 2; // 保留一半
            targetTextArea.deleteText(0, excess);
        }
    }

    @Override
    public void start() {
        // 在 Appender 启动时，必须调用 super.start()，并检查必要的依赖
        if (this.encoder == null) {
            addError("No encoder set for the appender named \"" + name + "\".");
            return;
        }
        // 初始化 encoder（非常重要！）
        this.encoder.start();
        super.start();
    }

    @Override
    public void stop() {
        // 停止时，也停止 encoder
        if (this.encoder != null) {
            this.encoder.stop();
        }
        super.stop();
    }
}