package com.xant;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.*;

/**
 * excel填充word模板
 *
 * @author xuhq
 */
@Slf4j
public class Excel2WordApplication {

    private static final String PROJECT_ROOT = System.getProperty("user.dir");

    public static void main(String[] args) {

        try {
            Properties properties = new Properties();
            File configFile = new File(PROJECT_ROOT, "config.properties");
            if (configFile.exists()) {
                properties.load(new FileReader(configFile));
            }

            String prefix = properties.getProperty("template.prefix", "${");
            String suffix = properties.getProperty("template.suffix", "}");
            String inputDir = properties.getProperty("input.dir", PROJECT_ROOT);
            String outputDir = properties.getProperty("output.dir", PROJECT_ROOT);
            String templateFile = properties.getProperty("template.file", PROJECT_ROOT + File.separator + "template.docx");
            log.info("当前读取excel文件的默认目录为：{}", inputDir);
            if (!FileUtil.exist(inputDir)) {
                log.error("输入excel目录不存在，结束执行，当前来源excel目录为：{}", inputDir);
                return;
            }
            log.info("当前输出word文件的默认目录为：{}", outputDir);
            if (!FileUtil.exist(outputDir)) {
                log.error("输出word目录不存在，结束执行，当前输出word目录为：{}", outputDir);
                return;
            }
            log.info("当前模板word文件的为：{}", templateFile);
            if (!FileUtil.exist(templateFile)) {
                log.error("模板文件不存在，结束执行，当前模板文件为：{}", templateFile);
                return;
            }

            // 配置word的模板语法
            Configure configure = Configure.builder().buildGramer(prefix, suffix).build();
            // 数据填充并生成新Word
            try (XWPFTemplate template = XWPFTemplate.compile(templateFile, configure)) {
                Files.walkFileTree(Paths.get(inputDir), new SimpleFileVisitor<Path>() {
                    @Override
                    public FileVisitResult visitFile(Path filePath, BasicFileAttributes attrs) throws IOException {
                        String extName = FileUtil.extName(filePath.getFileName().toString());
                        if (StrUtil.equalsAny(extName, "xls", "xlsx")) {
                            File excelFile = filePath.toFile();

                            List<String> sheetNameList = new ArrayList<>();
                            try (ExcelReader reader = ExcelUtil.getReader(excelFile)) {
                                sheetNameList = reader.getSheetNames();
                            }

                            Map<String, Object> dataModel = new LinkedHashMap<>();
                            List<String> finalSheetNameList = sheetNameList;
                            RowHandler rowHandler = new RowHandler() {
                                private int sheetIndexField = 0;

                                @Override
                                public void handle(int sheetIndex, long rowIndex, List<Object> rowCells) {
                                    sheetIndexField = sheetIndex;
                                    for (int i = 0; i < rowCells.size(); i++) {
                                        Object currentData = rowCells.get(i);
                                        String colName = ExcelUtil.indexToColName(i);
                                        // 替换表达式不能以数字开头
                                        int from1SheetIndex = sheetIndex + 1;
                                        long from1RowIndex = rowIndex + 1;
                                        dataModel.put("S" + from1SheetIndex + colName + from1RowIndex, currentData);
                                    }
                                }

                                @Override
                                public void doAfterAllAnalysed() {
                                    log.info("读取完Sheet: {}", finalSheetNameList.get(sheetIndexField));
                                }
                            };
                            ExcelUtil.readBySax(excelFile, -1, rowHandler);
                            String mainName = FileUtil.mainName(excelFile);
                            String outputFile = outputDir + File.separator + mainName + ".docx";
                            if (FileUtil.exist(outputFile)) {
                                FileUtil.del(outputFile);
                            }
                            try (FileOutputStream out = new FileOutputStream(outputFile)) {
                                template.render(dataModel).write(out);
                                log.info("已生成目标文件：{}", outputFile);
                            }
                        }
                        return FileVisitResult.CONTINUE;
                    }
                });
            }
        } catch (Exception e) {
            log.error("文件处理异常，请联系开发人员", e);
        } finally {
            log.info("程序执行完毕，按任意键退出...");
            try {
                System.in.read();
            } catch (IOException ignored) {
            }
        }
    }
}
