package com.xant.util;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.template.MetaTemplate;
import com.xant.entity.ConfigPO;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.*;

/**
 * 根据word模板，使用excel数据生成word文件
 *
 * @author xuhq
 */
@Slf4j
public class Excel2WordUtil {

    public static void generateWordFromExcel(ConfigPO configPO) {
        try {
            log.info("当前模板word文件的为：{}", configPO.getTemplateFile());
            if (StrUtil.isEmpty(configPO.getTemplateFile())) {
                log.error("模板文件为空，结束执行");
                return;
            }
            if (!FileUtil.exist(configPO.getTemplateFile())) {
                log.error("模板文件不存在，结束执行，当前模板文件为：{}", configPO.getTemplateFile());
                return;
            }
            log.info("当前读取excel文件的默认目录为：{}", configPO.getInputDir());
            if (StrUtil.isEmpty(configPO.getInputDir())) {
                log.error("输入excel目录为空，结束执行");
                return;
            }
            if (!FileUtil.exist(configPO.getInputDir())) {
                log.error("输入excel目录不存在，结束执行，当前来源excel目录为：{}", configPO.getInputDir());
                return;
            }
            log.info("当前输出word文件的默认目录为：{}", configPO.getOutputDir());
            if (StrUtil.isEmpty(configPO.getOutputDir())) {
                log.error("输出word目录为空，结束执行");
                return;
            }
            if (!FileUtil.exist(configPO.getOutputDir())) {
                log.error("输出word目录不存在，结束执行，当前输出word目录为：{}", configPO.getOutputDir());
                return;
            }

            // 配置word的模板语法
            Configure configure = Configure.builder().buildGramer(configPO.getTemplatePrefix(), configPO.getTemplateSuffix()).build();
            // 数据填充并生成新Word
            try (XWPFTemplate template = XWPFTemplate.compile(configPO.getTemplateFile(), configure)) {
                // region ========== 收集变量 ==========
                int prefixLength = StrUtil.length(configPO.getTemplatePrefix());
                int suffixLength = StrUtil.length(configPO.getTemplateSuffix());
                List<MetaTemplate> elementTemplateList = template.getElementTemplates();
                Set<String> variableSet = new HashSet<>();
                for (MetaTemplate metaTemplate : elementTemplateList) {
                    String metaTemplateVariable = metaTemplate.variable();
                    variableSet.add(StrUtil.sub(metaTemplateVariable, prefixLength, -suffixLength));
                }
                // endregion
                Files.walkFileTree(Paths.get(configPO.getInputDir()), new SimpleFileVisitor<Path>() {
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
                                        String variable = "S" + from1SheetIndex + colName + from1RowIndex;
                                        if (variableSet.contains(variable)) {
                                            dataModel.put(variable, currentData);
                                        }
                                    }
                                }

                                @Override
                                public void doAfterAllAnalysed() {
                                    log.info("读取完Sheet: {}", finalSheetNameList.get(sheetIndexField));
                                }
                            };
                            ExcelUtil.readBySax(excelFile, -1, rowHandler);
                            String mainName = FileUtil.mainName(excelFile);
                            String outputFile = configPO.getOutputDir() + File.separator + mainName + ".docx";
                            if (FileUtil.exist(outputFile)) {
                                FileUtil.del(outputFile);
                            }
                            try (FileOutputStream outputStream = new FileOutputStream(outputFile)) {
                                template.render(dataModel).write(outputStream);
                                log.info("已生成目标文件：{}", outputFile);
                            }
                        }
                        return FileVisitResult.CONTINUE;
                    }
                });
            }
        } catch (Exception e) {
            log.error("文件处理异常，请联系开发人员", e);
        }
    }
}
