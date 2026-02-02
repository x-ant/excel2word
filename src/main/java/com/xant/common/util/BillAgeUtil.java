package com.xant.common.util;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import com.xant.component.jdbc.TransactionUtil;
import com.xant.entity.YearBillConfigPO;
import com.xant.entity.YearBillPO;
import com.xant.manager.YearBillManager;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;

import java.math.BigDecimal;
import java.util.*;

/**
 * 账龄工具类
 *
 * @author xant
 */
@Slf4j
public class BillAgeUtil {

    public static void truncateYearBill(YearBillConfigPO configPO) {

        /*List<String> sheetNameList = new ArrayList<>();
        try (ExcelReader reader = ExcelUtil.getReader(configPO.getInputFile())) {
            sheetNameList = reader.getSheetNames();
        }

        Map<String, List<YearBillPO>> name2DataListMap = new HashMap<>();
        List<String> finalSheetNameList = sheetNameList;
        RowHandler rowHandler = new RowHandler() {
            private int sheetIndexField = 0;
            private String lastName = StrUtil.EMPTY;

            @Override
            public void handle(int sheetIndex, long rowIndex, List<Object> rowCells) {
                sheetIndexField = sheetIndex;
                if (rowIndex < configPO.getInputFileStartRow()) {
                    return;
                }

                YearBillPO yearBillPO = buildYearBillPO(rowCells, configPO);
                if (Objects.isNull(yearBillPO)) {
                    return;
                }
                List<YearBillPO> yearBillList = name2DataListMap.computeIfAbsent(yearBillPO.getCompany(), k -> new ArrayList<>());
                yearBillList.add(yearBillPO);
                if (configPO.getInputFileIsOrderByName() && !lastName.equals(yearBillPO.getCompany())) {
                    List<YearBillPO> truncateList = doTruncateYearBill(name2DataListMap.remove(lastName));
                    YearBillManager.getInstance().saveBatch(truncateList);
                    lastName = yearBillPO.getCompany();
                }
            }

            @Override
            public void doAfterAllAnalysed() {
                log.info("读取完Sheet: {}", finalSheetNameList.get(sheetIndexField));
            }
        };
        ExcelUtil.readBySax(configPO.getInputFile(), -1, rowHandler);

        // 处理最后一个，或者是整个都处理掉
        TransactionUtil.transactionWithRequired(() -> {
            Iterator<Map.Entry<String, List<YearBillPO>>> iterator = name2DataListMap.entrySet().iterator();
            while (iterator.hasNext()) {
                Map.Entry<String, List<YearBillPO>> name2DateListEntry = iterator.next();
                List<YearBillPO> truncateList = doTruncateYearBill(name2DateListEntry.getValue());
                YearBillManager.getInstance().saveBatch(truncateList);
                iterator.remove();
            }
            return Void.TYPE;
        });*/

        writeYearBill(configPO);
    }

    /**
     * 获取有收入没有被抵消的年份
     *
     * @param yearBillList 完整年份
     * @return 没有被抵消年份的列表
     */
    private static List<YearBillPO> doTruncateYearBill(List<YearBillPO> yearBillList) {
        // 如果只有名称，都没有年份数据，就是空列表，返回0
        if (CollUtil.isEmpty(yearBillList)) {
            return null;
        }
        yearBillList.sort(Comparator.comparing(YearBillPO::getYear));
        YearBillPO lastPO = CollUtil.getLast(yearBillList);
        BigDecimal balance = lastPO.getBalance();
        BigDecimal amountCount = BigDecimal.ZERO;
        List<YearBillPO> resultList = new ArrayList<>();
        for (int i = yearBillList.size() - 1; i >= 0; i--) {
            YearBillPO yearBillPO = yearBillList.get(i);
            resultList.add(yearBillPO);
            amountCount = amountCount.add(yearBillPO.getRecAmount());
            if (amountCount.compareTo(balance) >= 0) {
                BigDecimal truncateAmount = yearBillPO.getRecAmount().subtract(amountCount.subtract(balance));
                yearBillPO.setRecAmount(truncateAmount);
                break;
            }
        }
        return resultList;
    }

    private static void writeYearBill(YearBillConfigPO configPO) {
        try (ExcelWriter writer = ExcelUtil.getWriter(configPO.getOutputFile())) {
            writer.setSheet(configPO.getOutputFileSheetName());
            int outputFileBillAgeFillStartColIndex = ExcelUtil.colNameToIndex(configPO.getOutputFileBillAgeFillStartCol());
            int outputFileCompanyColIndex = ExcelUtil.colNameToIndex(configPO.getOutputFileCompanyCol());
            for (int i = 0; i < writer.getRowCount(); i++) {
                Cell nameCell = writer.getCell(i, outputFileCompanyColIndex);
                String company = nameCell.getStringCellValue();
                if (StrUtil.isEmpty(company)) {
                    continue;
                }
                List<YearBillPO> yearBillList = YearBillManager.getInstance().queryListByCompany(company);
                if (CollUtil.isEmpty(yearBillList)) {
                    continue;
                }
            }
        }
    }


    private static YearBillPO buildYearBillPO(List<Object> rowCellList, YearBillConfigPO configPO) {
        YearBillPO yearBillPO = new YearBillPO();
        if (CollUtil.isEmpty(rowCellList)) {
            return null;
        }

        int inputFileYearColIndex = ExcelUtil.colNameToIndex(configPO.getInputFileYearCol());
        String yearStr = StrUtil.toStringOrNull(CollUtil.get(rowCellList, inputFileYearColIndex));
        if (StrUtil.isEmpty(yearStr)) {
            log.warn("第{}行数据没有年份，跳过处理逻辑", rowCellList);
            return null;
        }
        yearBillPO.setYear(Integer.parseInt(yearStr));

        int inputFileCompanyColIndex = ExcelUtil.colNameToIndex(configPO.getInputFileCompanyCol());
        yearBillPO.setCompany(StrUtil.toStringOrNull(CollUtil.get(rowCellList, inputFileCompanyColIndex)));

        int inputFileAmountColIndex = ExcelUtil.colNameToIndex(configPO.getInputFileAmountCol());
        String amountStr = StrUtil.toStringOrNull(CollUtil.get(rowCellList, inputFileAmountColIndex));
        if (StrUtil.isEmpty(amountStr)) {
            yearBillPO.setRecAmount(BigDecimal.ZERO);
        } else {
            yearBillPO.setRecAmount(new BigDecimal(amountStr));
        }

        int inputFileBalanceColIndex = ExcelUtil.colNameToIndex(configPO.getInputFileBalanceCol());
        String balanceStr = StrUtil.toStringOrNull(CollUtil.get(rowCellList, inputFileBalanceColIndex));
        if (StrUtil.isEmpty(balanceStr)) {
            yearBillPO.setBalance(BigDecimal.ZERO);
        } else {
            yearBillPO.setBalance(new BigDecimal(balanceStr));
        }
        return yearBillPO;
    }


}
