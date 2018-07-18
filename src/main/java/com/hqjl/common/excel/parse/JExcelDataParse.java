package com.hqjl.common.excel.parse;

import com.hqjl.common.excel.bean.CellBean;
import com.hqjl.common.excel.bean.ExcelCellData;
import com.hqjl.common.excel.bean.ExcelSheetData;
import com.hqjl.common.excel.bean.ExcelTableData;
import com.hqjl.common.excel.bean.FormatType;
import com.hqjl.common.excel.bean.TableBean;
import com.hqjl.common.excel.service.ExportTableService;
import com.hqjl.common.excel.utils.ObjectHelper;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class JExcelDataParse {

    private Workbook workbook;
    private List<ExcelSheetData> excelSheetDatas;
    private Collection<CellBean> cellBeans;
    private int startRowCount;
    private int startColumnCount;
    private int rowCount;
    private int columnCount;

    private Map<String,Integer> customTypeTransform=new HashMap<>();

    public JExcelDataParse(Workbook workbook, List<ExcelSheetData> excelSheetDatas) {
        this.workbook = workbook;
        this.excelSheetDatas = excelSheetDatas;
        ini();
    }

    public JExcelDataParse(Workbook workbook,ExcelSheetData excelSheetData){
        this.workbook=workbook;
        List<ExcelSheetData> excelSheetDatas=new ArrayList<>();
        excelSheetDatas.add(excelSheetData);
        this.excelSheetDatas=excelSheetDatas;
        ini();
    }

    public JExcelDataParse(Workbook workbook,String sheetName,ExcelTableData excelTableData){
        this.workbook=workbook;
        List<ExcelTableData> excelTableDatas=new ArrayList<>();
        excelTableDatas.add(excelTableData);
        ExcelSheetData excelSheetData=new ExcelSheetData();
        excelSheetData.setSheetName(sheetName);
        excelSheetData.setTables(excelTableDatas);
        List<ExcelSheetData> excelSheetDatas=new ArrayList<>();
        excelSheetDatas.add(excelSheetData);
        this.excelSheetDatas=excelSheetDatas;
        ini();
    }


    private TableBean tableBean;
    private ExcelSheetData excelSheetData;
    private List<List<ExcelCellData>> headersDataList;
    private List<Map<String, Object>> rowsData;

    public void doParse() {

        if (excelSheetDatas != null && excelSheetDatas.size() > 0) {
            for (ExcelSheetData esd : excelSheetDatas) {
                this.excelSheetData = esd;
                doExcelSheetTableParse();
                this.tableBean = new TableBean(rowCount, columnCount);
                tableBean.setCellBeans(cellBeans);
                Sheet sheet = workbook.getSheet(excelSheetData.getSheetName());
                sheet=sheet==null?workbook.createSheet(excelSheetData.getSheetName()):sheet;
                new ExportTableService(sheet, tableBean).doExport();
                ini();
            }
        }

    }

    private void doExcelSheetTableParse() {

        for (ExcelTableData tableData : excelSheetData.getTables()) {
            ExcelCellData titleData = tableData.getTitle();
            startRowCount = tableData.getStartRow();
            startColumnCount = tableData.getStartColumn();
            if (titleData != null && titleData.getName() != null) {
                CellBean cellBean = new CellBean(titleData.getName(), tableData.getStartRow(), tableData.getStartColumn(), titleData.getxSize(), titleData.getySize());
                cellBeans.add(cellBean);

                startRowCount = startRowCount + titleData.getySize();

                calculateColumCount(startColumnCount + titleData.getxSize());
            }
            this.headersDataList = tableData.getHeaders();
            this.rowsData = tableData.getRows();

            doExcelSheetHeadParse();
            doExcelSheetDataParse(tableData.getNumType());
            if(rowCount < startRowCount){
                rowCount = startRowCount;
            }
        }

    }

    private void doExcelSheetHeadParse() {

        int maxRow = 0;
        for (List<ExcelCellData> headersData : headersDataList) {
            int rowColumnCount = startColumnCount;
            for (ExcelCellData headerData : headersData) {
                if (ObjectHelper.isNotEmpty(headerData.getName())) {
                    CellBean cellBean = new CellBean(headerData.getName(), startRowCount, rowColumnCount, headerData.getxSize(), headerData.getySize());
                    cellBeans.add(cellBean);
                }
                    if(maxRow < startRowCount+headerData.getySize()){
                        maxRow = startRowCount+headerData.getySize();
                    }

                    rowColumnCount = rowColumnCount + headerData.getxSize();

                    calculateColumCount(rowColumnCount);

            }
            startRowCount ++;

        }

        if(maxRow > startRowCount){
            startRowCount = maxRow;
        }

    }

    private void doExcelSheetDataParse(int numType) {

        List<String> keys = headersDataList.get(headersDataList.size() - 1).stream().map(e -> e.getKey())
                .collect(Collectors.toList());
        Map<String, Integer> xSizeMap = headersDataList.get(headersDataList.size() - 1).stream()
                .collect(Collectors.toMap(k -> k.getKey(), v -> v.getxSize(),(v,v2)->v2));

        for (Map<String, Object> rowData : rowsData) {
            int rowColumnCount = startColumnCount;
            for (String key : keys) {
                Object value = rowData.get(key);
                numType=customTypeTransform.containsKey(key)?customTypeTransform.get(key):numType;
                if(ObjectHelper.isNotEmpty(value)) {
                    if(value instanceof ExcelCellData){
                        ExcelCellData excelCellData = (ExcelCellData) value;
                        cellBeans.add(new CellBean(String.valueOf(excelCellData.getName()),startRowCount,rowColumnCount,
                            excelCellData.getxSize(),excelCellData.getySize(),numType));
                    }else {
                        CellBean cellBean = new CellBean(String.valueOf(value), startRowCount, rowColumnCount, xSizeMap.get(key), 1,numType);
                        cellBeans.add(cellBean);
                    }
                }
                rowColumnCount = rowColumnCount + xSizeMap.get(key);

                calculateColumCount(rowColumnCount);
            }
            startRowCount++;
        }

    }

    private void ini() {
        this.startRowCount = 0;
        this.startColumnCount = 0;
        this.rowCount = 0;
        this.columnCount = 0;
        this.cellBeans = new HashSet<>();
    }

    private void calculateColumCount(int compareColum){
        if(columnCount < compareColum){
            columnCount = compareColum;
        }
    }

    public Map<String, Integer> getCustomTypeTransform() {
        return customTypeTransform;
    }

    public void setCustomTypeTransform(Map<String, Integer> customTypeTransform) {
        this.customTypeTransform = customTypeTransform;
    }
    public JExcelDataParse putCustomTypeTransform(String key,FormatType formatType){
        customTypeTransform.put(key,formatType.getNumType());
        return this;
    }
}
