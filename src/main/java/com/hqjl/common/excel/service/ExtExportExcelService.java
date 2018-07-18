package com.hqjl.common.excel.service;

import com.hqjl.common.excel.annotation.ExcelColumnGroup;
import com.hqjl.common.excel.annotation.Lang;
import com.hqjl.common.excel.bean.BaseExcelVo;
import com.hqjl.common.excel.bean.CellBean;
import com.hqjl.common.excel.bean.ColumnBean;
import com.hqjl.common.excel.bean.GroupConfig;
import com.hqjl.common.excel.bean.TookPairs;
import com.hqjl.common.excel.exception.AdapterException;
import com.hqjl.common.excel.exception.ColumnErrorException;
import com.hqjl.common.excel.json.JsonUtil;
import com.hqjl.common.excel.utils.ObjectHelper;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @author xiang.li
 * create date 2018/7/5
 * description
 */
public class ExtExportExcelService<T extends BaseExcelVo> extends ExportExcelService {
    private Map<String,String[]> extFileName=new HashMap<>();

    public ExtExportExcelService(Object data, Sheet sheet, String[] fields, String title) throws Exception {
        super(data, sheet, fields, title);
    }


    public ExtExportExcelService(String sheetName, Object data, String[] fields, String title) throws Exception {
        super(sheetName, data, fields, title);
    }

    public ExtExportExcelService(Object data, Sheet sheet, String title) throws Exception {
        super(data, sheet, title);
    }

    public ExtExportExcelService putExtFileName(String fileName,String[] langValue){
        extFileName.put(fileName,langValue);
        return this;
    }
    @Override
    protected int createHead() {
        List<String> columnBeanJson = new ArrayList<String>(0);
        List<String> columns = new ArrayList<String>();
       List<CellBean> cellBeans=new ArrayList<>();
       int  extColumnIndex=0;
       int columnIndex=0;
        boolean isExt=false;
        for (Field field : dataBean.getFiledList()) {
            if (field.isAnnotationPresent(ExcelColumnGroup.class)) {
                ExcelColumnGroup columnGroup = field.getAnnotation(ExcelColumnGroup.class);
                GroupConfig group = (GroupConfig) groupConfig.get(field.getName());
                //List里面是基本类型
                if (!BaseExcelVo.class.isAssignableFrom(columnGroup.type())) {
                    cellBeans.add(new CellBean(language.translate(field), COLUMN_ROW,extColumnIndex,group.getLength(),1));
                    extColumnIndex++;
                    for (int i = 0; i < group.getLength(); i++) {
                        ColumnBean columnBean = new ColumnBean();
                        columnBean.setColumnName(field.getName());
                        columnBean.setSize(1);
                        columnBean.setLength(group.getLength());
                        TookPairs tookPairs =(TookPairs) tookMap.get(field.getName() + i);
                        if (tookPairs != null) {
                            String took = tookPairs.getValue();
                            if (!ObjectHelper.isNullOrEmptyString(took)) {
                                columnBean.setTookValue(took);
                            }
                        }
                        columnBeanJson.add(JsonUtil.toJSON(columnBean));
                        columns.add(group.getLangName(0, i));
                        cellBeans.add(new CellBean(group.getLangName(0, i), COLUMN_ROW +1,columnIndex,1,1));
                    }
                } else {//List里面是复杂的对象
                    isExt=true;
                    if (ObjectHelper.isEmpty(group.getFieldNames())) {
                        String[] filedNames = dataBean.getChildDataBean(field.getName()).getFiledNames();
                        group.setFieldNames(Arrays.asList(filedNames));
                    }
                    if (field.isAnnotationPresent(Lang.class)&&!extFileName.containsKey(field.getName())) {
                        int xSize = group.getLength() * group.getFieldNames().size();
                        cellBeans.add(new CellBean(language.translate(field), COLUMN_ROW, extColumnIndex, xSize, 1));
                        extColumnIndex+=xSize;
                    }

                    for (int i = 0; i < group.getLength(); i++) {
                        if (extFileName.containsKey(field.getName())) {
                            String[] langValues = extFileName.get(field.getName());
                            int xSize = group.getFieldNames().size();
                            cellBeans.add(new CellBean(langValues[i], COLUMN_ROW, extColumnIndex, xSize, 1));
                            extColumnIndex+=xSize;
                        }
                        for (int j = 0; j < group.getFieldNames().size(); j++) {
                            ColumnBean columnBean = new ColumnBean();
                            columnBean.setColumnName(field.getName());
                            columnBean.setSize(group.getGroupSize());
                            columnBean.setLength(group.getLength());
                            columnBean.setInnerColumn(group.getFieldNames().get(j));
                            TookPairs tookPairs = (TookPairs) tookMap.get(columnBean.getInnerColumn() + i);
                            if (tookPairs != null && !ObjectHelper.isNullOrEmptyString(tookPairs.getValue())) {
                                columnBean.setTookName(tookPairs.getSourceField());
                                columnBean.setTookValue(tookPairs.getValue());
                            }
                            columnBeanJson.add(JsonUtil.toJSON(columnBean));
                            String langName = group.getLangName(j, i);
                            columns.add(langName);
                            cellBeans.add(new CellBean(langName, COLUMN_ROW +1,columnIndex,1,1));
                            columnIndex++;
                        }
                    }

                }
            } else {
                ColumnBean columnBean = new ColumnBean();
                columnBean.setColumnName(field.getName());
                TookPairs tookPairs =(TookPairs) tookMap.get(field.getName() + 0);
                if (tookPairs != null && !ObjectHelper.isNullOrEmptyString(tookPairs.getValue())) {
                    columnBean.setTookValue(tookPairs.getValue());
                }
                columnBeanJson.add(JsonUtil.toJSON(columnBean));
                String LangValue = language.translate(field);
                columns.add(LangValue);
                cellBeans.add(new CellBean(LangValue, COLUMN_ROW,extColumnIndex,1,2));
                extColumnIndex++;
                columnIndex++;
            }
        }
        sheet.createRow(hashRow);
        addTitle(sheet, TITLE_ROW, columnBeanJson.size(), title);
        Row row = addRow(sheet, HIDDENFIELDHEAD, columnBeanJson.toArray(new String[]{}));
        if(isExt){
            createHead(cellBeans);
            START_ROW = COLUMN_ROW +2;
        }else {
            addRow(sheet, COLUMN_ROW, columns.toArray(new String[]{}));
        }
        row.setHeight(Short.valueOf("0"));
        return columnBeanJson.size();
    }

    private void createHead(Collection<CellBean> cellBeans) {
        if(ObjectHelper.isNotEmpty(cellBeans)){
            for(CellBean cellBean:cellBeans){
                if(cellBean.getXSize()>1||cellBean.getYSize()>1){
                    log.debug("有合并单元格：{}", JsonUtil.toJSON(cellBean));
                    CellRangeAddress range=new CellRangeAddress(cellBean.getRowIndex(),cellBean.getRowIndex()+cellBean.getYSize()-1,cellBean.getColumnIndex(),cellBean.getColumnIndex()+cellBean.getXSize()-1);
                    sheet.addMergedRegion(range);

                }
                log.debug("set row:{},column:{},content:{}",cellBean.getRowIndex(),cellBean.getColumnIndex(),cellBean.getContent());
                Row row = sheet.getRow(cellBean.getRowIndex());
                row=row==null?sheet.createRow(cellBean.getRowIndex()):row;
                Cell cell = row.getCell(cellBean.getColumnIndex());
                cell= cell == null ? row.createCell(cellBean.getColumnIndex()) : cell;
                CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
                //水平居中
                if(cellBean.isAlignCenter()){
                    cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
                }
                //垂直居中
                if(cellBean.isVerticalCenter()){
                    cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
                }
                cell.setCellStyle(cellStyle);
                try {
                    cell.setCellValue(Double.parseDouble (cellBean.getContent().trim ()));
                }catch (NumberFormatException e){
                    cell.setCellValue(cellBean.getContent());
                }
            }
        }
    }

    @Override
    public ExportExcelService doExport() throws AdapterException, ColumnErrorException {
        log.debug("准备生成表头。。。");
        columnLength = createHead();
        return super.createData();
    }

}
