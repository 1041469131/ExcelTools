/*
 * Copyright 2015 www.hyberbin.com.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * Email:hyberbin@qq.com
 */
package com.hqjl.common.excel;

import com.hqjl.common.excel.annotation.ExcelColumnGroup;
import com.hqjl.common.excel.annotation.ExcelVoConfig;
import com.hqjl.common.excel.annotation.Lang;
import com.hqjl.common.excel.annotation.input.InputDicConfig;
import com.hqjl.common.excel.annotation.output.OutputDicConfig;
import com.hqjl.common.excel.annotation.validate.DicValidateConfig;
import com.hqjl.common.excel.bean.BaseExcelVo;

import java.util.List;

/**
 *
 * @author Hyberbin
 */
@ExcelVoConfig//Excel导出的配置
public class SchoolCourse extends BaseExcelVo {

    @Lang(value = "ID")//Excel导出的配置
    private String id;
    @Lang(value = "课程名称")//Excel导出的配置
    private String courseName;

    @InputDicConfig(dicCode = "KCLX")//Excel导入的配置
    @OutputDicConfig(dicCode = "KCLX")//Excel导出的配置
    @DicValidateConfig(dicCode = "KCLX")//如果要导出下拉框就加这个
    @Lang(value = "课程类型")//Excel导出的配置
    private String type;
    @ExcelColumnGroup(type = String.class)
    private List<String> baseArray;
    @Lang(value = "教学班")
    @ExcelColumnGroup(type = InnerVo.class)
    private List<InnerVo> innerVoArray;

    @Lang(value = "教学班")
    @ExcelColumnGroup(type = InnerVo.class)
    private List<InnerVo> innerVoArray1;

    public SchoolCourse() {
    }

    public SchoolCourse(String id, String courseName, String type) {
        this.id = id;
        this.courseName = courseName;
        this.type = type;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getCourseName() {
        return courseName;
    }

    public void setCourseName(String courseName) {
        this.courseName = courseName;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public List<String> getBaseArray() {
        return baseArray;
    }

    public void setBaseArray(List<String> baseArray) {
        this.baseArray = baseArray;
    }

    public List<InnerVo> getInnerVoArray() {
        return innerVoArray;
    }

    public void setInnerVoArray(List<InnerVo> innerVoArray) {
        this.innerVoArray = innerVoArray;
    }

    public List<InnerVo> getInnerVoArray1() {
        return innerVoArray1;
    }

    public void setInnerVoArray1(List<InnerVo> innerVoArray1) {
        this.innerVoArray1 = innerVoArray1;
    }

    @Override
    public int getHashVal() {
        return 0;
    }

}
