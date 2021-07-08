package com.example.easypoi.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.handler.inter.IExcelDataModel;
import cn.afterturn.easypoi.handler.inter.IExcelModel;
import lombok.Data;

import javax.validation.constraints.NotBlank;
import javax.validation.constraints.Pattern;
import java.util.Date;

/**
 * @author wad
 * @date 2021年05月08日 16:05
 */
@Data
public class ExcelUser implements IExcelDataModel, IExcelModel {

    /**
     * 行号
     */
    private int rowNum;

    /**
     * 错误消息
     */
    private String errorMsg;

    @Excel(name = "姓名")
    @NotBlank(message = "[姓名]不能为空")
    private String name;

    @Excel(name = "性别", replace = {"男_0", "女_1"})
    @NotBlank(message = "[性别]不能为空")
    @Pattern(regexp = "[01]", message = "性别输入错误")
    private String sex;

    @Excel(name = "年龄", fixedIndex = 2)
    @NotBlank(message = "[年龄]不能为空")
    private String age;

    @Excel(name = "学历", fixedIndex = 2)
    private String education;

    @Excel(name = "就读学校", fixedIndex = 2)
    private String school;

    @Excel(name = "毕业时间")
    private String date;

    @Excel(name = "籍贯")
    private String place;

    @Excel(name = "备注")
    private String remark;


}
