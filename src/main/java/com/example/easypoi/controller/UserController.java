package com.example.easypoi.controller;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import com.example.easypoi.util.ExcelStyleUtil;
import com.example.easypoi.handler.UserVerifyHandler;
import com.example.easypoi.entity.ExcelUser;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author wad
 * @since 2021-03-31 17:28:40
 */

@RestController
@RequestMapping("/user")
public class UserController {


    @Autowired
    private UserVerifyHandler userVerifyHandler;

    @PostMapping("/upload")
    public String upload(@RequestParam("file") MultipartFile multipartFile) throws Exception {

        ImportParams params = new ImportParams();
        ExcelImportResult<ExcelUser> result;
        try {
            // 表头设置为1行
            params.setHeadRows(1);
            // 标题行设置为0行，默认是0，可以不设置
            params.setTitleRows(0);
            //开启检验
            params.setNeedVerfiy(true);
            params.setVerifyHandler(userVerifyHandler);
            result = ExcelImportUtil.importExcelMore
                    (multipartFile.getInputStream(), ExcelUser.class, params);
        } finally {
            // 清除threadLocal 防止内存泄漏
            ThreadLocal<List<ExcelUser>> threadLocal = userVerifyHandler.getThreadLocal();
            if (threadLocal != null) {
                threadLocal.remove();
            }
        }

        System.out.println(result.getFailList().toString());
        System.out.println(result.getList().toString());


        return null;

    }


    @GetMapping("/getList")
    public String getList(HttpServletRequest request, HttpServletResponse response) {
        List<ExcelUser> list = new ArrayList<>();
        ExcelUser excelUser = new ExcelUser();
        excelUser.setName("张三");
        excelUser.setAge("20");
        excelUser.setSex("男");
        ExcelUser excelUser1 = new ExcelUser();
        excelUser1.setName("李四");
        excelUser1.setAge("30");
        excelUser1.setSex("女");
        list.add(excelUser);
        list.add(excelUser1);
        ExportParams exportParams = new ExportParams("计算机一班学生", "学生");
        exportParams.setStyle(ExcelStyleUtil.class);
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams,
                ExcelUser.class, list);

        // 判断数据
        if (workbook == null) {
            return "fail";
        }
        // 设置excel的文件名称
        String excelName = "测试导出";
        // 重置响应对象
        response.reset();
        // 当前日期，用于导出文件名称


        // 指定下载的文件名--设置响应头
        response.setHeader("Content-Disposition", "attachment;filename=" + excelName + ".xls");
        response.setContentType("application/vnd.ms-excel;charset=UTF-8");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Cache-Control", "no-cache");
        response.setDateHeader("Expires", 0);
        // 写出数据输出流到页面
        try {
            OutputStream output = response.getOutputStream();
            BufferedOutputStream bufferedOutPut = new BufferedOutputStream(output);
            workbook.write(bufferedOutPut);
            bufferedOutPut.flush();
            bufferedOutPut.close();
            output.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "success";

    }

}