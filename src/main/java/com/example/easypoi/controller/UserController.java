package com.example.easypoi.controller;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
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
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

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

    /**
     * 模板导出方法
     *
     * @return
     */
    @GetMapping("templateDerive")
    public String templateDerive() throws IOException {
        //这里是模板的路径
        TemplateExportParams params = new TemplateExportParams(
                "excel/模板导出.xlsx");
        Map<String, Object> map = new HashMap<>();
        List<ExcelUser> list = new ArrayList<>();
        //设置日期格式
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd-hh-mm-ss");
        map.put("date", sdf.format(new Date()));
        //赋值制表人
        map.put("make", "天天");
        //赋值核表人
        map.put("check", "地地");
        //赋值标题
        map.put("title", "个人基本信息表格制作");
        ExcelUser excelUser1 = new ExcelUser();
        excelUser1.setName("小小");
        excelUser1.setAge("18");
        excelUser1.setSex("男");
        excelUser1.setEducation("本科");
        excelUser1.setPlace("中国");
        excelUser1.setSchool("中国大学");
        excelUser1.setRemark("这是一个备注");
        excelUser1.setDate(sdf.format(new Date()));
        ExcelUser excelUser2 = new ExcelUser();
        excelUser2.setName("大大");
        excelUser2.setAge("20");
        excelUser2.setSex("女");
        excelUser2.setEducation("本科");
        excelUser2.setPlace("北京");
        excelUser2.setSchool("北京大学");
        excelUser2.setRemark("这是第二个备注备注");
        excelUser2.setDate(sdf.format(new Date()));
        list.add(excelUser1);
        list.add(excelUser2);
        map.put("list", list);
        Workbook workbook = ExcelExportUtil.exportExcel(params, map);
        //存储位置
        File savefile = new File("D:/excel/");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("D:/excel/模板导出.xlsx");
        workbook.write(fos);
        fos.close();
        return "成功";

    }


}