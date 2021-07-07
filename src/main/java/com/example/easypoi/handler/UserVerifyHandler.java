package com.example.easypoi.handler;

import cn.afterturn.easypoi.excel.entity.result.ExcelVerifyHandlerResult;
import cn.afterturn.easypoi.handler.inter.IExcelVerifyHandler;
import com.example.easypoi.service.UserService;
import com.example.easypoi.entity.ExcelUser;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;
import java.util.StringJoiner;

/**
 * @author wad
 * @date 2021年05月08日 16:21
 */
@Component
public class UserVerifyHandler implements IExcelVerifyHandler<ExcelUser> {

    private final ThreadLocal<List<ExcelUser>> threadLocal = new ThreadLocal<>();

    @Autowired
    UserService userService;

    @Override
    public ExcelVerifyHandlerResult verifyHandler(ExcelUser excelProduct) {
        StringJoiner joiner = new StringJoiner(",");
        //检验名字是否存在
        boolean flag = userService.checkout(excelProduct.getName());
        if (flag) {
            joiner.add("该名字已存在");
        }

        List<ExcelUser> threadLocalVal = threadLocal.get();
        if (threadLocalVal == null) {
            threadLocalVal = new ArrayList<>();
        }

        threadLocalVal.forEach(e -> {
            if (e.getName().equals(excelProduct.getName())) {
                int lineNumber = e.getRowNum() + 1;
                joiner.add("名字与" + lineNumber + "行重复");
            }
        });
        // 添加本行数据对象到ThreadLocal中
        threadLocalVal.add(excelProduct);
        threadLocal.set(threadLocalVal);
        if (joiner.length() != 0) {
            return new ExcelVerifyHandlerResult(false, joiner.toString());
        }
        if (joiner.length() != 0) {
            return new ExcelVerifyHandlerResult(false, joiner.toString());
        }
        return new ExcelVerifyHandlerResult(true);

    }

    public ThreadLocal<List<ExcelUser>> getThreadLocal() {
        return threadLocal;
    }


}
