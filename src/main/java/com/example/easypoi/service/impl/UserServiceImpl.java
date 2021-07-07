package com.example.easypoi.service.impl;

import com.example.easypoi.service.UserService;
import org.springframework.stereotype.Service;

/**
 * @author wad
 * @date 2021年07月07日 11:49
 */
@Service
public class UserServiceImpl implements UserService {
    @Override
    public boolean checkout(String name) {
        return true;
    }
}
