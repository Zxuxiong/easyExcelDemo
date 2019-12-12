package com.xuxiong.easyexceldemo.util.easyexcel;

import com.alibaba.excel.annotation.ExcelProperty;

/**
 * @author 许兄
 * Created on 2019/12/11
 */
public class Demo {

    @ExcelProperty(value = "姓名", index = 0)
    private String name;

    @ExcelProperty(value = {"邮箱"})
    private String email;

    @ExcelProperty(value = {"通讯地址信息", "手机号"})
    private String phone;

    @ExcelProperty(value = {"通讯地址信息", "地址"})
    private String address;


    @ExcelProperty(value = {"年龄"})
    private Integer age;
    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }


    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }
}

