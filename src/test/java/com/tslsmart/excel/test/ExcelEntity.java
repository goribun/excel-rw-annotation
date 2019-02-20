package com.tslsmart.excel.test;

import cn.wangxs.excel.annotation.ExcelField;
import org.apache.poi.hssf.util.HSSFColor;

import java.util.Date;

/**
 * @author JakeWoo
 * @version 1.0
 * @date 2019-02-14 17:50:33
 */
public class ExcelEntity {
    @ExcelField(name = "姓名", order = 1)
    private String userName;
    @ExcelField(name = "密码", order = 2)
    private String passWord;
    @ExcelField(name = "网址", order = 3)
    private String url;
    @ExcelField(name = "注册时间", order = 5, format = "yyyy-MM-dd HH:mm:ss")
    private Date time;
    @ExcelField(name = "账户余额", order = 4, format = "#.###", color = HSSFColor.YELLOW.class, expression = "money<1000")
    private Double money;

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public String getPassWord() {
        return passWord;
    }

    public void setPassWord(String passWord) {
        this.passWord = passWord;
    }

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    public Date getTime() {
        return time;
    }

    public void setTime(Date time) {
        this.time = time;
    }

    public Double getMoney() {
        return money;
    }

    public void setMoney(Double money) {
        this.money = money;
    }
}
