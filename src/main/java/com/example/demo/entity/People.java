package com.example.demo.entity;

public class People {

    public String name;
    public String sex;
    public String age;
    public String phone;

    public People() {
    }

    public People(String name, String sex, String age, String phone) {
        this.name = name;
        this.sex = sex;
        this.age = age;
        this.phone = phone;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }
}
