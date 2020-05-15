package tests;

import com.github.chengyuxing.excel.core.Head;

public class User {
    @Head("姓名")
    private final String name;
    @Head("地址")
    private final String address;
    @Head("国家")
    private final String country;

    public User(String name, String address, String country) {
        this.name = name;
        this.address = address;
        this.country = country;
    }

    public String getName() {
        return name;
    }

    public String getAddress() {
        return address;
    }

    public String getCountry() {
        return country;
    }
}
