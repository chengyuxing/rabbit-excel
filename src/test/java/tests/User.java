package tests;

public class User {
    private final String name;
    private final String address;
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
