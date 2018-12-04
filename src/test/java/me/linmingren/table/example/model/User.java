package me.linmingren.table.example.model;

import lombok.Data;

import java.util.Date;

@Data
public class User {
    private String name;
    private String address;
    private int score;
    private Date createdAt;

    public User(String user, String address, int score, Date createdAt) {
        this.name = user;
        this.address = address;
        this.score = score;
        this.createdAt = createdAt;
    }
}
