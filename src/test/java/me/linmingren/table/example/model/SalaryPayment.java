package me.linmingren.table.example.model;

import lombok.Data;

@Data
public class SalaryPayment {
    private String userName;
    private Double baseSalary;
    private Double fullAttendanceBonus;
    private Double mealSupplement;
    private Double transportationAllowance;
    private Double sickLeave;
    private Double personalLeave;
    private Double actualPay;

    public SalaryPayment(String userName, Double baseSalary, Double fullAttendanceBonus, Double mealSupplement,
                         Double transportationAllowance, Double sickLeave, Double personalLeave,
                         Double actualPay) {
        this.userName = userName;
        this.baseSalary = baseSalary;
        this.fullAttendanceBonus = fullAttendanceBonus;
        this.mealSupplement = mealSupplement;
        this.transportationAllowance = transportationAllowance;
        this.sickLeave = sickLeave;
        this.personalLeave = personalLeave;
        this.actualPay = actualPay;
    }
}
